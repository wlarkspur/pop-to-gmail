"""가마감용 Daily Report 월초 작업 (HAS, YSKD 따로).

1. Gmail(회사메일 사본)에서 Irina의 지난달 Cash Daily Report 첨부 중 가장 늦게 받은 것을 고름
   (월초에 지난달분 정정본이 오면 그게 최신이라 자동으로 그걸 씀)
2. 구글드라이브 '@가마감용 DailyReport_YYYY_법인.xlsx'에 그달 시트를 맨 앞에 추가 (이미 있으면 건너뜀)
3. 합계, 잔액 흐름 검증 후 결과 메일 (실행 기록에는 상태만, 금액은 메일로만 - 공개 저장소)

사용: python3 daily_job.py [YYYY-MM] [HAS|YSKD]   (비우면 지난달, 두 법인 모두)
"""
from __future__ import annotations

import datetime as dt
import email
import email.header
import imaplib
import json
import os
import sys
import tempfile
from email.utils import parsedate_to_datetime

import daily_report as dr
import fx_job
from pop2gmail import _imap_quote, gmail_folders

KLD = dt.timezone(dt.timedelta(hours=2))
AUTO_FOLDER = "1Dssop-HfJKQb-qJVYL4_KNmEQ5JHPBt-"  # 업무 자동화 (daily_overrides.json = 사용자가 정한 분류, 이름으로 찾음)
TARGET = {  # 법인: (드라이브 폴더, 연도별 파일 이름, 원자료 첨부 이름, 처음 보는 출금의 임시 분류)
    "HAS": ("1kPGUNCzZSoH6Q2M-lDcoKfgj_R9xqwF_", "@가마감용 DailyReport_{y}_HAS.xlsx",
            lambda y, m: f"Daily Report {y}.{m:02d} (HAS)", "잡비"),
    "YSKD": ("1_zBWBXDD9cZJBoCxKfjWiVB2YfrS2lDy", "@가마감용 DailyReport_{y}_YSKD.xlsx",
             lambda y, m: f"Daily Report {m:02d}.{y} YSKD", "기타"),
}


def latest_attachment(y: int, m: int, base: str, out: str):
    """그달 리포트 첨부 중 가장 늦게 받은 것을 out에 저장. (받은 시각, 정정본 여부) 반환."""
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
    conn.select(_imap_quote(gmail_folders(conn)[0]), readonly=True)
    since = dt.date(y, m, 20)
    q = f'from:financial@hyauto.ru has:attachment after:{since:%Y/%m/%d}'
    typ, data = conn.uid("SEARCH", "X-GM-RAW", _imap_quote(q))
    best = None
    for uid in (data[0].split() if data and data[0] else []):
        _, msg = conn.uid("FETCH", uid, "(RFC822)")
        mm = email.message_from_bytes(msg[0][1])
        when = parsedate_to_datetime(mm["Date"])
        for part in mm.walk():
            fn = part.get_filename()
            if not fn:
                continue
            fn = str(email.header.make_header(email.header.decode_header(fn)))
            if fn.lower().startswith(base.lower()) and fn.lower().endswith(".xlsx") and (best is None or when > best[0]):
                best = (when, part.get_payload(decode=True))
    conn.logout()
    if not best:
        return None
    open(out, "wb").write(best[1])
    month_end = dt.datetime(y + (m == 12), m % 12 + 1, 1, tzinfo=KLD)
    return best[0], best[0] >= month_end


def find_file(s, folder: str, name: str):
    q = f"'{folder}' in parents and trashed = false and name = '{name}'"
    r = s.get(fx_job.API, params={"q": q, "fields": "files(id,name)", "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    f = r.json()["files"]
    return f[0]["id"] if f else None


def _download(s, fid, path):
    r = s.get(f"{fx_job.API}/{fid}", params={"alt": "media", "supportsAllDrives": "true"})
    r.raise_for_status()
    open(path, "wb").write(r.content)


def _upload(s, fid, path):
    r = s.patch(f"https://www.googleapis.com/upload/drive/v3/files/{fid}",
                params={"uploadType": "media", "supportsAllDrives": "true"}, data=open(path, "rb").read(),
                headers={"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    r.raise_for_status()


def run(ent: str, y: int, m: int, s, overrides: dict) -> list[str]:
    folder, fname, base, fallback = TARGET[ent]
    name = fname.format(y=y)
    fid = find_file(s, folder, name)
    if not fid:
        return [f"[{ent}] 드라이브에 '{name}' 파일이 없음 - 새해 파일은 처음 한 번 만들어 둬야 함 (Claude에게 요청)"]
    with tempfile.TemporaryDirectory() as tmp:
        got = latest_attachment(y, m, base(y, m), f"{tmp}/src.xlsx")
        if not got:
            late = dt.datetime.now(KLD).day >= 5  # 5일까지 안 오면 알림 (YSKD는 늦게 오는 편)
            return [f"[{ent}] {y}-{m:02d} Irina 리포트('{base(y, m)}') 메일을 아직 못 찾음" + (" - 5일까지 안 옴, 확인 필요" if late else "")]
        when, corrected = got
        _download(s, fid, f"{tmp}/in.xlsx")
        sheet = dr.sheet_name(y, m, ent)
        import openpyxl
        names = [n.strip() for n in openpyxl.load_workbook(f"{tmp}/in.xlsx", read_only=True).sheetnames]
        py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
        if sheet not in names and dr.sheet_name(py, pm, ent) not in names and m != 1:
            return [f"[{ent}] 직전 달 '{dr.sheet_name(py, pm, ent)}' 시트가 없어 대기 - 직전 달 원자료를 먼저 넣어야 함 (확인 필요)"]
        if sheet in names:
            return [f"[{ent}] '{sheet}' 시트가 이미 있음 - 건너뜀" + (" (단, 월초에 정정본이 왔으니 확인 필요)" if corrected else "")]
        src = dr.parse_source(f"{tmp}/src.xlsx")
        if not all(src["checks"].values()):
            return [f"[{ent}] 원자료 검증 실패 {src['checks']} - 파일 안 바꿈, 확인 필요"]
        hist = dr.history(f"{tmp}/in.xlsx", ent)
        cats, new = [], []
        for i in (x for x in src["items"] if x["kind"] == "out"):
            c, how = dr.classify(i["desc"], hist, overrides.get(ent))
            if c is None:
                c = fallback
                new.append(f"  - {i['day']}일 {i['amount']:,.2f} '{i['desc']}' -> 임시로 '{fallback}'")
            cats.append(c)
        dr.build(f"{tmp}/in.xlsx", f"{tmp}/out.xlsx", y, m, src, cats, ent)
        v = dr.verify(f"{tmp}/out.xlsx", y, m, src, ent)
        if not (v["입금 합계 일치"] and v["출금 합계 일치"]) or v["분류 빈 칸 행"] or v["목록에 없는 분류 행"]:
            return [f"[{ent}] 결과 검증 실패 {v} - 파일 안 바꿈, 확인 필요"]
        _upload(s, fid, f"{tmp}/out.xlsx")
    lines = [f"[{ent}] '{name}'에 '{sheet}' 시트 추가 (원자료 {when.astimezone(KLD):%m/%d %H:%M} 받은 메일"
             + (", 월초 정정본" if corrected else "") + ")",
             f"  입금 {src['in']:,.2f} / 출금 {src['out']:,.2f} / 월초 잔액 {src['open']:,.2f} -> 월말 {src['close']:,.2f} (검증 통과)"]
    if new:
        lines += ["  처음 보는 출금 (분류 확인 필요, 정하면 daily_overrides.json에 추가):"] + new
    return lines


def main() -> int:
    now = dt.datetime.now(KLD).date()
    if len(sys.argv) > 1 and sys.argv[1].strip():
        y, m = map(int, sys.argv[1].strip().split("-"))
    else:
        prev = now.replace(day=1) - dt.timedelta(days=1)
        y, m = prev.year, prev.month
    ents = [sys.argv[2].strip()] if len(sys.argv) > 2 and sys.argv[2].strip() else list(TARGET)
    s = fx_job.drive()
    oid = find_file(s, AUTO_FOLDER, "daily_overrides.json")
    overrides = {}
    if oid:
        r = s.get(f"{fx_job.API}/{oid}", params={"alt": "media", "supportsAllDrives": "true"})
        overrides = json.loads(r.content.decode("utf-8")) if r.ok else {}
    lines = []
    for ent in ents:
        try:
            lines += run(ent, y, m, s, overrides)
        except Exception as e:  # noqa: BLE001
            lines.append(f"[{ent}] 오류: {type(e).__name__}: {e}")
    done = [l for l in lines if "시트 추가" in l]
    print(f"{y}-{m:02d}: 추가 {len(done)}건, 메시지 {len(lines)}줄")  # 공개 실행 기록에는 상태만
    if done or any("확인 필요" in l or "오류" in l for l in lines):
        fx_job.mail(f"[가마감 Daily Report] {y}년 {m}월", "\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
