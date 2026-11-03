"""HAS 가마감 생성 작업 (GitHub). 결과는 드라이브 '가마감' 폴더의 'HAS 법인 YY년 M월 가마감(HDC 보고).xlsx'에 덮어씀.

입력: 전월 마감 파일(양식), Irina 'Отчет HAS за MM/YYYY' 첨부 3개 + 월말 Cash Daily Report(계좌 잔액),
      가마감용 Daily Report(8번), 하나은행 환율(fx_hana). 결과 칸 목록과 확인 필요 사항은 메일로 (실행 기록에는 상태만).
사용: python3 gamagam_job.py YYYY-MM
"""
from __future__ import annotations

import email
import email.header
import imaplib
import os
import sys
import tempfile
from email.utils import parsedate_to_datetime

import openpyxl

import daily_job
import daily_report as dr
import fx_hana
import fx_job
import has_gamagam as h
from pop2gmail import _imap_quote, gmail_folders

GM_FOLDER = "1zr_4mbUyOe0VxqTbpz3MHS2ZtoBoUYMq"      # HAS 손익보고/가마감
DR_FOLDER = "1kPGUNCzZSoH6Q2M-lDcoKfgj_R9xqwF_"      # 가마감용 Daily Report_HAS
FEE_CAT = "은행수수료 - 환전/송금, 기타 수수료"


def irina_report(y: int, m: int, tmp: str) -> dict:
    """'Отчет HAS за MM/YYYY' 최신 메일의 첨부 3개를 저장."""
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
    conn.select(_imap_quote(gmail_folders(conn)[0]), readonly=True)
    q = f"from:financial@hyauto.ru subject:HAS has:attachment after:{y}/{m:02d}/15"  # IMAP 검색은 영문만
    _, data = conn.uid("SEARCH", "X-GM-RAW", _imap_quote(q))
    found = {}  # 첨부 종류별로 가장 늦게 받은 것 (정정 메일에 일부 첨부만 다시 오는 경우)
    for uid in (data[0].split() if data and data[0] else []):
        _, msg = conn.uid("FETCH", uid, "(RFC822)")
        mm = email.message_from_bytes(msg[0][1])
        when = parsedate_to_datetime(mm["Date"])
        subj = str(email.header.make_header(email.header.decode_header(mm["Subject"] or "")))
        if f"{m:02d}/{y}" not in subj or "отчет" not in subj.lower():
            continue
        for part in mm.walk():
            fn = part.get_filename()
            if not fn:
                continue
            fn = str(email.header.make_header(email.header.decode_header(fn)))
            key = "monthly" if "monthly" in fn.lower() else "fs" if "재무제표" in fn else "ost" if "остатки" in fn.lower() else None
            if key and (key not in found or when > found[key][0]):
                found[key] = (when, part.get_payload(decode=True))
    conn.logout()
    missing = {"monthly", "fs", "ost"} - set(found)
    if missing:
        raise RuntimeError(f"첨부가 모자람: {missing}")
    out = {}
    for key, (when, data_) in found.items():
        out[key] = f"{tmp}/{key}.xlsx"
        open(out[key], "wb").write(data_)
    out["when"] = max(w for w, _ in found.values())
    out["dates"] = {k: w for k, (w, _) in found.items()}
    return out


def spb_statements(y: int, m: int, tmp: str) -> list[str]:
    """'Обороты по банку СПБ за MM/YYYY' 메일의 은행 명세서 첨부 (가장 최근 메일)."""
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
    conn.select(_imap_quote(gmail_folders(conn)[0]), readonly=True)
    _, data = conn.uid("SEARCH", "X-GM-RAW", _imap_quote(f"from:financial@hyauto.ru filename:statement after:{y}/{m:02d}/15"))
    best = None
    for uid in (data[0].split() if data and data[0] else []):
        _, msg = conn.uid("FETCH", uid, "(RFC822)")
        mm = email.message_from_bytes(msg[0][1])
        subj = str(email.header.make_header(email.header.decode_header(mm["Subject"] or "")))
        when = parsedate_to_datetime(mm["Date"])
        if "СПБ" in subj.upper() and f"{m:02d}/{y}" in subj and (best is None or when > best[0]):
            best = (when, mm)
    conn.logout()
    paths = []
    if best:
        for i, part in enumerate(p for p in best[1].walk() if p.get_filename()):
            paths.append(f"{tmp}/spb{i}.xlsx")
            open(paths[-1], "wb").write(part.get_payload(decode=True))
    return paths


def load_overrides(s, y: int, m: int) -> dict:
    import json
    oid = find_by_name(s, "gamagam_overrides.json", daily_job.AUTO_FOLDER)
    if not oid:
        return {}
    r = s.get(f"{fx_job.API}/{oid}", params={"alt": "media", "supportsAllDrives": "true"})
    r.raise_for_status()
    return json.loads(r.content.decode("utf-8")).get("HAS", {}).get(f"{y}-{m:02d}", {})


def apply_overrides(s, y: int, m: int, src: dict) -> list[str]:
    """드라이브 '업무 자동화/gamagam_overrides.json'에 적힌 원자료 수정분을 받은 첨부에 덮어씀 (입력 파일만, 양식은 그대로)."""
    ov = load_overrides(s, y, m)
    notes = []
    for key, cells in ov.items():
        if key not in src:
            continue
        wb = openpyxl.load_workbook(src[key])
        for ref, val in cells.items():
            if ref.startswith("_"):
                continue
            sheet, cell = ref.split("!")
            wb[sheet][cell] = val
        wb.save(src[key])
        notes.append(f"원자료 수정분 적용: {key} {', '.join(k for k in cells if not k.startswith('_'))} ({cells.get('_출처', '')})")
    return notes


def find_by_name(s, name: str, folder: str | None = None):
    q = f"name = '{name}' and trashed = false" + (f" and '{folder}' in parents" if folder else "")
    r = s.get(fx_job.API, params={"q": q, "fields": "files(id,name)", "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    f = r.json()["files"]
    return f[0]["id"] if f else None


def main() -> int:
    y, m = map(int, sys.argv[1].split("-"))
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    s = fx_job.drive()
    out_name = f"HAS 법인 {y % 100}년 {m}월 가마감(HDC 보고).xlsx"
    out_id = find_by_name(s, out_name, GM_FOLDER)
    tmpl_id = find_by_name(s, f"HAS 법인 {py % 100}년 {pm}월 마감(HDC 보고).xlsx") or \
        find_by_name(s, f"HAS 법인 {py % 100}년 {pm}월 가마감(HDC 보고).xlsx")
    if not out_id or not tmpl_id:
        raise RuntimeError(f"드라이브 파일 없음 (결과 자리 {bool(out_id)}, 전월 양식 {bool(tmpl_id)}) - 결과 자리는 처음 한 번 만들어 둬야 함")
    with tempfile.TemporaryDirectory() as tmp:
        src = irina_report(y, m, tmp)
        fixes = apply_overrides(s, y, m, src)
        daily_job._download(s, tmpl_id, f"{tmp}/tmpl.xlsx")
        dr_id = find_by_name(s, f"@가마감용 DailyReport_{y}_HAS.xlsx", DR_FOLDER)
        daily_job._download(s, dr_id, f"{tmp}/dr.xlsx")
        drs = h.read_daily_sheet(f"{tmp}/dr.xlsx", dr.sheet_name(y, m, "HAS"))
        got = daily_job.latest_attachment(y, m, f"Daily Report {y}.{m:02d} (HAS)", f"{tmp}/cash.xlsx")
        acc = dr.parse_source(f"{tmp}/cash.xlsx")["accounts"] if got else None
        j8 = fx_hana.month_end_rate(y, m)["미화환산율"]
        avg = fx_hana.month_avg_rate(y, m)["대미환산율"]
        old = openpyxl.load_workbook(f"{tmp}/tmpl.xlsx")[h.S61]["J8"].value

        b = h.Book(f"{tmp}/tmpl.xlsx")
        h.fill_153(b, src["monthly"], m)
        h.fill_fs(b, src["fs"])
        h.fill_rates(b, y, m, j8, avg)
        h.fill_cashflow(b, m, old, j8, drs, FEE_CAT)
        notes = list(fixes)
        if acc:
            g = lambda k: round(next((v for n, v in acc.items() if k in n.upper()), 0), 2)
            spb = g("SPB")
            st = spb_statements(y, m, tmp)
            if st:
                vals = h.spb_balances(st)
                same = abs(sum(vals) - spb) < 0.01
                notes.append(f"11번 SPB: 은행 명세서 기말잔액 {' + '.join(f'{v:,.2f}' for v in vals)} = {sum(vals):,.2f}"
                             + (" (Daily Report와 일치)" if same else f" - Daily Report {spb:,.2f}와 다름, 명세서 값 사용 (확인 필요)"))
                spb = vals
            else:
                notes.append("11번 SPB: 은행 명세서 메일을 못 찾아 Daily Report 잔액 사용 (확인 필요)")
            h.fill_cash(b, {"cash": g("CASH"), "woori": g("WOORI"), "entb": g("ENTB"), "spb": spb})
            notes.append("11번 Woori, ENTB, 현금: Irina 월말 Cash Daily Report 잔액 (은행 캡처와 대사는 안 함)")
        else:
            notes.append("11번 현금시재표: 월말 Cash Daily Report를 못 찾아 비워 둠 (확인 필요)")
        notes += h.fill_ostatki(b, h.read_ostatki(src["ost"], y), y, m)
        out_ov = load_overrides(s, y, m).get("out", {})
        if out_ov:
            done = h.apply_out_overrides(b, out_ov)
            notes.append(f"사용자 지정 칸 {len(done)}개 반영: {', '.join(done)}")
        b.save(f"{tmp}/out.xlsx")
        daily_job._upload(s, out_id, f"{tmp}/out.xlsx")
    print(f"{y}-{m:02d} HAS 가마감: {len(b.log)}칸, 확인 사항 {len(notes)}건")
    lines = [f"'{out_name}' 생성 (전월 마감 양식 + Irina {src['when']:%m/%d} 자료)",
             f"환율: 월말 {j8} (전월 {old}), 월평균 {avg}", f"바꾼 칸 {len(b.log)}개", "", "확인 필요:"] + [f"- {n}" for n in notes]
    fx_job.mail(f"[가마감] HAS {y}년 {m}월 초안", "\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
