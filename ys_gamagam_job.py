"""YSKD 가마감 생성 작업 (GitHub). 결과는 드라이브 'YS 손익보고' 폴더의 그달 파일 3개에 덮어씀.

입력: 전월 '누적 마감_YSKD', '채권_채무_관리_YSKD' (양식), Irina 'отчет КЛД за MM/YYYY' 첨부 2개 (monthly_MM.YYYY, YSH KLD (RUB)),
      marina@ 'дебиторка КЛД MM/YYYY' 첨부(1C 채권채무 잔액), CBR EUR 환율(월평균, 월말).
출력: 'YYYY년 M월_누적 마감_YSKD.xlsx', 'YYYY년 M월_채권_채무_관리_YSKD.xlsx', 'KLD 재무제표 (RUB) YY년M월.xlsx'
      (서비스 계정은 파일을 새로 만들 수 없어 같은 이름의 빈 자리 파일이 미리 있어야 함. 없으면 그 파일만 건너뛰고 메일로 알림)
사용: python3 ys_gamagam_job.py YYYY-MM
"""
from __future__ import annotations

import calendar
import datetime as dt
import email
import email.header
import imaplib
import os
import re
import statistics
import sys
import tempfile
from email.utils import parsedate_to_datetime

import daily_job
import fx_cbr
import fx_job
import ys_gamagam as g
from pop2gmail import _imap_quote, gmail_folders

YS_FOLDER = "1i2jZhYpQTYL4RFG3WZuK_gPUVsjsb2br"  # YS 손익보고
RU_MONTH = {1: "янв", 2: "фев", 3: "мар", 4: "апр", 5: "ма", 6: "июн", 7: "июл", 8: "авг", 9: "сен", 10: "окт", 11: "ноя", 12: "дек"}


def _mail_attachments(y: int, m: int, gmail_q: str, subj_ok, pick, tmp: str) -> dict:
    """gmail_q로 찾은 메일 중 subj_ok(제목)인 것들의 첨부를 pick(파일명)->키 로 분류, 키별로 가장 늦게 받은 것 저장."""
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
    conn.select(_imap_quote(gmail_folders(conn)[0]), readonly=True)
    _, data = conn.uid("SEARCH", "X-GM-RAW", _imap_quote(gmail_q))
    found = {}
    for uid in (data[0].split() if data and data[0] else []):
        _, msg = conn.uid("FETCH", uid, "(RFC822)")
        mm = email.message_from_bytes(msg[0][1])
        subj = str(email.header.make_header(email.header.decode_header(mm["Subject"] or "")))
        if not subj_ok(subj.lower()):
            continue
        when = parsedate_to_datetime(mm["Date"])
        for part in mm.walk():
            fn = part.get_filename()
            if not fn:
                continue
            fn = str(email.header.make_header(email.header.decode_header(fn)))
            key = pick(fn.lower())
            if key and (key not in found or when > found[key][0]):
                found[key] = (when, part.get_payload(decode=True), fn)
    conn.logout()
    out = {}
    for key, (when, blob, fn) in found.items():
        out[key] = f"{tmp}/{key}.xlsx"
        open(out[key], "wb").write(blob)
        out[key + "_name"], out[key + "_when"] = fn, when
    return out


def irina_report(y: int, m: int, tmp: str) -> dict:
    tag = f"{m:02d}/{y}"
    src = _mail_attachments(y, m, f"from:financial@hyauto.ru has:attachment after:{y}/{m:02d}/20",
                            lambda s: "клд" in s and tag in s and "отчет" in s,
                            lambda f: "monthly" if "monthly" in f else "kld" if "kld" in f else None, tmp)
    missing = {"monthly", "kld"} - set(src)
    if missing:
        raise RuntimeError(f"Irina 'отчет КЛД за {tag}' 첨부가 모자람: {missing}")
    return src


def arap_raw(y: int, m: int, tmp: str) -> dict:
    tag = f"{m:02d}/{y}"
    return _mail_attachments(y, m, f"has:attachment filename:xlsx after:{y}/{m:02d}/20",
                             lambda s: "дебиторка" in s and "клд" in s and tag in s,
                             lambda f: "raw" if f.endswith(".xlsx") else None, tmp)


def find_latest(s, pattern: str, before: tuple[int, int]):
    """폴더 안에서 'YYYY년 M월_...' 이름 중 before(y, m)보다 앞선 가장 최근 달 파일 (id, y, m). 대상 월 자리 파일(빈 사본)은 제외."""
    r = s.get(fx_job.API, params={"q": f"'{YS_FOLDER}' in parents and trashed = false", "fields": "files(id,name)",
                                  "supportsAllDrives": "true", "includeItemsFromAllDrives": "true", "pageSize": "200"})
    r.raise_for_status()
    best = None
    for f in r.json()["files"]:
        mt = re.match(r"(\d{4})년 (\d{1,2})월_" + re.escape(pattern) + r"\.xlsx$", f["name"])
        if mt and (int(mt[1]), int(mt[2])) < before and (best is None or (int(mt[1]), int(mt[2])) > best[1:]):
            best = (f["id"], int(mt[1]), int(mt[2]))
    return best


def main() -> int:
    y, m = map(int, sys.argv[1].split("-"))
    s = fx_job.drive()
    names = {"cum": f"{y}년 {m}월_누적 마감_YSKD.xlsx", "arap": f"{y}년 {m}월_채권_채무_관리_YSKD.xlsx",
             "kld": f"KLD 재무제표 (RUB) {y % 100}년{m}월.xlsx"}
    out_ids = {k: daily_job.find_file(s, YS_FOLDER, n) for k, n in names.items()}
    notes, lines = [], []
    with tempfile.TemporaryDirectory() as tmp:
        src = irina_report(y, m, tmp)
        raw = arap_raw(y, m, tmp)
        rates = fx_cbr.daily(y, m)
        month_end = rates[max(rates)]
        avg = {mon: statistics.mean(fx_cbr.daily(y, mon).values()) for mon in range(1, m + 1)}
        # 1) 누적 마감
        t = find_latest(s, "누적 마감_YSKD", (y, m))
        if not t:
            raise RuntimeError("드라이브에 전월 누적 마감 파일이 없음")
        daily_job._download(s, t[0], f"{tmp}/cum_tmpl.xlsx")
        log, n = g.fill_cum(f"{tmp}/cum_tmpl.xlsx", src["monthly"], y, m, avg, f"{tmp}/cum.xlsx")
        notes += n
        lines.append(f"누적 마감: 양식 {t[1]}년 {t[2]}월 파일 + monthly {src['monthly_when']:%m/%d} -> 바꾼 칸 {len(log)}개, EUR 월평균 {avg[m]:.4f}")
        if out_ids["cum"]:
            daily_job._upload(s, out_ids["cum"], f"{tmp}/cum.xlsx")
        else:
            notes.append(f"'{names['cum']}' 자리 파일이 없어 올리지 못함 (드라이브에 같은 이름 파일을 만들어 두고 다시 실행)")
        # 2) 채권채무
        if "raw" in raw:
            t2 = find_latest(s, "채권_채무_관리_YSKD", (y, m))
            if not t2:
                raise RuntimeError("드라이브에 전월 채권채무 파일이 없음")
            daily_job._download(s, t2[0], f"{tmp}/ar_tmpl.xlsx")
            base = g.raw_date(raw["raw"])
            exp = dt.date(y, m, calendar.monthrange(y, m)[1])
            if base != exp:
                notes.append(f"1C 채권채무 파일 기준일 {base} (기대 {exp}) - 파일은 {exp} 기준으로 처리")
            log2, n2 = g.fill_arap(f"{tmp}/ar_tmpl.xlsx", raw["raw"], exp, month_end, f"{tmp}/ar.xlsx")
            notes += n2
            lines.append(f"채권채무: 양식 {t2[1]}년 {t2[2]}월 파일 + 1C 잔액 '{raw['raw_name']}' ({raw['raw_when']:%m/%d}) -> {', '.join(log2)}, EUR 월말 {month_end}")
            if out_ids["arap"]:
                daily_job._upload(s, out_ids["arap"], f"{tmp}/ar.xlsx")
            else:
                notes.append(f"'{names['arap']}' 자리 파일이 없어 올리지 못함")
        else:
            notes.append(f"1C 채권채무 잔액 메일('дебиторка КЛД {m:02d}/{y}')을 못 찾아 채권채무 파일은 안 만듦")
        # 3) KLD 재무제표 = Irina 첨부 그대로
        if out_ids["kld"]:
            daily_job._upload(s, out_ids["kld"], src["kld"])
            lines.append(f"KLD 재무제표: Irina 첨부 '{src['kld_name']}' 그대로 올림")
        else:
            notes.append(f"'{names['kld']}' 자리 파일이 없어 올리지 못함")
    print(f"{y}-{m:02d} YSKD 가마감: 확인 사항 {len(notes)}건")
    body = lines + ["", "확인 필요:"] + [f"- {x}" for x in notes] + ["", "다음 달을 위해: 드라이브 YS 손익보고 폴더에 다음 달 이름의 빈 자리 파일 3개(이번 달 파일 복사)를 만들어 둘 것"]
    fx_job.mail(f"[가마감] YSKD {y}년 {m}월 초안", "\n".join(body))
    return 0


if __name__ == "__main__":
    sys.exit(main())
