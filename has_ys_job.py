"""HAS 손익보고 YS(YSH) 보고용 초안 작업 (GitHub). 결과는 드라이브 'HAS 손익보고 / YY년 (YS) / M월' 폴더의 파일 3개에 덮어씀.

입력: 앞 달 폴더의 3개 파일(양식), Irina 'Отчет HAS за MM/YYYY' 첨부 3개(monthly, 재무제표(RUB), Конечные остатки -
      정정 메일이 오면 종류별로 가장 늦은 것), CBR EUR 월평균.
출력: 'HAS법인 YY년 M월 마감(YSH 보고).xlsx', 'HAS YYYY년 재무제표 (KR).xlsx', 'HAS YYYY M월 ARAP 차입대여금 현황.xlsx'
      (서비스 계정은 파일을 새로 못 만들어서 같은 이름 자리 파일이 그달 폴더에 있어야 함. 없으면 그 파일만 건너뛰고 알림)
사용: python3 has_ys_job.py YYYY-MM
"""
from __future__ import annotations

import calendar
import datetime as dt
import re
import statistics
import sys
import tempfile

import daily_job
import fx_cbr
import fx_job
import gamagam_job
import has_ys as h
import tg

HAS_ROOT = "1FyYT6b5j0YltEZDGDpt_IddI382OIPIn"  # HAS 손익보고


def _children(s, folder: str) -> list[dict]:
    r = s.get(fx_job.API, params={"q": f"'{folder}' in parents and trashed = false", "fields": "files(id,name,mimeType)",
                                  "supportsAllDrives": "true", "includeItemsFromAllDrives": "true", "pageSize": "200"})
    r.raise_for_status()
    return r.json()["files"]


def month_folder(s, y: int, m: int):
    yf = next((f["id"] for f in _children(s, HAS_ROOT) if f["name"].replace(" ", "") == f"{y % 100}년(YS)"), None)
    return next((f["id"] for f in _children(s, yf) if f["name"] == f"{m}월"), None) if yf else None


def names(y: int, m: int) -> dict:
    return {"ysh": f"HAS법인 {y % 100}년 {m}월 마감(YSH 보고).xlsx", "kr": f"HAS {y}년 재무제표 (KR).xlsx",
            "arap": f"HAS {y} {m}월 ARAP 차입대여금 현황.xlsx"}


PATTERN = {"ysh": r"마감\(YSH 보고\)\.xlsx$", "kr": r"재무제표 \(KR\)\.xlsx$", "arap": r"ARAP 차입대여금 현황\.xlsx$"}


def find_template(s, key: str, y: int, m: int):
    """대상 월보다 앞선 가장 최근 달 폴더에서 양식 파일 (id, y, m)."""
    for _ in range(12):
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
        fid = month_folder(s, y, m)
        if fid:
            for f in _children(s, fid):
                if re.search(PATTERN[key], f["name"]):
                    return f["id"], y, m
    return None


def main() -> int:
    y, m = map(int, sys.argv[1].split("-"))
    base = dt.date(y, m, calendar.monthrange(y, m)[1])
    s = fx_job.drive()
    nm = names(y, m)
    folder = month_folder(s, y, m)
    out_ids = {k: (daily_job.find_file(s, folder, n) if folder else None) for k, n in nm.items()}
    notes, lines = [], []
    if not folder:
        notes.append(f"드라이브 'HAS 손익보고/{y % 100}년 (YS)/{m}월' 폴더가 없어 결과를 올리지 못함")
    if m == 1:
        notes.append("1월: 연도가 바뀌어 재무제표(KR) 이름과 마감 파일의 연도 시트를 확인해야 함 (양식은 작년 12월 파일)")
    with tempfile.TemporaryDirectory() as tmp:
        src = gamagam_job.irina_report(y, m, tmp)
        rate = {mon: statistics.mean(fx_cbr.daily(y, mon).values()) for mon in range(1, m + 1)}
        jobs = [("ysh", lambda t, o: h.fill_ysh(t, src["monthly"], y, m, rate, o)),
                ("kr", lambda t, o: h.fill_kr(t, src["fs"], base, rate[m], o)),
                ("arap", lambda t, o: h.fill_arap(t, src["ost"], base, o))]
        for key, fn in jobs:
            t = find_template(s, key, y, m)
            if not t:
                notes.append(f"{nm[key]}: 앞 달 양식 파일을 못 찾음")
                continue
            daily_job._download(s, t[0], f"{tmp}/{key}_t.xlsx")
            log, n = fn(f"{tmp}/{key}_t.xlsx", f"{tmp}/{key}.xlsx")
            notes += [f"[{key}] {x}" for x in n]
            lines.append(f"{nm[key]}: 양식 {t[1]}년 {t[2]}월 + Irina {src['when']:%m/%d} 자료 -> 바꾼 칸 {len(log)}개")
            if out_ids[key]:
                daily_job._upload(s, out_ids[key], f"{tmp}/{key}.xlsx")
                tg.send_file(f"{tmp}/{key}.xlsx", f"HAS(YS) {y}년 {m}월 초안", name=nm[key])
            elif folder:
                notes.append(f"'{nm[key]}' 자리 파일이 {m}월 폴더에 없어 올리지 못함")
    print(f"{y}-{m:02d} HAS YS 보고 초안: 확인 사항 {len(notes)}건")
    body = lines + [f"EUR 월평균(CBR) {rate[m]:.6f}", "", "확인 필요:"] + [f"- {x}" for x in notes] + \
        ["", "다음 달을 위해: 'HAS 손익보고/YY년 (YS)/다음 달' 폴더와 다음 달 이름의 자리 파일 3개(이번 달 파일 복사)를 둘 것"]
    fx_job.mail(f"[가마감] HAS(YS 보고) {y}년 {m}월 초안", "\n".join(body))
    return 0


if __name__ == "__main__":
    sys.exit(main())
