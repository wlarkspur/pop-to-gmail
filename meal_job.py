"""매월 1일 08:00: 급여 폴더 식대 정산 xlsx에 그달 시트를 만들고 메일로 알린다 (규칙은 meal_sheet.py)."""
from __future__ import annotations

import datetime as dt
import os
import re
import sys
import tempfile

import fx_job
import meal_sheet
from salary_job import KLD, PAY_FOLDER_ID, _download, _upload

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def find_meal_file(s):
    q = f"'{PAY_FOLDER_ID}' in parents and trashed = false and name contains '식대 정산' and mimeType = '{XLSX}'"
    r = s.get(fx_job.API, params={"q": q, "fields": "files(id,name)", "supportsAllDrives": "true",
                                  "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    files = r.json()["files"]
    if not files:
        raise RuntimeError("급여 폴더에 '식대 정산' 파일이 없음")
    year = lambda f: int(m.group(1)) if (m := re.search(r"(\d\d)년", f["name"])) else 0  # noqa: E731
    f = max(files, key=year)
    return f["id"], f["name"]


def main() -> int:
    if len(sys.argv) > 1:
        y, m = map(int, sys.argv[1].split("-"))
    else:
        t = dt.datetime.now(KLD).date()
        y, m = t.year, t.month
    s = fx_job.drive()
    fid, fname = find_meal_file(s)
    with tempfile.TemporaryDirectory() as tmp:
        _download(s, fid, f"{tmp}/in.xlsx")
        res = meal_sheet.build(f"{tmp}/in.xlsx", f"{tmp}/built.xlsx", y, m)
        if res["상태"] != "이미 있음":
            from xlsx_patch import focus
            focus(f"{tmp}/built.xlsx", f"{tmp}/out.xlsx", res["시트"], "E3")  # 열면 새 시트 첫 칸
        print("식대 시트:", res["상태"])  # 공개 저장소 실행 기록이라 금액은 안 찍음 (상세는 결과 메일)
        if res["상태"] == "이미 있음":
            return 0
        _upload(s, fid, f"{tmp}/out.xlsx", XLSX)
        if res["상태"] != "생성":
            return 0
    if os.environ.get("DRY_RUN") == "1":
        return 0
    b1, b2 = res["1차"], res["2차"]
    lines = [f"{fname}에 '{res['시트']}' 시트를 만들었습니다 ('{res['복사한 시트']}' 복사).", "",
             f"1차 {b1[4]}: {b1[0]} ~ {b1[1]}, 근무일 {b1[2]}일 x 750 = {b1[3]:,} RUB/인",
             f"2차 {b2[4]}: {b2[0]} ~ {b2[1]}, 근무일 {b2[2]}일 x 750 = {b2[3]:,} RUB/인",
             f"대상: {', '.join(res['인원'])} (전원 750으로 채움)", "",
             "휴가, 출장 등으로 빠지는 날은 직접 지워 주세요. 공휴일, 대체휴일은 러시아 생산 달력 기준으로 이미 뺐습니다.",
             "인원이 바뀌었으면 시트에서 직접 고쳐 주세요 (지난달 시트 인원을 그대로 가져옴)."]
    fx_job.mail(f"[식대] {res['시트']} 시트 생성 - 1차 {b1[2]}일 / 2차 {b2[2]}일", "\n".join(lines))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        fx_job.mail("[식대] 새 달 시트 생성 실패", f"오류: {e}")
        raise
