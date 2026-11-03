"""급여환율: 기준일에 x-rates 월평균으로 EUR, USD, CNY 급여환율을 정하고 파일 두 곳을 채운다.

- 급여일 10일, 25일 (쉬는 날이면 직전 근무일). 기준일 = 급여일 포함 근무일 5일째 (러시아 생산 달력)
- 3개월 = 기준일이 속한 달의 전전달, 전달(전체) + 이번 달(기준일까지)
- 기준일 08:00부터 30분마다 확인. 이번 달 집계 일수가 기준일 날짜와 같아야 (= 기준일까지 업데이트됨) 진행
  아니면 다음 회차에 재시도, 12:00 회차까지 안 되면 알림 메일
- 월평균 값: EUR, USD는 x-rates 화면 소수 6자리 그대로, CNY는 소수 4자리에서 자름 (기존 파일 방식)
- 급여환율은 각 파일의 기존 수식이 계산: EUR, USD = ROUNDUP(평균, 0), CNY = ROUNDUP(평균, 2)
- 기록 위치
  EUR: 환율 폴더 환율!.xls '{YY}년' 시트 급여환율 칸 (W96:Y97)
  USD, CNY: 급여 폴더 '{YY}년 한국인 급여 정리...xlsx '환율' 시트 (USD/RUB, CNY/RUB 블록)
- 이미 같은 값이 있으면 다시 쓰지 않음

사용: python3 salary_job.py            -> 오늘(칼리닌그라드)이 기준일이면 실행
      python3 salary_job.py 2026-11-03 -> 그 날짜를 기준일로 보고 실행 (수동, 재실행)
"""
from __future__ import annotations

import datetime as dt
import os
import sys
import tempfile
from decimal import ROUND_CEILING, ROUND_DOWN, Decimal

import re

import xlrd

import ru_calendar
import xrates

KLD = dt.timezone(dt.timedelta(hours=2))
PAY_FOLDER_ID = "1DRZnDLuXnkkYipHPC7A4KT_VunR5feDV"


def find_payday(today: dt.date):
    nxt = today + dt.timedelta(days=31)
    for y, m in ((today.year, today.month), (nxt.year, nxt.month)):
        for nominal, pay, rate in ru_calendar.schedule(y, m):
            if rate == today:
                return nominal, pay, rate
    return None


def three_months(rate: dt.date):
    out = []
    for back in (2, 1, 0):
        y, m = rate.year, rate.month - back
        if m <= 0:
            y, m = y - 1, m + 12
        out.append((y, m))
    return out


def salary_rates(rate: dt.date) -> dict:
    months = three_months(rate)
    res = {"월": [f"{y}-{m:02d}" for y, m in months]}
    for cur in ("EUR", "USD", "CNY"):
        pages = {y: xrates.monthly(cur, y) for y in {y for y, _ in months}}
        vals = []
        for y, m in months:
            if m not in pages[y]:
                raise RuntimeError(f"x-rates {cur} {y}-{m:02d} 값 없음")
            v, days = pages[y][m]
            if (y, m) == (rate.year, rate.month) and days != rate.day:
                raise RuntimeError(f"x-rates {cur} {m}월이 {days}일까지만 반영됨 (기준일 {rate.day}일) - 업데이트 전")
            if cur == "CNY":
                v = v.quantize(Decimal("0.0001"), rounding=ROUND_DOWN)
            vals.append(v)
        avg = sum(vals) / 3
        places = Decimal("0.01") if cur == "CNY" else Decimal("1")
        final = avg.quantize(places, rounding=ROUND_CEILING)
        res[cur] = {"월평균": [str(v) for v in vals], "3개월 평균": str(avg.quantize(Decimal("0.000001"))),
                    "급여환율": str(final)}
    return res


# ---------------- 파일 쓰기 ----------------

def write_eur_xls(src, dst, rate: dt.date, vals) -> dict:
    """환율!.xls '{YY}년' 시트의 급여환율 칸(EUR/RUB)."""
    from xls_patch import patch_xls
    sheet = f"{rate.year % 100:02d}년"
    s = xlrd.open_workbook(src).sheet_by_name(sheet)
    pos = None
    for r in range(s.nrows - 2):
        for c in range(s.ncols - 3):
            if str(s.cell_value(r, c)).strip() == "급여환율" and str(s.cell_value(r + 1, c)).strip() == "EUR/RUB":
                pos = (r + 1, c + 1)
    if not pos:
        raise RuntimeError(f"환율!.xls {sheet} 시트에서 급여환율(EUR/RUB) 칸을 찾지 못함")
    r, c = pos
    (_, m1), (_, m2), _ = three_months(rate)
    label = f"{rate.month}월(~{rate.day})"
    want = {(r, c): float(m1), (r, c + 1): float(m2), (r, c + 2): label,
            (r + 1, c): float(vals[0]), (r + 1, c + 1): float(vals[1]), (r + 1, c + 2): float(vals[2])}
    where = f"환율!.xls {sheet} {xlrd.cellname(r, c)}:{xlrd.cellname(r + 1, c + 2)}"
    if all(s.cell_value(*k) == v for k, v in want.items()):
        return {"상태": "이미 기록됨", "위치": where}
    patch_xls(src, dst, sheet, want)
    return {"상태": "기록", "위치": where}


def _col(c: str, k: int) -> str:
    return chr(ord(c) + k)


def write_pay_xlsx(src, dst, rate: dt.date, usd, cny, year: int, shots: dict | None) -> dict:
    """급여 정리 xlsx '환율' 시트: USD/RUB, CNY/RUB 블록 (머리글 행 = 표 이름 행, 값 = 다음 행)."""
    import openpyxl  # 위치 찾기만 (쓰기는 xlsx_patch)
    from xlsx_patch import read_values, set_cells
    ws = openpyxl.load_workbook(src, read_only=True)["환율"]
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=60, max_col=10):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() in ("USD/RUB", "CNY/RUB"):
                pos[cell.value.strip()] = (cell.column_letter, cell.row)
    if set(pos) != {"USD/RUB", "CNY/RUB"}:
        raise RuntimeError("급여 파일 '환율' 시트에서 USD/RUB, CNY/RUB 표를 찾지 못함")
    (_, m1), (_, m2), _ = three_months(rate)
    label = f"{rate.month}월\n({rate.month}/1~{rate.day})"
    uc, ur = pos["USD/RUB"]
    cc, cr = pos["CNY/RUB"]
    d = lambda k, r: f"{_col(uc if r in (ur, ur + 1) else cc, k + 1)}{r}"  # noqa: E731
    values = {d(0, ur): m1, d(1, ur): m2, d(2, ur): label}
    values.update({d(k, ur + 1): float(v) for k, v in enumerate(usd["월평균"])})
    values.update({d(k, cr + 1): float(v) for k, v in enumerate(cny["월평균"])})
    cached = {d(3, ur + 1): float(usd["급여환율"]), d(3, cr + 1): float(cny["급여환율"]),
              d(0, cr): m1, d(1, cr): m2, d(2, cr): label}
    # 표 위 연도 칸 '(26년)' -> 그해
    for row in ws.iter_rows(min_row=1, max_row=ur, max_col=10):
        for cell in row:
            if isinstance(cell.value, str) and re.fullmatch(r"\(\d\d년\)", cell.value.strip()):
                values[cell.coordinate] = f"({year % 100:02d}년)"
    where = f"'환율' 시트 {d(0, ur)}:{d(2, ur + 1)}, {d(0, cr + 1)}:{d(2, cr + 1)}"
    now = read_values(src, "환율", list(values))
    if all((now[k] == v) if isinstance(v, str) else (now[k] is not None and abs(now[k] - v) < 1e-9) for k, v in values.items()):
        return {"상태": "이미 기록됨", "위치": where}
    set_cells(src, dst + ".1", "환율", values, cached)
    if shots:
        from xlsx_patch import replace_pictures
        replace_pictures(dst + ".1", dst + ".2", "환율", [shots["USD"], shots["CNY"]])  # 왼쪽 USD, 오른쪽 CNY
        os.remove(dst + ".1")
        os.replace(dst + ".2", dst + ".1")
        where += ", 캡처 2장 교체"
    else:
        where += ", 캡처 실패 - 그림은 이전 것 그대로"
    from xlsx_patch import focus
    focus(dst + ".1", dst, "환율", d(3, ur + 1))  # 열면 '환율' 시트 USD 급여환율 칸
    os.remove(dst + ".1")
    return {"상태": "기록", "위치": where}


def _download(s, fid, path):
    import fx_job
    r = s.get(f"{fx_job.API}/{fid}", params={"alt": "media", "supportsAllDrives": "true"})
    r.raise_for_status()
    open(path, "wb").write(r.content)


def _upload(s, fid, path, mime):
    import os
    if os.environ.get("DRY_RUN") == "1":
        print(f"[시험 실행] 업로드 생략: {fid}")
        return
    up = s.patch(f"https://www.googleapis.com/upload/drive/v3/files/{fid}",
                 params={"uploadType": "media", "supportsAllDrives": "true"},
                 data=open(path, "rb").read(), headers={"Content-Type": mime})
    up.raise_for_status()


def _list_pay(s, year: int):
    import fx_job
    q = (f"'{PAY_FOLDER_ID}' in parents and trashed = false and name contains '{year % 100:02d}년 한국인 급여 정리' "
         "and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'")
    r = s.get(fx_job.API, params={"q": q, "fields": "files(id,name)", "supportsAllDrives": "true",
                                  "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    return r.json()["files"]


def find_pay_file(s, year: int) -> tuple[str, str, bool]:
    """(파일 id, 이름, 새로 만들었는지). 그해 파일이 없으면 전년 파일을 드라이브에서 복사해 만든다."""
    import fx_job
    files = _list_pay(s, year)
    if len(files) == 1:
        return files[0]["id"], files[0]["name"], False
    if files:
        raise RuntimeError(f"급여 폴더에 '{year % 100:02d}년 한국인 급여 정리' 파일이 {len(files)}개 (1개여야 함)")
    prev = _list_pay(s, year - 1)
    if len(prev) != 1:
        raise RuntimeError(f"'{year % 100:02d}년 급여 파일도, 복사할 '{(year - 1) % 100:02d}년 파일도 없음")
    name = prev[0]["name"].replace(f"{(year - 1) % 100:02d}년", f"{year % 100:02d}년", 1)
    if os.environ.get("DRY_RUN") == "1":
        print(f"[시험 실행] 새 파일 생성 생략: {name}")
        return prev[0]["id"], name, True
    r = s.post(f"{fx_job.API}/{prev[0]['id']}/copy", params={"supportsAllDrives": "true"},
               json={"name": name, "parents": [PAY_FOLDER_ID]})
    r.raise_for_status()
    return r.json()["id"], name, True


def main() -> int:
    import fx_job
    if len(sys.argv) > 1 and sys.argv[1].strip():  # 예약 실행은 빈 문자열이 넘어옴 -> 오늘
        today = dt.date.fromisoformat(sys.argv[1].strip())
        hit = find_payday(today) or (None, None, today)
        if len(sys.argv) > 2 and sys.argv[2]:  # 시험용: 지급일 지정
            p = dt.date.fromisoformat(sys.argv[2])
            hit = (p, p, today)
    else:
        today = dt.datetime.now(KLD).date()
        hit = find_payday(today)
        if not hit:
            print(f"{today}: 급여환율 기준일 아님")
            return 0
    nominal, pay, rate = hit
    rates = salary_rates(rate)
    print("x-rates 월평균 확인 완료")  # 공개 저장소 실행 기록이라 값은 결과 메일로만
    s = fx_job.drive()
    notes = []
    with tempfile.TemporaryDirectory() as tmp:
        fid = fx_job.find_file(s)
        _download(s, fid, f"{tmp}/fx.xls")
        res = write_eur_xls(f"{tmp}/fx.xls", f"{tmp}/fx_out.xls", rate, rates["EUR"]["월평균"])
        if res["상태"] == "기록":
            _upload(s, fid, f"{tmp}/fx_out.xls", "application/vnd.ms-excel")
        notes.append(f"EUR: {res['위치']} ({res['상태']})")
        year = (nominal or rate).year
        pid, pname, created = find_pay_file(s, year)
        if created:
            notes.append(f"새 급여 파일 생성: {pname} (전년 파일 복사)")
        _download(s, pid, f"{tmp}/pay.xlsx")
        try:
            import xrates_shot
            shots = xrates_shot.capture(("USD", "CNY"), rate.year)
        except Exception as e:  # noqa: BLE001
            print(f"캡처 실패: {e}")
            shots = None
        res = write_pay_xlsx(f"{tmp}/pay.xlsx", f"{tmp}/pay_out.xlsx", rate, rates["USD"], rates["CNY"], year, shots)
        if shots and os.environ.get("DRY_RUN") == "1":
            for k, v in shots.items():
                open(f"shot_{k}.png", "wb").write(v)
        if res["상태"] == "기록":
            _upload(s, pid, f"{tmp}/pay_out.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        notes.append(f"USD, CNY: {pname} {res['위치']} ({res['상태']})")
    print("기록:", ", ".join(n.split("(")[-1].rstrip(")") for n in notes))  # 상태만
    if all("이미 기록됨" in n for n in notes if "생성" not in n) and os.environ.get("DRY_RUN") != "1" \
            and os.environ.get("REBUILD_NOTICE") != "1":  # 공지 양식을 고쳤을 때만 대기열 공지를 다시 만듦
        return 0
    if os.environ.get("DRY_RUN") == "1" and os.environ.get("PREVIEW_TEST") != "1":
        return 0
    title = f"{nominal:%m/%d} 급여" if nominal else "급여"
    if pay and (os.environ.get("DRY_RUN") != "1" or os.environ.get("PREVIEW_TEST") == "1"):
        import notice_queue  # 공지 메일 3통을 제어판 발송 대기열에 올림 (비밀번호 승인 후 발송)
        import salary_notice
        test = os.environ.get("DRY_RUN") == "1"
        mails = salary_notice.build(rates, rate, pay)
        summary = [f"{c}/RUB {rates[c]['급여환율']} (3개월 평균 {rates[c]['3개월 평균']})" for c in ("EUR", "USD", "CNY")]
        key = f"salary-{nominal:%Y%m%d}" + ("-test" if test else "")
        notice_queue.add_notice(s, key, ("[테스트] " if test else "") + f"{title} 급여환율 공지 3통", mails,
                                [f"기준일 {rate}, 지급일 {pay}"] + summary, test=test)
        notes.append("공지 메일 3통을 제어판 발송 대기열에 올림 - 제어판에서 확인 후 비밀번호로 발송"
                     + (f": {notice_queue.PANEL_URL}" if notice_queue.PANEL_URL else ""))
    lines = [f"{title} 환율 (기준일 {rate}, 지급일 {pay or '-'}) - x-rates 월평균 {', '.join(rates['월'])}", ""]
    for cur in ("EUR", "USD", "CNY"):
        x = rates[cur]
        rule = "소수 2자리 올림" if cur == "CNY" else "올림"
        lines.append(f"{cur}/RUB: {x['급여환율']}  (3개월 평균 {x['3개월 평균']}, {rule}) <- {' / '.join(x['월평균'])}")
    lines += ["", "기록:"] + [f"- {n}" for n in notes]
    lines.append("수식은 엑셀에서 열 때 다시 계산됩니다. 이전 버전은 드라이브 버전 관리에 남아 있습니다.")
    subject = f"[급여환율] {title}: EUR {rates['EUR']['급여환율']} / USD {rates['USD']['급여환율']} / CNY {rates['CNY']['급여환율']}"
    if os.environ.get("DRY_RUN") != "1":
        fx_job.mail(subject, "\n".join(lines))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        import fx_job
        now = dt.datetime.now(KLD)
        last_try = len(sys.argv) > 1 or now.hour >= 12 or (now.hour == 11 and now.minute >= 45)
        if last_try:  # 12시 회차까지 안 되면 알림
            fx_job.mail("[급여환율] 12시까지 처리 못함 - 확인 필요",
                        f"오류: {e}\n\n08:00부터 30분마다 확인했지만 12:00 회차까지 처리하지 못했습니다.\n"
                        "x-rates가 늦게 업데이트되면 Actions > 급여환율 > Run workflow에 오늘 날짜를 넣어 다시 실행하거나 직접 확인해 주세요.")
            raise
        print(f"다음 회차(30분 뒤)에 다시 시도: {e}")  # 중간 회차는 실패로 표시 안 함 (GitHub 실패 메일 방지)
        sys.exit(0)
