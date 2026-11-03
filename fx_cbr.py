"""러시아 중앙은행(CBR) EUR 공식환율로 환율!.xls의 지난달 일별 칸을 채운다.

규칙 (기존 파일 273일치로 확인): 날짜 D의 값 = D에 적용되는 CBR 고시환율.
주말, 공휴일은 고시가 없으므로 직전 고시값을 그대로 쓴다.
시트 = "{YY}년", 해당 날짜가 적힌 칸의 오른쪽 두 번째 칸(EURO 다음)이 환율 칸.
이미 값이 있는 칸은 덮어쓰지 않고, CBR과 다르면 목록으로 알린다.
새 해 시트: 12월분을 채울 때(1월 1~3일 실행) 전년 시트를 복사해 만든다 (날짜 +1년, 환율과 수기 숫자 비움, 수식 유지).
"""
from __future__ import annotations
import calendar, datetime as dt, json, re, sys, urllib.error, urllib.request
import xlrd
import calendar as _cal
import shutil
from xls_patch import add_year_sheet_xls, patch_xls

def cbr_eur(d1: dt.date, d2: dt.date) -> dict:
    """CBR 공식 사이트, 막히면 같은 자료를 그대로 옮겨두는 cbr-xml-daily.ru 아카이브."""
    try:
        return _cbr_site(d1, d2)
    except Exception as e:  # noqa: BLE001
        print(f"CBR 사이트 실패, 아카이브로 대체: {e}", file=sys.stderr)
        return _cbr_archive(d1, d2)


def _cbr_archive(d1: dt.date, d2: dt.date) -> dict:
    out, d = {}, d1
    while d <= d2:
        u = f"https://www.cbr-xml-daily.ru/archive/{d:%Y/%m/%d}/daily_json.js"
        try:
            j = json.loads(urllib.request.urlopen(urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"}), timeout=30).read())
            out[dt.date.fromisoformat(j["Date"][:10])] = float(j["Valute"]["EUR"]["Value"])
        except urllib.error.HTTPError as e:
            if e.code != 404:  # 404 = 그날 고시 없음 (주말, 공휴일)
                raise
        d += dt.timedelta(days=1)
    if not out:
        raise RuntimeError("아카이브에도 환율 없음")
    return out


def _cbr_site(d1: dt.date, d2: dt.date) -> dict:
    u = ("https://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=%s&date_req2=%s&VAL_NM_RQ=R01239"
         % (d1.strftime("%d/%m/%Y"), d2.strftime("%d/%m/%Y")))
    t = urllib.request.urlopen(urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"}), timeout=60).read().decode("cp1251")
    out = {dt.datetime.strptime(a, "%d.%m.%Y").date(): float(b.replace(",", "."))
           for a, b in re.findall(r'Date="([\d.]+)"[^>]*>.*?<Value>([\d,]+)</Value>', t)}
    if not out:
        raise RuntimeError("CBR 응답에 환율 없음")
    return out

def daily(year, month):
    last = calendar.monthrange(year, month)[1]
    first = dt.date(year, month, 1)
    rates = cbr_eur(first - dt.timedelta(days=15), dt.date(year, month, last))
    days = [first + dt.timedelta(i) for i in range(last)]
    if max(rates) < days[-1] - dt.timedelta(days=5):
        raise RuntimeError("CBR 자료가 월말까지 없음")
    return {d: rates[max(k for k in rates if k <= d)] for d in days}

def ensure_year_sheet(path, year) -> bool:
    """'{YY}년' 시트가 없으면 전년 시트를 복사해 만든다. 만들었으면 True."""
    name, prev = f"{year % 100:02d}년", f"{(year - 1) % 100:02d}년"
    if name in xlrd.open_workbook(path, on_demand=True).sheet_names():
        return False
    if _cal.isleap(year) != _cal.isleap(year - 1):
        raise RuntimeError(f"{name}: 윤년이라 날짜 칸 수가 달라 자동 생성 불가 - 직접 만들어 주세요")
    shift = (dt.date(year, 1, 1) - dt.date(year - 1, 1, 1)).days
    tmp = path + ".new"
    add_year_sheet_xls(path, tmp, prev, name, shift, f"{(year - 1) % 100:02d}Y", f"{year % 100:02d}Y")
    shutil.move(tmp, path)
    return True


def report_block(path, sheet, month) -> dict:
    """시트 하단 '*월간 보고서 환율' 블록의 월 칸 3개를 (지난달-2, 지난달-1, 지난달)로 바꾼다.
    월 평균은 원래 수식(INDEX/MATCH)이 그 월 번호로 가져오고, 3개월 평균도 수식이 계산한다."""
    s = xlrd.open_workbook(path).sheet_by_name(sheet)
    pos = None
    for r in range(s.nrows):
        for c in range(s.ncols):
            if str(s.cell_value(r, c)).strip() == "*월간 보고서 환율" and r + 2 < s.nrows \
                    and str(s.cell_value(r + 1, c)).strip() == "CBR" and str(s.cell_value(r + 3, c)).strip() == "3개월 평균":
                pos = (r + 1, c + 1)  # 마지막으로 찾은 것 = 시트 맨 아래 블록
    if not pos:
        return {"상태": "블록 못 찾음"}
    if month < 3:
        return {"상태": f"{month}월은 전년 달이 섞여 자동 갱신 안 함 - 확인 필요"}
    want = [month - 2, month - 1, month]
    r, c = pos
    now = [s.cell_value(r, c + k) for k in range(3)]
    if now == want:
        return {"상태": "이미 최신", "월": want}
    patch_xls(path, path + ".3", sheet, {(r, c + k): v for k, v in enumerate(want)}, focus=(r + 2, c))  # 열면 3개월 평균 칸
    shutil.move(path + ".3", path)
    return {"상태": "갱신", "월": want, "바꿈": True}


def run(src, dst, year, month):
    """src를 읽어 dst에 결과를 쓴다. 바뀐 게 없으면 dst를 만들지 않는다."""
    work = dst + ".work"
    shutil.copy(src, work)
    created = []
    if ensure_year_sheet(work, year):          # 1월분인데 새 해 시트가 없을 때
        created.append(f"{year % 100:02d}년")
    vals = daily(year, month)
    sheet = f"{year % 100:02d}년"
    s = xlrd.open_workbook(work).sheet_by_name(sheet)
    where = {}
    for r in range(s.nrows):
        for c in range(s.ncols):
            x = s.cell(r, c)
            if x.ctype == 3:
                d = dt.date(*xlrd.xldate_as_tuple(x.value, 0)[:3])
                if d in vals:
                    where[d] = (r, c + 2)
    missing = [str(d) for d in vals if d not in where]
    if missing:
        raise RuntimeError(f"{sheet} 시트에 날짜 칸이 없음: {missing[:3]}...")
    cells, kept, diff = {}, 0, []
    for d, (r, c) in where.items():
        cur = s.cell(r, c)
        if cur.ctype == 2:
            kept += 1
            if abs(cur.value - vals[d]) > 1e-9:
                diff.append(f"{d} 파일 {cur.value} / CBR {vals[d]}")
        else:
            cells[(r, c)] = vals[d]
    if cells:
        patch_xls(work, work + ".2", sheet, cells)
        shutil.move(work + ".2", work)
    report = report_block(work, sheet, month)
    if month == 12 and ensure_year_sheet(work, year + 1):   # 12월분 채울 때 다음 해 시트 미리
        created.append(f"{(year + 1) % 100:02d}년")
    if cells or created or report.get("바꿈"):
        shutil.move(work, dst)
    else:
        import os
        os.remove(work)
    avg = sum(vals.values()) / len(vals)
    return {"시트": sheet, "월": f"{year}-{month:02d}", "채운 칸": len(cells), "이미 있던 칸": kept,
            "새로 만든 시트": created, "월간 보고서 환율": report, "CBR과 다른 기존 값": diff,
            "월평균(일수 평균)": round(avg, 6), "말일 환율": vals[max(vals)]}


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    if len(sys.argv) > 3:
        y, m = map(int, sys.argv[3].split("-"))
    else:
        t = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        y, m = t.year, t.month
    print(json.dumps(run(src, dst, y, m), ensure_ascii=False, default=str))
