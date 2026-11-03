"""러시아 생산 달력(공휴일, 대체 휴일, 주말)과 급여환율 기준일 계산.

출처: isdayoff.ru (러시아 공식 생산 달력 기반, 1 = 쉬는 날), 실패하면 xmlcalendar.ru (컨설턴트플러스 달력).
급여일: 매달 10일, 25일. 쉬는 날이면 직전 근무일에 지급.
급여환율 기준일: 급여일을 1일째로 세어 근무일 기준 5일째 되는 날 (예: 2026-11-10 → 11-03, 11-04 공휴일 제외).
"""
from __future__ import annotations

import datetime as dt
import json
import urllib.request

_cache: dict[int, set] = {}


def _get(url: str) -> bytes:
    return urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30).read()


def holidays(year: int) -> set:
    """그해 쉬는 날(주말 포함) 집합."""
    if year in _cache:
        return _cache[year]
    try:
        flags = _get(f"https://isdayoff.ru/api/getdata?year={year}&cc=ru").decode().strip()
        days = (dt.date(year, 12, 31) - dt.date(year, 1, 1)).days + 1
        if len(flags) != days or set(flags) - {"0", "1"}:
            raise ValueError("형식 다름")
        out = {dt.date(year, 1, 1) + dt.timedelta(i) for i, f in enumerate(flags) if f == "1"}
    except Exception:  # noqa: BLE001
        j = json.loads(_get(f"https://xmlcalendar.ru/data/ru/{year}/calendar.json"))
        out = set()
        for m in j["months"]:
            for x in m["days"].split(","):
                if "*" in x:  # 단축 근무일 = 근무일
                    continue
                out.add(dt.date(year, m["month"], int(x.strip("+"))))
    _cache[year] = out
    return out


def is_workday(d: dt.date) -> bool:
    return d not in holidays(d.year)


def pay_day(year: int, month: int, day: int) -> dt.date:
    d = dt.date(year, month, day)
    while not is_workday(d):
        d -= dt.timedelta(days=1)
    return d


def rate_day(pay: dt.date, n: int = 5) -> dt.date:
    d, k = pay, 1
    while k < n:
        d -= dt.timedelta(days=1)
        if is_workday(d):
            k += 1
    return d


def schedule(year: int, month: int):
    """그달 급여 2회: [(명목 급여일, 실제 지급일, 환율 기준일)]"""
    return [(dt.date(year, month, n), pay_day(year, month, n), rate_day(pay_day(year, month, n))) for n in (10, 25)]


if __name__ == "__main__":
    import sys
    y = int(sys.argv[1]) if len(sys.argv) > 1 else dt.date.today().year
    for m in range(1, 13):
        for nominal, pay, rate in schedule(y, m):
            print(nominal, "지급", pay, "환율 기준일", rate)
