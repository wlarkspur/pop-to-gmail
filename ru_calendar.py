"""러시아 생산 달력(공휴일, 대체 휴일, 주말)과 급여환율 기준일 계산.

출처: 저장소의 ru_holidays.json (연도별 쉬는 날 목록, calendar.yml이 1년에 한 번 받아 커밋). 파일에 그 해가 없을 때만
      isdayoff.ru (러시아 공식 생산 달력 기반, 1 = 쉬는 날), 실패하면 xmlcalendar.ru (컨설턴트플러스 달력)에 접속.
급여일: 매달 10일, 25일. 쉬는 날이면 직전 근무일에 지급.
급여환율 기준일: 급여일을 1일째로 세어 근무일 기준 5일째 되는 날 (예: 2026-11-10 → 11-03, 11-04 공휴일 제외).
"""
from __future__ import annotations

import datetime as dt
import json
import pathlib
import urllib.request

_cache: dict[int, set] = {}
FILE = pathlib.Path(__file__).with_name("ru_holidays.json")


def _load_file() -> dict[int, set]:
    if not FILE.exists():
        return {}
    j = json.loads(FILE.read_text(encoding="utf-8"))
    return {int(y): {dt.date.fromisoformat(d) for d in days} for y, days in j.items()}


def has_year(year: int) -> bool:
    """저장 파일에 그 해 달력이 있는가 (네트워크 없이 판단 가능한가)."""
    return year in _load_file()


def update_file(years) -> list[int]:
    """사이트에서 받아 ru_holidays.json에 저장 (있는 해는 덮어씀). 받은 해 목록 반환."""
    data = {str(y): sorted(d.isoformat() for d in days) for y, days in _load_file().items()}
    got = []
    for y in years:
        try:
            days = _fetch(y)
        except Exception as e:  # noqa: BLE001
            print(f"{y}년 달력 못 받음: {e}")
            continue
        data[str(y)] = sorted(d.isoformat() for d in days)
        got.append(y)
    FILE.write_text(json.dumps(dict(sorted(data.items())), ensure_ascii=False, indent=0), encoding="utf-8")
    return got


def _get(url: str) -> bytes:
    return urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30).read()


def holidays(year: int) -> set:
    """그해 쉬는 날(주말 포함) 집합. 저장 파일 우선, 없으면 사이트."""
    if year in _cache:
        return _cache[year]
    saved = _load_file()
    if year in saved:
        _cache[year] = saved[year]
        return saved[year]
    _cache[year] = _fetch(year)
    return _cache[year]


def _fetch(year: int) -> set:
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
    if len(out) < 100:
        raise ValueError(f"{year}년 쉬는 날이 {len(out)}일뿐 - 자료 이상")
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
    if len(sys.argv) > 1 and sys.argv[1] == "--update":
        ys = [int(a) for a in sys.argv[2:]] or [dt.date.today().year, dt.date.today().year + 1]
        print("저장:", update_file(ys))
        sys.exit(0)
    y = int(sys.argv[1]) if len(sys.argv) > 1 else dt.date.today().year
    for m in range(1, 13):
        for nominal, pay, rate in schedule(y, m):
            print(nominal, "지급", pay, "환율 기준일", rate)
