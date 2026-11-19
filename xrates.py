"""x-rates.com 월평균 환율 (급여환율용). 화면에 보이는 소수 6자리 값 그대로 쓴다."""
from __future__ import annotations

import re
import urllib.request
from decimal import Decimal

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def monthly(cur: str, year: int, to: str = "RUB") -> dict[int, tuple[Decimal, int]]:
    """{월: (월평균, 집계 일수)}. 이번 달은 오늘까지 평균."""
    url = f"https://www.x-rates.com/average/?from={cur}&to={to}&amount=1&year={year}"
    html = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=60).read().decode("utf-8")
    # 2026-11 사이트 개편 후: <td class="... avgMonth ...">Jan</td><td class="... avgRate ...">91.70</td><td ... avgDays>– 31 days</td>
    rows = re.findall(r'<td[^>]*\bavgMonth\b[^>]*>(\w+)</td>\s*<td[^>]*\bavgRate\b[^>]*>([\d.]+)</td>\s*'
                      r'<td[^>]*\bavgDays\b[^>]*>[^\d<]*(\d+) days?</td>', html)
    if not rows:  # 예전 구조
        rows = re.findall(r'class="avgMonth">(\w+)</span>.*?class="avgRate">([\d.]+)</span>.*?class="avgDays">[^\d<]*(\d+) days?</span>', html)
    if not rows:
        raise RuntimeError(f"x-rates {cur}/{to} {year}: 월평균 표를 찾지 못함 (사이트 구조 변경 가능성)")
    return {MONTHS.index(m) + 1: (Decimal(v), int(n)) for m, v, n in rows}
