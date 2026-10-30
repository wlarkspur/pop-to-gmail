"""하나은행 고시환율에서 HAS 손익보고용 환율 두 가지를 가져온다. 고시회차는 둘 다 최초(1회차).

1) 월말 환율 (6-1. 재무제표(USD) J8)
   현재환율 화면: 통화 RUB, 일자 = 해당 월 마지막 날, 값 = 미화환산율.
   그날 고시가 없으면(주말, 공휴일) 하나은행이 직전 영업일 고시를 돌려준다. 실제 기준일을 같이 출력한다.
2) 월평균 환율 (15-4. 26년 실적(USD) 3행 해당 월)
   평균환율 화면: 월평균, 해당 월, 러시아 RUB, 값 = 대미환산율.

사용: python3 fx_hana.py 2026-09          -> 두 값 모두 (9월 말일, 9월 평균)
      python3 fx_hana.py 2026-09 USD      -> 다른 통화
출력(JSON 한 줄)
"""

from __future__ import annotations

import calendar
import http.cookiejar
import json
import re
import sys
import urllib.parse
import urllib.request

BASE = "https://www.kebhana.com/cms/rate"
NAMES = {"RUB": "러시아 RUB", "USD": "미국 USD", "EUR": "유로 EUR", "CNY": "중국 CNY"}


def _opener(content: str):
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    opener.addheaders = [("User-Agent", "Mozilla/5.0"), ("X-Requested-With", "XMLHttpRequest"),
                         ("Referer", f"{BASE}/index.do?contentUrl=/cms/rate/{content}")]
    opener.open(f"{BASE}/index.do?contentUrl=/cms/rate/{content}", timeout=30).read()
    return opener


def _post(opener, endpoint: str, fields: dict) -> str:
    raw = opener.open(f"{BASE}/{endpoint}", data=urllib.parse.urlencode(fields).encode(), timeout=30).read()
    for enc in ("euc-kr", "utf-8"):
        try:
            return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw.decode(enc)))
        except UnicodeDecodeError:
            continue
    raise RuntimeError("하나은행 응답을 읽지 못함")


def _row(text: str, cur: str, count: int) -> list[float]:
    m = re.search(re.escape(NAMES.get(cur, cur)) + r"((?: [\d.,]+){%d})" % count, text)
    if not m:
        raise RuntimeError("하나은행 응답에서 환율 표를 찾지 못함 (사이트 구조 변경 가능성)")
    return [float(x.replace(",", "")) for x in m.group(1).split()]


def month_end_rate(year: int, month: int, cur: str = "RUB") -> dict:
    """현재환율: 해당 월 말일, 최초 고시, 미화환산율."""
    date = f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"
    text = _post(_opener("wpfxd651_01i.do"), "wpfxd651_01i_01.do", {
        "tmpInqStrDt": date, "pbldDvCd": "1", "pbldSqn": "", "curCd": cur,
        "inqStrDt": date.replace("-", ""), "inqKindCd": "1",
    })
    base = re.search(r"기준일 : (\d{4})년(\d{2})월(\d{2})일", text)
    notice = re.search(r"고시일시 : (\d{4}년\d{2}월\d{2}일 \d{2}시\d{2}분\d{2}초) \((\d+)회차\)", text)
    if not (base and notice):
        raise RuntimeError("월말 환율: 기준일/고시회차를 찾지 못함")
    if notice.group(2) != "1":
        raise RuntimeError(f"월말 환율: 최초 고시가 아님 ({notice.group(2)}회차)")
    nums = _row(text, cur, 10)  # 현찰(사실,스프레드,파실,스프레드) 송금(보낼,받을) 수표 매매기준율 환가료율 미화환산율
    return {"요청일": date, "기준일": "-".join(base.groups()),
            "고시": f"{notice.group(1)} ({notice.group(2)}회차)",
            "매매기준율": nums[7], "미화환산율": nums[9]}


def month_avg_rate(year: int, month: int, cur: str = "RUB") -> dict:
    """평균환율: 월평균, 최초 고시, 대미환산율."""
    last = calendar.monthrange(year, month)[1]
    text = _post(_opener("wpfxd651_06i.do"), "wpfxd651_06i_01.do", {
        "inqDvCd": "2", "tmpInqStrDtY_m": f"{year:04d}", "tmpInqStrDtM_m": f"{month:02d}",
        "tmpPbldDvCd": "1", "pbldDvCd": "1",
        "inqStrDt": f"{year:04d}{month:02d}01", "inqEndDt": f"{year:04d}{month:02d}{last:02d}",
    })
    if "대미 환산율" not in text:
        raise RuntimeError("월평균 환율: 표 머리글(대미 환산율)을 찾지 못함")
    nums = _row(text, cur, 8)  # 현찰(사실,파실) 송금(보낼,받을) 수표 매매기준율 연환가료율 대미환산율
    return {"기간": f"{year:04d}-{month:02d}-01 ~ {year:04d}-{month:02d}-{last:02d}",
            "고시회차": "최초", "매매기준율": nums[5], "대미환산율": nums[7]}


if __name__ == "__main__":
    if len(sys.argv) < 2 or not re.fullmatch(r"\d{4}-\d{2}", sys.argv[1]):
        raise SystemExit(__doc__)
    y, m = map(int, sys.argv[1].split("-"))
    cur = (sys.argv[2] if len(sys.argv) > 2 else "RUB").upper()
    print(json.dumps({"통화": cur,
                      "월말(6-1 J8)": month_end_rate(y, m, cur),
                      "월평균(15-4 3행)": month_avg_rate(y, m, cur)}, ensure_ascii=False))
