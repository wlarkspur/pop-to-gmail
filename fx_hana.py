"""하나은행 고시환율에서 통화의 '미화환산율'을 가져온다 (HAS 6-1 J8 월말 환율용).

규칙: 통화 RUB, 일자 = 가마감 해당 월 마지막 날, 고시회차 = 최초(1회차), 값 = 미화환산율.
그날 고시가 없으면(주말, 공휴일) 하나은행이 직전 영업일 고시를 돌려준다. 실제 기준일을 같이 출력한다.

사용: python3 fx_hana.py 2026-09            -> 2026-09-30 기준
      python3 fx_hana.py 2026-09-30 RUB
출력(JSON): {"요청일": ..., "기준일": ..., "고시": ..., "통화": ..., "미화환산율": ...}
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


def month_end(text: str) -> str:
    if re.fullmatch(r"\d{4}-\d{2}", text):
        y, m = map(int, text.split("-"))
        return f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return text
    raise SystemExit("날짜 형식: YYYY-MM 또는 YYYY-MM-DD")


def fetch(date: str, cur: str = "RUB") -> dict:
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    opener.addheaders = [("User-Agent", "Mozilla/5.0"), ("X-Requested-With", "XMLHttpRequest"),
                         ("Referer", f"{BASE}/index.do?contentUrl=/cms/rate/wpfxd651_01i.do")]
    opener.open(f"{BASE}/index.do?contentUrl=/cms/rate/wpfxd651_01i.do", timeout=30).read()
    data = urllib.parse.urlencode({
        "tmpInqStrDt": date, "pbldDvCd": "1", "pbldSqn": "", "curCd": cur,
        "inqStrDt": date.replace("-", ""), "inqKindCd": "1",
    }).encode()
    raw = opener.open(f"{BASE}/wpfxd651_01i_01.do", data=data, timeout=30).read()
    for enc in ("euc-kr", "utf-8"):
        try:
            html = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))

    base = re.search(r"기준일 : (\d{4})년(\d{2})월(\d{2})일", text)
    notice = re.search(r"고시일시 : (\d{4}년\d{2}월\d{2}일 \d{2}시\d{2}분\d{2}초) \((\d+)회차\)", text)
    row = re.search(re.escape(NAMES.get(cur, cur)) + r"((?: [\d.,]+){10})", text)
    if not (base and notice and row):
        raise RuntimeError("하나은행 응답에서 환율 표를 찾지 못함 (사이트 구조 변경 가능성)")
    nums = row.group(1).split()
    if notice.group(2) != "1":
        raise RuntimeError(f"최초 고시가 아님 ({notice.group(2)}회차)")
    return {
        "요청일": date,
        "기준일": "-".join(base.groups()),
        "고시": f"{notice.group(1)} ({notice.group(2)}회차)",
        "통화": cur,
        "매매기준율": float(nums[7].replace(",", "")),
        "미화환산율": float(nums[9].replace(",", "")),
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    result = fetch(month_end(sys.argv[1]), (sys.argv[2] if len(sys.argv) > 2 else "RUB").upper())
    print(json.dumps(result, ensure_ascii=False))
