"""HAS 손익보고 YS(YSH) 보고용 3개 파일: 양식은 그대로 두고 값만 바꿈.

1) 마감(YSH 보고)  'HAS법인 YY년 M월 마감(YSH 보고).xlsx'
   - '26년 실적(RUB)' 월 열(F=1월 ... Q=12월)에 Irina monthly 파일 '9-1 상세(RUB)'(G=1월 ...)의 같은 항목 값을 넣음.
     YS 보고는 매달 1~이번 달 전체를 monthly 기준으로 다시 맞춤 (지난달 수정분도 반영). 과거 월이 바뀐 칸은 목록으로.
   - '환율정보' 44행 = CBR EUR 월평균, summary B1 = 보고 월
2) 재무제표 (KR)  'HAS 2026년 재무제표 (KR).xlsx' '재무현황' D열 = Irina 'HAS 재무제표(RUB)' Balance Sheets / Income Statement J열
3) ARAP 차입대여금 현황  'HAS YYYY M월 ARAP 차입대여금 현황.xlsx'
   - 기준일, 5.AP 금액, 6.차입금 잔액 = Irina 'Конечные остатки' (연도 시트)
   - 3.AR(채권)은 송장 단위 수기 목록이라 바꾸지 않고, 거래처별 합계가 Конечные остатки와 다르면 목록으로
   - 6.차입금에 없는 차입처가 Конечные остатки에 생기면 목록으로 (행 추가는 사용자 판단)
"""
from __future__ import annotations

import datetime as dt
import re

import openpyxl

Y_RUB, Y_FX, Y_SUM, M_SHEET = "26년 실적(RUB)", "환율정보", "summary ", "9-1 상세(RUB)"
# 마감(YSH) '26년 실적(RUB)' 행 -> monthly '9-1 상세(RUB)' 행 (입력 행만, 소계/합계는 양식 수식)
ROWMAP = {9: 8, 10: 9, 11: 10, 12: 11, 14: 13, 15: 14, 16: 15, 17: 16, 18: 17, 19: 18, 20: 19, 21: 20, 22: 21,
          23: 22, 24: 23, 25: 24, 28: 27, 29: 28, 30: 29, 31: 30, 32: 31, 33: 32, 34: 33, 35: 34, 36: 37, 37: 38,
          38: 39, 39: 40, 40: 42, 41: 43, 42: 44, 43: 45, 44: 46, 45: 47, 46: 48, 47: 49, 48: 50, 49: 51, 50: 52,
          51: 53, 52: 54, 53: 55, 54: 56, 55: 57, 56: 58, 57: 59, 58: 60, 59: 61, 60: 62, 61: 63, 62: 64, 63: 65,
          64: 66, 65: 67, 66: 68, 67: 69, 68: 70, 69: 71, 70: 72, 71: 73, 72: 74, 73: 75, 74: 76, 75: 77, 76: 78,
          77: 79, 78: 80, 80: 82, 81: 83, 84: 86, 85: 87, 86: 88, 87: 89, 89: 91, 90: 92, 91: 93, 92: 94, 93: 101,
          94: 95, 95: 96, 96: 97, 97: 98, 98: 99, 99: 100, 103: 105, 104: 106, 105: 107}
# 합계 대조: (마감 입력 행의 합) = monthly 합계 행
CHECKS = {"매출원가": ([14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25], 12),
          "판관비": ([r for r in ROWMAP if 28 <= r <= 81], 26), "영업외수익": ([84, 85, 86, 87], 85),
          "영업외비용": ([r for r in ROWMAP if 89 <= r <= 99], 90), "법인세": ([103, 104, 105], 104)}


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def fill_ysh(tmpl: str, monthly: str, y: int, m: int, avg_rates: dict[int, float], out: str):
    wb = openpyxl.load_workbook(tmpl)
    rub, fx, sm = wb[Y_RUB], wb[Y_FX], wb[Y_SUM]
    k = openpyxl.load_workbook(monthly, data_only=True)[M_SHEET]
    log, notes = [], []
    if _num(k["E1"].value) != y:
        notes.append(f"monthly '{M_SHEET}' 연도 표시가 {k['E1'].value} (기대 {y})")
    for cr, kr in ROWMAP.items():  # 양식이 바뀌었는지: 영문 항목명 대조
        a, b = _norm(rub.cell(cr, 3).value), _norm(k.cell(kr, 3).value)
        if a and b and a[:14] != b[:14]:
            notes.append(f"항목명 불일치: 마감 {cr}행 '{rub.cell(cr, 3).value}' vs monthly {kr}행 '{k.cell(kr, 3).value}'")
    for mon in range(1, m + 1):
        cc, kc = 5 + mon, 6 + mon
        for cr, kr in ROWMAP.items():
            kv = round(_num(k.cell(kr, kc).value), 2)
            cur = rub.cell(cr, cc)
            if isinstance(cur.value, str) and cur.value.startswith("="):
                if kv:
                    notes.append(f"{Y_RUB}!{cur.coordinate} 가 수식이라 monthly 값 {kv:,.2f} 을 못 넣음")
                continue
            old = round(_num(cur.value), 2)
            if kv != old:
                cur.value = kv if kv else None
                log.append(cur.coordinate)
                if mon < m:
                    notes.append(f"{mon}월 {Y_RUB}!{cur.coordinate} ({rub.cell(cr, 5).value or rub.cell(cr, 3).value}) "
                                 f"{old:,.2f} → {kv:,.2f} (monthly 수정분 반영)")
        for name, (rows, tot) in CHECKS.items():
            s = round(sum(_num(k.cell(ROWMAP[r], kc).value) for r in rows), 2)
            t = round(_num(k.cell(tot, kc).value), 2)
            if abs(s - t) > 0.01:
                notes.append(f"{mon}월 {name}: 넣은 행 합 {s:,.2f} ≠ monthly 합계 {t:,.2f} (매핑 안 된 항목이 있음)")
        fc = fx.cell(44, 2 + mon)  # 지난 달 환율은 이미 적힌 값을 그대로 둠 (비어 있을 때만 채움)
        if mon in avg_rates and (mon == m or not _num(fc.value)):
            if round(_num(fc.value), 6) != round(avg_rates[mon], 6):
                fc.value = avg_rates[mon]
                log.append(f"{Y_FX}!{fc.coordinate}")
    if sm["B1"].value != m:
        sm["B1"] = m
        log.append(f"{Y_SUM.strip()}!B1")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return log, notes


# ─────────────────────────── 재무제표 (KR)
K_SHEET = "재무현황"
# '재무현황' D행 -> (시트, J열 행 목록) 합계
KRMAP = {6: ("Balance Sheets", [13]), 7: ("Balance Sheets", [15]), 8: ("Balance Sheets", list(range(23, 35))),
         9: ("Balance Sheets", [35]), 11: ("Balance Sheets", [45]), 12: ("Balance Sheets", [72]),
         13: ("Balance Sheets", [89]), 18: ("Balance Sheets", [103]), 20: ("Balance Sheets", [100]),
         21: ("Balance Sheets", list(range(105, 117))), 24: ("Balance Sheets", [120, 121, 122]),
         27: ("Balance Sheets", [130]), 30: ("Balance Sheets", [137]), 31: ("Balance Sheets", [146]),
         32: ("Balance Sheets", [152]), 39: ("Income Statement", [11]), 41: ("Income Statement", [20]),
         43: ("Income Statement", [29]), 46: ("Income Statement", [33]), 47: ("Income Statement", [36, 37]),
         48: ("Income Statement", [42]), 50: ("Income Statement", [44]), 51: ("Income Statement", [47, 48]),
         52: ("Income Statement", list(range(51, 57))), 54: ("Income Statement", [75])}


def fill_kr(tmpl: str, fs: str, base: dt.date, rate: float, out: str):
    wb = openpyxl.load_workbook(tmpl)
    ws = wb[K_SHEET]
    src = openpyxl.load_workbook(fs, data_only=True)
    log, notes = [], []
    for sheet in ("Balance Sheets", "Income Statement"):
        hdr = " ".join(str(src[sheet].cell(r, 10).value or "") for r in (9, 10))
        if base.strftime("%Y. %m.%d") not in hdr.replace(" ", " ") and f"{base:%Y}. {base:%m}.{base:%d}" not in hdr:
            notes.append(f"재무제표 '{sheet}' 기준일 표시 '{hdr.strip()}' (기대 {base:%Y. %m.%d})")
    for r, (sheet, rows) in KRMAP.items():
        v = round(sum(_num(src[sheet].cell(x, 10).value) for x in rows), 2)
        c = ws.cell(r, 4)
        if isinstance(c.value, str) and c.value.startswith("=") and "[" not in c.value:
            notes.append(f"{K_SHEET}!D{r} 가 수식이라 그대로 둠")
            continue
        c.value = v  # 외부 파일 연결 수식([1]...)은 값으로 바꿈 (사용자 10월 파일과 같은 방식)
        log.append(f"D{r}")
    bs = src["Balance Sheets"]
    if not bs["J135"].value is True and str(bs["J135"].value).upper() != "TRUE":
        notes.append("재무제표 Balance Sheets 대차 일치 표시(J135)가 TRUE가 아님")
    ws["D2"] = f"{base:%Y.%m.%d}"
    ws["D60"] = rate
    log += ["D2", "D60"]
    notes.append(f"{K_SHEET}!D59 평균환율은 근거를 몰라 그대로 둠 ({_num(ws['D59'].value):.4f}) - 수식에는 안 쓰임")
    wb._external_links = []
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return log, notes


# ─────────────────────────── ARAP 차입대여금 현황
def read_ost(path: str, year: int) -> dict[str, list[tuple[str, float]]]:
    """Конечные остатки 연도 시트 -> {구역 제목: [(항목, RUB)]} (구역은 A열 번호 줄 기준)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[str(year)] if str(year) in wb.sheetnames else wb.worksheets[-1]
    out, cur = {}, None
    for r in range(1, ws.max_row + 1):
        a, b, c = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
        if a and re.match(r"^\d+\.$", str(a).strip()):
            cur = str(b or "").strip() + f" #{str(a).strip()}"
            out[cur] = []
        elif cur and b and str(b).strip() != "Итого" and isinstance(c, (int, float)):
            out[cur].append((str(b).strip(), float(c)))
    out["_date"] = ws["C1"].value
    return out


def _sec(o, word, num=None):
    for k, v in o.items():
        if k != "_date" and word.lower() in k.lower() and (num is None or k.endswith(f"#{num}.")):
            return v
    return []


def _sum(items, *words, exclude=()):
    return round(sum(v for n, v in items if all(w.lower() in n.lower() for w in words)
                     and not any(x.lower() in n.lower() for x in exclude)), 2)


# 5.AP(채무) 행: (Конечные остатки 구역 번호, 포함할 단어들, 제외 단어)
AP_MAP = {7: (3, ("Hyunday",), ()), 8: (3, ("Youngsan",), ()), 9: (4, ("ЭЛЕКТРЭЛ", "аренда"), ()),
          10: (4, ("ЭЛЕКТРЭЛ", "коммунал"), ()), 11: (7, ("Hyunday",), ()), 12: (7, ("Youngsan",), ()),
          13: (7, ("Ёнгсан КЛД",), ())}
# 6.차입금 거래처명(B열) 키워드 -> Конечные остатки에서 찾는 방법
LOAN_KEYS = {"byun": ("Пён",), "jeong": ("Чжон",), "elias": ("ЭЛИАС",)}


def fill_arap(tmpl: str, ost: str, base: dt.date, out: str):
    wb = openpyxl.load_workbook(tmpl)
    o = read_ost(ost, base.year)
    log, notes = [], []
    if isinstance(o.get("_date"), dt.datetime) and o["_date"].date() != base:
        notes.append(f"Конечные остатки 기준일 {o['_date']:%Y-%m-%d} (기대 {base})")
    sec = {int(k.rsplit("#", 1)[1].rstrip(".")): v for k, v in o.items() if k != "_date"}
    ar = wb["3. AR(채권)"]
    ar["B3"] = dt.datetime(base.year, base.month, base.day)
    log.append("3. AR(채권)!B3")
    # AP
    ap = wb["5. AP(채무)"]
    for r, (s, words, ex) in AP_MAP.items():
        v = _sum(sec.get(s, []), *words, exclude=ex)
        old = _num(ap.cell(r, 6).value)
        if abs(old - v) > 0.005:
            ap.cell(r, 6).value = v
            log.append(f"5. AP(채무)!F{r}")
    # 차입금: 장기(6) + 단기(8)
    ln = wb["6. 차입금(LOAN FROM)"]
    long_, short = sec.get(6, []), sec.get(8, [])
    hdc = [v for n, v in long_ if "Hyunday" in n and "2026" not in n]
    used = set()  # 표에 행이 있는 개인/기타 차입처
    for r in range(7, ln.max_row + 1):
        name = str(ln.cell(r, 2).value or "")
        if not name or ln.cell(r, 5).value in (None, ""):
            continue
        cur = ln.cell(r, 5).value
        if isinstance(cur, str):
            continue
        nl = name.lower()
        v = None
        if "hyundai" in nl:
            d = ln.cell(r, 3).value
            if isinstance(d, dt.datetime) and d.year == 2026:
                v = _sum(long_, "Hyunday", "2026")
            else:  # 2021년 본 차입, 2024년 추가 차입 순서
                older = sorted((ln.cell(x, 3).value, x) for x in range(7, ln.max_row + 1)
                               if "hyundai" in str(ln.cell(x, 2).value or "").lower()
                               and isinstance(ln.cell(x, 3).value, dt.datetime) and ln.cell(x, 3).value.year < 2026)
                idx = [x for _, x in older].index(r)
                v = hdc[idx] if idx < len(hdc) else None
        elif "youngsan glonet" in nl:
            # CNY 차입(2024-01) / USD 차입(2024-10): Конечные остатки 순서와 같음
            ys = [v for n, v in long_ if "Youngsan" in n]
            rows = [x for x in range(7, ln.max_row + 1) if "youngsan glonet" in str(ln.cell(x, 2).value or "").lower()]
            v = ys[rows.index(r)] if rows.index(r) < len(ys) else None
        elif "youngsan kld" in nl:
            v = _sum(long_, "Ёнгсан КЛД") + _sum(short, "Ёнгсан КЛД")
        else:
            for key, words in LOAN_KEYS.items():
                if key in nl:
                    v = _sum(short, *words)
                    used.add(words[0])
        if v is None:
            notes.append(f"6.차입금 {r}행 '{name}' 에 맞는 Конечные остатки 항목을 못 찾아 그대로 둠")
            continue
        if abs(_num(cur) - v) > 0.005:
            ln.cell(r, 5).value = v
            log.append(f"6. 차입금!E{r}")
    known = ("Ёнгсан КЛД",) + tuple(used)
    for n, v in short:
        if not any(k.lower() in n.lower() for k in known):
            notes.append(f"Конечные остатки 단기차입에 6.차입금 표에 없는 항목: {n} {v:,.2f} (행 추가 판단 필요)")
    for n, v in long_:
        if not any(k in n for k in ("Hyunday", "Youngsan", "Ёнгсан")):
            notes.append(f"Конечные остатки 장기차입에 6.차입금 표에 없는 항목: {n} {v:,.2f} (행 추가 판단 필요)")
    # AR 대조 (거래처별 합계)
    deb = sec.get(2, [])
    rows = [(str(ar.cell(r, 3).value or ""), _num(ar.cell(r, 6).value if not isinstance(ar.cell(r, 6).value, str)
             else float(ar.cell(r, 6).value) if re.match(r"^-?[\d.]+$", ar.cell(r, 6).value or "") else 0))
            for r in range(7, ar.max_row + 1) if ar.cell(r, 3).value]
    alias = {"АВТОТОР": "avtotor", "ЗПД": "pcf"}
    for word, key in alias.items():
        a = _sum(deb, word)
        b = round(sum(v for n, v in rows if key in n.lower()), 2)
        if abs(a - b) > 0.005:
            notes.append(f"3.AR {key.upper()} 합계 {b:,.2f} ≠ Конечные остатки {a:,.2f} - 송장 행 추가/정리 필요")
    for n, v in deb:
        if not any(w.lower() in n.lower() for w in alias):
            notes.append(f"Конечные остатки 채권에 새 거래처: {n} {v:,.2f} (3.AR 반영 판단 필요)")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return log, notes
