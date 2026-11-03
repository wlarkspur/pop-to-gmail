"""HAS 가마감 (HDC 보고 엑셀): 지난달 가마감 파일을 양식으로, 이번 달 숫자만 바꿔 넣는다.

매핑 근거: 보고서 자동화 등록양식 'HAS매핑' 탭 (9월 원자료로 9월 파일을 재현해 확인)
- 15-3 실적(RUB) 당월 열  <- monthly '9-1 상세(RUB)' 당월 열 (행 오프셋 표 SEG)
- 7-1/7-2/7-3 당기(J열)  <- HAS 재무제표(RUB) Balance Sheets / Income Statement / G&A (같은 주소, 7-2만 일부 행 재배치)
- 6-1 J8 월말환율, H5 결산일 / 15-4 3행 당월 월평균환율 / 1번 C5 보고월   <- 하나은행 (fx_hana.py)
- 8번 자금수지: 모든 달을 '이번 달 월말환율'로 다시 환산(양식 관례) + 당월 열은 가마감용 Daily Report에서
- 11번 현금시재표: 계좌별 잔액 (은행 자료가 없으면 Daily Report 월말 잔액)
- 9/12/13/14번: Конечные остатки (별도 함수)
수식 셀은 건드리지 않고, 엑셀이 열 때 다시 계산(fullCalcOnLoad).
"""
from __future__ import annotations

import calendar
import re
import zipfile
from decimal import ROUND_HALF_UP, Decimal

import openpyxl

from daily_report import _put
from xlsx_patch import _sheet_path

SEG = [(8, 11, 1), (13, 24, 1), (27, 34, 1), (37, 40, -1), (42, 83, -2), (86, 89, -2), (91, 94, -2),
       (95, 100, -1), (101, 101, -8), (102, 102, -2), (105, 107, -2)]
MONTHLY_TO_153 = {r: r + o for a, b, o in SEG for r in range(a, b + 1)}   # monthly 행 -> 15-3 행
IS_TO_72 = {49: 56, 50: 49, 51: 50, 52: 51, 53: 52, 54: 53, 55: 54, 56: 55}  # 7-2 행 -> 원본 행 (그 외 같은 행)
S153, S154 = "15-3. 26년 실적(RUB)", "15-4. 26년 실적(USD)"
S61, S71, S72, S73 = "6-1. 재무제표(USD)", "7-1. 재무제표(RUB)", "7-2. 손익계산서(RUB)", "7-3. 판매관리비(RUB)"
S8, S11, S1 = "8. 법인 자금수지현황", "11. 현금시재표", "1. 손익개황"


def col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Book:
    """원본 xlsx의 시트 XML만 고쳐서 다시 쓰는 도구 (서식, 그림, 다른 시트는 바이트 그대로)."""

    def __init__(self, path: str):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.xml: dict[str, str] = {}
        self.log: list[tuple[str, str, object]] = []
        self.formula = openpyxl.load_workbook(path)  # 수식 여부 확인용 (읽기만)

    def _x(self, sheet):
        p = _sheet_path(self.z, sheet)
        if p not in self.xml:
            self.xml[p] = self.z.read(p).decode("utf-8")
        return p

    def is_formula(self, sheet, ref) -> bool:
        v = self.formula[sheet][ref].value
        return isinstance(v, str) and v.startswith("=")

    def value(self, sheet, ref):
        return self.formula[sheet][ref].value

    def set(self, sheet, ref, value):
        if self.is_formula(sheet, ref):
            raise ValueError(f"{sheet}!{ref}는 수식 칸이라 값으로 덮지 않음")
        p = self._x(sheet)
        self.xml[p] = _put(self.xml[p], ref, value)
        self.log.append((sheet, ref, value))

    def save(self, dst: str):
        wb = self.z.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        repl = {**self.xml, "xl/workbook.xml": wb}
        # 수식 캐시는 그대로 두면 엑셀 외 프로그램에서 옛 값이 보이므로, 바꾼 시트의 캐시는 비움
        for p in self.xml:
            repl[p] = re.sub(r"(</f>|<f [^>]*/>)<v>[^<]*</v>", r"\1", repl[p])
        with zipfile.ZipFile(dst, "w") as out:
            for item in self.z.infolist():
                data = repl[item.filename].encode("utf-8") if item.filename in repl else self.z.read(item.filename)
                out.writestr(item, data, compress_type=item.compress_type)


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def fill_153(b: Book, monthly: str, m: int):
    ws = openpyxl.load_workbook(monthly, data_only=True).worksheets[0]
    if str(ws.cell(4, 6 + m).value or "").strip()[:3].lower() != calendar.month_abbr[m].lower():
        raise RuntimeError(f"monthly {col(6 + m)}4 머리글이 {calendar.month_abbr[m]}가 아님")
    tc = col(5 + m)
    for src, dst in MONTHLY_TO_153.items():
        ref = f"{tc}{dst}"
        if b.is_formula(S153, ref):
            continue
        b.set(S153, ref, _num(ws.cell(src, 6 + m).value))


def fill_fs(b: Book, fs: str):
    wb = openpyxl.load_workbook(fs, data_only=True)
    plans = [(S71, "Balance Sheets", range(13, 163), {}),
             (S72, "Income Statement", list(range(13, 57)) + [76], IS_TO_72),
             (S73, "general&adminstrative expenses", range(10, 47), {})]
    for tname, sname, rows, rmap in plans:
        s = wb[sname]
        for r in rows:
            ref = f"J{r}"
            if b.is_formula(tname, ref):
                continue
            v = _num(s.cell(rmap.get(r, r), 10).value)
            if v is None and b.value(tname, ref) is None:
                continue
            b.set(tname, ref, v)


def fill_rates(b: Book, y: int, m: int, month_end: float, month_avg: float):
    last = calendar.monthrange(y, m)[1]
    b.set(S61, "J8", month_end)
    b.set(S61, "H5", f"{y}. {m:02d}. {last:02d}")
    b.set(S154, f"{col(5 + m)}3", month_avg)
    b.set(S1, "C5", str(m) if isinstance(b.value(S1, "C5"), str) else m)


def _r2(x) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


INCOME_ROW = [("deposit interest", 11), ("hyundai", 15), ("elias", 17), ("youngsan kld", 16),
              ("interest-free loan", 16)]  # 나머지 입금은 18행(기타)


def fill_cashflow(b: Book, m: int, old_rate: float, new_rate: float, dr: dict, bank_fee_cat: str):
    """8번: 지난 달들을 새 월말환율로 다시 환산(RUB = 값/옛 환율) + 당월 열 채움."""
    ws = b.formula[S8]
    for c in range(4, 4 + m - 1):  # D(1월) ~ 전월
        for r in range(8, 31):
            ref = f"{col(c)}{r}"
            v = ws[ref].value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0:
                b.set(S8, ref, _r2(v / old_rate) * new_rate)
    tc = col(3 + m)
    rows = {r: 0.0 for r in (9, 11, 15, 16, 17, 18, 24, 25)}
    for i in dr["items"]:
        if i["kind"] == "in":
            d = i["desc"].lower()
            r = next((row for key, row in INCOME_ROW if key in d), 18)
            rows[r] += i["amount"]
    fee = sum(o["amount"] for o, c in zip([x for x in dr["items"] if x["kind"] == "out"], dr["cats"]) if c == bank_fee_cat)
    rows[25] = fee
    rows[24] = dr["out"] - fee
    for r, rub in rows.items():
        b.set(S8, f"{tc}{r}", _r2(rub) * new_rate)
    return rows


def fill_cash(b: Book, balances: dict):
    """11번: balances = {'cash','woori','entb','spb'} RUB."""
    b.set(S11, "F8", balances["cash"])
    b.set(S11, "D9", balances["woori"])
    b.set(S11, "D10", balances["entb"])
    b.set(S11, "D11", balances["spb"])


def read_daily_sheet(path: str, sheet: str) -> dict:
    """가마감용 Daily Report 시트(B/D 입금, F/H/I 출금)를 8번용 형태로."""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet.strip() if sheet in [] else sheet]
    items, cats = [], []
    for r in range(4, 60):
        b, d = ws.cell(r, 2).value, ws.cell(r, 4).value
        if isinstance(b, (int, float)) and b:
            items.append({"kind": "in", "amount": b, "desc": str(d or "")})
        f, h, i = ws.cell(r, 6).value, ws.cell(r, 8).value, ws.cell(r, 9).value
        if isinstance(f, (int, float)) and f and isinstance(h, str):
            items.append({"kind": "out", "amount": f, "desc": h})
            cats.append(i)
    out = round(sum(x["amount"] for x in items if x["kind"] == "out"), 2)
    return {"items": items, "cats": cats, "out": out, "in": round(sum(x["amount"] for x in items if x["kind"] == "in"), 2)}
