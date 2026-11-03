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

    def set(self, sheet, ref, value, like: str | None = None):
        """값 쓰기. '='로 시작하는 글자는 수식으로 씀 (빈 칸에만). like = 서식을 빌려올 칸 (새로 채우는 행)."""
        if self.is_formula(sheet, ref) and not (isinstance(value, str) and value.startswith("=")):
            raise ValueError(f"{sheet}!{ref}는 수식 칸이라 값으로 덮지 않음")
        p = self._x(sheet)
        x = _put(self.xml[p], ref, value if not (isinstance(value, str) and value.startswith("=")) else None)
        style = None
        if like:
            m = re.search(r'<c r="%s"[^>]*? s="(\d+)"' % like, x)
            style = m.group(1) if m else None
        m = re.search(r'<c r="%s"((?: [^>]*?)?)(/>|>(.*?)</c>)' % ref, x, re.S)
        cell = m.group(0)
        if style:
            cell = re.sub(r' s="\d+"', "", cell, count=1).replace(f'<c r="{ref}"', f'<c r="{ref}" s="{style}"', 1)
        if isinstance(value, str) and value.startswith("="):
            from xml.sax.saxutils import escape as esc
            sm = re.search(r' s="\d+"', cell)
            cell = f'<c r="{ref}"{sm.group(0) if sm else ""}><f>{esc(value[1:])}</f></c>'
        self.xml[p] = x[:m.start()] + cell + x[m.end():]
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
    spb = balances["spb"]  # 양식은 '=계좌+법인카드' 수식 칸이라 같은 방식으로 씀
    b.set(S11, "D11", "=" + "+".join(f"{x:.2f}" for x in (spb if isinstance(spb, (list, tuple)) else [spb])))


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


S9, S12, S13, S14 = "9. 법인 차입현황", "12. 외상매출금명세서", "13. 외상매입금명세서", "14. 재고자산명세서"


def read_ostatki(path: str, y: int) -> dict:
    """Конечные остатки 그해 시트를 구역 번호별 {제목, 줄[(이름, 루블, 외화, 통화)]}로."""
    ws = openpyxl.load_workbook(path, data_only=True)[str(y)]
    secs, cur = {}, None
    for r in range(1, ws.max_row + 1):
        a, bv = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(a, str) and re.fullmatch(r"\d+\.", a.strip()):
            cur = {"title": str(bv or "").strip(), "rows": [], "row": r}
            secs[int(a.strip()[:-1])] = cur
            continue
        if cur and bv:
            cur["rows"].append((str(bv).strip(), ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value, r))
    secs["_ws"] = ws
    return secs


def fill_ostatki(b: Book, ost: dict, y: int, m: int) -> list[str]:
    """9/12/13/14번. 양식에 없는 새 거래처/대출은 빈 다음 행에 추가하고 '확인 필요'로 돌려줌."""
    notes = []
    ws = ost["_ws"]
    by_title = [(v["title"].lower(), v) for k, v in ost.items() if k != "_ws"]  # 같은 제목 구역이 여럿 있음

    def sec(prefix, nth=0):
        hits = [v for t, v in by_title if t.startswith(prefix)]
        hits.sort(key=lambda v: v["row"])
        return hits[nth] if len(hits) > nth else None

    # 14번 재고: 기타자재 기초/입고/출고, 폴리프로필렌 금액
    mat = sec("прочие материалы")
    vals = {n.strip().split()[0]: v for n, v, *_ in mat["rows"]}
    b.set(S14, "H12", vals["начальный"])
    b.set(S14, "N12", vals["приход"])
    b.set(S14, "T12", vals["расход"])
    b.set(S14, "D8", f"{y}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}")
    pp = {}
    for r in range(mat["row"], mat["row"] + 6):
        g = str(ws.cell(r, 7).value or "")
        if "M7392T" in g:
            pp["T"] = (ws.cell(r, 9).value, ws.cell(r + 1, 9).value)
        if "M7392USK" in g:
            pp["USK"] = (ws.cell(r, 9).value, ws.cell(r + 1, 9).value)
    b.set(S14, "H10", pp["T"][0])
    b.set(S14, "H11", pp["USK"][0])
    if b.value(S14, "X10") not in (pp["T"][1], str(pp["T"][1])):
        b.set(S14, "X10", pp["T"][1])
    x11 = b.value(S14, "X11")
    if float(x11 or 0) != float(pp["USK"][1] or 0):
        notes.append(f"14번 X11 수량 {x11} ≠ Остатки {pp['USK'][1]} (양식 값 유지, 확인 필요)")

    # 12번 매출채권: 기존 행(AVT 임대) 유지, 새 채권은 빈 행에 추가
    ar = sec("дебиторская")
    known = sum(float(b.value(S12, f"F{r}") or 0) for r in range(7, 35) if isinstance(b.value(S12, f"F{r}"), (int, float, str)) and str(b.value(S12, f"F{r}")).replace(".", "", 1).isdigit())
    first_free = next(r for r in range(7, 35) if b.value(S12, f"F{r}") in (None, ""))
    for name, rub, fx, cur, _ in ar["rows"]:
        if name.lower().startswith("итого") or not rub:
            continue
        if any(str(b.value(S12, f"C{r}") or "").strip() == name for r in range(7, 35)):
            continue  # 이미 넣은 채권
        if "АВТОТОР" in name:
            if abs(rub - known) > 0.01:
                notes.append(f"12번 AVT 임대 채권 Остатки {rub:,.2f} ≠ 양식 {known:,.2f} (확인 필요)")
            continue
        r = first_free
        first_free += 1
        b.set(S12, f"C{r}", name, like="C11")
        b.set(S12, f"D{r}", "RUB", like="D11")
        b.set(S12, f"E{r}", f"=$L$1*F{r}", like="E11")
        b.set(S12, f"F{r}", rub, like="F11")
        b.set(S12, f"L{r}", f"{fx:,.2f} {cur}" if fx else "", like="L11")
        notes.append(f"12번 {r}행에 새 매출채권 추가: {name} {rub:,.2f} RUB - 발생일/비고 확인 필요")

    # 13번 매입채무 (HDC, YSG): 외화는 그대로, 장부금액(RUB)만 갱신
    ap = sec("кредиторская задолженность")
    for name, rub, fx, cur, _ in ap["rows"]:
        key = "HDC" if "hyunday" in name.lower() else "YSG" if "youngsan" in name.lower() else None
        if not key:
            continue
        r = 7 if key == "HDC" else 8
        b.set(S13, f"G{r}", rub)
        if abs(float(b.value(S13, f"F{r}") or 0) - (fx or 0)) > 0.005:
            b.set(S13, f"F{r}", fx)
            notes.append(f"13번 {key} 외화금액 변경 {fx:,.2f} (확인 필요)")

    # 9번 단기차입: 대출처별 합계로 대조, 늘어난 만큼만 새 행 추가 (금액 쪼개기는 양식 방식 유지)
    st = sec("задолженность по краткосрочным займам")
    lend = lambda n: "KLD" if ("КЛД" in n or n.upper() == "KLD") else "ELIAS" if ("ЭЛИАС" in n or "ELIAS" in n.upper()) \
        else "BYUN" if ("Пён" in n or "변용필" in n) else "JSM" if "Чжон" in n else n
    tmpl, free = {}, None
    for r in range(32, 45):
        v, k = b.value(S9, f"E{r}"), str(b.value(S9, f"K{r}") or "")
        parts = [float(x) for x in re.findall(r"[\d.]+", v)] if isinstance(v, str) and v.startswith("=") else \
            [float(v)] if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().replace(".", "", 1).isdigit()) else []
        if parts:
            tmpl.setdefault(lend(k), []).extend(parts)
        elif free is None:
            free = r
    srcs = {}
    for name, rub, *_ in st["rows"]:
        if not name.lower().startswith("итого") and rub:
            srcs.setdefault(lend(name), []).append((name, float(rub)))
    for who, loans in sorted(srcs.items(), key=lambda kv: kv[0] not in tmpl):  # 양식에 있던 대출처 먼저
        have = list(tmpl.get(who, []))
        if abs(sum(x for _, x in loans) - sum(have)) < 0.01:
            continue
        for name, rub in loans:
            if rub in have:
                have.remove(rub)
                continue
            date = re.search(r"(\d\d)\.(\d\d)\.(\d{4})", name)
            b.set(S9, f"E{free}", rub, like="E35")
            b.set(S9, f"F{free}", 0.01 if who == "KLD" else 0 if "беспроцент" in name else None, like="F35")
            b.set(S9, f"G{free}", f"{date.group(3)}.{date.group(2)}.{date.group(1)}" if date else None, like="G35")
            b.set(S9, f"K{free}", {"KLD": "KLD", "ELIAS": "ELIAS RUS", "JSM": "정석민 과장", "BYUN": "변용필 상무"}.get(who, name), like="K35")
            b.set(S9, f"M{free}", "개인 Loan, 무이자" if "беспроцент" in name else "", like="M36")
            notes.append(f"9번 {free}행에 새 단기차입 추가: {name} {rub:,.0f} RUB - 만기 확인 필요")
            free += 1
    return notes


def apply_out_overrides(b: Book, cells: dict) -> list[str]:
    """사용자가 정한 결과 칸 값 (gamagam_overrides.json의 out). {"시트!칸": 값 또는 {"value","type":"date","like"}}"""
    import datetime as _dt
    done = []
    for ref, v in cells.items():
        if ref.startswith("_"):
            continue
        sheet, cell = ref.split("!")
        like = None
        if isinstance(v, dict):
            like = v.get("like")
            if v.get("type") == "date":
                d = _dt.date.fromisoformat(v["value"])
                v = (d - _dt.date(1899, 12, 30)).days
            else:
                v = v.get("value")
        b.set(sheet, cell, v, like=like)
        done.append(ref)
    return done


def spb_balances(paths: list[str]) -> list[float]:
    """SPB 은행 명세서(xlsx)들의 'Исходящий остаток'(기말 잔액). 큰 것(계좌) 먼저."""
    vals = []
    for p in paths:
        ws = openpyxl.load_workbook(p, data_only=True).worksheets[0]
        got = None
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "исходящий остаток" in c.value.lower():
                    right = [x.value for x in row if x.column > c.column and isinstance(x.value, (int, float))]
                    below = ws.cell(c.row + 1, c.column).value
                    got = right[0] if right else below if isinstance(below, (int, float)) else None
                    if isinstance(got, str):
                        got = float(got.replace(" ", "").replace(",", "."))
            if got is not None:
                break
        if got is None:
            raise RuntimeError(f"{p}: 'Исходящий остаток'을 못 찾음")
        vals.append(round(float(got), 2))
    return sorted(vals, reverse=True)
