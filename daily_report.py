"""가마감용 Daily Report: Irina 월말 Cash Daily Report(일별 시트)를 모아 가마감용 파일에 그달 시트를 추가.

- 원자료: 'Daily Report YYYY.MM (HAS).xlsx' (일자별 시트 '01'~'30', Incoming/Outgoing 구간)
- 결과: '@가마감용 DailyReport_YYYY.MM.xlsx' = 지난달 파일 + 맨 앞에 'YY DD MM' 시트 (직전 달 시트를 복사해 양식, 수식 유지)
  Incoming: B(금액) D(내용) / Outgoing: F(금액) H(내용) I(분류) / K47~ 입금 내용별 합계 목록
- 출금 분류는 지난 시트들의 같은 내용 분류를 따르고, 처음 보는 내용은 '확인 필요'로 돌려줌
- 검증: 원자료 일별 시트 Total 합 = 옮긴 합계, 시작 잔액 + 입금 - 출금 = 월말 잔액
"""
from __future__ import annotations

import calendar
import re
import zipfile
from html import escape, unescape

import openpyxl

# 법인별 가마감용 파일 양식 (요약 수식 범위와 같게)
LAYOUT = {
    "HAS": {"in": range(4, 11), "out": range(4, 55), "in_list": range(47, 54), "cat_list": range(5, 33),
            "name": lambda y, m, d: f"{y % 100:02d} {d:02d} {m:02d}"},          # 예: '26 31 10'
    "YSKD": {"in": range(5, 58), "out": range(5, 59), "in_list": range(42, 51), "cat_list": range(6, 34),
             "name": lambda y, m, d: f"{y % 100:02d} {m} {d:02d}"},             # 예: '26 10 31'
}


# ---------- 원자료 읽기 ----------
def parse_source(path: str) -> dict:
    """일별 시트에서 입출금 행을 모두 읽음. 머리글 위치는 시트에서 찾음 (HAS, YSKD 양식 차이 대응)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    items, sheet_tot, first, last = [], {"in": 0.0, "out": 0.0}, None, None
    for ws in wb:
        hdr = next(r for r in range(1, 15) if any(str(ws.cell(r, c).value or "").strip() == "Cash RUB Account" for c in range(3, 8)))
        c0 = next(c for c in range(3, 8) if str(ws.cell(hdr, c).value or "").strip() == "Cash RUB Account")
        ctot = next(c for c in range(c0, c0 + 10) if str(ws.cell(hdr, c).value or "").strip() == "Total")
        cdesc = next(c for c in range(ctot, ctot + 6) for r in range(hdr, hdr + 12)
                     if str(ws.cell(r, c).value or "").startswith("Contents"))
        bal = [r for r in range(hdr + 1, hdr + 5) if isinstance(ws.cell(r, ctot).value, (int, float))]
        if first is None:
            first = ws.cell(bal[-2], ctot).value  # 첫 시트의 '전일' 잔액 = 월초 잔액
        last = ws.cell(bal[-1], ctot).value       # 마지막 시트의 '오늘' 잔액 = 월말 잔액
        accounts = {str(ws.cell(hdr, c).value).strip(): ws.cell(bal[-1], c).value or 0 for c in range(c0, ctot)}
        sec = None
        for r in range(hdr + 1, ws.max_row + 1):
            b = str(ws.cell(r, 2).value or "").strip().lower()
            if b in ("incoming", "outgoing"):
                sec = "in" if b == "incoming" else "out"
            elif b == "total" and sec:
                sheet_tot[sec] += ws.cell(r, ctot).value or 0
                sec = None
                continue
            elif b.startswith("business"):
                break
            if not sec:
                continue
            amt = sum(v for c in range(c0, ctot) if isinstance(v := ws.cell(r, c).value, (int, float)))
            desc = str(ws.cell(r, cdesc).value or "").strip()
            if abs(amt) > 1e-9:
                items.append({"day": ws.title, "kind": sec, "amount": round(amt, 2), "desc": desc})
    got = {k: round(sum(i["amount"] for i in items if i["kind"] == k), 2) for k in ("in", "out")}
    checks = {
        "일별 Total 합과 일치": all(abs(got[k] - sheet_tot[k]) < 0.01 for k in got),
        "잔액 흐름 일치": first is not None and abs(first + got["in"] - got["out"] - last) < 0.01,
    }
    return {"items": items, "in": got["in"], "out": got["out"], "open": first, "close": last, "checks": checks,
            "accounts": accounts}  # 계좌별 월말 잔액 (마지막 시트 머리글 이름 그대로)


# ---------- 분류 ----------
def _norm(s: str) -> str:
    s = str(s).lower().replace("\\~", "~")
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[0-9'’./~,#№-]+", " ", s)
    s = re.sub(r"\b(jan|feb|mar|apr|may|jun|june|jul|july|aug|august|sep|september|oct|october|nov|november|dec|december|for)\b", " ", s)
    return " ".join(s.split())


def history(path: str, ent: str = "HAS") -> dict:
    """기존 시트들의 출금 내용(H) -> 분류(I). 분류 칸이 '=K7' 같은 수식이면 그 칸 글자로."""
    wb = openpyxl.load_workbook(path)
    hist = {}
    for ws in wb:
        for r in LAYOUT[ent]["out"]:
            h, i = ws.cell(r, 8).value, ws.cell(r, 9).value
            if isinstance(i, str) and (m := re.fullmatch(r"=\$?([A-Z]+)\$?(\d+)", i.strip())):
                i = ws[f"{m.group(1)}{m.group(2)}"].value
            if h and i:
                d = hist.setdefault(_norm(h), {})
                d[i] = d.get(i, 0) + 1
    return hist


def classify(desc: str, hist: dict, overrides: dict | None = None) -> tuple[str | None, str]:
    for k, v in (overrides or {}).items():
        if k.lower() in desc.lower():
            return v, "지정"
    k = _norm(desc)
    if k in hist:
        return max(hist[k], key=hist[k].get), "기존"
    return None, "확인 필요"


# ---------- xlsx 쓰기 (원본 양식 유지, 시트 XML만 고침) ----------
def _put(xml: str, ref: str, value) -> str:
    """셀 값 교체 (서식 유지). 행에 셀이 없으면 열 순서에 맞게 넣음. value None = 빈 칸."""
    col, row = re.fullmatch(r"([A-Z]+)(\d+)", ref).groups()
    rm = re.search(r'<row r="%s"[^>]*?(?:/>|>(.*?)</row>)' % row, xml, re.S)
    if not rm:
        raise ValueError(f"행 없음: {row}")
    row_xml = rm.group(0)
    m = re.search(r'<c r="%s"((?: [^>]*?)?)(/>|>(.*?)</c>)' % ref, row_xml, re.S)
    style = ""
    if m:
        s = re.search(r' s="\d+"', m.group(1))
        style = s.group(0) if s else ""
    if value is None:
        new = f'<c r="{ref}"{style}/>'
    elif isinstance(value, str):
        new = f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    else:
        new = f'<c r="{ref}"{style}><v>{value}</v></c>'
    if m:
        new_row = row_xml[:m.start()] + new + row_xml[m.end():]
    else:
        def ci(c):
            n = 0
            for ch in c:
                n = n * 26 + ord(ch) - 64
            return n
        cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', row_xml))
        after = next((c for c in cells if ci(c.group(1)) > ci(col)), None)
        if after:
            new_row = row_xml[:after.start()] + new + row_xml[after.start():]
        elif row_xml.endswith("/>"):
            new_row = row_xml[:-2] + ">" + new + "</row>"
        else:
            new_row = row_xml[:-6] + new + "</row>"
    return xml[:rm.start()] + new_row + xml[rm.end():]


def sheet_name(year: int, month: int, ent: str = "HAS") -> str:
    return LAYOUT[ent]["name"](year, month, calendar.monthrange(year, month)[1])


def build(tmpl: str, dst: str, year: int, month: int, src: dict, cats: list[str | None], ent: str = "HAS") -> dict:
    """tmpl(지난달 가마감용 파일)에 그달 시트를 맨 앞에 추가해 dst로 저장."""
    L = LAYOUT[ent]
    IN_ROWS, OUT_ROWS, IN_LIST = L["in"], L["out"], L["in_list"]
    name = sheet_name(year, month, ent)
    ins = [i for i in src["items"] if i["kind"] == "in"]
    outs = [i for i in src["items"] if i["kind"] == "out"]
    in_kinds = list(dict.fromkeys(i["desc"] for i in ins))
    if len(ins) > len(IN_ROWS) or len(outs) > len(OUT_ROWS) or len(in_kinds) > len(IN_LIST):
        raise RuntimeError(f"행이 모자람: 입금 {len(ins)}/{len(IN_ROWS)}, 출금 {len(outs)}/{len(OUT_ROWS)}, 입금 종류 {len(in_kinds)}/{len(IN_LIST)}")
    with zipfile.ZipFile(tmpl) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        sheets = re.findall(r'<sheet name="([^"]+)" sheetId="(\d+)" r:id="(rId\d+)"/>', wb)
        if any(unescape(n).strip() == name for n, _, _ in sheets):
            raise RuntimeError(f"이미 '{name}' 시트가 있음")
        _, _, trid = sheets[0]  # 맨 앞 = 직전 달
        path_of = lambda rid: "xl/" + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1).lstrip("/").replace("xl/", "")
        xml = z.read(path_of(trid)).decode("utf-8")

        for k, r in enumerate(IN_ROWS):
            i = ins[k] if k < len(ins) else None
            xml = _put(xml, f"B{r}", i["amount"] if i else None)
            xml = _put(xml, f"D{r}", i["desc"] if i else None)
        for k, r in enumerate(OUT_ROWS):
            o = outs[k] if k < len(outs) else None
            xml = _put(xml, f"F{r}", o["amount"] if o else None)
            xml = _put(xml, f"H{r}", o["desc"] if o else None)
            xml = _put(xml, f"I{r}", (cats[k] or "") if o else None)
        for k, r in enumerate(IN_LIST):
            xml = _put(xml, f"K{r}", in_kinds[k] if k < len(in_kinds) else None)
        xml = re.sub(r"(</f>|<f [^>]*/>)<v>[^<]*</v>", r"\1", xml)  # 수식 캐시 비움 (엑셀이 열 때 계산)
        first = f"F{OUT_ROWS[0]}"
        xml = re.sub(r'<selection [^>]*/>', f'<selection activeCell="{first}" sqref="{first}"/>', xml, count=1)
        if ' tabSelected="1"' not in xml:
            xml = re.sub(r"<sheetView ", '<sheetView tabSelected="1" ', xml, count=1)
        xml = re.sub(r'(<pageSetup[^>]*?) r:id="rId\d+"', r"\1", xml)
        xml = re.sub(r' xr:uid="\{[0-9A-F-]+\}"', "", xml, count=1)

        nums = [int(n) for n in re.findall(r"worksheets/sheet(\d+)\.xml", "\n".join(z.namelist()))]
        new_path = f"xl/worksheets/sheet{max(nums) + 1}.xml"
        rid_n = max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1
        sid = max(int(s) for _, s, _ in sheets) + 1
        rels2 = rels.replace("</Relationships>", f'<Relationship Id="rId{rid_n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{new_path.split("/")[-1]}"/></Relationships>')
        rels2 = re.sub(r'<Relationship Id="rId\d+" Type="[^"]*/calcChain" Target="calcChain.xml"/>', "", rels2)
        wb2 = wb.replace("<sheets>", f'<sheets><sheet name="{escape(name)}" sheetId="{sid}" r:id="rId{rid_n}"/>', 1)
        wb2 = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{int(m.group(1)) + 1}"', wb2)  # 시트 순서 한 칸씩 밀림
        wb2 = re.sub(r'(<workbookView[^>]*?) activeTab="\d+"', r"\1", wb2)
        wb2 = re.sub(r'(<workbookView[^>]*?) firstSheet="\d+"', r"\1", wb2)
        if "fullCalcOnLoad" not in wb2:
            wb2 = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb2, count=1)
        repl = {"xl/workbook.xml": wb2, "xl/_rels/workbook.xml.rels": rels2}
        for _, _, rid in sheets:  # 다른 시트 선택 표시 끔
            p = path_of(rid)
            x = z.read(p).decode("utf-8")
            if ' tabSelected="1"' in x:
                repl[p] = x.replace(' tabSelected="1"', "", 1)
        ct = z.read("[Content_Types].xml").decode("utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
        repl["[Content_Types].xml"] = ct.replace("</Types>", f'<Override PartName="/{new_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        if "docProps/app.xml" in z.namelist():
            app = z.read("docProps/app.xml").decode("utf-8")
            n = len(sheets) + 1
            app = re.sub(r"(<vt:lpstr>(?:워크시트|Worksheets)</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+", rf"\g<1>{n}", app)
            app = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)', lambda m: f"{m.group(1)}{int(m.group(2)) + 1}", app)
            app = re.sub(r'(<TitlesOfParts><vt:vector size="\d+" baseType="lpstr">)', rf"\1<vt:lpstr>{escape(name)}</vt:lpstr>", app)
            repl["docProps/app.xml"] = app
        with zipfile.ZipFile(dst, "w") as out:
            for item in z.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue
                data = repl[item.filename].encode("utf-8") if item.filename in repl else z.read(item.filename)
                out.writestr(item, data, compress_type=item.compress_type)
            out.writestr(zipfile.ZipInfo(new_path, date_time=z.getinfo(path_of(trid)).date_time), xml.encode("utf-8"),
                         compress_type=zipfile.ZIP_DEFLATED)
    return {"시트": name, "입금": len(ins), "출금": len(outs), "입금 종류": in_kinds}


def verify(dst: str, year: int, month: int, src: dict, ent: str = "HAS") -> dict:
    """만든 파일을 다시 읽어 원자료와 대조 (값 셀 기준)."""
    L = LAYOUT[ent]
    IN_ROWS, OUT_ROWS = L["in"], L["out"]
    ws = openpyxl.load_workbook(dst)[sheet_name(year, month, ent)]
    b = round(sum(ws.cell(r, 2).value or 0 for r in IN_ROWS), 2)
    f = round(sum(ws.cell(r, 6).value or 0 for r in OUT_ROWS), 2)
    empty_cat = [r for r in OUT_ROWS if ws.cell(r, 6).value and not ws.cell(r, 9).value]
    cats = {ws.cell(r, 11).value for r in L["cat_list"]}
    unknown_cat = [r for r in OUT_ROWS if ws.cell(r, 9).value and ws.cell(r, 9).value not in cats]
    return {"입금 합계 일치": abs(b - src["in"]) < 0.01, "출금 합계 일치": abs(f - src["out"]) < 0.01,
            "분류 빈 칸 행": empty_cat, "목록에 없는 분류 행": unknown_cat, "입금": b, "출금": f}
