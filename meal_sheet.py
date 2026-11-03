"""식대 정산 xlsx에 새 달 시트를 만든다 ('YY M월'). 직전 달 시트를 그대로 복사해 날짜와 750만 바꾼다.

규칙 (사용자 확인)
- 식대 = 근무일(러시아 생산 달력, 주말/공휴일/대체휴일 제외)마다 1인 750 RUB
- 1차(25일 지급분) / 2차(다음 달 10일 지급분)로 나눔. 나누는 날은 13~16일 사이에서
  두 구간 합계가 50으로 끝나지 않게 (= 구간 근무일 수가 짝수) 고르고, 안 되면 1차를 짝수로, 그다음 15일에 가까운 날
- 사람 목록, 칸 서식, 수식(T열 SUM)은 직전 시트 그대로. 직전 시트의 메모(' => ...') 내용은 비움
- 휴가 등 예외는 알 수 없으므로 전원 750으로 채우고 메일로 확인 요청
- 다른 시트와 파일 내용은 그대로 (openpyxl, LibreOffice로 다시 저장하지 않음)
"""
from __future__ import annotations

import datetime as dt
import re
import zipfile
from html import unescape
from xml.sax.saxutils import escape

import ru_calendar

COLS = [chr(c) for c in range(ord("E"), ord("S") + 1)]  # 날짜 칸 E~S (15칸)
EXCEL_EPOCH = dt.date(1899, 12, 30)


def workdays(y: int, m: int) -> list[dt.date]:
    d, out = dt.date(y, m, 1), []
    while d.month == m:
        if ru_calendar.is_workday(d):
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def split(days: list[dt.date]) -> int:
    """1차 구간 근무일 수."""
    cands = [k for k in range(1, len(days)) if 13 <= days[k - 1].day <= 16] or \
            [min(range(1, len(days)), key=lambda k: abs(days[k - 1].day - 15))]

    def score(k):
        a, b = k, len(days) - k
        return (not (a % 2 == 0 and b % 2 == 0), a % 2 != 0, abs(days[k - 1].day - 15), -k)
    return min(cands, key=score)


def _sheet_name_date(name: str):
    m = re.fullmatch(r"(\d\d) (\d{1,2})월", name)
    return (2000 + int(m.group(1)), int(m.group(2))) if m else None


def _sst(z) -> list[str]:
    xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    return [unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))) for si in re.findall(r"<si>(.*?)</si>", xml, re.S)]


def _cell_text(cell_xml: str, sst) -> str | None:
    v = re.search(r"<v>(.*?)</v>", cell_xml)
    if 't="s"' in cell_xml and v:
        return sst[int(v.group(1))]
    if 't="inlineStr"' in cell_xml:
        return unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", cell_xml, re.S)))
    return None


def _set(row_xml: str, ref: str, value) -> str:
    """행 XML 안의 셀 값을 바꿈 (서식 유지). value None = 빈 칸."""
    pat = re.compile(r'<c r="%s"((?: [^>]*?)?)(/>|>(.*?)</c>)' % ref, re.S)
    m = pat.search(row_xml)
    if not m:
        raise ValueError(f"셀 없음: {ref}")
    style = re.search(r' s="\d+"', m.group(1))
    style = style.group(0) if style else ""
    if value is None:
        new = f'<c r="{ref}"{style}/>'
    elif isinstance(value, str):
        new = f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    else:
        new = f'<c r="{ref}"{style}><v>{value}</v></c>'
    return row_xml[:m.start()] + new + row_xml[m.end():]


def _ensure_active(z, src, dst, wb, sheets, name) -> dict:
    """이미 있는 시트를 열 때 보이는 탭으로 (아니면 아무것도 안 함)."""
    idx = [unescape(n) for n, _, _ in sheets].index(name)
    cur = re.search(r'<workbookView[^>]*activeTab="(\d+)"', wb)
    if cur and int(cur.group(1)) == idx:
        return {"상태": "이미 있음", "시트": name}
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    repl = {}
    for i, (_, _, rid) in enumerate(sheets):
        p = "xl/" + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1).lstrip("/").replace("xl/", "")
        x = z.read(p).decode("utf-8")
        if i == idx and ' tabSelected="1"' not in x:
            repl[p] = re.sub(r"<sheetView ", '<sheetView tabSelected="1" ', x, count=1)
        elif i != idx and ' tabSelected="1"' in x:
            repl[p] = x.replace(' tabSelected="1"', "", 1)
    if cur:
        repl["xl/workbook.xml"] = re.sub(r'(<workbookView[^>]*?activeTab=")\d+', rf"\g<1>{idx}", wb, count=1)
    else:
        repl["xl/workbook.xml"] = re.sub(r"<workbookView ", f'<workbookView activeTab="{idx}" ', wb, count=1)
    with zipfile.ZipFile(dst, "w") as out:
        for item in z.infolist():
            data = repl[item.filename].encode("utf-8") if item.filename in repl else z.read(item.filename)
            out.writestr(item, data, compress_type=item.compress_type)
    return {"상태": "활성 탭 변경", "시트": name}


def build(src: str, dst: str, year: int, month: int) -> dict:
    name = f"{year % 100:02d} {month}월"
    with zipfile.ZipFile(src) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        sheets = re.findall(r'<sheet name="([^"]+)" sheetId="(\d+)" r:id="(rId\d+)"/>', wb)
        if any(unescape(n) == name for n, _, _ in sheets):
            return _ensure_active(z, src, dst, wb, sheets, name)
        dated = [(d, n, rid) for n, _, rid in sheets if (d := _sheet_name_date(unescape(n))) and d < (year, month)]
        if not dated:
            raise RuntimeError("복사할 직전 달 시트를 찾지 못함")
        (_, tname, trid) = max(dated)
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        tpath = "xl/" + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % trid, rels).group(1).lstrip("/").replace("xl/", "")
        xml = z.read(tpath).decode("utf-8")
        sst = _sst(z)

        days = workdays(year, month)
        k = split(days)
        blocks = [days[:k], days[k:]]
        nxt = (year + (month == 12), month % 12 + 1)
        labels = [f"{month}/25일 지급분", f"{nxt[1]}/10일 지급분"]

        rows = re.findall(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', xml, re.S)
        row_xml = {int(r): m.group(0) for r, m in zip(rows, re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', xml, re.S))}
        headers = [r for r, x in sorted(row_xml.items())
                   if (c := re.search(r'<c r="B%d"[^>]*?(?:/>|>.*?</c>)' % r, x, re.S)) and _cell_text(c.group(0), sst) == "이름"]
        if len(headers) != 2:
            raise RuntimeError(f"{tname}: '이름' 머리글 행이 {len(headers)}개 (2개여야 함)")
        people, notes = [], 0
        new_rows = dict(row_xml)
        for bi, h in enumerate(headers):
            x = new_rows[h]
            for i, col in enumerate(COLS):
                x = _set(x, f"{col}{h}", (blocks[bi][i] - EXCEL_EPOCH).days if i < len(blocks[bi]) else None)
            new_rows[h] = _set(x, f"T{h}", labels[bi])
            r = h + 1
            end = headers[1] if bi == 0 else max(row_xml) + 1
            while r < end:
                x = new_rows.get(r)
                if x is None:
                    r += 1
                    continue
                c = re.search(r'<c r="B%d"[^>]*?(?:/>|>.*?</c>)' % r, x, re.S)
                text = _cell_text(c.group(0), sst) if c else None
                if text and text.strip().startswith("=>"):
                    new_rows[r] = _set(x, f"B{r}", None)  # 지난달 메모 비움
                    notes += 1
                elif text and re.search(r'<c r="T%d"[^>]*>\s*<f' % r, x):
                    for i, col in enumerate(COLS):
                        x = _set(x, f"{col}{r}", 750 if i < len(blocks[bi]) else None)
                    x = re.sub(r'(<c r="T%d"[^>]*>(?:<f[^>]*/>|<f[^>]*>.*?</f>))<v>[^<]*</v>' % r, r"\1", x, flags=re.S)
                    new_rows[r] = x
                    if bi == 0:
                        people.append(text)
                r += 1
        body = xml
        for r, x in row_xml.items():
            body = body.replace(x, new_rows[r], 1)
        if ' tabSelected="1"' not in body:  # 새 시트를 선택된 탭으로
            body = re.sub(r"<sheetView ", '<sheetView tabSelected="1" ', body, count=1)
        body = re.sub(r'(<pageSetup[^>]*?) r:id="rId\d+"', r"\1", body)
        body = re.sub(r' xr:uid="\{[0-9A-F-]+\}"', "", body, count=1)

        # 새 시트 등록
        nums = [int(n) for n in re.findall(r"worksheets/sheet(\d+)\.xml", "\n".join(z.namelist()))]
        new_path = f"xl/worksheets/sheet{max(nums) + 1}.xml"
        rid_n = max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1
        sid = max(int(s) for _, s, _ in sheets) + 1
        rels2 = rels.replace("</Relationships>", f'<Relationship Id="rId{rid_n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{new_path.split("/")[-1]}"/></Relationships>')
        wb2 = wb.replace("</sheets>", f'<sheet name="{escape(name)}" sheetId="{sid}" r:id="rId{rid_n}"/></sheets>')
        # 파일을 열면 새 시트가 보이게 (현재 탭 = 새 시트, 다른 시트의 선택 표시는 끔)
        if re.search(r"<workbookView[^>]*activeTab=", wb2):
            wb2 = re.sub(r'(<workbookView[^>]*?activeTab=")\d+', rf"\g<1>{len(sheets)}", wb2, count=1)
        else:
            wb2 = re.sub(r"<workbookView ", f'<workbookView activeTab="{len(sheets)}" ', wb2, count=1)
        unselect = {}
        for n_, _, rid_ in sheets:
            p_ = "xl/" + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid_, rels).group(1).lstrip("/").replace("xl/", "")
            x_ = z.read(p_).decode("utf-8")
            if ' tabSelected="1"' in x_:
                unselect[p_] = x_.replace(' tabSelected="1"', "", 1)
        if "fullCalcOnLoad" not in wb2:
            wb2 = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb2, count=1)
        ct = z.read("[Content_Types].xml").decode("utf-8").replace(
            "</Types>", f'<Override PartName="/{new_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in z.namelist() else None
        if app:
            n = len(sheets) + 1
            app = re.sub(r"(<vt:lpstr>워크시트</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+", rf"\g<1>{n}", app)
            app = re.sub(r'(<TitlesOfParts><vt:vector size=")\d+', rf"\g<1>{n}", app)
            app = app.replace("</vt:vector></TitlesOfParts>", f"<vt:lpstr>{escape(name)}</vt:lpstr></vt:vector></TitlesOfParts>")
        repl = {"xl/workbook.xml": wb2, "xl/_rels/workbook.xml.rels": rels2, "[Content_Types].xml": ct, **unselect}
        if app:
            repl["docProps/app.xml"] = app
        with zipfile.ZipFile(dst, "w") as out:
            for item in z.infolist():
                data = repl[item.filename].encode("utf-8") if item.filename in repl else z.read(item.filename)
                out.writestr(item, data, compress_type=item.compress_type)
            out.writestr(zipfile.ZipInfo(new_path, date_time=z.getinfo(tpath).date_time), body.encode("utf-8"),
                         compress_type=zipfile.ZIP_DEFLATED)
    return {"상태": "생성", "시트": name, "복사한 시트": unescape(tname), "인원": people,
            "1차": [str(blocks[0][0]), str(blocks[0][-1]), len(blocks[0]), len(blocks[0]) * 750, labels[0]],
            "2차": [str(blocks[1][0]), str(blocks[1][-1]), len(blocks[1]), len(blocks[1]) * 750, labels[1]],
            "비운 메모": notes}
