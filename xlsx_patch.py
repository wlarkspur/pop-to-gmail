"""xlsx에서 지정한 셀 값만 바꾸고 나머지는 바이트 그대로 둔다 (openpyxl, LibreOffice로 다시 저장하지 않음).

- 바꾸는 건 해당 시트 XML 안의 <c r="..."> 요소와 workbook.xml의 calcPr(fullCalcOnLoad) 뿐
- 셀 서식(s 속성)은 유지. 숫자는 <v>, 글자는 inlineStr로 씀
- 수식 셀은 수식은 두고 캐시값(<v>)만 바꿀 수 있음 (cached=True)
- 다른 zip 항목은 내용 그대로 복사
"""
from __future__ import annotations

import re
import zipfile
from html import unescape
from xml.sax.saxutils import escape


def _sheet_path(z: zipfile.ZipFile, name: str) -> str:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(escape(name, {'"': "&quot;"})), wb)
    if not m:
        raise ValueError(f"시트 없음: {name}")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    t = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels) or \
        re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % m.group(1), rels)
    target = t.group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _num(v: float) -> str:
    s = repr(float(v))
    return s[:-2] if s.endswith(".0") else s


def _replace_cell(xml: str, ref: str, value, cached: bool) -> str:
    pat = re.compile(r'<c r="%s"((?: [^>]*?)?)(/>|>(.*?)</c>)' % ref, re.S)
    m = pat.search(xml)
    if not m:
        raise ValueError(f"셀 없음: {ref} (새 셀 추가는 안 함)")
    attrs, inner = m.group(1), m.group(3) or ""
    style = re.search(r' s="\d+"', attrs)
    style = style.group(0) if style else ""
    if cached:
        f = re.search(r"<f[^>]*>.*?</f>|<f[^>]*/>", inner, re.S)
        if not f:
            raise ValueError(f"{ref}: 수식 셀이 아님")
        t = ' t="str"' if isinstance(value, str) else ""
        v = escape(value) if isinstance(value, str) else _num(value)
        new = f'<c r="{ref}"{style}{t}>{f.group(0)}<v>{v}</v></c>'
    elif isinstance(value, str):
        new = f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    else:
        new = f'<c r="{ref}"{style}><v>{_num(value)}</v></c>'
    if "<f" in inner and not cached:
        raise ValueError(f"{ref}: 수식 셀은 값으로 덮어쓰지 않음")
    return xml[:m.start()] + new + xml[m.end():]


def read_values(path: str, sheet: str, refs: list[str]) -> dict:
    """셀의 현재 저장값(글자는 공유 문자열까지 풀어서)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read(_sheet_path(z, sheet)).decode("utf-8")
        sst = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode("utf-8"), re.S):
                sst.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    out = {}
    for ref in refs:
        m = re.search(r'<c r="%s"((?: [^>]*?)?)(/>|>(.*?)</c>)' % ref, xml, re.S)
        if not m or not m.group(3):
            out[ref] = None
            continue
        inner, attrs = m.group(3), m.group(1)
        if 't="s"' in attrs:
            out[ref] = unescape(sst[int(re.search(r"<v>(.*?)</v>", inner).group(1))]).replace("\r\n", "\n")
        elif 't="inlineStr"' in attrs:
            out[ref] = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S))).replace("\r\n", "\n")
        else:
            v = re.search(r"<v>(.*?)</v>", inner)
            out[ref] = float(v.group(1)) if v else None
    return out


def set_cells(src: str, dst: str, sheet: str, values: dict, cached: dict | None = None):
    cached = cached or {}
    with zipfile.ZipFile(src) as zin:
        path = _sheet_path(zin, sheet)
        xml = zin.read(path).decode("utf-8")
        for ref, v in values.items():
            xml = _replace_cell(xml, ref, v, cached=False)
        for ref, v in cached.items():
            xml = _replace_cell(xml, ref, v, cached=True)
        wb = zin.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        with zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                if item.filename == path:
                    zout.writestr(item, xml.encode("utf-8"), compress_type=item.compress_type)
                elif item.filename == "xl/workbook.xml":
                    zout.writestr(item, wb.encode("utf-8"), compress_type=item.compress_type)
                else:  # 내용 그대로 (압축 방식도 원본과 같게)
                    zout.writestr(item, zin.read(item.filename), compress_type=item.compress_type)


def replace_pictures(src: str, dst: str, sheet: str, pngs: list[bytes]):
    """시트에 붙어 있는 그림들을 왼쪽부터 순서대로 새 PNG로 바꾼다 (위치 유지, 크기는 새 그림 픽셀에 맞춤).
    그림 개수와 pngs 개수가 같아야 한다."""
    import struct
    with zipfile.ZipFile(src) as zin:
        spath = _sheet_path(zin, sheet)
        srels = zin.read(spath.replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
        d = re.search(r'Type="[^"]*/drawing" Target="([^"]+)"', srels) or re.search(r'Target="([^"]+)"[^>]*Type="[^"]*/drawing"', srels)
        if not d:
            raise ValueError(f"{sheet}: 그림 없음")
        dpath = "xl/" + d.group(1).replace("../", "")
        dxml = zin.read(dpath).decode("utf-8")
        drels = zin.read(dpath.replace("drawings/", "drawings/_rels/") + ".rels").decode("utf-8")
        anchors = list(re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor)\b.*?</xdr:\1>", dxml, re.S))
        pics = []
        for a in anchors:
            col = int(re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>", a.group(0)).group(1))
            rid = re.search(r'r:embed="(rId\d+)"', a.group(0)).group(1)
            target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, drels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, drels)
            pics.append((col, a, "xl/" + target.group(1).replace("../", "")))
        if len(pics) != len(pngs):
            raise ValueError(f"{sheet}: 그림 {len(pics)}개, 새 그림 {len(pngs)}개")
        pics.sort(key=lambda x: x[0])
        new_media, new_xml = {}, dxml
        for (col, a, media), png in sorted(zip(pics, pngs), key=lambda x: -x[0][1].start()):  # 뒤에서부터 바꿔야 위치가 안 밀림
            w, h = struct.unpack(">II", png[16:24])
            cx, cy = w * 9525, h * 9525
            frm = re.search(r"<xdr:from>.*?</xdr:from>", a.group(0), re.S).group(0)
            pic = re.search(r"<xdr:pic>.*?</xdr:pic>", a.group(0), re.S).group(0)
            pic = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', f'<a:ext cx="{cx}" cy="{cy}"/>', pic, count=1)
            one = f'<xdr:oneCellAnchor>{frm}<xdr:ext cx="{cx}" cy="{cy}"/>{pic}<xdr:clientData/></xdr:oneCellAnchor>'
            new_xml = new_xml[:a.start()] + one + new_xml[a.end():]
            new_media[media] = png
        wb = zin.read("xl/workbook.xml").decode("utf-8")
        with zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                if item.filename == dpath:
                    zout.writestr(item, new_xml.encode("utf-8"), compress_type=item.compress_type)
                elif item.filename in new_media:
                    zout.writestr(item, new_media[item.filename], compress_type=item.compress_type)
                else:
                    zout.writestr(item, zin.read(item.filename), compress_type=item.compress_type)


def focus_xml(workbook_xml: str, sheet_xmls: dict, sheet_order: list, sheet_path: str, cell: str):
    """xlsx: 파일을 열면 sheet_path 시트의 cell이 선택돼 보이게.
    sheet_xmls = {경로: xml} (모든 시트), sheet_order = 탭 순서대로 경로. 바뀐 것만 dict로 돌려준다."""
    out = {}
    idx = sheet_order.index(sheet_path)
    m = re.match(r"([A-Z]+)(\d+)", cell)
    col_n = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1))))
    row_n = int(m.group(2))
    top = f"{_colname(max(1, col_n - 6))}{max(1, row_n - 12)}"
    for p in sheet_order:
        x = sheet_xmls[p]
        if p == sheet_path:
            v = re.search(r"<sheetView\b[^>]*?(/>|>)", x)
            tag = v.group(0)
            new = re.sub(r' (tabSelected|topLeftCell)="[^"]*"', "", tag)
            new = new.replace("<sheetView", f'<sheetView tabSelected="1" topLeftCell="{top}"', 1)
            if tag.endswith("/>"):
                new = new[:-2] + f'><selection activeCell="{cell}" sqref="{cell}"/></sheetView>'
                x = x.replace(tag, new, 1)
            else:
                end = x.index("</sheetView>", v.end())
                inner = x[v.end():end]
                if "<selection" in inner:
                    inner = re.sub(r'activeCell="[^"]*"', f'activeCell="{cell}"', inner)
                    inner = re.sub(r'sqref="[^"]*"', f'sqref="{cell}"', inner)
                    inner = re.sub(r' activeCellId="\d+"', "", inner)
                else:
                    inner += f'<selection activeCell="{cell}" sqref="{cell}"/>'
                x = x[:v.start()] + new + inner + x[end:]
            out[p] = x
        elif ' tabSelected="1"' in x:
            out[p] = x.replace(' tabSelected="1"', "", 1)
    wb = workbook_xml
    if re.search(r"<workbookView[^>]*activeTab=", wb):
        wb = re.sub(r'(<workbookView[^>]*?activeTab=")\d+', rf"\g<1>{idx}", wb, count=1)
    else:
        wb = re.sub(r"<workbookView ", f'<workbookView activeTab="{idx}" ', wb, count=1)
    fs = re.search(r'<workbookView[^>]*firstSheet="(\d+)"', wb)
    if fs and int(fs.group(1)) > idx:
        wb = re.sub(r'(<workbookView[^>]*?firstSheet=")\d+', rf"\g<1>{idx}", wb, count=1)
    out["xl/workbook.xml"] = wb
    return out


def _colname(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_paths(z: zipfile.ZipFile) -> tuple[list[str], list[str]]:
    """(탭 순서대로 시트 이름, 경로)"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    names, paths = [], []
    for n, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1).lstrip("/")
        names.append(unescape(n))
        paths.append(t if t.startswith("xl/") else "xl/" + t)
    return names, paths


def focus(src: str, dst: str, sheet: str, cell: str):
    """xlsx 파일을 열면 sheet 시트의 cell이 선택돼 보이게 (내용은 그대로)."""
    with zipfile.ZipFile(src) as z:
        names, paths = sheet_paths(z)
        xmls = {p: z.read(p).decode("utf-8") for p in paths}
        repl = focus_xml(z.read("xl/workbook.xml").decode("utf-8"), xmls, paths, paths[names.index(sheet)], cell)
        with zipfile.ZipFile(dst, "w") as out:
            for item in z.infolist():
                data = repl[item.filename].encode("utf-8") if item.filename in repl else z.read(item.filename)
                out.writestr(item, data, compress_type=item.compress_type)
