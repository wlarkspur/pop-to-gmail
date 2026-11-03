"""BIFF8 .xls 파일에서 지정한 셀의 숫자만 바꾸고 나머지 바이트는 그대로 둔다.

- 서식, 수식, 다른 시트, 문서 속성은 건드리지 않는다 (엑셀이나 LibreOffice로 다시 저장하지 않음)
- 바꿀 수 있는 셀: 이미 있는 빈 셀(BLANK, MULBLANK 안의 칸), 숫자 셀(NUMBER, RK). 셀 서식(XF)은 원래 것을 유지
- 셀 레코드가 아예 없는 칸은 거부한다 (행 구조를 새로 만들어야 해서 위험)
- 바꾼 뒤 DBCELL, INDEX 오프셋을 다시 계산하고, RECALCID를 낮춰 엑셀이 열 때 모든 수식을 다시 계산하게 한다
- 복합 문서(OLE) 컨테이너는 원래 스트림 3개(Workbook, SummaryInformation, DocumentSummaryInformation)를 그대로 다시 담는다
"""

from __future__ import annotations

import struct

import olefile

BOF, EOF_, BOUNDSHEET, ROW, DBCELL, INDEX = 0x809, 0x0A, 0x85, 0x208, 0xD7, 0x20B
BLANK, MULBLANK, NUMBER, RK, RECALCID, MULRK = 0x201, 0xBE, 0x203, 0x27E, 0x1C1, 0xBD
DEFCOLWIDTH, EXTSST, SST, LABELSST, LABEL, FORMULA, STRING, WINDOW2 = 0x55, 0xFF, 0xFC, 0xFD, 0x204, 0x06, 0x207, 0x23E
CELL_TYPES = {0x06, 0x201, 0x203, 0x204, 0x205, 0x27E, 0xBD, 0xBE, 0xFD}  # 행 번호로 시작하는 셀 레코드만


def _records(buf: bytes):
    out, p = [], 0
    while p < len(buf):
        t, n = struct.unpack_from("<HH", buf, p)
        out.append([t, bytes(buf[p + 4:p + 4 + n])])
        p += 4 + n
    return out


def _join(recs) -> bytes:
    return b"".join(struct.pack("<HH", t, len(d)) + d for t, d in recs)


def _sheet_ranges(recs):
    """[(이름, 시작 인덱스, 끝 인덱스)] - 시트 서브스트림의 레코드 인덱스 범위."""
    names = []
    for t, d in recs:
        if t == BOUNDSHEET:
            cch, flag = d[6], d[7]
            names.append(d[8:8 + cch * 2].decode("utf-16-le") if flag & 1 else d[8:8 + cch].decode("latin-1"))
        if t == EOF_:
            break
    bofs = [i for i, (t, _) in enumerate(recs) if t == BOF]
    out = []
    for name, start in zip(names, bofs[1:]):
        end = next(i for i in range(start, len(recs)) if recs[i][0] == EOF_)
        out.append((name, start, end))
    return out


def _cell(r, c, xf, value):
    """숫자면 NUMBER, 글자면 LABEL(직접 문자열) 레코드."""
    if isinstance(value, str):
        return [LABEL, struct.pack("<HHHHB", r, c, xf, len(value), 1) + value.encode("utf-16-le")]
    return [NUMBER, struct.pack("<HHHd", r, c, xf, float(value))]


def _set_cell(recs, lo, hi, row, col, value) -> int:
    """시트 레코드 [lo, hi] 안에서 (row, col)을 숫자나 글자로 바꾼다 (셀 서식 유지). 늘어난 레코드 수를 돌려준다."""
    for i in range(lo, hi + 1):
        t, d = recs[i]
        if t in (BLANK, NUMBER, RK, LABELSST, LABEL):
            r, c, xf = struct.unpack_from("<HHH", d)
            if (r, c) == (row, col):
                recs[i] = _cell(r, c, xf, value)
                return 0
        elif t == MULRK:
            r, c1 = struct.unpack_from("<HH", d)
            c2 = struct.unpack_from("<H", d, len(d) - 2)[0]
            if r == row and c1 <= col <= c2:  # 칸별 RK 레코드로 풀어서 다시 시도
                parts = [[RK, struct.pack("<HH", r, c1 + k) + d[4 + 6 * k:10 + 6 * k]] for k in range(c2 - c1 + 1)]
                recs[i:i + 1] = parts
                return len(parts) - 1 + _set_cell(recs, i, hi + len(parts) - 1, row, col, value)
        elif t == MULBLANK:
            r, c1 = struct.unpack_from("<HH", d)
            c2 = struct.unpack_from("<H", d, len(d) - 2)[0]
            if r == row and c1 <= col <= c2:
                xfs = struct.unpack_from("<%dH" % (c2 - c1 + 1), d, 4)
                new = []
                if col > c1:
                    new.append(_blanks(r, c1, xfs[:col - c1]))
                new.append(_cell(r, col, xfs[col - c1], value))
                if col < c2:
                    new.append(_blanks(r, col + 1, xfs[col - c1 + 1:]))
                recs[i:i + 1] = new
                return len(new) - 1
    raise ValueError(f"셀 레코드가 없음 (행 {row + 1}, 열 {col + 1}) - 바꾸지 않음")


def _blanks(r, c1, xfs):
    if len(xfs) == 1:
        return [BLANK, struct.pack("<HHH", r, c1, xfs[0])]
    return [MULBLANK, struct.pack("<HH", r, c1) + struct.pack("<%dH" % len(xfs), *xfs) + struct.pack("<H", c1 + len(xfs) - 1)]


def _fix_dbcells(recs, lo, hi, base):
    """시트의 DBCELL, INDEX를 실제 위치로 다시 계산한다. base = 시트 BOF의 스트림 위치."""
    pos, p = [], base
    for t, d in recs[lo:hi + 1]:
        pos.append(p)
        p += 4 + len(d)
    idx = list(range(lo, hi + 1))
    db_positions = []
    block_rows = []  # 현재 블록의 ROW 레코드 (행번호, 위치)
    first_cell = {}
    for k, i in enumerate(idx):
        t, d = recs[i]
        if t == ROW:
            block_rows.append((struct.unpack_from("<H", d)[0], pos[k], len(d)))
        elif t in CELL_TYPES:
            r = struct.unpack_from("<H", d)[0]
            first_cell.setdefault(r, pos[k])
        elif t == DBCELL:
            first_row_pos = block_rows[0][1]
            anchor = block_rows[1][1] if len(block_rows) > 1 else first_row_pos + 4 + block_rows[0][2]
            # 빈 행의 위치 = 다음 셀 있는 행의 첫 셀 위치 (엑셀이 쓰는 방식, 원본으로 확인)
            starts, nxt = [], pos[k]
            for r, _, _ in reversed(block_rows):
                nxt = first_cell.get(r, nxt)
                starts.append(nxt)
            offs = []
            for st in reversed(starts):
                offs.append(st - anchor)
                anchor = st
            recs[i] = [DBCELL, struct.pack("<I", pos[k] - first_row_pos) + struct.pack("<%dH" % len(offs), *offs)]
            db_positions.append(pos[k])
            block_rows, first_cell = [], {}
    defcol = next((pos[k] for k, i in enumerate(idx) if recs[i][0] == DEFCOLWIDTH), 0)
    for k, i in enumerate(idx):
        if recs[i][0] == INDEX:
            d = recs[i][1]
            n = (len(d) - 16) // 4
            if n != len(db_positions):
                raise ValueError("INDEX와 DBCELL 개수가 다름")
            recs[i] = [INDEX, d[:12] + struct.pack("<I", defcol) + struct.pack("<%dI" % n, *db_positions)]


def _finalize(recs, sst_shift=0):
    """모든 시트의 BOUNDSHEET 위치, DBCELL, INDEX를 다시 계산. sst_shift = SST 앞에서 늘어난 바이트 수."""
    ranges = _sheet_ranges(recs)
    pos, p = [], 0
    for t, d in recs:
        pos.append(p)
        p += 4 + len(d)
    bs = [i for i, (t, _) in enumerate(recs[:ranges[0][1]]) if t == BOUNDSHEET]
    for i, (_, lo, _) in zip(bs, ranges):
        recs[i] = [BOUNDSHEET, struct.pack("<I", pos[lo]) + recs[i][1][4:]]
    if sst_shift:
        for i, (t, d) in enumerate(recs[:ranges[0][1]]):
            if t == EXTSST:
                n = (len(d) - 2) // 8
                items = [struct.unpack_from("<IHH", d, 2 + 8 * k) for k in range(n)]
                recs[i] = [EXTSST, d[:2] + b"".join(struct.pack("<IHH", a + sst_shift, b, c) for a, b, c in items)]
    for _, lo, hi in ranges:
        _fix_dbcells(recs, lo, hi, pos[lo])


def _recalc_on_open(recs):
    for rec in recs:
        if rec[0] == RECALCID:  # 엑셀이 열 때 전체 다시 계산
            rec[1] = rec[1][:4] + struct.pack("<I", 1)


def self_check(wb: bytes) -> bool:
    """원본에 대해 오프셋 재계산을 돌려 바이트가 그대로인지 (구조를 제대로 이해했는지) 확인."""
    recs = [list(r) for r in _records(wb)]
    _finalize(recs)
    return _join(recs) == wb


def _focus(recs, sheet: str, row: int, col: int):
    """파일을 열면 sheet 시트의 (row, col) 칸이 선택돼 보이게 (WINDOW1 현재 탭, WINDOW2 스크롤, SELECTION)."""
    ranges = _sheet_ranges(recs)
    idx = [n for n, _, _ in ranges].index(sheet)
    for i, (t, d) in enumerate(recs[:ranges[0][1]]):
        if t == 0x3D:  # WINDOW1: itabCur, itabFirst, ctabSel
            first = min(struct.unpack_from("<H", d, 12)[0], idx)
            recs[i] = [t, d[:10] + struct.pack("<HHH", idx, first, 1) + d[16:]]
    for k, (name, lo, hi) in enumerate(ranges):
        for i in range(lo, hi + 1):
            t, d = recs[i]
            if t == WINDOW2:
                flags = struct.unpack_from("<H", d)[0]
                if k == idx:
                    d = struct.pack("<HHH", flags | 0x0600, max(0, row - 12), max(0, col - 6)) + d[6:]
                else:
                    d = struct.pack("<H", flags & ~0x0600) + d[2:]
                recs[i] = [t, d]
            elif t == 0x1D and k == idx:  # SELECTION
                recs[i] = [t, struct.pack("<BHHHH", d[0], row, col, 0, 1) + struct.pack("<HHBB", row, row, col, col)]


def patch_workbook_stream(wb: bytes, sheet: str, cells: dict[tuple[int, int], float | str], focus=None) -> bytes:
    recs = [list(r) for r in _records(wb)]
    target = [r for r in _sheet_ranges(recs) if r[0] == sheet]
    if not target:
        raise ValueError(f"시트 없음: {sheet}")
    _, lo, hi = target[0]
    n_sst = sum(1 for t, _ in recs if t == LABELSST)
    for (r, c), v in sorted(cells.items()):
        hi += _set_cell(recs, lo, hi, r, c, v if isinstance(v, str) else float(v))
    removed = n_sst - sum(1 for t, _ in recs if t == LABELSST)
    if removed:  # SST 참조 개수(cstTotal) 맞추기
        i = next(i for i, (t, _) in enumerate(recs) if t == SST)
        recs[i] = [SST, struct.pack("<I", struct.unpack_from("<I", recs[i][1])[0] - removed) + recs[i][1][4:]]
    if focus:
        _focus(recs, sheet, *focus)
    _finalize(recs)
    _recalc_on_open(recs)
    return _join(recs)


def add_year_sheet(wb: bytes, src_name: str, new_name: str, day_shift: int, label_from: str, label_to: str):
    """src 시트를 복사해 맨 뒤에 new 시트를 만든다.
    - 날짜 상수: day_shift일 더함 (수식 날짜는 앞 칸 +1 이라 엑셀이 다시 계산)
    - 숫자 상수: 1~12 정수(월 머리글)만 남기고 빈 칸으로 (환율, 수기 입력값 비움, 셀 서식은 유지)
    - 수식: 그대로, 캐시값만 비움 (날짜 수식은 캐시값도 day_shift 반영)
    - 글자: label_from이 들어간 칸만 label_to로 바꿔 직접 문자열(LABEL)로 씀
    """
    import xlrd as _x  # 날짜 서식 판별용
    recs = [list(r) for r in _records(wb)]
    ranges = _sheet_ranges(recs)
    if any(n == new_name for n, _, _ in ranges):
        raise ValueError(f"이미 있음: {new_name}")
    src = [r for r in ranges if r[0] == src_name]
    if not src:
        raise ValueError(f"시트 없음: {src_name}")
    _, lo, hi = src[0]
    book = _x.open_workbook(file_contents=wb, on_demand=True)  # xlrd는 Workbook 스트림 바이트도 읽음
    sh = book.sheet_by_name(src_name)
    sst = book._sharedstrings
    copy, n_sst = [], 0
    k = lo
    while k <= hi:
        t, d = recs[k]
        d = bytes(d)
        if t in (NUMBER, RK):
            r, c, xf = struct.unpack_from("<HHH", d)
            v = sh.cell_value(r, c)
            if sh.cell_type(r, c) == 3:
                v += day_shift
                d = struct.pack("<HHHd", r, c, xf, v)
                t = NUMBER
            elif not (float(v).is_integer() and 1 <= v <= 12):
                t, d = BLANK, struct.pack("<HHH", r, c, xf)
        elif t == FORMULA:
            r, c, xf = struct.unpack_from("<HHH", d)
            if sh.cell_type(r, c) == 3:
                res = struct.pack("<d", sh.cell_value(r, c) + day_shift)
            else:
                res = bytes([3, 0, 0, 0, 0, 0, 0xFF, 0xFF])  # 빈 문자열 결과
            d = d[:6] + res + d[14:]
            if k + 1 <= hi and recs[k + 1][0] == STRING:
                copy.append([t, d])
                k += 2
                continue
        elif t == LABELSST:
            r, c, xf, isst = struct.unpack_from("<HHHI", d)
            text = sst[isst]
            if label_from in text:
                text = text.replace(label_from, label_to)
                t, d = LABEL, struct.pack("<HHHHB", r, c, xf, len(text), 1) + text.encode("utf-16-le")
            else:
                n_sst += 1
        elif t == WINDOW2:  # 새 시트를 선택된 탭으로
            flags = struct.unpack_from("<H", d)[0] | 0x0600
            d = struct.pack("<H", flags) + d[2:]
        copy.append([t, d])
        k += 1
    last_bs = max(i for i, (t, _) in enumerate(recs[:ranges[0][1]]) if t == BOUNDSHEET)
    name = new_name.encode("utf-16-le")
    hs = recs[last_bs][1][4:6]  # 보이기 상태, 시트 종류는 마지막 시트와 같게
    new_bs = [BOUNDSHEET, struct.pack("<I", 0) + hs + bytes([len(new_name), 1]) + name]
    recs.insert(last_bs + 1, new_bs)
    for i, (t, d) in enumerate(recs):
        if t == SST:
            recs[i] = [SST, struct.pack("<I", struct.unpack_from("<I", d)[0] + n_sst) + d[4:]]
            break
    # 파일을 열면 새 시트가 보이게: 다른 시트의 선택 표시를 끄고 WINDOW1 현재 탭을 새 시트로
    for i, (t, d) in enumerate(recs):
        if t == WINDOW2:
            recs[i] = [t, struct.pack("<H", struct.unpack_from("<H", d)[0] & ~0x0600) + d[2:]]
        elif t == 0x3D:  # WINDOW1
            recs[i] = [t, d[:10] + struct.pack("<HHH", len(ranges), struct.unpack_from("<H", d, 12)[0], 1) + d[16:]]
    recs.extend(copy)
    _finalize(recs, sst_shift=4 + len(new_bs[1]))
    _recalc_on_open(recs)
    return _join(recs)


# ---------------- OLE 복합 문서 쓰기 (스트림 3개, 모두 4096바이트 이상 → 일반 섹터) ----------------

def _dir_entry(name, typ, color, left, right, child, clsid, start, size):
    n = (name + "\0").encode("utf-16-le") if name else b""
    return (n.ljust(64, b"\0") + struct.pack("<HBBIII", len(n), typ, color, left, right, child)
            + clsid + b"\0" * 4 + b"\0" * 16 + struct.pack("<II", start, size) + b"\0" * 4)


def write_ole(path, streams: list[tuple[str, bytes]], root_clsid: bytes):
    """streams: [(이름, 데이터)] 이름 정렬 규칙대로 이진 트리를 만든다 (모두 4096바이트 이상이어야 함)."""
    SS, FREE, END, FATSECT = 512, 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
    for name, data in streams:
        if len(data) < 4096:
            raise ValueError("작은 스트림(미니 스트림)은 미지원")
    chains, data_sectors = [], []
    for name, data in streams:
        n = -(-len(data) // SS)
        chains.append((len(data_sectors), n))
        data_sectors += [data[i * SS:(i + 1) * SS].ljust(SS, b"\0") for i in range(n)]
    dir_sects = 1  # 엔트리 4개 = 512바이트
    total_wo_fat = len(data_sectors) + dir_sects
    nfat = 1
    while nfat * 128 < total_wo_fat + nfat:
        nfat += 1
    if nfat > 109:
        raise ValueError("DIFAT 확장 미지원")
    fat = [FREE] * (nfat * 128)
    for start, n in chains:
        for k in range(n):
            fat[start + k] = start + k + 1 if k < n - 1 else END
    dir_start = len(data_sectors)
    fat[dir_start] = END
    fat_start = dir_start + 1
    for k in range(nfat):
        fat[fat_start + k] = FATSECT
    # 디렉터리 트리: 이름 (길이, 대문자) 순으로 정렬해 가운데를 루트 자식으로
    order = sorted(range(len(streams)), key=lambda i: (len(streams[i][0]), streams[i][0].upper()))

    def build(ids):
        if not ids:
            return FREE, {}
        m = len(ids) // 2
        lft, a = build(ids[:m])
        rgt, b = build(ids[m + 1:])
        return ids[m] + 1, {**a, **b, ids[m] + 1: (lft, rgt)}
    child, links = build(order)
    ents = [_dir_entry("Root Entry", 5, 1, FREE, FREE, child, root_clsid, END, 0)]
    for i, (name, data) in enumerate(streams):
        l, r = links[i + 1]
        ents.append(_dir_entry(name, 2, 1, l, r, FREE, b"\0" * 16, chains[i][0], len(data)))
    directory = b"".join(ents).ljust(SS, b"\0")
    difat = list(range(fat_start, fat_start + nfat)) + [FREE] * (109 - nfat)
    header = (bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16 + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6)
              + b"\0" * 6 + struct.pack("<IIIIIIIII", 0, nfat, dir_start, 0, 4096, END, 0, END, 0)
              + struct.pack("<109I", *difat))
    assert len(header) == 512
    fat_bytes = struct.pack("<%dI" % len(fat), *fat)
    with open(path, "wb") as f:
        f.write(header + b"".join(data_sectors) + directory + fat_bytes)


def _read(src):
    ole = olefile.OleFileIO(src)
    names = ["Workbook", "\x05SummaryInformation", "\x05DocumentSummaryInformation"]
    have = ["/".join(e) for e in ole.listdir()]
    if sorted(have) != sorted(names):
        raise ValueError(f"예상과 다른 스트림 구성: {have}")
    streams = [(n, ole.openstream(n).read()) for n in names]
    clsid = bytes.fromhex(ole.root.clsid.replace("-", ""))
    clsid = clsid[:4][::-1] + clsid[4:6][::-1] + clsid[6:8][::-1] + clsid[8:]
    ole.close()
    if not self_check(streams[0][1]):
        raise ValueError("오프셋 재계산이 원본과 다름 - 구조가 예상과 달라 중단")
    return streams, clsid


def patch_xls(src, dst, sheet, cells, focus=None):
    """focus=(행, 열): 파일을 열면 이 칸이 선택된 채로 보이게 (기본: 마지막으로 바꾼 칸)."""
    streams, clsid = _read(src)
    streams[0] = ("Workbook", patch_workbook_stream(streams[0][1], sheet, cells, focus or (max(cells) if cells else None)))
    write_ole(dst, streams, clsid)


def add_year_sheet_xls(src, dst, src_name, new_name, day_shift, label_from, label_to):
    streams, clsid = _read(src)
    streams[0] = ("Workbook", add_year_sheet(streams[0][1], src_name, new_name, day_shift, label_from, label_to))
    write_ole(dst, streams, clsid)
