"""YSKD(ЁНГСАН КЛД) 가마감 작업의 알맹이: 양식은 그대로 두고 값만 바꿈.

1) 누적 마감 (YYYY년 M월_누적 마감_YSKD.xlsx)
   - '26년 실적(RUB)' 시트의 월 열(F=1월 ... Q=12월)에 Irina monthly 파일 'KLD 2026' 시트의 같은 항목 월 값을 넣음
     (두 시트는 같은 뿌리의 양식이라 행이 1:1로 대응. ROWMAP = 누적 마감 행 -> KLD 2026 행)
   - '환율정보' 44행(EUR/RUB 월평균)에 CBR 일별 고시의 달력일 평균을 넣음 (사용자 10월 파일로 소수 끝자리까지 일치 확인)
   - summary B1 = 보고 월
   - 비어 있는 과거 월 열도 같이 채움 (예: 9월을 건너뛴 상태에서 10월을 만들 때). 이미 값이 있는 과거 월은 안 바꾸고 KLD와 다르면 목록으로.
2) 채권채무 관리 (YYYY년 M월_채권_채무_관리_YSKD.xlsx)
   - 전월 파일 + 1C 채권채무 잔액(Кредиторская-дебиторская задолженность, marina@youngsan.com 메일) -> 6.Rawdata대사 행 교체,
     3.채권/4.채무의 Rawdata 연동 칸 갱신, 기준일, EUR 월말 환율(CBR). 집계키는 전월 것을 계정+거래처로 이어받고 신규 행은 '관리대상 외' + 비고.
3) KLD 재무제표 (KLD 재무제표 (RUB) YY년M월.xlsx) = Irina 첨부 그대로 (이름만)
"""
from __future__ import annotations

import copy
import datetime as dt
import re

import openpyxl

S_RUB, S_FX, S_SUM = "26년 실적(RUB)", "환율정보", "summary "
K_SHEET = "KLD 2026"
# 누적 마감 '26년 실적(RUB)' 행 -> Irina monthly 'KLD 2026' 행 (판관비, 영업외, 법인세 입력 행만. 소계/합계는 양식 수식)
ROWMAP = {47: 23, 48: 24, 49: 25, 50: 26, 51: 27, 55: 28, 56: 29, 57: 30, 58: 31, 59: 32,
          60: 33, 61: 34, 62: 35, 63: 36, 64: 37, 65: 38, 66: 39, 67: 40, 68: 41, 69: 42, 70: 43, 71: 44,
          72: 45, 73: 46, 74: 47, 75: 48, 76: 49, 77: 50, 78: 51, 79: 52, 80: 53, 81: 54,
          82: 55, 83: 56, 84: 57, 85: 58, 86: 59, 87: 60, 88: 61, 89: 62, 90: 63, 91: 64, 92: 65,
          93: 66, 94: 67, 95: 68, 96: 69, 97: 70, 99: 71, 113: 72, 114: 73,
          117: 76, 118: 77, 119: 78, 120: 79,
          122: 81, 123: 82, 124: 83, 125: 84, 126: 91, 127: 86, 128: 87, 129: 85, 130: 88, 131: 89, 132: 90,
          136: 95, 137: 96, 138: 97}
# 합계 대조: (누적 마감 입력 행들의 합) = KLD 합계 행
CHECKS = {"판관비": ([r for r in ROWMAP if 47 <= r <= 114], 22), "영업외수익": ([117, 118, 119, 120], 75),
          "영업외비용": ([r for r in ROWMAP if 122 <= r <= 132], 80), "법인세": ([136, 137, 138], 94)}


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def _norm(s):
    return re.sub(r"[^a-z0-9가-힣]", "", str(s or "").lower())


def fill_cum(tmpl: str, monthly: str, y: int, m: int, avg_rates: dict[int, float], out: str) -> tuple[list[str], list[str]]:
    """전월 누적 마감 양식에 KLD 2026 월 값을 넣어 out에 저장. (바꾼 칸, 확인 사항) 반환."""
    wb = openpyxl.load_workbook(tmpl)
    rub, fx, sm = wb[S_RUB], wb[S_FX], wb[S_SUM]
    k = openpyxl.load_workbook(monthly, data_only=True)[K_SHEET]
    log, notes = [], []
    if _num(k["F1"].value) != y:
        notes.append(f"monthly 파일 KLD 2026 연도 표시가 {k['F1'].value} (기대 {y})")
    for cr, kr in ROWMAP.items():  # 양식이 바뀌었는지: 영문 항목명 대조
        a, b = _norm(rub.cell(cr, 3).value), _norm(k.cell(kr, 4).value)
        if a and b and a[:12] != b[:12]:
            notes.append(f"항목명 불일치: 누적 마감 {cr}행 '{rub.cell(cr, 3).value}' vs KLD 2026 {kr}행 '{k.cell(kr, 4).value}'")
    for mon in range(1, m + 1):
        cc, kc = 5 + mon, 7 + mon
        empty = all(_num(rub.cell(r, cc).value) == 0 for r in ROWMAP)  # 0, 빈칸, 수식만 있으면 아직 안 채운 달
        write = mon == m or empty
        for cr, kr in ROWMAP.items():
            kv = round(_num(k.cell(kr, kc).value), 2)
            cur = rub.cell(cr, cc)
            if isinstance(cur.value, str) and cur.value.startswith("="):
                if kv:
                    notes.append(f"{S_RUB}!{cur.coordinate} 가 수식이라 KLD 값 {kv:,.2f} 을 못 넣음")
                continue
            if write:
                if kv != round(_num(cur.value), 2):
                    cur.value = kv if kv else None
                    log.append(cur.coordinate)
            elif kv != round(_num(cur.value), 2):
                notes.append(f"{mon}월 {S_RUB}!{cur.coordinate} 기존 {_num(cur.value):,.2f} vs KLD {kv:,.2f} (과거 월이라 안 바꿈)")
        if write:
            for name, (rows, tot) in CHECKS.items():
                s = round(sum(_num(k.cell(ROWMAP[r], kc).value) for r in rows), 2)
                t = round(_num(k.cell(tot, kc).value), 2)
                if abs(s - t) > 0.01:
                    notes.append(f"{mon}월 {name}: 매핑 행 합 {s:,.2f} ≠ KLD 합계 {t:,.2f} (매핑 안 된 항목이 있음)")
        # 환율 (환율정보 44행: C=1월 ... N=12월)
        fc = fx.cell(44, 2 + mon)
        if mon in avg_rates and (write or fc.value in (None, "")):
            if round(_num(fc.value), 6) != round(avg_rates[mon], 6):
                fc.value = avg_rates[mon]
                log.append(f"{S_FX}!{fc.coordinate}")
    if sm["B1"].value != m:
        sm["B1"] = m
        log.append(f"{S_SUM.strip()}!B1")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return log, notes


# ─────────────────────────── 채권채무 관리
R_SHEET, AR_SHEET, AP_SHEET, SUM_SHEET = "6.Rawdata대사", "3.채권", "4.채무", "1.거래처_종합"
R_FIRST = 6  # Rawdata 첫 데이터 행 (5행 머리글)
WORDS = ("YOUNGSAN GLONET", "ЭЙЧ УАЙ", "ХАНДЕЛЬС", "МСК", "КВЦ", "ЭЛЛАДА", "ГЕФЕСТ", "ЭЛИАС", "КАРМЕЛ", "РОСГОССТРАХ",
         "СОГЛАСИЕ", "МЕГАФОН", "ЛУКОЙЛ", "ЛИНДЕ", "ТЕРЕЩЕНКО", "ФТС", "НАЛОГИ", "НДФЛ", "НДС", "НА ПРИБЫЛЬ", "ЕНС",
         "СТРАХОВЫЕ", "ОПЛАТЕ ТРУДА", "ГЛОВИС", "АРГЕНТУМ", "СТРАХОВАНИЕ", "ЮАП", "ЛОГИСТИК", "СЕБИКО", "СДЭК", "КОЛОТИЛИН", "ЧЖОН")
# 3.채권/4.채무에서 Rawdata 값을 그대로 받는 칸: (시트, 칸, 구분, 계정, 거래처 핵심어, 'rub'|'eur')
CELLMAP = [
    (AR_SHEET, "F6", "AR", "76.22", "YOUNGSAN", "rub"), (AR_SHEET, "F7", "AR", "76.02", "YOUNGSAN", "rub"),
    (AR_SHEET, "F14", "AR", "76.09", "МСК", "rub"),
    (AR_SHEET, "F15", "AR", "58.03", "ХАНДЕЛЬС", "rub"), (AR_SHEET, "F16", "AR", "76.09", "ХАНДЕЛЬС", "rub"),
    (AR_SHEET, "F17", "AR", "58.03", "ЭЙЧ", "rub"), (AR_SHEET, "F18", "AR", "76.09", "ЭЙЧ", "rub"),
    (AP_SHEET, "F21", "AP", "76.05", "ЭЛЛАДА", "rub"), (AP_SHEET, "F22", "AP", "60.01", "ПРЕТЕНЗ", "rub"),
    (AP_SHEET, "F23", "AP", "60.01", "КОММУН", "rub"),
    (AP_SHEET, "G42", "AP", "67.23", "КАРМЕЛ", "eur"), (AP_SHEET, "G43", "AP", "67.24", "КАРМЕЛ", "eur"),
    (AP_SHEET, "F44", "AP", "67.03", "ЭЛИАС", "rub"), (AP_SHEET, "F45", "AP", "67.04", "ЭЛИАС", "rub"),
    (AP_SHEET, "F46", "AP", "67.03", "ГЕФЕСТ", "rub"), (AP_SHEET, "F47", "AP", "67.04", "ГЕФЕСТ", "rub"),
]


def read_raw(path: str) -> list[list]:
    """1C 채권채무 잔액 파일 -> [구분, 계정, 거래처, RUB, EUR|None]. 계정 칸이 비면 위 행 계정을 이어받음."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ar = next(ws for ws in wb if "дебитор" in ws.title.lower())
    ap = next(ws for ws in wb if "кредитор" in ws.title.lower())
    out = []
    for ws, kind in ((ar, "AR"), (ap, "AP")):
        acct = None
        for r in range(3, ws.max_row + 1):
            a, name, rub, eur = (ws.cell(r, c).value for c in (2, 3, 4, 5))
            if a not in (None, ""):
                acct = str(a).strip()
            if not name or not isinstance(name, str) or name.strip().startswith("="):
                continue
            if name.strip() == "Налоги" and a in (None, "") and not isinstance(rub, (int, float)):
                continue  # 세금 묶음 제목 줄
            out.append([kind, acct, name.strip(), round(_num(rub), 2), (round(eur, 2) if isinstance(eur, (int, float)) else None)])
    return out


def raw_date(path: str) -> dt.date | None:
    """Дебиторская 시트 제목의 'на DD.MM.YYYY' (Кредиторская 쪽 제목은 갱신이 늦어 안 믿음)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next(ws for ws in wb if "дебитор" in ws.title.lower())
    for row in ws.iter_rows(max_row=3):
        for c in row:
            mt = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(c.value or ""))
            if mt:
                return dt.date(int(mt[3]), int(mt[2]), int(mt[1]))
    return None


def _key(acct, name):
    n = name.upper().replace("ООО", "").replace("(HAS)", "").replace('"', "")
    base = acct.split(".")[0] if acct else ""
    for w in WORDS:
        if w in n:
            extra = "ПРЕТЕНЗ" if "ПРЕТЕНЗ" in n else ("АРЕНДА" if "АРЕНД" in n and "КВЦ" in n else "")
            return (acct if acct and acct.startswith(("58", "66", "67", "76", "60.2")) else base, w, extra)
    return (acct, n[:20], "")


def _fmt(v):
    return f"{v:,.2f}".rstrip("0").rstrip(".") if isinstance(v, float) and v != int(v) else f"{v:,.0f}"


def fill_arap(tmpl: str, raw_path: str, base: dt.date, rate: float, out: str) -> tuple[list[str], list[str]]:
    """전월 채권채무 파일 + 1C 잔액 -> 이번 달 파일. (바꾼 칸, 확인 사항) 반환."""
    tag = f"{base:%y.%m}"
    rows_raw = read_raw(raw_path)
    wb = openpyxl.load_workbook(tmpl)
    ws = wb[R_SHEET]
    log, notes = [], []
    # 전월 Rawdata 행 (합계행 전까지)
    last_old = R_FIRST
    while ws.cell(last_old + 1, 2).value in ("AR", "AP"):
        last_old += 1
    prev_tag = str(wb[SUM_SHEET]["L2"].value)[:10]
    prev = {}
    for r in range(R_FIRST, last_old + 1):
        pv = ws.cell(r, 5).value
        prev[_key(str(ws.cell(r, 3).value), str(ws.cell(r, 4).value))] = (ws.cell(r, 7).value, ws.cell(r, 8).value,
                                                                           None if isinstance(pv, str) else _num(pv))  # 수식이면 비교 안 함
    rows = []
    for kind, acct, name, rub, eur in rows_raw:
        gkey, note, prub = prev.get(_key(acct, name), (None, None, None))
        if gkey is None:
            gkey = "관리대상 외"
            note = "잔액 없음" if rub == 0 else f"{tag} 신규 행 - 3.채권/4.채무 반영 여부 판단 필요"
            if rub:
                notes.append(f"Rawdata 신규: {kind} {acct} {name} {_fmt(rub)} RUB -> '관리대상 외'로 둠 (판단 필요)")
        else:
            b = re.split(r" \(\d\d\.\d\d |\s-\s\d\d\.\d\d ", note or "")[0].strip() if note else ""
            if prub is not None and abs(prub - rub) > 0.005:
                chg = f"{prev_tag[2:7].replace('-', '.')} {_fmt(prub)} → {_fmt(rub)}"
                if eur is not None:
                    chg += ", EUR 재평가" if acct not in ("67.24",) else f", EUR {_fmt(eur)}"
                note = f"{b} ({chg})" if b else chg
            else:
                note = b or None
        rows.append([kind, acct, name, rub, eur, gkey, note])
    n_old, n_new = last_old - R_FIRST + 1, len(rows)
    last_new = R_FIRST + n_new - 1
    if n_new > n_old:
        ws.insert_rows(last_old + 1, n_new - n_old)
        for r in range(last_old + 1, last_new + 1):
            for c in range(2, 9):
                ws.cell(r, c)._style = copy.copy(ws.cell(last_old, c)._style)
            ws.row_dimensions[r].height = ws.row_dimensions[last_old].height
    elif n_new < n_old:
        ws.delete_rows(last_new + 1, n_old - n_new)
    for i, vals in enumerate(rows):
        for c, v in zip(range(2, 9), vals):
            ws.cell(R_FIRST + i, c).value = v
    tot = last_new + 1
    ws.cell(tot, 5).value = f"=SUM(E{R_FIRST}:E{last_new})"
    ws.cell(tot, 6).value = f"=SUM(F{R_FIRST}:F{last_new})"
    for r in range(R_FIRST, R_FIRST + 20):  # 우측 대사표 SUMIFS 범위
        v = ws.cell(r, 11).value
        if isinstance(v, str) and "$6:$" in v:
            ws.cell(r, 11).value = re.sub(r"\$([EGB])\$6:\$\1\$\d+", lambda mt: f"${mt[1]}$6:${mt[1]}${last_new}", v)
    log.append(f"{R_SHEET} {n_new}행")
    # 60.21 YSG EUR 잔액 -> 4.채무 Diff 행 참조, 대사표 환산차이 비고
    r6021 = next((R_FIRST + i for i, x in enumerate(rows) if x[1] == "60.21"), None)
    if r6021 and rows[r6021 - R_FIRST][4]:
        book = rows[r6021 - R_FIRST][3] / rows[r6021 - R_FIRST][4]
        for r in range(R_FIRST, R_FIRST + 20):
            if ws.cell(r, 10).value == "AP_YSG_*":
                ws.cell(r, 14).value = f"Rawdata RUB는 장부환율(≈{book:.2f}), 상세는 EUR×CBR 환율({rate}) → 환산차이만 발생"
    ap = wb[AP_SHEET]
    for r in range(3, ap.max_row + 1):
        v = ap.cell(r, 7).value
        if isinstance(v, str) and R_SHEET in v and r6021:
            ap.cell(r, 7).value = re.sub(r"\$F\$\d+", f"$F${r6021}", v)
    # 3.채권 / 4.채무 연동 칸
    def find(kind, acct, word):
        for k_, a, n, rub, eur, *_ in rows:
            if k_ == kind and a == acct and word.upper() in n.upper():
                return rub, eur
        return None
    for sheet, cell, kind, acct, word, cur in CELLMAP:
        hit = find(kind, acct, word)
        if hit is None:
            notes.append(f"{sheet}!{cell}: Rawdata에 {acct} {word} 행이 없어 그대로 둠")
            continue
        new = hit[0] if cur == "rub" else hit[1]
        c = wb[sheet][cell]
        old = _num(c.value) if not (isinstance(c.value, str) and c.value.startswith("=")) else None
        if old is None:
            notes.append(f"{sheet}!{cell} 가 수식이라 Rawdata 값 {_fmt(new)} 을 못 넣음")
            continue
        if abs(old - new) > 0.005:
            c.value = new
            log.append(f"{sheet}!{cell}")
            n = wb[sheet].cell(c.row, 14)
            n.value = f"{tag} Rawdata {acct}: {_fmt(old)} → {_fmt(new)}" + (" (EUR 재평가)" if cur == "rub" and hit[1] is not None else "")
    for sheet in (AP_SHEET,):
        wb[sheet]["K1"] = rate
    wb[SUM_SHEET]["L2"] = dt.datetime(base.year, base.month, base.day)
    log += [f"{AP_SHEET}!K1", f"{SUM_SHEET}!L2"]
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return log, notes
