"""POP3 메일함의 새 메일을 Gmail 받은편지함으로 복사한다.

- 회사 서버의 메일은 절대 지우지 않는다 (DELE 안 씀, Outlook과 같이 써도 안전).
- 원본 메일을 IMAP APPEND로 그대로 넣기 때문에 보낸 사람, 첨부, 날짜가 원본과 같다.
- 상태 파일 없이 Gmail(전체보관함, 휴지통, 스팸)에 같은 Message-ID가 있는지 확인해서 중복을 막는다.
- 메일 한 통이 실패해도 나머지는 계속 처리한다.
- 로그에는 건수만 남기고 제목, 주소 같은 내용은 남기지 않는다.

환경변수
  POP_HOST, POP_USER, POP_PASS          회사 메일 서버 접속 정보 (필수)
  POP_PORT=110, POP_SSL=0               포트, SSL 여부 (SSL이 아니면 STARTTLS를 먼저 시도)
  GMAIL_USER, GMAIL_APP_PASSWORD        Gmail 주소와 앱 비밀번호 (필수)
  GMAIL_LABEL=회사메일                  가져온 메일에 붙일 Gmail 라벨 (빈 값이면 안 붙임)
  MAX_SCAN=400                          매번 헤더를 확인할 최대 메일 수 (넘으면 앞뒤 절반씩)
  MAX_AGE_HOURS=48                      서버 도착 시각이 이보다 오래된 메일은 건너뜀 (첫 실행 폭주 방지)
"""

from __future__ import annotations

import base64
import hashlib
import imaplib
import os
import poplib
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from email.parser import BytesHeaderParser
from email.utils import parsedate_to_datetime

poplib._MAXLINE = 4 * 1024 * 1024  # 긴 줄이 있는 메일에서 poplib이 죽지 않게

SAFE_MID = re.compile(r"^<[\x21-\x7e]+>$")  # 공백, 괄호, 따옴표 없는 ASCII만


def env(name: str, default: str | None = None, required: bool = False) -> str:
    value = os.environ.get(name) or default  # 빈 시크릿은 기본값으로
    if required and not value:
        sys.exit(f"[설정 오류] 환경변수 {name} 가 비어 있습니다")
    return value or ""


# ---------- POP ----------

def pop_connect(host: str, port: int, use_ssl: bool, user: str, password: str, timeout: int = 60):
    if use_ssl:
        conn = poplib.POP3_SSL(host, port, timeout=timeout)
    else:
        conn = poplib.POP3(host, port, timeout=timeout)
        try:
            if "STLS" in conn.capa():
                conn.stls()  # 지원하면 암호화해서 비밀번호 전송
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
            conn = poplib.POP3(host, port, timeout=timeout)  # 암호화 실패 시 평문으로 재접속
    conn.user(user)
    conn.pass_(password)
    return conn


def pop_headers(conn, num: int) -> bytes:
    """메일 헤더만 가져온다. TOP 미지원 서버면 전체를 받아 헤더만 자른다."""
    try:
        _, lines, _ = conn.top(num, 0)
        return b"\r\n".join(lines) + b"\r\n\r\n"
    except poplib.error_proto:
        raw = pop_retr(conn, num)
        end = raw.find(b"\r\n\r\n")
        return raw if end < 0 else raw[: end + 4]


def pop_retr(conn, num: int) -> bytes:
    _, lines, _ = conn.retr(num)
    # POP의 dot-stuffing은 poplib이 이미 풀어준다
    return b"\r\n".join(lines) + b"\r\n"


# ---------- 헤더 처리 ----------

def _header(msg, name: str) -> str:
    value = msg.get(name)
    return "" if value is None else str(value)


def message_id_of(headers: bytes) -> tuple[str, bool]:
    """(Gmail 검색에 안전한 Message-ID, 원본 그대로인지).

    원본 ID가 없거나 검색에 쓰기 위험한 형태면, 헤더 내용으로 고정 ID를 만든다.
    같은 메일이면 매번 같은 ID가 나오고, POP 번호가 바뀌어도 영향이 없다.
    """
    msg = BytesHeaderParser().parsebytes(headers)
    mid = "".join(_header(msg, "Message-ID").split())  # 접힌 줄, 공백 제거
    if mid and SAFE_MID.match(mid):
        return mid, True
    basis = "\n".join(
        [mid]
        + [_header(msg, h) for h in ("From", "To", "Date", "Subject")]
        + [str(r) for r in (msg.get_all("Received") or [])[:1]]
    )
    digest = hashlib.sha1(basis.encode("utf-8", "replace")).hexdigest()
    return f"<{digest}@pop2gmail.local>", False


def _parse_date(text: str) -> datetime | None:
    try:
        dt = parsedate_to_datetime(text.strip())
    except (TypeError, ValueError, IndexError):
        return None
    if dt is None:
        return None
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def arrival_of(headers: bytes) -> datetime | None:
    """서버 도착 시각. 맨 위 Received 헤더(마지막 서버)를 우선, 없으면 Date."""
    msg = BytesHeaderParser().parsebytes(headers)
    for received in (msg.get_all("Received") or [])[:1]:
        text = str(received)
        if ";" in text:
            dt = _parse_date(text.rsplit(";", 1)[1])
            if dt:
                return dt
    date = _header(msg, "Date")
    return _parse_date(date) if date else None


def add_message_id(raw: bytes, mid: str) -> bytes:
    return f"Message-ID: {mid}\r\n".encode("ascii") + raw


# ---------- Gmail IMAP ----------

def gmail_connect(user: str, app_password: str):
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(user, app_password)
    return conn


def _imap_quote(text: str) -> str:
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _utf7(text: str) -> str:
    """IMAP 폴더/라벨 이름용 수정 UTF-7 인코딩 (한글 라벨 지원)."""
    out, buf = [], []

    def flush():
        if buf:
            b64 = base64.b64encode("".join(buf).encode("utf-16-be")).decode("ascii")
            out.append("&" + b64.rstrip("=").replace("/", ",") + "-")
            buf.clear()

    for ch in text:
        if 0x20 <= ord(ch) <= 0x7E:
            flush()
            out.append("&-" if ch == "&" else ch)
        else:
            buf.append(ch)
    flush()
    return "".join(out)


def gmail_folders(conn) -> list[str]:
    """중복 확인할 폴더: 전체보관함, 휴지통, 스팸 (언어 설정과 무관하게 플래그로 찾음)."""
    found = {}
    typ, data = conn.list()
    if typ == "OK":
        for line in data or []:
            text = line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line)
            for flag in ("\\All", "\\Trash", "\\Junk"):
                if flag in text and flag not in found:
                    m = re.search(r'"([^"]+)"\s*$', text) or re.search(r"(\S+)\s*$", text)
                    if m:
                        found[flag] = m.group(1)
    found.setdefault("\\All", "[Gmail]/All Mail")
    return [found[f] for f in ("\\All", "\\Trash", "\\Junk") if f in found]


def gmail_existing(conn, folders: list[str], mids: list[str]) -> set[str]:
    """이미 Gmail에 있는 Message-ID 집합."""
    present: set[str] = set()
    for folder in folders:
        typ, _ = conn.select(_imap_quote(folder), readonly=True)
        if typ != "OK":
            continue
        for mid in mids:
            if mid in present:
                continue
            query = _imap_quote(f"rfc822msgid:{mid.strip('<>')}")
            typ, data = conn.uid("SEARCH", "X-GM-RAW", query)
            if typ != "OK":
                raise RuntimeError("Gmail 검색 실패")
            if data and data[0] and data[0].split():
                present.add(mid)
    return present


def gmail_append(conn, raw: bytes, when: datetime | None, label: str) -> bool:
    """INBOX에 넣고 라벨을 붙인다. 라벨까지 성공하면 True."""
    date_arg = imaplib.Time2Internaldate(when.timestamp() if when else time.time())
    typ, data = conn.append("INBOX", None, date_arg, raw)
    if typ != "OK":
        raise RuntimeError("Gmail 저장 실패")
    if not label:
        return True
    m = re.search(rb"APPENDUID \d+ (\d+)", data[0] if data and data[0] else b"")
    if not m:
        return False
    conn.select("INBOX")
    typ, _ = conn.uid("STORE", m.group(1).decode(), "+X-GM-LABELS", _imap_quote(_utf7(label)))
    return typ == "OK"


# ---------- 실행 ----------

def run() -> int:
    pop_host = env("POP_HOST", required=True)
    pop_port = int(env("POP_PORT", "110"))
    pop_ssl = env("POP_SSL", "0") in ("1", "true", "True", "yes")
    pop_user = env("POP_USER", required=True)
    pop_pass = env("POP_PASS", required=True)
    gmail_user = env("GMAIL_USER", required=True)
    gmail_pw = env("GMAIL_APP_PASSWORD", required=True).replace(" ", "")
    label = env("GMAIL_LABEL", "회사메일")
    max_scan = int(env("MAX_SCAN", "400"))
    cutoff = datetime.now(timezone.utc) - timedelta(hours=float(env("MAX_AGE_HOURS", "48")))

    try:
        pop = pop_connect(pop_host, pop_port, pop_ssl, pop_user, pop_pass)
    except poplib.error_proto as e:
        text = str(e).lower()
        if "lock" in text or "in use" in text or "busy" in text:
            print("[건너뜀] 메일함이 다른 프로그램(Outlook 등)에서 사용 중. 다음 실행 때 다시 시도")
            return 0
        print("[오류] 회사 메일 로그인 실패 (아이디/비밀번호 확인)")
        return 1
    except OSError:
        print("[오류] 회사 메일 서버에 접속할 수 없음 (서버 주소, 포트, 방화벽 확인)")
        return 1

    imported = skipped_old = skipped_dup = failed = label_failed = 0
    count = checked = 0
    seen_dates: list[datetime] = []
    gmail = None
    try:
        count, _ = pop.stat()
        # 서버마다 메일 번호 순서(오래된 순/최신 순)가 달라서 순서를 가정하지 않는다
        if count <= max_scan:
            numbers = list(range(1, count + 1))
        else:
            half = max_scan // 2
            numbers = list(range(1, half + 1)) + list(range(count - half + 1, count + 1))
        checked = len(numbers)

        # 1) 후보 고르기: 헤더만 받아서 기간, Message-ID 확인
        candidates = []  # (번호, mid, 원본 ID 여부, 도착 시각)
        for num in numbers:
            try:
                headers = pop_headers(pop, num)
                when = arrival_of(headers)
                if when:
                    seen_dates.append(when)
                if when and when < cutoff:
                    skipped_old += 1
                    continue
                mid, had_mid = message_id_of(headers)
                candidates.append((num, mid, had_mid, when))
            except Exception:
                failed += 1

        # 도착 시각 순으로 넣어서 Gmail에서도 순서가 맞게
        candidates.sort(key=lambda c: c[3] or datetime.now(timezone.utc))
        if candidates:
            gmail = gmail_connect(gmail_user, gmail_pw)
            # 2) 이미 Gmail에 있는지 한 번에 확인 (전체보관함, 휴지통, 스팸)
            present = gmail_existing(gmail, gmail_folders(gmail), [c[1] for c in candidates])
            # 3) 없는 것만 원본 받아서 넣기 (오래된 것부터)
            for num, mid, had_mid, when in candidates:
                if mid in present:
                    skipped_dup += 1
                    continue
                try:
                    raw = pop_retr(pop, num)
                    if not had_mid:
                        raw = add_message_id(raw, mid)
                    if not gmail_append(gmail, raw, when, label):
                        label_failed += 1
                    present.add(mid)
                    imported += 1
                except imaplib.IMAP4.abort:
                    raise  # 연결 끊김은 다음 실행에서 다시
                except Exception:
                    failed += 1
    except imaplib.IMAP4.error:
        print("[오류] Gmail 로그인 또는 처리 실패 (앱 비밀번호 확인)")
        return 1
    except (poplib.error_proto, OSError, RuntimeError) as e:
        print(f"[오류] 처리 중 실패: {type(e).__name__}. 다음 실행 때 다시 시도")
        return 1
    finally:
        if gmail is not None:
            try:
                gmail.logout()
            except Exception:
                pass
        try:
            pop.quit()  # DELE를 안 했으니 서버의 메일은 그대로 남는다
        except Exception:
            pass

    if seen_dates:
        fmt = "%Y-%m-%d %H:%M UTC"
        print(f"[참고] 확인한 메일 도착 시각 범위: {min(seen_dates).astimezone(timezone.utc):{fmt}} ~ {max(seen_dates).astimezone(timezone.utc):{fmt}}")
    print(f"[완료] 서버 메일 {count}통 중 {checked}통 확인 / 새로 가져옴 {imported} / "
          f"이미 있음 {skipped_dup} / 기간 지나 건너뜀 {skipped_old} / 실패 {failed} / 라벨 실패 {label_failed}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(run())
