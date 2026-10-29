"""가짜 POP 서버 + 가짜 Gmail로 전체 흐름을 검증한다. 실행: python3 test_pop2gmail.py"""

import os
import socketserver
import threading
from datetime import datetime, timedelta, timezone
from email.utils import format_datetime

import pop2gmail as p

NOW = datetime.now(timezone.utc)


def make_mail(i, hours_ago, mid="default", attach=False, received_hours_ago=None, big=False):
    lines = []
    if received_hours_ago is not None:
        lines.append(f"Received: from mx by mail.youngsan.com; "
                     f"{format_datetime(NOW - timedelta(hours=received_hours_ago))}")
    lines += [
        "From: financial@hyauto.ru",
        "To: someone@hyundaicorp.com",
        "Cc: smjeong@youngsan.com",
        f"Subject: test {i}",
        f"Date: {format_datetime(NOW - timedelta(hours=hours_ago))}",
    ]
    if mid == "default":
        lines.append(f"Message-ID: <msg{i}@hyauto.ru>")
    elif mid is not None:
        lines.append(mid)
    if attach:
        lines += ["MIME-Version: 1.0", 'Content-Type: multipart/mixed; boundary="B"', "", "--B",
                  "Content-Type: text/plain", "", "body", ".leading dot line", "--B",
                  'Content-Type: application/vnd.ms-excel; name="r.xlsx"',
                  "Content-Transfer-Encoding: base64", "", "UEsDBAo=", "--B--"]
    else:
        lines += ["", f"body {i}" + (" BIG" if big else ""), ".leading dot line"]
    return "\r\n".join(lines).encode("utf-8") + b"\r\n"


MAILS = [
    make_mail(1, 200),                                         # 오래됨 -> 건너뜀
    make_mail(2, 5),                                           # 가져옴
    make_mail(3, 1, mid=None),                                 # Message-ID 없음 -> 생성
    make_mail(4, 0.5, attach=True),                            # 첨부
    make_mail(5, 300, received_hours_ago=1),                   # Date는 옛날, 도착은 최근 -> 가져옴
    make_mail(6, 1, mid="Message-ID: <한글@example.com>"),     # 비ASCII ID -> 생성 ID로 대체
    make_mail(7, 1, mid="Message-ID:\r\n <folded@x.ru>"),      # 접힌 ID
    make_mail(8, 1, big=True),                                 # Gmail 저장 실패 -> 다음 메일은 계속
    make_mail(9, 0.1),                                         # 마지막 정상 메일
]
DELETED = []


class FakePOP(socketserver.StreamRequestHandler):
    def send(self, s):
        self.wfile.write(s if isinstance(s, bytes) else s.encode())

    def multiline(self, data: bytes):
        self.send(b"+OK\r\n")
        for line in data.split(b"\r\n")[:-1]:
            self.send((b"." + line if line.startswith(b".") else line) + b"\r\n")
        self.send(b".\r\n")

    def handle(self):
        self.send("+OK fake\r\n")
        for raw in self.rfile:
            cmd = raw.decode().strip().split()
            op = cmd[0].upper()
            if op == "CAPA":
                self.multiline(b"TOP\r\nUIDL\r\nUSER\r\n")
            elif op in ("USER", "PASS"):
                self.send("-ERR auth\r\n" if op == "PASS" and cmd[1] != "pw" else "+OK\r\n")
            elif op == "STAT":
                self.send(f"+OK {len(MAILS)} 100\r\n")
            elif op == "TOP":
                m = MAILS[int(cmd[1]) - 1]
                self.multiline(m[: m.find(b"\r\n\r\n") + 2])
            elif op == "RETR":
                self.multiline(MAILS[int(cmd[1]) - 1])
            elif op in ("DELE", "RSET"):
                DELETED.append(cmd)
                self.send("+OK\r\n")
            elif op == "QUIT":
                self.send("+OK bye\r\n")
                return
            else:
                self.send("-ERR\r\n")


class FakeGmail:
    def __init__(self):
        self.store = []  # (raw, label)
        self.selected = []

    def list(self):
        return "OK", [b'(\\HasNoChildren \\All) "/" "[Gmail]/&yATMtLz0rQDVaA-"',
                      b'(\\HasNoChildren \\Trash) "/" "[Gmail]/Trash"',
                      b'(\\HasNoChildren \\Junk) "/" "[Gmail]/Spam"',
                      b'(\\HasNoChildren) "/" "INBOX"']

    def select(self, box, readonly=False):
        self.selected.append(box)
        return "OK", [b"1"]

    def uid(self, cmd, *args):
        if cmd == "SEARCH":
            q = args[1]
            assert "\r" not in q and "\n" not in q
            q.encode("ascii")  # 비ASCII면 imaplib가 죽으므로 여기서도 확인
            mid = q.split("rfc822msgid:")[1].rstrip('"')
            hit = [str(i).encode() for i, (r, _) in enumerate(self.store, 1) if mid.encode() in r]
            return "OK", [b" ".join(hit)]
        if cmd == "STORE":
            i = int(args[0])
            self.store[i - 1] = (self.store[i - 1][0], args[2])
            return "OK", [b""]

    def append(self, box, flags, date, raw):
        if b" BIG" in raw:
            return "NO", [b"too big"]
        self.store.append((raw, None))
        return "OK", [f"[APPENDUID 1 {len(self.store)}] done".encode()]

    def logout(self):
        pass


def main(reverse=False):
    if reverse:
        MAILS.reverse()  # 최신 메일이 1번인 서버
    server = socketserver.ThreadingTCPServer(("127.0.0.1", 0), FakePOP)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    fake = FakeGmail()
    p.gmail_connect = lambda u, pw: fake
    os.environ.update(POP_HOST="127.0.0.1", POP_PORT=str(server.server_address[1]), POP_USER="u",
                      POP_PASS="pw", GMAIL_USER="g", GMAIL_APP_PASSWORD="abcd efgh", POP_SSL="")

    # 1회차: 메일 8 저장 실패 -> 실패 코드지만 나머지는 다 들어가야 함
    assert p.run() == 1
    raws = [r for r, _ in fake.store]
    subjects = [next(s for s in (f"test {i}" for i in range(1, 10)) if f"Subject: {s}\r".encode() in r)
                for r in raws]
    assert sorted(subjects) == ["test 2", "test 3", "test 4", "test 5", "test 6", "test 7", "test 9"], subjects
    assert subjects[0] == "test 2" and subjects[-1] == "test 9", "도착 순서대로 넣지 않음"
    raws = {s: r for s, r in zip(subjects, raws)}
    assert not DELETED, "서버 메일을 지우거나 되돌리면 안 됨"
    assert raws["test 3"].startswith(b"Message-ID: <") and raws["test 6"].startswith(b"Message-ID: <"), "ID 생성 안 됨"
    assert b"\r\n.leading dot line" in raws["test 2"], "dot-stuffing 복원 오류"
    assert b'name="r.xlsx"' in raws["test 4"] and b"UEsDBAo=" in raws["test 4"], "첨부 손실"
    assert b"From: financial@hyauto.ru" in raws["test 2"], "보낸 사람 손실"
    assert all(lbl == '"&1ozArLpUx3w-"' for _, lbl in fake.store), "라벨 인코딩 오류"
    assert {'"[Gmail]/Trash"', '"[Gmail]/Spam"'} <= set(fake.selected), "휴지통/스팸 확인 안 함"

    # 2회차: 중복 없이 0건 (실패한 8번만 다시 시도)
    assert p.run() == 1
    assert len(fake.store) == 7, "중복으로 다시 가져옴"

    # 비밀번호 틀림 -> 실패 코드
    os.environ["POP_PASS"] = "bad"
    assert p.run() == 1
    print("모든 테스트 통과")


if __name__ == "__main__":
    import sys
    main(reverse="--reverse" in sys.argv)
