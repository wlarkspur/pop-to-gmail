"""제어판 발송 안전장치 검증: 가짜 드라이브 + 가짜 SMTP로 process()를 돌려, 어떤 경우에 회사 메일 서버에 접속하는지 센다.

실행: python3 test_notice_queue.py
- 테스트 항목은 비밀번호가 맞으면 3통 모두 본인 주소(보내는 사람) 하나에게만 (헤더, 실제 전송 대상 둘 다)
- 비밀번호가 틀리면 SMTP 접속 0회
- 대조군: 진짜 항목 + 맞는 비밀번호면 3통 발송 (가짜 서버가 실제로 잡아낸다는 증명)
"""
import copy
import hashlib
import json
import os
import smtplib

os.environ["NOTICE_PASSWORD"] = "pw-정답"
ME = "me@example.com"
os.environ["NOTICE_CONFIG_JSON"] = json.dumps({  # 실제 수신자는 드라이브 설정 파일에만 있음. 검증은 가짜 주소로
    "from": ME, "sender_name": "보내는사람", "signature": "<p>서명</p>",
    "group": {"to": [["", "a1@example.com"], ["", "a2@example.com"]], "cc": [["", "c1@example.com"]]},
    "has": {"to": [["", "b1@example.com"]], "cc": [["", "c2@example.com"]]},
    "choi": {"to": [["", "d1@example.com"]], "cc": []}})
os.environ["POP_USER"], os.environ["POP_PASS"] = "u", "p"

import fx_job  # noqa: E402
import notice_queue as nq  # noqa: E402
import salary_notice  # noqa: E402

SMTP_LOG = []


class FakeSMTP:
    def __init__(self, host, port, timeout=None):
        SMTP_LOG.append(("connect", host, port))

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def ehlo(self):
        pass

    def has_extn(self, x):
        return False

    def login(self, u, p):
        SMTP_LOG.append(("login", u))

    def send_message(self, msg, to_addrs=None):
        SMTP_LOG.append(("send", msg["To"], msg["Subject"], tuple(to_addrs or ()), msg["Cc"]))


smtplib.SMTP = FakeSMTP  # 회사 메일 서버 대신 기록만 하는 가짜


class Resp:
    def __init__(self, content=b"", js=None):
        self.content, self._js = content, js

    def raise_for_status(self):
        pass

    def json(self):
        return self._js


class FakeDrive:
    def __init__(self, queue, confirms, cancels=()):
        self.queue, self.confirms, self.cancels = queue, confirms, list(cancels)

    def get(self, url, params=None):
        if url.endswith(nq.QUEUE_ID):
            return Resp(json.dumps(self.queue).encode())
        if "q" in (params or {}):
            if "cancel_" in params["q"]:
                return Resp(js={"files": [{"id": n, "name": n, "createdTime": n} for n in self.cancels]})
            return Resp(js={"files": [{"id": cid, "name": f"confirm_{c['key']}_{c['ts']}.json", "createdTime": cid}
                                      for cid, c in self.confirms.items()]})
        cid = url.rsplit("/", 1)[1]
        return Resp(json.dumps(self.confirms[cid]).encode())

    def patch(self, url, params=None, data=None, headers=None):
        self.queue = json.loads(data)
        return Resp()


def confirm(key, password, ts="1793630000000"):
    return {"key": key, "ts": ts, "hash": hashlib.sha256(f"{password}|{key}|{ts}".encode()).hexdigest()}


MAILS = salary_notice.build(
    {"월": ["2026-09", "2026-10", "2026-11"],
     "EUR": {"월평균": ["97.803116", "93.087456", "90.048853"], "3개월 평균": "93.646475", "급여환율": "94"},
     "USD": {"월평균": ["84.905856", "83.000432", "80.614262"], "3개월 평균": "82.840183", "급여환율": "83"},
     "CNY": {"월평균": ["12.6533", "12.4105", "12.1141"], "3개월 평균": "12.392633", "급여환율": "12.40"}},
    __import__("datetime").date(2026, 11, 2), __import__("datetime").date(2026, 11, 10))


def run(key, test, password, cancel_ts=None):
    SMTP_LOG.clear()
    q = {"notices": [{"key": key, "title": key, "created": "-", "status": "pending", "test": test,
                      "summary": [], "mails": copy.deepcopy(MAILS), "history": []}], "processed": []}
    drive = FakeDrive(q, {"c1": confirm(key, password)}, [f"cancel_{key}_{cancel_ts}.json"] if cancel_ts else [])
    fx_job.drive = lambda: drive
    fx_job.mail = lambda *a: None
    nq.process()
    return drive.queue["notices"][0]["status"], [x for x in SMTP_LOG if x[0] == "connect"], [x for x in SMTP_LOG if x[0] == "send"]


cases = [
    ("테스트 항목(test=true, -test) + 맞는 비밀번호", "salary-20261110-test", True, "pw-정답"),
    ("test 표시 없이 이름만 -test + 맞는 비밀번호", "salary-20261110-test", False, "pw-정답"),
    ("test=true 표시만 + 맞는 비밀번호", "salary-20261110", True, "pw-정답"),
    ("진짜 항목 + 틀린 비밀번호", "salary-20261110", False, "pw-오답"),
    ("대조군: 진짜 항목 + 맞는 비밀번호", "salary-20261110", False, "pw-정답"),
]
for name, key, test, pw in cases:
    status, conns, sends = run(key, test, pw)
    print(f"{name:<40} -> 상태 {status:<15} 회사 SMTP 접속 {len(conns)}회, 발송 {len(sends)}통")
    if "대조군" in name:
        assert status == "sent" and len(sends) == 3, "대조군에서 발송이 잡혀야 함"
        for _, to, subj, env, cc in sends:
            print(f"    (가짜 서버가 받은 것) 실제 전송 대상: {', '.join(env)} / 제목: {subj}")
    elif status == "test_ok":
        assert len(sends) == 3
        for _, to, subj, env, cc in sends:
            assert env == (ME,) and to == ME and cc is None and subj.startswith("[테스트]")
            print(f"    (가짜 서버가 받은 것) 실제 전송 대상: {', '.join(env)} / 참조: {cc} / 제목: {subj}")
    else:
        assert not conns and not sends, f"{name}: 발송되면 안 됨"
status, conns, sends = run("salary-20261110", False, "pw-정답", cancel_ts="1793630000500")
print(f"{'진짜 항목 + 맞는 비밀번호 + 승인 취소':<40} -> 상태 {status:<15} 회사 SMTP 접속 {len(conns)}회, 발송 {len(sends)}통")
assert status == "pending" and not conns, "취소한 승인은 발송되면 안 됨"
status, conns, sends = run("salary-20261110", False, "pw-정답", cancel_ts="1793629999000")
assert status == "sent", "승인보다 먼저 한 취소는 새 승인을 막으면 안 됨"
print("모든 검증 통과: 테스트는 본인 주소로만, 틀린 비밀번호는 서버 접속 없음, 진짜 항목만 원래 수신자에게")
