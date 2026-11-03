"""자동화 제어판 발송 대기열 (구글드라이브 '업무 자동화/발송대기' 폴더).

- queue.json (사용자 소유, 서비스 계정이 내용만 고침): 발송 대기 메일 목록과 처리 기록
- confirm_<key>_<ts>.json (제어판 페이지가 사용자 권한으로 만듦): 발송 승인 요청 {key, ts, hash}
  hash = sha256("<비밀번호>|<key>|<ts>") - 비밀번호 원문은 어디에도 저장하지 않음, 확인은 여기(서버)에서만
- process(): 처리 안 된 승인 요청마다 비밀번호 확인 -> 회사 메일로 발송 -> queue.json에 결과 기록 -> 결과 메일

비밀번호는 GitHub Secret NOTICE_PASSWORD. 테스트 항목(test=true)은 3통을 본인(본인 회사 주소)에게만 보냄.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import os
import sys

import fx_job

FOLDER_ID = "1Er1KuWJQuINftS1esmy6HN6eskSUmCs-"
QUEUE_ID = "1zBI43kJMXSB6QvfTXYNEc3jnNEwteRXl"
PANEL_URL = os.environ.get("PANEL_URL", "https://claude.ai/artifact/3tU19JPD6u8MoM6PfPaZWr")
KLD = dt.timezone(dt.timedelta(hours=2))


def _now() -> str:
    return dt.datetime.now(KLD).strftime("%Y-%m-%d %H:%M")


def load(s) -> dict:
    r = s.get(f"{fx_job.API}/{QUEUE_ID}", params={"alt": "media", "supportsAllDrives": "true"})
    r.raise_for_status()
    return json.loads(r.content.decode("utf-8") or "{}") or {"notices": [], "processed": []}


def save(s, q: dict, changed: bool = True):
    if changed:
        q["updated"] = _now()  # 마지막으로 내용이 바뀐 시각
    q["checked"] = _now()  # 서버가 마지막으로 확인한 시각 (매 실행)
    r = s.patch(f"https://www.googleapis.com/upload/drive/v3/files/{QUEUE_ID}",
                params={"uploadType": "media", "supportsAllDrives": "true"},
                data=json.dumps(q, ensure_ascii=False, indent=1).encode("utf-8"),
                headers={"Content-Type": "text/plain; charset=utf-8"})
    r.raise_for_status()


def verifier(password: str, key: str) -> str:
    """제어판이 비밀번호를 바로 확인할 수 있게 하는 값 (PBKDF2, 원문 복원 불가). 실제 발송 승인은 process()가 다시 확인."""
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), f"has-panel|{key}".encode("utf-8"), 200000).hex()


def add_notice(s, key: str, title: str, mails: list[dict], summary: list[str], test: bool = False):
    """발송 대기 항목 추가 (같은 key가 아직 안 보낸 상태면 내용 교체)."""
    q = load(s)
    q.setdefault("notices", [])
    old = next((n for n in q["notices"] if n["key"] == key), None)
    if old and old.get("status") == "sent":
        return False
    item = {"key": key, "title": title, "created": _now(), "status": "pending", "test": test,
            "summary": summary, "mails": mails,
            "verifier": verifier(os.environ["NOTICE_PASSWORD"], key) if os.environ.get("NOTICE_PASSWORD") else None,
            "history": (old or {}).get("history", []) + [f"{_now()} 대기열에 올림"]}
    q["notices"] = [n for n in q["notices"] if n["key"] != key] + [item]
    q["notices"] = q["notices"][-20:]  # 최근 20건만
    save(s, q)
    return True


def _list_confirms(s, prefix: str = "confirm_") -> list[dict]:
    q = f"'{FOLDER_ID}' in parents and trashed = false and name contains '{prefix}'"
    r = s.get(fx_job.API, params={"q": q, "fields": "files(id,name,createdTime)", "pageSize": 100,
                                  "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    return sorted(r.json()["files"], key=lambda f: f["createdTime"])


def _check(password: str, key: str, ts: str, h: str) -> bool:
    want = hashlib.sha256(f"{password}|{key}|{ts}".encode("utf-8")).hexdigest()
    return hmac.compare_digest(want, (h or "").lower())


def process() -> str:
    import salary_notice
    s = fx_job.drive()
    q = load(s)
    q.setdefault("processed", [])
    password = os.environ.get("NOTICE_PASSWORD", "")
    done, changed = [], False
    cancels = {}  # 제어판 '승인 취소': cancel_<key>_<ts>.json -> 그 시각 이전 승인 요청은 보내지 않음
    for f in _list_confirms(s, "cancel_"):
        try:
            key, ts = f["name"][len("cancel_"):-len(".json")].rsplit("_", 1)
            cancels[key] = max(cancels.get(key, 0), int(ts))
        except ValueError:
            pass
    for f in _list_confirms(s):
        if f["id"] in q["processed"]:
            continue
        q["processed"] = (q["processed"] + [f["id"]])[-200:]
        changed = True
        try:
            ckey, cts = f["name"][len("confirm_"):-len(".json")].rsplit("_", 1)
            if cancels.get(ckey, 0) > int(cts):
                n = next((x for x in q["notices"] if x["key"] == ckey), None)
                if n:
                    n["history"].append(f"{_now()} 승인 취소됨 - 발송 안 함")
                continue
        except ValueError:
            pass
        r = s.get(f"{fx_job.API}/{f['id']}", params={"alt": "media", "supportsAllDrives": "true"})
        r.raise_for_status()
        try:
            c = json.loads(r.content.decode("utf-8"))
        except ValueError:
            continue
        n = next((x for x in q["notices"] if x["key"] == c.get("key")), None)
        if not n:
            continue
        if n.get("status") == "sent":
            n["history"].append(f"{_now()} 이미 발송된 건이라 무시")
            continue
        if not password:
            n["status"], msg = "error", "서버에 비밀번호(NOTICE_PASSWORD)가 설정되지 않음"
        elif not _check(password, n["key"], str(c.get("ts")), c.get("hash")):
            n["status"], msg = "wrong_password", "비밀번호가 맞지 않음 - 발송 안 함"
        elif n.get("test") or n["key"].endswith("-test"):  # 테스트: 본인(본인 회사 주소)에게만
            try:
                for m in n["mails"]:
                    salary_notice._send_company(m, test=True)
                n["status"], msg = "test_ok", f"테스트: 비밀번호 확인, {len(n['mails'])}통을 본인(본인 회사 주소)에게만 발송"
            except Exception as e:  # noqa: BLE001
                n["status"], msg = "error", f"테스트 발송 중 오류: {e}"
        else:
            try:
                for m in n["mails"]:
                    salary_notice._send_company(m)
                n["status"], msg = "sent", f"{len(n['mails'])}통 발송 완료 (본인 회사 주소)"
            except Exception as e:  # noqa: BLE001
                n["status"], msg = "error", f"발송 중 오류: {e}"
        n["history"].append(f"{_now()} {msg}")
        done.append((n, msg))
    save(s, q, changed)  # 확인 시각은 매번, 변경 시각은 실제로 바뀐 때만
    for n, msg in done:
        lines = [f"{n['title']}: {msg}", ""]
        if n["status"] == "sent":
            for m in n["mails"]:
                to = ", ".join(a for _, a in m["to"])
                cc = ", ".join(a for _, a in m["cc"])
                lines.append(f"- {m['name']}: {m['subject']} / 받는 사람 {to}" + (f" / 참조 {cc}" if cc else ""))
        if PANEL_URL:
            lines += ["", f"제어판: {PANEL_URL}"]
        tag = {"sent": "발송 완료", "test_ok": "테스트 확인", "wrong_password": "비밀번호 틀림"}.get(n["status"], "오류")
        fx_job.mail(f"[제어판] {n['title']} - {tag}", "\n".join(lines))
    return "; ".join(f"{n['key']}: {msg}" for n, msg in done) or "처리할 승인 요청 없음"


if __name__ == "__main__":
    print(process())
    sys.exit(0)
