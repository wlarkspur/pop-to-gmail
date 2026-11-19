"""자동화 제어판 발송 대기열 (구글드라이브 '업무 자동화/발송대기' 폴더).

- queue.json (사용자 소유, 서비스 계정이 내용만 고침): 발송 대기 메일 목록과 처리 기록
- confirm_<key>_<ts>.json (제어판 페이지가 사용자 권한으로 만듦): 발송 승인 요청 {key, ts, hash}
  hash = sha256("<비밀번호>|<key>|<ts>") - 비밀번호 원문은 어디에도 저장하지 않음, 확인은 여기(서버)에서만
- process(): 처리 안 된 승인 요청마다 비밀번호 확인 -> 회사 메일로 발송 -> queue.json에 결과 기록 -> 결과 메일

- 텔레그램 승인(2단계): 봇 웹훅(Cloudflare Worker)이 이 작업의 입력 TG_CONFIRM="key|ts|mac"으로 넘김, 취소는 TG_CANCEL="key|ts".
  저장소가 공개라 실행 입력이 보일 수 있어서 mac = HMAC-SHA256(TG_TOKEN, "notice|비밀번호|key|ts") - 봇 토큰 없이는 비밀번호를 대입해 볼 수 없음.

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
    import tg
    tg.send("\n".join([f"📨 발송 대기: {item['title']}", ""] + list(summary) + ["", f"[{key}]"]),
            buttons=tg.approve_buttons(key, PANEL_URL))
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


def _check_tg(password: str, key: str, ts: str, mac: str) -> bool:
    tok = os.environ.get("TG_TOKEN", "")
    if not tok:
        return False
    want = hmac.new(tok.encode(), f"notice|{password}|{key}|{ts}".encode("utf-8"), hashlib.sha256).hexdigest()
    return hmac.compare_digest(want, (mac or "").lower())


RETRY_MAX = 13  # 연결 오류면 5분 간격(시계)으로 최대 12번 자동 재시도 = 약 1시간 (실행할 때마다 다른 서버가 걸림)


def runner_host() -> str:
    """실행 서버가 어느 클라우드인지 (AWS는 25번 포트를 막음). IP 같은 건 안 남김."""
    try:
        v = open("/sys/class/dmi/id/sys_vendor").read().strip()
    except OSError:
        return "알 수 없음"
    return "AWS" if "Amazon" in v else "Microsoft" if "Microsoft" in v else v or "알 수 없음"


AWS_ALERT = (10, 3)  # 최근 10번 발송 시도 중 AWS가 3번 이상이면 텔레그램 경고 (평소 1번 안팎)


def _track(q: dict, host: str, ok: bool) -> str | None:
    """발송 시도마다 서버 종류 기록. AWS 비율이 커지면 경고 문구 돌려줌 (한 번만)."""
    log = (q.get("runner_log", []) + [f"{host}:{'ok' if ok else 'fail'}"])[-AWS_ALERT[0]:]
    q["runner_log"] = log
    aws = sum(x.startswith("AWS") for x in log)
    if len(log) >= AWS_ALERT[0] and aws >= AWS_ALERT[1]:
        if not q.get("aws_alerted"):
            q["aws_alerted"] = True
            return (f"⚠️ 메일 발송 서버 경고: 최근 {len(log)}번 시도 중 {aws}번이 AWS 서버(25번 포트 막힘)에 걸림. "
                    "GitHub가 AWS 비중을 늘리는 중일 수 있어요. 계속되면 발송 방식을 바꿔야 해요.")
    elif aws < AWS_ALERT[1]:
        q["aws_alerted"] = False
    return None


def _deliver(n: dict) -> str:
    """승인된 진짜 항목 발송. 보낸 통은 sent_idx에 기록해서 다시 시도해도 중복 발송 없음.
    서버 연결 오류(일시적)면 상태를 retry로 두고 다음 실행 때 자동으로 다시 보냄 (비밀번호 다시 안 물음)."""
    import salary_notice
    sent = n.setdefault("sent_idx", [])
    try:
        for i, m in enumerate(n["mails"]):
            if i in sent:
                continue
            salary_notice._send_company(m)
            sent.append(i)
        n["status"] = "sent"
        return f"{len(n['mails'])}통 발송 완료 (본인 회사 주소)" + (f" - 자동 재시도 {n['retries']}번째에 성공 ({runner_host()} 서버)" if n.get("retries") else "")
    except Exception as e:  # noqa: BLE001
        n["retries"] = n.get("retries", 0) + 1
        part = f"{len(sent)}/{len(n['mails'])}통 보낸 뒤"
        if n["retries"] < RETRY_MAX:
            n["status"] = "retry"
            return f"발송 중 오류 ({part}, {runner_host()} 서버): {e} - 5분 뒤 다른 실행 서버로 자동 재시도 ({n['retries']}/{RETRY_MAX - 1})"
        n["status"] = "error"
        return f"발송 중 오류 ({part}, {runner_host()} 서버): {e} - 자동 재시도 끝, 다시 승인하면 안 보낸 통만 보냄"


def process() -> str:
    import salary_notice
    import tg
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
    tg_cancel = os.environ.get("TG_CANCEL", "").strip()
    if tg_cancel:
        try:
            key, ts = tg_cancel.split("|")
            cancels[key] = max(cancels.get(key, 0), int(ts))
            n = next((x for x in q["notices"] if x["key"] == key), None)
            if n:
                n["history"].append(f"{_now()} 텔레그램에서 승인 취소")
                changed = True
        except ValueError:
            pass
    for n in q["notices"]:
        if n.get("status") != "retry":
            continue
        if cancels.get(n["key"], 0) > int(n.get("retry_since", 0)):
            n["status"] = "pending"
            n["history"].append(f"{_now()} 자동 재시도 중 승인 취소 - 발송 멈춤")
            changed = True
            continue
        msg = _deliver(n)
        warn = _track(q, runner_host(), n["status"] == "sent")
        if warn:
            tg.send(warn)
        if n["status"] != "retry" or n["retries"] in (3, 6, 9):  # 성공, 최종 실패, 중간 경과만 기록, 알림
            n["history"].append(f"{_now()} {msg}")
            done.append((n, msg))
        changed = True
    confirms = []  # (처리 id, 승인 내용 dict | None, 파일 id | None)
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
            confirms.append(json.loads(r.content.decode("utf-8")))
        except ValueError:
            continue
    tg_confirm = os.environ.get("TG_CONFIRM", "").strip()
    if tg_confirm:
        try:
            key, ts, h = tg_confirm.split("|")
            if cancels.get(key, 0) <= int(ts):
                confirms.append({"key": key, "ts": ts, "hash": h, "via": "텔레그램"})
                changed = True
        except ValueError:
            pass
    for c in confirms:
        n = next((x for x in q["notices"] if x["key"] == c.get("key")), None)
        if not n:
            continue
        if n.get("status") == "sent":
            n["history"].append(f"{_now()} 이미 발송된 건이라 무시")
            continue
        if not password:
            n["status"], msg = "error", "서버에 비밀번호(NOTICE_PASSWORD)가 설정되지 않음"
        elif not (_check_tg if c.get("via") == "텔레그램" else _check)(password, n["key"], str(c.get("ts")), c.get("hash")):
            n["status"], msg = "wrong_password", "비밀번호가 맞지 않음 - 발송 안 함"
        elif n.get("test") or n["key"].endswith("-test"):  # 테스트: 본인(본인 회사 주소)에게만
            try:
                for m in n["mails"]:
                    salary_notice._send_company(m, test=True)
                n["status"], msg = "test_ok", f"테스트: 비밀번호 확인, {len(n['mails'])}통을 본인(본인 회사 주소)에게만 발송"
            except Exception as e:  # noqa: BLE001
                n["status"], msg = "error", f"테스트 발송 중 오류: {e}"
        else:
            n["retries"] = 0
            import time
            n["retry_since"] = int(time.time() * 1000)
            msg = _deliver(n)
            warn = _track(q, runner_host(), n["status"] == "sent")
            if warn:
                tg.send(warn)
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
