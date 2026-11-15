"""텔레그램 알림 (보내기 전용). 시크릿 TG_TOKEN, TG_CHAT_ID가 없으면 아무것도 안 함.
알림은 부가 기능이라 실패해도 작업을 멈추지 않음. 토큰, chat_id는 로그에 찍지 않음.

사용: tg.send("텍스트"), tg.send_file("/path/x.xlsx", "설명")
      python3 tg.py --fail "<작업 이름>" "<실행 URL>"   (워크플로 실패 단계에서)
"""
from __future__ import annotations

import os
import sys

import requests

API = "https://api.telegram.org/bot{}/{}"
LIMIT = 4000


def _cfg():
    t, c = os.environ.get("TG_TOKEN"), os.environ.get("TG_CHAT_ID")
    return (t, c) if t and c else (None, None)


def send(text: str) -> bool:
    t, c = _cfg()
    if not t:
        return False
    try:
        for i in range(0, max(len(text), 1), LIMIT):
            r = requests.post(API.format(t, "sendMessage"), json={"chat_id": c, "text": text[i:i + LIMIT],
                              "disable_web_page_preview": True}, timeout=20)
            r.raise_for_status()
        return True
    except Exception as e:  # noqa: BLE001
        print("텔레그램 전송 실패:", type(e).__name__)
        return False


def send_file(path: str, caption: str = "", name: str | None = None) -> bool:
    t, c = _cfg()
    if not t:
        return False
    try:
        with open(path, "rb") as f:
            r = requests.post(API.format(t, "sendDocument"), data={"chat_id": c, "caption": caption[:1000]},
                              files={"document": (name or os.path.basename(path), f)}, timeout=60)
        r.raise_for_status()
        return True
    except Exception as e:  # noqa: BLE001
        print("텔레그램 파일 전송 실패:", type(e).__name__)
        return False


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--fail":
        job = sys.argv[2] if len(sys.argv) > 2 else "작업"
        url = sys.argv[3] if len(sys.argv) > 3 else ""
        send(f"⚠️ 실패: {job}\n{url}")
    elif len(sys.argv) >= 2:
        send(" ".join(sys.argv[1:]))
