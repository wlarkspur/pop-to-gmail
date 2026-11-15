"""배포 후 텔레그램에 웹훅 주소, 명령 목록을 등록 (bot_deploy.yml에서 실행). 토큰은 출력하지 않음."""
import json
import os
import sys
import urllib.request

T = os.environ["TG_TOKEN"]


def call(method: str, body: dict):
    req = urllib.request.Request(f"https://api.telegram.org/bot{T}/{method}", data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    j = json.load(urllib.request.urlopen(req, timeout=20))
    print(method, j.get("ok"), j.get("description", ""))
    if not j.get("ok"):
        sys.exit(1)


call("setWebhook", {"url": sys.argv[1], "secret_token": os.environ["WEBHOOK_SECRET"],
                    "allowed_updates": ["message", "callback_query"], "drop_pending_updates": True})
call("setMyCommands", {"commands": [
    {"command": "status", "description": "최근 자동 작업 결과"},
    {"command": "gamagam", "description": "가마감 초안 (예: HAS 2026-11)"},
    {"command": "daily", "description": "가마감용 Daily Report"},
    {"command": "cbr", "description": "CBR EUR 환율 채우기"},
    {"command": "salary", "description": "급여환율 수동 실행"},
    {"command": "meal", "description": "식대 시트"},
    {"command": "help", "description": "도움말"}]})
call("sendMessage", {"chat_id": os.environ["TG_CHAT_ID"], "text": "🤖 봇 배포 완료. /help 로 명령어를 볼 수 있어요."})
