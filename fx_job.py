"""매월 1~3일 실행: 구글드라이브 '환율' 폴더의 환율!.xls에 지난달 CBR EUR 일별 환율을 채워 새 버전으로 올린다.

- 이미 채워져 있으면 아무것도 안 함 (1~3일 중 처음 성공한 날만 반영)
- 드라이브는 덮어쓰기 전 버전을 자동 보관 (파일 > 버전 관리에서 되돌리기 가능)
- 결과는 Gmail로 보고 (월평균, 말일 환율, 기존 값과 CBR 차이)

필요한 Secrets: GDRIVE_SA_JSON (서비스 계정 키 JSON 전체), GMAIL_USER, GMAIL_APP_PASSWORD
"""

from __future__ import annotations

import json
import os
import smtplib
import sys
import tempfile
from email.message import EmailMessage

from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession

import fx_cbr

FOLDER_ID = os.environ.get("FX_FOLDER_ID", "1EWFs9zymtT-anJZebJHLdxFfZEoLFJNy")
FILE_NAME = os.environ.get("FX_FILE_NAME", "환율!.xls")
API = "https://www.googleapis.com/drive/v3/files"


def drive():
    info = json.loads(os.environ["GDRIVE_SA_JSON"])
    cred = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive"])
    return AuthorizedSession(cred)


def find_file(s) -> str:
    q = f"'{FOLDER_ID}' in parents and name = '{FILE_NAME}' and trashed = false"
    r = s.get(API, params={"q": q, "fields": "files(id,name)", "supportsAllDrives": "true",
                           "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    files = r.json()["files"]
    if len(files) != 1:
        raise RuntimeError(f"'{FILE_NAME}' 파일이 {len(files)}개 - 1개여야 함 (폴더 공유 확인)")
    return files[0]["id"]


def mail(subject: str, body: str):
    user, pw = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD")
    if not (user and pw):
        return
    m = EmailMessage()
    m["From"] = m["To"] = user
    m["Subject"] = subject
    m.set_content(body)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(user, pw.replace(" ", ""))
        smtp.send_message(m)


def main() -> int:
    month = sys.argv[1] if len(sys.argv) > 1 else None
    s = drive()
    fid = find_file(s)
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = f"{tmp}/in.xls", f"{tmp}/out.xls"
        r = s.get(f"{API}/{fid}", params={"alt": "media", "supportsAllDrives": "true"})
        r.raise_for_status()
        open(src, "wb").write(r.content)
        if month:
            y, m = map(int, month.split("-"))
        else:
            import datetime as dt
            t = dt.date.today().replace(day=1) - dt.timedelta(days=1)
            y, m = t.year, t.month
        res = fx_cbr.run(src, dst, y, m)
        print("채운 칸:", res["채운 칸"])  # 공개 저장소 실행 기록이라 상세는 결과 메일로만
        if res["채운 칸"] == 0 and not res["새로 만든 시트"] and not res["월간 보고서 환율"].get("바꿈"):
            print("이미 채워져 있음 - 올리지 않음")
            return 0
        up = s.patch(f"https://www.googleapis.com/upload/drive/v3/files/{fid}",
                     params={"uploadType": "media", "supportsAllDrives": "true"},
                     data=open(dst, "rb").read(), headers={"Content-Type": "application/vnd.ms-excel"})
        up.raise_for_status()
    lines = [f"{res['월']} CBR EUR 일별 환율을 '{FILE_NAME}' {res['시트']} 시트에 채웠습니다.",
             f"- 채운 칸: {res['채운 칸']}일",
             f"- 월평균(일수 평균): {res['월평균(일수 평균)']}",
             f"- 말일 환율: {res['말일 환율']}",
             "- 평균, 3개월 평균 등 수식은 엑셀에서 열 때 다시 계산됩니다.",
             "- 이전 버전은 드라이브 > 파일 > 버전 관리에 남아 있습니다."]
    rb = res["월간 보고서 환율"]
    lines.append(f"- 월간 보고서 환율: {rb['상태']}" + (f" ({', '.join(str(m) + '월' for m in rb['월'])})" if rb.get("월") else ""))
    if res["새로 만든 시트"]:
        lines.append(f"- 새 시트 생성: {', '.join(res['새로 만든 시트'])} (전년 시트 복사, 날짜 +1년, 환율과 수기 숫자는 비움). 엑셀에서 한 번 열어 확인해 주세요.")
    if res["CBR과 다른 기존 값"]:
        lines.append("- 기존에 적혀 있던 값 중 CBR과 다른 것 (그대로 둠): " + "; ".join(res["CBR과 다른 기존 값"]))
    mail(f"[환율] {res['월']} CBR EUR 반영 완료 (월평균 {res['월평균(일수 평균)']:.4f})", "\n".join(lines))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        mail("[환율] 자동 반영 실패", f"오류: {e}\n\nGitHub Actions 로그를 확인하세요.")
        raise
