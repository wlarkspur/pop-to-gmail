# pop-to-gmail

회사 POP3 메일함의 새 메일을 30분마다 Gmail 받은편지함으로 복사합니다. PC가 꺼져 있어도 GitHub Actions에서 돌아갑니다.

- 회사 서버 메일은 지우지 않습니다 (Outlook과 같이 써도 안전)
- 원본을 그대로 넣기 때문에 보낸 사람, 첨부, 날짜가 원본과 같습니다
- Gmail에 `회사메일` 라벨이 붙습니다
- 같은 메일은 두 번 넣지 않습니다 (Message-ID 기준)
- 로그에는 건수만 남습니다

## 설정 (Settings > Secrets and variables > Actions > New repository secret)

| 이름 | 값 |
|---|---|
| `POP_HOST` | mail.youngsan.com |
| `POP_PORT` | 110 (비워두면 110) |
| `POP_USER` | 회사 메일 로그인 아이디 |
| `POP_PASS` | 회사 메일 비밀번호 |
| `GMAIL_USER` | 받을 Gmail 주소 |
| `GMAIL_APP_PASSWORD` | Gmail 앱 비밀번호 16자리 |

## 실행

- 자동: 30분마다
- 수동: Actions 탭 > `회사메일 → Gmail` > Run workflow
- 실패하면 GitHub가 알림 메일을 보냅니다

## 첫 실행

최근 48시간 메일만 가져옵니다 (`MAX_AGE_HOURS`로 조절). 그 이전 메일은 건너뜁니다.

## 테스트

```
python3 test_pop2gmail.py
```
