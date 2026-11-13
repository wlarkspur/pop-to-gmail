"""외부 예약 서비스(cron-job.org)가 5분마다 부르는 '시계'. 칼리닌그라드 시각을 보고 지금 할 작업만 실행 신호를 보낸다.

GitHub 자체 예약(schedule)은 몇 시간씩 밀려서, 정시 실행은 이 파일이 맡고 GitHub 예약은 예비로 남겨 둔다.
각 작업은 이미 했으면 그냥 끝나게 되어 있어서 두 번 불려도 문제없다.
"""
from __future__ import annotations

import datetime as dt
import os
import sys

import requests

import ru_calendar

KLD = dt.timezone(dt.timedelta(hours=2))


def due(now: dt.datetime) -> list[str]:
    h, m, d = now.hour, now.minute, now.day
    slot = m % 30 < 5  # 정시, 30분 (5분 간격 호출이 조금 늦어도 한 번만 걸리게)
    jobs = []
    if 7 <= h < 21:
        jobs.append("notice_send.yml")  # 제어판 승인 확인: 5분마다
    if slot:
        jobs.append("pop2gmail.yml")  # 회사 메일 가져오기: 30분마다
    if slot and (8 <= h < 12 or (h == 12 and m < 5)) and salary_due(now.date()):
        jobs.append("salary_fx.yml")  # 급여환율: 기준일에만 08:00~12:00 30분마다
    if d <= 3 and h == 6 and m < 5:
        jobs.append("fx_cbr.yml")  # CBR 일별 환율: 매월 1~3일 06:00
    if d <= 3 and h == 8 and m < 5:
        jobs.append("meal.yml")  # 식대 시트: 매월 1~3일 08:00
    if d <= 5 and h == 7 and m < 5:
        jobs.append("daily.yml")  # 가마감용 Daily Report: 매월 1~5일 07:00 (YSKD가 늦게 오는 편, 이미 있으면 건너뜀)
    return jobs


def salary_due(today: dt.date) -> bool:
    """오늘이 급여환율 기준일인가. 저장소 달력(ru_holidays.json)만 보고 판단, 그 해 달력이 없으면 작업 쪽에서 판단하게 True."""
    nxt = today + dt.timedelta(days=31)
    if not (ru_calendar.has_year(today.year) and ru_calendar.has_year(nxt.year)):
        return True
    return any(rate == today for y, m in {(today.year, today.month), (nxt.year, nxt.month)}
               for _, _, rate in ru_calendar.schedule(y, m))


def main() -> int:
    now = dt.datetime.now(KLD)
    jobs = due(now)
    repo, token = os.environ["GITHUB_REPOSITORY"], os.environ["GITHUB_TOKEN"]
    for w in jobs:
        r = requests.post(f"https://api.github.com/repos/{repo}/actions/workflows/{w}/dispatches",
                          headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"},
                          json={"ref": "main"}, timeout=30)
        print(w, r.status_code)
        r.raise_for_status()
    print(f"{now:%m/%d %H:%M} 칼리닌그라드: {', '.join(jobs) or '할 작업 없음'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
