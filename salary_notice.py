"""급여환율 공지 메일 3통 만들기와 회사 메일 서버(본인 회사 주소) 발송.

- 형식은 2026-10-19 공지 메일 그대로 (수신, 참조, 제목, 문구, 서명). 바뀌는 건 날짜, 월, 환율뿐
- build(): 3통 만들기 -> notice_queue로 제어판 대기열에 올림 -> 비밀번호 승인 후 _send_company()로 발송
"""
from __future__ import annotations

import datetime as dt
import json
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from html import escape

CONFIG_ID = "1Oqe2oY_H1MJDVrk4pGIh83a4ei95u8ZQ"  # 구글드라이브 '업무 자동화/notice_config.json'
_CFG: dict | None = None


def config() -> dict:
    """보내는 사람, 수신자/참조, 서명은 저장소(공개)가 아니라 드라이브 설정 파일에 둔다."""
    global _CFG
    if _CFG is None:
        if os.environ.get("NOTICE_CONFIG_JSON"):  # 검증 테스트용
            _CFG = json.loads(os.environ["NOTICE_CONFIG_JSON"])
        else:
            import fx_job
            r = fx_job.drive().get(f"{fx_job.API}/{CONFIG_ID}", params={"alt": "media", "supportsAllDrives": "true"})
            r.raise_for_status()
            _CFG = json.loads(r.content.decode("utf-8"))
    return _CFG


P = 'style="margin:0;font-family:\'맑은 고딕\',\'Malgun Gothic\',sans-serif;font-size:11pt"'
TD = "border:1px solid #000;padding:4px 10px;text-align:center;font-family:'맑은 고딕','Malgun Gothic',sans-serif"


def _fmt4(v: str) -> str:
    return f"{float(v):.4f}"


def _eur_table(months, vals, final, label3) -> str:
    W = "width:95px;min-width:95px;height:27px"  # 원래 메일 간격: 칸 95px, 세 행 높이 같게 24px
    h = "".join(f'<td style="{TD};font-size:10pt;{W}">{escape(m)}</td>' for m in months[:2] + [label3])
    v = "".join(f'<td style="{TD};font-size:10pt">{_fmt4(x)}</td>' for x in vals)
    return (f'<table style="border-collapse:collapse;margin-left:30px">'
            f'<tr><td style="{TD};font-size:10pt;{W}">EUR/RUB</td>{h}</tr>'
            f'<tr><td style="{TD};font-size:10pt;height:27px">월 평균</td>{v}</tr>'
            f'<tr><td style="{TD};font-size:9pt;font-weight:bold;background:#FFFF00;height:27px">급여환율</td>'
            f'<td colspan="3" style="{TD};font-size:10pt;font-weight:bold;background:#FFFF00">{escape(final)}</td></tr></table>')


def _avg_table(pair, yy, months, vals, final, label3) -> str:
    W = "width:100px;min-width:100px;height:40px;padding:2px 10px;line-height:1.2"  # 원래 메일 간격: 칸 100px, 1행(두 줄)과 2행 높이 같게 40px
    h = "".join(f'<td style="{TD};font-size:10.5pt;{W}">{m}</td>' for m in [escape(months[0]), escape(months[1]), label3])
    v = "".join(f'<td style="{TD};font-size:10.5pt">{_fmt4(x)}</td>' for x in vals)
    thick = "border:2px solid #000"
    return (f'<p {P}><span style="font-size:10pt">({yy:02d}년)</span></p>'
            f'<table style="border-collapse:collapse">'
            f'<tr><td style="{TD};font-size:10.5pt;{W}">{pair}</td>{h}<td style="{TD};{thick};font-size:10.5pt;{W}">평균</td></tr>'
            f'<tr><td style="{TD};font-size:10.5pt;height:40px;padding:2px 10px">월 평균</td>{v}<td style="{TD};{thick};font-size:10.5pt;font-weight:bold">{escape(final)}</td></tr></table>')


def build(rates: dict, rate_day: dt.date, pay_day: dt.date) -> list[dict]:
    """rates: salary_job.salary_rates() 결과. 3통의 메일 정보(dict) 목록."""
    yy = pay_day.year % 100
    months = [f"{int(m[5:])}월" for m in rates["월"]]
    m3 = rate_day.month
    eur_label = f"{m3}월(~{rate_day.day})"
    avg_label = f"{m3}월<br>({m3}/1~{rate_day.day})"
    pay_txt = f"{pay_day.month}/{pay_day.day}"
    c = config()
    me, SIGNATURE = escape(c["sender_name"]), c["signature"]
    GROUP, HAS, CHOI = c["group"], c["has"], c["choi"]
    eur = rates["EUR"]
    group_body = f"""<p {P}>러시아 내 한국 임직원에 대한 ‘{yy}년도 급여 환율을 아래와 같이 송부 드리니 업무에 참고 부탁드립니다.</p>
<p {P}>&nbsp;</p><p {P} align="center">-&nbsp;&nbsp;&nbsp;&nbsp;아&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;래&nbsp;&nbsp;&nbsp;&nbsp;-</p><p {P}>&nbsp;</p>
<ol style="margin-top:0;font-family:'맑은 고딕','Malgun Gothic',sans-serif;font-size:11pt">
<li>적용 시점: {yy}년 {pay_day.month}월 {pay_day.day}일 급여(당월 포함 3개월 평균 환율)</li>
<li>적용 법인: Gefest SP / YSKD / Elias Rus</li>
<li>적용 환율: €1= {escape(eur['급여환율'])}RUB (소수점 절상)</li>
<li>환율 기준: <a href="https://www.x-rates.com/average/?from=EUR&amp;to=RUB&amp;amount=1&amp;year={rate_day.year}">https://www.x-rates.com/average/?from=EUR&amp;to=RUB&amp;amount=1&amp;year={rate_day.year}</a></li>
</ol>
{_eur_table(months, eur['월평균'], eur['급여환율'], eur_label)}
<p {P}>&nbsp;</p><p {P}>이상입니다.</p><p {P}>{me} 올림</p><p {P}>&nbsp;</p>{SIGNATURE}"""

    def has_body(table):
        return f"""<p {P}>수신 제위,</p><p {P}>&nbsp;</p><p {P}>안녕하세요, {me}입니다.</p>
<p {P}><span style="font-size:10pt">‘{yy}년 {pay_txt}일 급여 지급분에 대한 환율을 아래와 같이 공유 드리오니, 업무에 참조 해주시기 바랍니다.</span></p>
<p {P}>&nbsp;</p>{table}<p {P}>&nbsp;</p><p {P}>이상입니다.</p><p {P}>{me} 올림.</p><p {P}>&nbsp;</p>{SIGNATURE}"""

    usd, cny = rates["USD"], rates["CNY"]
    return [
        {"name": "그룹 공지 (EUR)", **{k: GROUP[k] for k in ('to', 'cc')}, "subject": f"러시아 내 한국 임직원 '{yy}년도 급여 환율 송부", "html": group_body},
        {"name": "HAS 법인 (USD)", **{k: HAS[k] for k in ('to', 'cc')}, "subject": "HAS 법인 급여환율 공지의 건",
         "html": has_body(_avg_table("USD/RUB", yy, months, usd["월평균"], usd["급여환율"], avg_label))},
        {"name": "CNY 공지", **{k: CHOI[k] for k in ('to', 'cc')}, "subject": "HAS 법인 급여환율 공지의 건",
         "html": has_body(_avg_table("CNY/RUB", yy, months, cny["월평균"], cny["급여환율"], avg_label))},
    ]


def _addr(lst):
    return ", ".join(formataddr((n, a)) if n else a for n, a in lst)


def _addr_html(lst):
    return escape("; ".join(f"'{n}' <{a}>" if n else a for n, a in lst)) or "-"


def _send_company(m: dict, test: bool = False):
    """회사 메일 서버로 1통 발송. test=True면 받는 사람을 무조건 본인(보내는 사람 주소) 하나로 고정
    (헤더뿐 아니라 실제 전송 대상 목록도 본인만), 제목에 [테스트], 본문 맨 위에 원래 받는 사람/참조를 적는다."""
    host = os.environ.get("SMTP_HOST", "mail.youngsan.com")
    port = int(os.environ.get("SMTP_PORT", "25"))
    FROM = config()["from"]
    msg = EmailMessage()
    msg["From"] = FROM
    html = m["html"]
    if test:
        msg["To"] = FROM
        msg["Subject"] = "[테스트] " + m["subject"]
        note = (f'<p style="margin:0 0 10px;padding:8px;border:1px dashed #c00;color:#c00;font-size:10pt">'
                f'[테스트 발송 - 본인에게만] 실제 발송 시 받는 사람: {escape(_addr(m["to"]))}'
                f' / 참조: {escape(_addr(m["cc"])) or "-"}</p>')
        html = note + html
        to_addrs = [FROM]
    else:
        msg["To"] = _addr(m["to"])
        if m["cc"]:
            msg["Cc"] = _addr(m["cc"])
        msg["Subject"] = m["subject"]
        to_addrs = [a for _, a in m["to"]] + [a for _, a in m["cc"]]
    msg["Message-ID"] = make_msgid(domain="youngsan.com")
    msg.set_content("HTML 메일입니다.")
    msg.add_alternative(f"<html><body>{html}</body></html>", subtype="html")
    with smtplib.SMTP(host, port, timeout=60) as s:  # Outlook 설정과 같은 서버, 포트, 인증
        s.ehlo()
        if s.has_extn("starttls"):
            import ssl
            s.starttls(context=ssl.create_default_context())
            s.ehlo()
        s.login(os.environ["POP_USER"], os.environ["POP_PASS"])
        s.send_message(msg, to_addrs=to_addrs)
