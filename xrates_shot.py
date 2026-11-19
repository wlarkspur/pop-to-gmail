"""x-rates 월평균 화면(Monthly Average 상자)을 그대로 캡처한다. 급여 파일 '환율' 시트 증빙용."""
from __future__ import annotations


def _stack(*pngs: bytes) -> bytes:
    import io
    from PIL import Image
    ims = [Image.open(io.BytesIO(p)).convert("RGB") for p in pngs]
    out = Image.new("RGB", (max(i.width for i in ims), sum(i.height for i in ims)), "white")
    y = 0
    for i in ims:
        out.paste(i, (0, y))
        y += i.height
    buf = io.BytesIO()
    out.save(buf, "PNG")
    return buf.getvalue()


def capture(curs=("USD", "CNY"), year: int | None = None) -> dict[str, bytes]:
    import datetime as dt
    from playwright.sync_api import sync_playwright
    year = year or dt.date.today().year
    out = {}
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1280, "height": 1000}, device_scale_factor=1)
        for cur in curs:
            pg.goto(f"https://www.x-rates.com/average/?from={cur}&to=RUB&amount=1&year={year}",
                    wait_until="networkidle", timeout=60000)
            pad = 8
            if pg.locator("td.avgMonth").count():  # 2026-11 개편 후 화면
                box = pg.evaluate("""() => {  // 표 상자 (페이지 좌표)
                    const r = document.querySelector('td.avgMonth').closest('div.rounded-card').getBoundingClientRect();
                    return {x: r.x + scrollX, y: r.y + scrollY, width: r.width, height: r.height};
                }""")
                # 제목 'Monthly Average' + 부제 'Russian Ruble per 1 US Dollar ...' 묶음
                head = pg.evaluate("""() => {
                    const h = [...document.querySelectorAll('h1')].find(e => e.textContent.trim() === 'Monthly Average');
                    const a = h.getBoundingClientRect();  // 제목 위 ~ 부제(바로 다음 요소) 아래
                    const b = (h.nextElementSibling || h).getBoundingClientRect();
                    return {x: a.x + scrollX, y: a.y + scrollY, width: a.width, height: b.bottom - a.top + 14};
                }""")
                top = pg.screenshot(full_page=True, clip={"x": box["x"], "y": head["y"] - pad, "width": box["width"],
                                                          "height": head["height"] + 2 * pad})
                body = pg.screenshot(full_page=True, clip={"x": box["x"] - 1, "y": box["y"] - 1,
                                                           "width": box["width"] + 2, "height": box["height"] + 2})
            else:  # 예전 화면
                sel = "div.moduleContent:has(h1.monthlyaverageH1)"
                el = pg.locator(sel).first
                el.scroll_into_view_if_needed()
                # 숫자 부분만: 제목(통화) + 월별 평균, 일수 목록. 그래프 자리는 뺌
                title = pg.locator(f"{sel} h1.monthlyaverageH1").first.bounding_box()
                lst = pg.locator(f"{sel} ul.OutputLinksAvg").first.bounding_box()
                box = el.bounding_box()
                top = pg.screenshot(clip={"x": box["x"], "y": title["y"] - pad, "width": box["width"],
                                          "height": title["height"] + pad})
                body = pg.screenshot(clip={"x": box["x"], "y": lst["y"] - pad, "width": box["width"],
                                           "height": lst["height"] + 2 * pad})
            out[cur] = _stack(top, body)
        b.close()
    return out


if __name__ == "__main__":
    for cur, png in capture().items():
        open(f"{cur}.png", "wb").write(png)
        print(cur, len(png))
