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
            sel = "div.moduleContent:has(h1.monthlyaverageH1)"
            el = pg.locator(sel).first
            el.scroll_into_view_if_needed()
            # 숫자 부분만: 제목(통화) + 월별 평균, 일수 목록. 그래프 자리는 뺌
            title = pg.locator(f"{sel} h1.monthlyaverageH1").first.bounding_box()
            lst = pg.locator(f"{sel} ul.OutputLinksAvg").first.bounding_box()
            box = el.bounding_box()
            pad = 8
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
