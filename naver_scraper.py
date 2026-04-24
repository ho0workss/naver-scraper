"""
naver_scraper.py — 네이버 스마트스토어/브랜드스토어 경쟁사 상품 정보 수집

수집 항목:
    - 상품명
    - 등재가     (소비자가 / 정가)
    - 즉시할인가  (현재 판매가)
    - 쿠폰가     (다운로드 쿠폰 적용가, 미로그인 시 미표시 가능)
    - 적립금     (N페이 포인트 등)
    - 관심고객수  (찜 수)
    - 리뷰 수 / 평점

수집 전략 (우선순위):
    1. 네트워크 인터셉트  — API JSON 응답에서 구조화 데이터 추출 (가장 안정적)
    2. JSON-LD           — <script type="application/ld+json"> 파싱
    3. JS 전역 상태       — window.__PRELOADED_STATE__ 추출
    4. DOM 선택자        — 폴백, 클래스명이 배포마다 바뀔 수 있음

사용법:
    pip install playwright pandas openpyxl
    playwright install chromium
    python src/naver_scraper.py
    python src/naver_scraper.py --output result.xlsx --delay 4 --visible

주의:
    - 쿠폰가는 네이버 로그인 없이 수집 시 일부 미표시될 수 있음.
    - --delay 를 너무 낮추면 IP 차단 위험. 기본 3초 권장.
    - 멀티시트 상품(옵션 있는 상품)은 기본 옵션 기준 가격 수집.
"""

import argparse
import asyncio
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    Response,
    TimeoutError as PWTimeout,
    async_playwright,
)

# ── 경쟁사 상품 목록 ─────────────────────────────────────────────────────────
PRODUCTS: list[tuple[str, str]] = [
    ("스텐팟",          "https://brand.naver.com/stenpot/products/7215155280"),
    ("스테나",          "https://brand.naver.com/stena/products/9205111057"),
    ("조지루시 4L 신형", "https://smartstore.naver.com/smilelab/products/11001858893"),
    ("조지루시 4L 구형", "https://smartstore.naver.com/krbmall/products/9256095682"),
    ("한일 가열식",      "https://smartstore.naver.com/hanilshop/products/11033431188"),
    ("다룸",            "https://brand.naver.com/da_room/products/10908055537"),
    ("끌리젠",          "https://brand.naver.com/kklizen/products/11062212950"),
    ("롯데 가전",        "https://smartstore.naver.com/bfkr/products/11214521788"),
    ("에디르 가열식",    "https://brand.naver.com/edir/products/11012835013"),
    ("웰포트",          "https://brand.naver.com/wellpot24/products/12404396176"),
    ("에어메이드 가열식","https://brand.naver.com/airmade/products/11129902190"),
    ("스테나 하토",      "https://brand.naver.com/stena/products/12996768969"),
    ("르젠 퓨어",       "https://brand.naver.com/lezen/products/7406607504"),
    ("르젠 숨시내",     "https://smartstore.naver.com/chun/products/9199391132"),
    ("케어팟",          "https://brand.naver.com/carepod/products/9182523514"),
    ("에어메이드",      "https://brand.naver.com/airmade/products/10843229255"),
    ("에디르 복합식",   "https://brand.naver.com/edir/products/12427283875"),
    ("에디르 초음파",   "https://brand.naver.com/edir/products/12390164692"),
    ("엔퍼센트",        "https://brand.naver.com/npercent/products/9240525458"),
    ("에디르",          "https://brand.naver.com/edir/products/7286491267"),
    ("가야",            "https://smartstore.naver.com/eg/products/4886752011"),
    ("에어메이드(2)",   "https://brand.naver.com/airmade/products/8592490353"),
    ("한양테크",        "https://smartstore.naver.com/hysp/products/9999620135"),
    ("신일",            "https://smartstore.naver.com/sh-mall/products/8178325903"),
    ("네이처플",        "https://smartstore.naver.com/natureple/products/8771208358"),
    ("한솔",            "https://smartstore.naver.com/healthgoodcare365/products/5510477386"),
    ("윈세이",          "https://smartstore.naver.com/winsay/products/8848644972"),
    ("파센느",          "https://brand.naver.com/passene/products/11750987501"),
    ("엔와이샵",        "https://smartstore.naver.com/daon12/products/6618214226"),
    ("보국",            "https://brand.naver.com/bokuk/products/10302502125"),
    ("신일(공식)",      "https://smartstore.naver.com/shinilofficialstore/products/10150114086"),
    ("끌리젠(2)",       "https://brand.naver.com/kklizen/products/10058839033"),
    ("시즘",            "https://smartstore.naver.com/sizm/products/11754768389"),
    ("매직쉐프z",       "https://smartstore.naver.com/illusionyoutong/products/11819211038"),
    ("매직쉐프2구",     "https://smartstore.naver.com/coscompany/products/11868744314"),
    ("딜팩토리",        "https://brand.naver.com/dealfactory/products/8359347791"),
    ("비비엔다",        "https://brand.naver.com/vivienda/products/5601069502"),
    ("포몽드",          "https://brand.naver.com/formongde/products/11987843269"),
    ("포그망(돗투돗)",  "https://smartstore.naver.com/dtdglobal/products/7538154062"),
    ("꾸망",            "https://smartstore.naver.com/ggumang/products/10258940779"),
    ("시온",            "https://smartstore.naver.com/_sion/products/10217384508"),
    ("꼼띠아",          "https://smartstore.naver.com/comtia/products/5489273018"),
    ("하우스터",        "https://smartstore.naver.com/hauster/products/4952800841"),
    ("릴리브",          "https://brand.naver.com/relievlab/products/12645818553"),
    ("아이닉",          "https://brand.naver.com/inic/products/11387540771"),
    ("폴레드",          "https://brand.naver.com/poled/products/8816502347"),
    ("에코따숨",        "https://brand.naver.com/ecotasum/products/4743358704"),
    ("웰비오",          "https://brand.naver.com/wellvio/products/7236387055"),
    ("가디브",          "https://smartstore.naver.com/guadiv/products/8190843770"),
    ("더나은",          "https://smartstore.naver.com/goatddang/products/4868964440"),
    ("디메디",          "https://brand.naver.com/dimedikorea/products/4967931561"),
    ("아리핏",          "https://smartstore.naver.com/unboxingtop/products/5209004384"),
    ("센시아",          "https://smartstore.naver.com/cotin/products/6219464808"),
    ("한일",            "https://smartstore.naver.com/markette/products/6396198476"),
    ("한경희",          "https://smartstore.naver.com/180donet/products/8677078481"),
    ("디스크랩",        "https://smartstore.naver.com/barnshop/products/2466977391"),
    ("리빙선생",        "https://smartstore.naver.com/clearmart/products/8294125046"),
    ("더마울트라",      "https://smartstore.naver.com/haveagoodfarm/products/2208339960"),
    ("시카케어",        "https://smartstore.naver.com/haveagoodfarm/products/2403358459"),
    ("힐텀스트레치",    "https://smartstore.naver.com/venkod/products/11493068418"),
]

# 네이버 스마트스토어/브랜드스토어 내부 API 응답 URL 패턴
_API_PATTERNS = (
    "/i/v1/products/",
    "/i/v2/products/",
    "/i/v1/channels/",
    "/i/v2/channels/",
    "smartstore.naver.com/i/",
    "brand.naver.com/i/",
)

# 가격 관련 JSON 키 우선순위
_PRICE_KEYS = [
    "discountedSalePrice",   # 할인 후 최종가
    "salePrice",             # 판매가
    "immediateDiscountPrice",
    "benefitPrice",
]
_ORIGIN_PRICE_KEYS = [
    "consumerPrice",         # 소비자가(등재가)
    "originalPrice",
    "retailPrice",
]
_WISH_KEYS = ["wishCount", "favoriteCount", "interestCount", "likeCount"]
_COUPON_KEYS = ["couponPrice", "couponDiscountPrice", "bestCouponPrice"]
_POINT_KEYS = ["accumulationAmount", "pointAmount", "rewardAmount", "mileageAmount"]


# ── 유틸 ────────────────────────────────────────────────────────────────────

def _to_int(value) -> int | None:
    if value is None:
        return None
    cleaned = re.sub(r"[^\d]", "", str(value))
    return int(cleaned) if cleaned else None


def _deep_search(data, keys: list[str]) -> int | None:
    """중첩 dict/list에서 첫 번째로 발견되는 키 값을 반환"""
    if isinstance(data, dict):
        for k in keys:
            if k in data and data[k] is not None:
                return _to_int(data[k])
        for v in data.values():
            result = _deep_search(v, keys)
            if result is not None:
                return result
    elif isinstance(data, list):
        for item in data:
            result = _deep_search(item, keys)
            if result is not None:
                return result
    return None


# ── 추출 함수들 ──────────────────────────────────────────────────────────────

def _extract_from_api_json(api_data: dict, record: dict) -> None:
    """API 응답 JSON에서 필드 추출 (네트워크 인터셉트 결과)"""
    if record["즉시할인가"] is None:
        record["즉시할인가"] = _deep_search(api_data, _PRICE_KEYS)
    if record["등재가"] is None:
        record["등재가"] = _deep_search(api_data, _ORIGIN_PRICE_KEYS)
    if record["관심고객수"] is None:
        record["관심고객수"] = _deep_search(api_data, _WISH_KEYS)
    if record["쿠폰가"] is None:
        record["쿠폰가"] = _deep_search(api_data, _COUPON_KEYS)
    if record["적립금"] is None:
        record["적립금"] = _deep_search(api_data, _POINT_KEYS)

    # 상품명
    if record["상품명"] is None:
        for k in ("name", "productName", "title"):
            val = _deep_search(api_data, [k])
            if val and isinstance(val, str) and len(val) > 2:
                record["상품명"] = str(api_data.get(k, val))
                break


async def _extract_json_ld(page: Page, record: dict) -> None:
    """JSON-LD 구조화 데이터 파싱"""
    try:
        scripts = await page.query_selector_all('script[type="application/ld+json"]')
        for script in scripts:
            try:
                data = json.loads(await script.inner_html())
                if isinstance(data, list):
                    data = next((d for d in data if d.get("@type") == "Product"), None)
                if not data or data.get("@type") != "Product":
                    continue
                if record["상품명"] is None:
                    record["상품명"] = data.get("name")
                rating = data.get("aggregateRating", {})
                if record["리뷰수"] is None:
                    record["리뷰수"] = _to_int(rating.get("reviewCount"))
                if record["평점"] is None:
                    record["평점"] = rating.get("ratingValue")
                offers = data.get("offers", {})
                if isinstance(offers, list):
                    offers = offers[0] if offers else {}
                if record["즉시할인가"] is None and offers.get("price"):
                    record["즉시할인가"] = _to_int(str(offers["price"]))
                break
            except (json.JSONDecodeError, Exception):
                pass
    except Exception:
        pass


async def _extract_js_state(page: Page, record: dict) -> None:
    """window.__PRELOADED_STATE__ 탐색"""
    try:
        state = await page.evaluate("""
            () => {
                for (const key of ['__PRELOADED_STATE__', '__INITIAL_STATE__', '__STATE__']) {
                    if (window[key] && typeof window[key] === 'object') {
                        return JSON.parse(JSON.stringify(window[key]));
                    }
                }
                return null;
            }
        """)
        if state:
            _extract_from_api_json(state, record)
    except Exception:
        pass


async def _extract_dom(page: Page, record: dict) -> None:
    """DOM 선택자 폴백 (클래스명은 배포마다 변경될 수 있음)"""

    async def try_selectors(selectors: list[str]) -> str | None:
        for sel in selectors:
            try:
                el = await page.query_selector(sel)
                if el:
                    text = (await el.inner_text()).strip()
                    if text:
                        return text
            except Exception:
                pass
        return None

    # 즉시할인가 — 현재 굵게 표시된 주요 판매가
    if record["즉시할인가"] is None:
        text = await try_selectors([
            "strong._2GtSVsvhLM",
            "strong[class*='salePrice']",
            "[class*='price_area'] strong",
            "._3hnF8DXwkE strong",
            "[class*='PriceReview'] strong",
        ])
        if text:
            record["즉시할인가"] = _to_int(text)

    # 등재가 — 취소선 표시 원가
    if record["등재가"] is None:
        text = await try_selectors([
            "del._1LY7DqCnwR",
            "del[class*='price']",
            "[class*='consumer'] del",
            "span[class*='origin'] del",
        ])
        if text:
            record["등재가"] = _to_int(text)

    # 쿠폰가
    if record["쿠폰가"] is None:
        text = await try_selectors([
            "[class*='coupon'] [class*='price'] strong",
            "[class*='couponPrice']",
            "[class*='CouponArea'] strong",
        ])
        if text:
            record["쿠폰가"] = _to_int(text)

    # 적립금
    if record["적립금"] is None:
        text = await try_selectors([
            "[class*='benefit'] [class*='point'] strong",
            "[class*='npay'] [class*='point']",
            "[class*='AccumulationInfo'] strong",
        ])
        if text:
            record["적립금"] = _to_int(text)

    # 관심고객수 (찜)
    if record["관심고객수"] is None:
        text = await try_selectors([
            "button[class*='wish'] em",
            "button[class*='Wish'] em",
            "[class*='wishCount']",
            "[class*='wish_count']",
            "[class*='interest'] em",
        ])
        if text:
            record["관심고객수"] = _to_int(text)

    # 리뷰 수
    if record["리뷰수"] is None:
        text = await try_selectors([
            "[class*='review'] [class*='count']",
            "a[href*='review'] em",
            "[class*='ReviewCount']",
        ])
        if text:
            record["리뷰수"] = _to_int(text)


# ── 메인 스크래핑 로직 ────────────────────────────────────────────────────────

async def scrape_product(page: Page, brand: str, url: str, retries: int = 2) -> dict:
    record: dict = {
        "브랜드":    brand,
        "URL":       url,
        "수집시각":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "상품명":    None,
        "등재가":    None,
        "즉시할인가": None,
        "쿠폰가":    None,
        "적립금":    None,
        "관심고객수": None,
        "리뷰수":    None,
        "평점":      None,
        "오류":      None,
    }

    captured_api: list[dict] = []

    async def on_response(resp: Response) -> None:
        if resp.status != 200:
            return
        if not any(p in resp.url for p in _API_PATTERNS):
            return
        if "json" not in resp.headers.get("content-type", ""):
            return
        try:
            captured_api.append(await resp.json())
        except Exception:
            pass

    page.on("response", on_response)

    last_error: str | None = None
    for attempt in range(1, retries + 2):
        captured_api.clear()
        try:
            try:
                await page.goto(url, wait_until="networkidle", timeout=35_000)
            except PWTimeout:
                await page.goto(url, wait_until="domcontentloaded", timeout=35_000)

            await page.wait_for_timeout(2_500)

            for api_data in captured_api:
                _extract_from_api_json(api_data, record)
            await _extract_json_ld(page, record)
            await _extract_js_state(page, record)
            await _extract_dom(page, record)

            record["오류"] = None
            break  # 성공 시 루프 탈출

        except PWTimeout:
            last_error = "타임아웃"
        except Exception as e:
            last_error = f"{type(e).__name__}: {str(e)[:100]}"

        if attempt <= retries:
            wait = 5 * attempt
            print(f"(재시도 {attempt}/{retries}, {wait}초 대기) ", end="", flush=True)
            await asyncio.sleep(wait)
    else:
        record["오류"] = last_error

    page.remove_listener("response", on_response)
    return record


# ── 실행 ────────────────────────────────────────────────────────────────────

async def run(output_path: Path, delay: float, visible: bool) -> None:
    async with async_playwright() as pw:
        browser: Browser = await pw.chromium.launch(
            headless=not visible,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
            ],
        )
        context: BrowserContext = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
            locale="ko-KR",
            timezone_id="Asia/Seoul",
        )
        # webdriver 플래그 숨김
        await context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
        )

        page: Page = await context.new_page()
        records: list[dict] = []

        print(f"총 {len(PRODUCTS)}개 상품 수집 시작 (딜레이 {delay}초)\n")

        for i, (brand, url) in enumerate(PRODUCTS, 1):
            print(f"[{i:02d}/{len(PRODUCTS):02d}] {brand} ", end="", flush=True)
            record = await scrape_product(page, brand, url)
            records.append(record)

            price = record["즉시할인가"]
            wish  = record["관심고객수"]
            err   = record["오류"]
            if err:
                print(f"→ 오류: {err}")
            else:
                print(f"→ 판매가: {price:,}원" if price else f"→ 가격 미수집  찜: {wish}")

            if i < len(PRODUCTS):
                await asyncio.sleep(delay)

        await browser.close()

    # ── Excel 저장 ──────────────────────────────────────────────────────────
    df = pd.DataFrame(records, columns=[
        "브랜드", "상품명", "등재가", "즉시할인가", "쿠폰가",
        "적립금", "관심고객수", "리뷰수", "평점", "수집시각", "URL", "오류",
    ])

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="경쟁사 가격")
        ws = writer.sheets["경쟁사 가격"]
        for col_cells in ws.columns:
            length = max(len(str(c.value or "")) for c in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = min(length, 60)

    ok    = sum(1 for r in records if not r["오류"])
    fail  = len(records) - ok
    print(f"\n{'─'*50}")
    print(f"저장: {output_path}")
    print(f"성공: {ok}개  실패: {fail}개  합계: {len(records)}개")


def main() -> None:
    parser = argparse.ArgumentParser(description="네이버 경쟁사 상품 정보 수집")
    parser.add_argument(
        "--output",
        default="",
        help="출력 xlsx 경로 (기본: 스크립트 위치에 competitor_YYYYMMDD_HHMMSS.xlsx)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=3.0,
        help="상품 간 수집 간격(초), 기본 3.0 (너무 낮추면 IP 제한 위험)",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="브라우저를 보이는 상태로 실행 (디버깅용)",
    )
    args = parser.parse_args()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = args.output or str(Path(__file__).parent.parent / f"competitor_{ts}.xlsx")
    output_path = Path(out).resolve()

    asyncio.run(run(output_path, args.delay, args.visible))


if __name__ == "__main__":
    main()
