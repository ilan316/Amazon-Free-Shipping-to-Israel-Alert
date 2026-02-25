"""
Amazon Free Shipping to Israel Checker
Uses Playwright (headed Chromium).

Architecture:
  - Location is set ONCE via setup_location_once() → saved in browser_profile
  - Regular checks (check_all_products) just navigate to each product and read
    the shipping block — no location modal interaction per check
  - The persistent browser profile remembers the Israel location across sessions

Confirmed Amazon free-shipping text format (date part changes):
  "FREE delivery Thursday, March 5 to Israel on eligible orders over $49"

Detection: all three fixed strings must appear in the delivery block:
  - "free delivery"
  - "to israel"
  - "eligible orders"
"""

import sys
import os

# Ensure Playwright finds Chromium regardless of how the app is launched.
# When frozen (PyInstaller exe), prefer a bundled browsers/ dir next to the exe;
# otherwise fall back to the standard ms-playwright location in LOCALAPPDATA.
if getattr(sys, "frozen", False):
    _base = os.path.dirname(sys.executable)
    _browsers = os.path.join(_base, "browsers")
    if os.path.isdir(_browsers):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _browsers
    else:
        # No bundled browsers — point to the system-installed ms-playwright dir
        _ms_pw = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _ms_pw
else:
    # Not frozen: still ensure we use the standard location, not the package dir
    if "PLAYWRIGHT_BROWSERS_PATH" not in os.environ:
        _ms_pw = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
        if os.path.isdir(_ms_pw):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _ms_pw

import asyncio
import random
import re
import logging
from dataclasses import dataclass
from enum import Enum

from playwright.async_api import (
    async_playwright,
    Page,
    TimeoutError as PWTimeout,
)

logger = logging.getLogger(__name__)

# ------------------------------------------------------------------
# Data types
# ------------------------------------------------------------------

class ShippingStatus(Enum):
    FREE    = "FREE"
    PAID    = "PAID"
    NO_SHIP = "NO_SHIP"
    UNKNOWN = "UNKNOWN"
    ERROR   = "ERROR"


@dataclass
class CheckResult:
    asin: str
    status: ShippingStatus
    raw_text: str = ""
    error_message: str = ""
    product_name: str = ""


# ------------------------------------------------------------------
# Selectors
# ------------------------------------------------------------------

DELIVER_TO_SELECTORS = [
    "#nav-global-location-popover-link",
    "#glow-ingress-line2",
    "#glow-ingress-line1",
]

COUNTRY_DROPDOWN_SELECTORS = [
    "#GLUXCountryList",
    "select.a-native-dropdown[name='countryCode']",
]

ZIP_INPUT_SELECTORS = [
    "#GLUXZipUpdateInput",
    "input[placeholder='ZIP Code']",
    "input[name='zipCode']",
]

# "Apply" button — for ZIP code submission only
ZIP_APPLY_BTN_SELECTORS = [
    "#GLUXZipUpdate input",
    "span[data-action='GLUXZipUpdate'] input",
    "span[data-action='GLUXSaveChanges'] input",
    "#GLUXSaveBtn input",
]

# "Done" button — after selecting a country from the dropdown
DONE_BTN_SELECTORS = [
    "#GLUXConfirmClose input",
    "span[data-action='GLUXConfirmClose'] input",
    "input.a-button-input[aria-labelledby*='GLUXConfirm']",
    ".a-popover-footer input.a-button-input",
    "#GLUXSaveBtn input",
    "input[value='Done']",
    "input[value='Save']",
    "text=Done",
    "text=Save",
]

DELIVERY_BLOCK_SELECTORS = [
    "#mir-layout-DELIVERY_BLOCK",
    "#ddmDeliveryMessage",
    "#deliveryMessageMirId",
    "#delivery-message",
    "#price-shipping-message",
    "#exports_feature_div",
    "#shippingMessageInsideBuyBox_feature_div",
    "#buybox",
    "#buyBoxInner",
]

CAPTCHA_SELECTORS = [
    "form[action='/errors/validateCaptcha']",
    "input#captchacharacters",
]

# "See All Buying Options" button — appears when there's no direct Add-to-Cart
SEE_ALL_BUYING_SELECTORS = [
    "#buybox-see-all-buying-choices a",
    "#buybox-see-all-buying-choices",
    "a#aod-ingress-link",
    ".a-button-buybox a",
    "text=See All Buying Options",
]

# All Offers Display (AOD) panel that slides in after clicking the button
AOD_OFFER_SELECTORS = [
    "#aod-offer-list",
    "#aod-container",
    "#aod-offer",
]

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

async def _pause(min_s: float, max_s: float):
    await asyncio.sleep(random.uniform(min_s, max_s))


async def _type_human(element, text: str):
    for ch in text:
        await element.type(ch, delay=random.randint(60, 160))


async def _first(page: Page, selectors: list, timeout: int = 4000):
    """Returns the first matching element, or None."""
    for sel in selectors:
        try:
            el = await page.wait_for_selector(sel, timeout=timeout, state="attached")
            if el:
                return el
        except PWTimeout:
            continue
    return None


async def _is_captcha(page: Page) -> bool:
    for sel in CAPTCHA_SELECTORS:
        try:
            if await page.query_selector(sel):
                return True
        except Exception:
            pass
    title = (await page.title()).lower()
    return "robot" in title or "captcha" in title or "validateCaptcha" in page.url


# ------------------------------------------------------------------
# Browser context (shared between setup and checks)
# ------------------------------------------------------------------

async def _create_context(config: dict):
    browser_cfg = config.get("browser", {})
    pw = await async_playwright().start()
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=browser_cfg.get("user_data_dir", "browser_profile"),
        headless=browser_cfg.get("headless", False),
        slow_mo=browser_cfg.get("slow_mo_ms", 120),
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ],
        ignore_default_args=["--enable-automation"],
        locale="en-US",
        timezone_id="America/New_York",
        viewport={"width": 1280, "height": 800},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
    )
    await context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    )
    return pw, context


# ------------------------------------------------------------------
# One-time location setup  (called from GUI "Setup Location" button)
# ------------------------------------------------------------------

async def _set_location_on_page(page: Page, country_code: str, zip_code: str) -> bool:
    """
    Opens the Amazon location modal and sets country + zip.
    Returns True if the location was saved successfully.

    Flow:
      1. Click the "Deliver to" nav link
      2. In the modal: select country from "ship outside the US" dropdown → click Done
      3. If no country dropdown: type the zip code → click Apply
    """
    # Open modal
    deliver_btn = await _first(page, DELIVER_TO_SELECTORS, timeout=8000)
    if not deliver_btn:
        logger.warning("Could not find 'Deliver to' button.")
        return False

    await deliver_btn.click()
    await _pause(1.5, 2.5)

    # Wait for the popover to actually appear before interacting
    await _first(page, [
        "#GLUXCountryList",
        "#GLUXZipUpdateInput",
        ".a-popover-content",
        "#nav-global-location-popover-link-heading",
    ], timeout=6000)

    # Select country from dropdown — Amazon applies automatically
    dropdown = await _first(page, COUNTRY_DROPDOWN_SELECTORS, timeout=5000)
    if dropdown:
        try:
            await dropdown.select_option(value=country_code)
            await _pause(2.0, 3.0)
            logger.info(f"Location set to {country_code}.")
            return True
        except Exception as e:
            logger.debug(f"Country dropdown approach failed: {e}")

    logger.warning("Could not set delivery location automatically.")
    return False


async def setup_location_once(config: dict) -> bool:
    """
    One-time setup: opens Amazon, sets the delivery location to Israel,
    then closes the browser. The persistent browser_profile saves the preference.

    Called from the GUI "Setup Location" button.
    """
    delivery = config.get("delivery", {})
    country_code = delivery.get("country_code", "IL")
    zip_code     = delivery.get("zip", "6100000")

    logger.info(f"Setting up delivery location: country={country_code}, zip={zip_code}")
    pw, context = await _create_context(config)
    try:
        page = await context.new_page()
        await page.goto("https://www.amazon.com", wait_until="domcontentloaded", timeout=20000)
        await _pause(2.0, 3.5)

        if await _is_captcha(page):
            logger.error("CAPTCHA on Amazon homepage during setup.")
            return False

        success = await _set_location_on_page(page, country_code, zip_code)
        if success:
            logger.info("Location setup complete — saved to browser profile.")
        return success
    finally:
        try:
            await context.close()
        except Exception as _e:
            logger.warning(f"Browser context close error (ignored): {_e}")
        try:
            await pw.stop()
        except Exception as _e:
            logger.warning(f"Playwright stop error (ignored): {_e}")


# ------------------------------------------------------------------
# Shipping text classification
# ------------------------------------------------------------------

def _classify(text: str) -> ShippingStatus:
    """
    FREE condition — all three fixed strings must appear (date is ignored):
      "FREE delivery [DATE] to Israel on eligible orders over $49"
    """
    t = text.lower()

    if "free delivery" in t and "to israel" in t and "eligible orders" in t:
        return ShippingStatus.FREE

    no_ship = [
        "doesn't ship to israel",
        "does not ship to israel",
        "cannot be shipped to israel",
        "not available for shipping to israel",
        "this item does not ship to your selected location",
        "item can't be shipped to your selected delivery location",
        "this item cannot be shipped to your selected delivery location",
    ]
    if any(p in t for p in no_ship):
        return ShippingStatus.NO_SHIP

    paid_pat = re.compile(
        r'\$[\d]+\.[\d]{2}.{0,40}israel|israel.{0,40}\$[\d]+\.[\d]{2}',
        re.IGNORECASE,
    )
    if paid_pat.search(text) and "israel" in t:
        return ShippingStatus.PAID

    if "israel" in t:
        return ShippingStatus.UNKNOWN

    return ShippingStatus.UNKNOWN


# ------------------------------------------------------------------
# Per-product check  (no location modal — profile already set)
# ------------------------------------------------------------------

async def _read_delivery_text(page: Page) -> str:
    for sel in DELIVERY_BLOCK_SELECTORS:
        try:
            el = await page.wait_for_selector(sel, timeout=3000, state="attached")
            if el:
                text = await el.inner_text()
                if text.strip():
                    return text.strip()
        except (PWTimeout, Exception):
            continue
    return ""


async def _check_all_buying_options(page: Page, asin: str) -> str:
    """
    If a 'See All Buying Options' button is present, click it to open the
    AOD (All Offers Display) panel and return the combined offer text.
    Returns '' if the button is absent or the panel fails to load.
    """
    btn = await _first(page, SEE_ALL_BUYING_SELECTORS, timeout=3000)
    if not btn:
        return ""
    try:
        await btn.click()
        await _pause(2.0, 3.0)
        # Use state="visible" — the container attaches to the DOM before offers
        # are actually rendered, so "attached" reads empty text too early.
        for sel in AOD_OFFER_SELECTORS:
            try:
                el = await page.wait_for_selector(sel, timeout=8000, state="visible")
                if el:
                    await _pause(1.0, 1.5)   # let offer rows finish rendering
                    text = await el.inner_text()
                    if text.strip():
                        logger.info(f"[{asin}] AOD panel: {text[:120]!r}")
                        return text.strip()
            except PWTimeout:
                continue
    except Exception as exc:
        logger.debug(f"[{asin}] All Buying Options check failed: {exc}")
    return ""


async def check_product(page: Page, asin: str, url: str) -> CheckResult:
    """
    Navigates to a product page and reads the delivery block.
    If the main page shows no free-shipping text but has a
    'See All Buying Options' button, checks the offers panel too.
    Does NOT touch the location modal — location is assumed to be pre-set
    in the browser profile via setup_location_once().
    """
    try:
        await page.goto(f"{url}?psc=1&th=1", wait_until="domcontentloaded", timeout=30000)
        product_name = ""
        try:
            await page.wait_for_selector("#productTitle", timeout=12000)
            el = await page.query_selector("#productTitle")
            if el:
                product_name = (await el.inner_text()).strip()
        except PWTimeout:
            logger.warning(f"[{asin}] productTitle not found.")

        if await _is_captcha(page):
            return CheckResult(asin, ShippingStatus.ERROR, error_message="CAPTCHA detected",
                               product_name=product_name)

        await _pause(0.8, 1.8)

        raw_text = await _read_delivery_text(page)
        status = _classify(raw_text) if raw_text else ShippingStatus.UNKNOWN

        # If the main page didn't confirm FREE, try "See All Buying Options" panel
        if status != ShippingStatus.FREE:
            aod_text = await _check_all_buying_options(page, asin)
            if aod_text:
                aod_status = _classify(aod_text)
                # The AOD panel shows offers in the context of the already-set
                # Israel location, so "to Israel" may not appear explicitly.
                # Relaxed check: "free delivery" + "eligible orders" is enough.
                if aod_status != ShippingStatus.FREE:
                    t = aod_text.lower()
                    if "free delivery" in t and "eligible orders" in t:
                        aod_status = ShippingStatus.FREE
                logger.info(f"[{asin}] AOD status: {aod_status.value}")
                # Use AOD result if it's FREE, or if the main page had nothing
                if aod_status == ShippingStatus.FREE or not raw_text:
                    raw_text = aod_text
                    status = aod_status

        if not raw_text:
            logger.warning(f"[{asin}] No delivery text found.")
            return CheckResult(asin, ShippingStatus.UNKNOWN, error_message="No delivery block found",
                               product_name=product_name)

        logger.info(f"[{asin}] {status.value} | {raw_text[:120]!r}")
        return CheckResult(asin, status, raw_text=raw_text, product_name=product_name)

    except PWTimeout as e:
        return CheckResult(asin, ShippingStatus.ERROR, error_message=f"Timeout: {e}")
    except Exception as e:
        return CheckResult(asin, ShippingStatus.ERROR, error_message=str(e))


# ------------------------------------------------------------------
# Main entry point for scheduled checks
# ------------------------------------------------------------------

async def check_all_products(config: dict, state: dict) -> list:
    """
    Checks all products sequentially.
    Sets the delivery location to Israel at the start of every check session
    so that Amazon always shows Israel-specific shipping information.
    """
    delivery     = config.get("delivery", {})
    country_code = delivery.get("country_code", "IL")
    zip_code     = delivery.get("zip", "6100000")

    pw, context = await _create_context(config)
    results = []

    try:
        page = await context.new_page()

        logger.info("Opening Amazon...")
        await page.goto("https://www.amazon.com", wait_until="domcontentloaded", timeout=20000)
        await _pause(1.5, 3.0)

        if await _is_captcha(page):
            logger.warning("CAPTCHA on Amazon homepage — skipping location setup.")
        else:
            logger.info("Setting delivery location to Israel...")
            await _set_location_on_page(page, country_code, zip_code)
            await _pause(1.5, 2.5)

        products = config.get("products", [])
        for i, product in enumerate(products):
            asin = product["asin"]
            name = product.get("name", asin)
            logger.info(f"[{i+1}/{len(products)}] {name}")

            result = await check_product(page, asin, product["url"])
            results.append(result)

            if i < len(products) - 1:
                await _pause(5.0, 12.0)

    finally:
        try:
            await context.close()
        except Exception as _e:
            logger.warning(f"Browser context close error (ignored): {_e}")
        try:
            await pw.stop()
        except Exception as _e:
            logger.warning(f"Playwright stop error (ignored): {_e}")

    return results
