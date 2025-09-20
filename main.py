"""
Highly-stealth Playwright Price Bot — upgraded

What I changed in this version (high level):
- Advanced human-like interaction: precise mouse movement to element with randomized path and small jitter, click with random offset, randomized scrolling, and realistic typing via page.type.
- Stealth & fingerprint mitigation: randomized viewport, locale, user-agent, extra headers (Accept-Language), and injection of init scripts to remove `navigator.webdriver` and emulate some browser properties.
- Improved waits, retries, and robust frame-aware operations.
- Clear logging and per-ID status; saves to Excel at the end.

Important notes / ethics:
- This script improves human-like behavior and reduces automated fingerprints, but no technique is 100% stealth. Use in accordance with the target site's terms of service and applicable law.

Usage:
- Put your IDs in `input_ids.xlsx` (or .csv) with header `id`.
- Activate venv and ensure `playwright`, `pandas`, `openpyxl` are installed.
- Run: `python playwright_price_bot.py`
"""

import asyncio
from typing import List, Optional  # noqa: F401
import sys
import subprocess
import importlib
import warnings
import os
import platform
import random
import math  # noqa: F401
import logging

# Suppress noisy pandas warnings
warnings.simplefilter(action="ignore", category=UserWarning)

# Ensure required modules are installed
required_modules = ["pandas", "playwright", "openpyxl"]
for module_name in required_modules:
    try:
        importlib.import_module(module_name)
    except ModuleNotFoundError:
        print(f"Module '{module_name}' not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])

import pandas as pd  # noqa: E402
from playwright.async_api import async_playwright, Page, TimeoutError, Frame  # noqa: E402, F401

# ---------------- CONFIG ----------------
LOG_LEVEL = logging.INFO
logging.basicConfig(level=LOG_LEVEL, format="[%(levelname)s] %(message)s")

TARGET_URL = "https://portal.bsh-partner.com/portal(bD1lbiZjPTA2MA==)/regionframe.htm"
COUNTRY_SELECTOR = "/html/body/div[2]/div/div[1]/div[2]/div[3]/div[9]/a"
LOGIN_INPUT_SELECTOR = "#PORTAL_LOGINNAME"
LOGIN_PASSWORD_SELECTOR = "#PORTAL_PASSWORD"
LOGIN_BUTTON_SELECTOR = "#loginsubmitbtn"
SERVIS_MENU_SELECTOR = "body > div:nth-child(1) > a"
QUICKFINDER_SELECTOR = "body > div:nth-child(4) > a"
INPUT_SELECTOR = "/html/body/form/div[3]/div[2]/table/tbody/tr[2]/td[2]/input"
PRICE_SELECTOR = "/html/body/form/table[5]/tbody/tr[3]/td[7]"

LOGIN_USERNAME = "SERAZ_BSC"
LOGIN_PASSWORD = "AAss1234*"

RESULT_FILE = "results.xlsx"
IDS_FILE = "input_ids.xlsx"  # or .csv
IDS_COLNAME = "id"
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15",
]
LOCALES = ["en-US", "tr-TR", "en-GB"]
VIEWPORTS = [(1200, 800), (1366, 768), (1440, 900), (1600, 900)]
TYPING_DELAY_RANGE = (0.04, 0.14)
CLICK_DELAY_RANGE = (0.08, 0.25)
MAX_RETRIES = 5
WAIT_TIMEOUT = 12000  # ms
SCROLL_VARIANCE = 100
ACCEPT_LANG_POOL = ["en-US,en;q=0.9", "tr-TR,tr;q=0.9,en;q=0.8", "en-GB,en;q=0.9"]


def normalize_selector(sel: str) -> str:
    s = sel.strip()
    if s.startswith("xpath=") or s.startswith("css="):
        return s
    if s.startswith("/") or s.startswith("//"):
        return f"xpath={s}"
    return s


async def find_locator(page: Page, selector: str, timeout: int = 2000):
    sel = normalize_selector(selector)
    try:
        await page.wait_for_selector(sel, timeout=timeout)
        return page.locator(sel)
    except Exception:
        for frame in page.frames:
            try:
                await frame.wait_for_selector(sel, timeout=timeout)
                return frame.locator(sel)
            except Exception:
                continue
    return None


async def human_move_and_click(page: Page, locator, click_offset=(0, 0)) -> bool:
    """Move mouse along a randomized path to the element center + offset, then click."""
    try:
        box = await locator.bounding_box()
        if not box:
            return False
        target_x = box["x"] + box["width"] / 2 + click_offset[0] + random.uniform(-5, 5)
        target_y = (
            box["y"] + box["height"] / 2 + click_offset[1] + random.uniform(-5, 5)
        )

        steps = random.randint(8, 18)
        vp = page.viewport_size or {"width": 1200, "height": 800}
        cur_x = vp["width"] / 2
        cur_y = vp["height"] / 2
        for i in range(1, steps + 1):
            t = i / steps
            ease = 1 - (1 - t) ** 3
            x = cur_x + (target_x - cur_x) * ease + random.uniform(-2, 2)
            y = cur_y + (target_y - cur_y) * ease + random.uniform(-2, 2)
            await page.mouse.move(x, y, steps=1)
            await asyncio.sleep(random.uniform(0.01, 0.03))
        await asyncio.sleep(random.uniform(0.05, 0.15))
        await page.mouse.click(target_x, target_y)
        await asyncio.sleep(random.uniform(*CLICK_DELAY_RANGE))
        return True
    except Exception as e:
        logging.debug(f"human_move_and_click failed: {e}")
        return False


async def safe_click(page: Page, selector: str, retries: int = MAX_RETRIES) -> bool:
    sel = normalize_selector(selector)  # noqa: F841
    for attempt in range(1, retries + 1):
        try:
            locator = await find_locator(page, selector, timeout=WAIT_TIMEOUT)
            if locator:
                ok = await human_move_and_click(page, locator)
                if ok:
                    return True
                try:
                    await locator.click()
                    await asyncio.sleep(random.uniform(*CLICK_DELAY_RANGE))
                    return True
                except Exception:
                    pass
        except Exception as e:
            logging.debug(f"safe_click attempt {attempt} error: {e}")
        await asyncio.sleep(0.4 * attempt)
    logging.error(f"Failed to click selector after {retries} attempts: {selector}")
    return False


async def safe_type(
    page: Page, selector: str, text: str, retries: int = MAX_RETRIES
) -> bool:
    sel = normalize_selector(selector)  # noqa: F841
    for attempt in range(1, retries + 1):
        try:
            locator = await find_locator(page, selector, timeout=WAIT_TIMEOUT)
            if locator:
                await locator.focus()
                await locator.fill("")
                await locator.type(text, delay=random.uniform(*TYPING_DELAY_RANGE))
                return True
        except Exception as e:
            logging.debug(f"safe_type attempt {attempt} error: {e}")
        await asyncio.sleep(0.25 * attempt)
    logging.error(f"Failed to type into selector after {retries} attempts: {selector}")
    return False


async def safe_get_text(page: Page, selector: str, retries: int = 4) -> str:
    normalize_selector(selector)  # pyright: ignore[reportUnusedVariable]
    for attempt in range(1, retries + 1):
        try:
            locator = await find_locator(page, selector, timeout=WAIT_TIMEOUT)
            if locator:
                txt = await locator.inner_text()
                return txt.strip()
        except Exception as e:
            logging.debug(f"safe_get_text attempt {attempt} error: {e}")
        await asyncio.sleep(0.2 * attempt)
    return ""


async def read_ids(path: str, colname: str) -> List[str]:
    try:
        if path.lower().endswith(".csv"):
            df = pd.read_csv(path, dtype=str)
        else:
            df = pd.read_excel(path, dtype=str)
        if colname not in df.columns:
            raise ValueError(f"Column '{colname}' not found in {path}.")
        return df[colname].dropna().astype(str).tolist()
    except Exception as e:
        logging.error(f"Error reading IDs: {e}")
        return []


async def save_results_to_excel(df: pd.DataFrame):
    try:
        df.to_excel(RESULT_FILE, index=False)
        logging.info(f"Saved results to {RESULT_FILE}")
        try:
            if platform.system() == "Windows":
                os.startfile(RESULT_FILE)
            elif platform.system() == "Darwin":
                subprocess.call(["open", RESULT_FILE])
            else:
                subprocess.call(["xdg-open", RESULT_FILE])
        except Exception as e:
            logging.warning(f"Could not open Excel automatically: {e}")
    except Exception as e:
        logging.error(f"Could not save results to Excel: {e}")


async def run_bot():
    ids = await read_ids(IDS_FILE, IDS_COLNAME)
    if not ids:
        logging.error(
            "No IDs to process. Place values in input_ids.xlsx (column 'id')."
        )
        return

    async with async_playwright() as pw:
        ua = random.choice(USER_AGENTS)
        vp = random.choice(VIEWPORTS)
        locale = random.choice(LOCALES)
        headers = {"Accept-Language": random.choice(ACCEPT_LANG_POOL)}

        context_args = dict(
            user_agent=ua,
            viewport={"width": vp[0], "height": vp[1]},
            locale=locale,
            java_script_enabled=True,
            extra_http_headers=headers,
        )
        browser = await pw.chromium.launch(
            headless=False, args=["--disable-blink-features=AutomationControlled"]
        )
        context = await browser.new_context(**context_args)
        await context.add_init_script("""
            // Pass the Chrome Test.
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            // Mock plugins and languages
            Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});

        """)
        page = await context.new_page()
        logging.info("Opening portal region page...")

        await page.goto(TARGET_URL)
        # login in baksh  bary  code hast kebtone vard  besdh
        ok = await safe_click(page, COUNTRY_SELECTOR)
        if not ok:
            logging.error("Could not select country. Aborting.")
            await browser.close()
            return
        if not await safe_type(page, LOGIN_INPUT_SELECTOR, LOGIN_USERNAME):
            logging.error("Login username field not found. Aborting.")
            await browser.close()
            return
        await asyncio.sleep(random.uniform(0.3, 0.6))
        if not await safe_type(page, LOGIN_PASSWORD_SELECTOR, LOGIN_PASSWORD):
            logging.error("Login password field not found. Aborting.")
            await browser.close()
            return
        await asyncio.sleep(random.uniform(0.2, 0.5))
        await safe_click(page, LOGIN_BUTTON_SELECTOR)

        await asyncio.sleep(random.uniform(1.0, 2.2))

        await safe_click(page, SERVIS_MENU_SELECTOR)
        await asyncio.sleep(random.uniform(0.3, 0.7))
        await safe_click(page, QUICKFINDER_SELECTOR)

        results = []
        total = len(ids)
        logging.info(f"Processing {total} IDs...")

        idx = 0
        for id_value in ids:
            idx += 1
            logging.info(f"[{idx}/{total}] Entering ID: {id_value}")
            try:
                if not await safe_type(page, INPUT_SELECTOR, id_value):
                    logging.warning(f"Could not type ID {id_value}; skipping")
                    results.append(
                        {"ID": id_value, "Price": "", "Status": "typing_failed"}
                    )
                    continue
                await page.press(normalize_selector(INPUT_SELECTOR), "Enter")

                # wait for price cell to update - do multiple attempts with small backoff
                price_text = ""
                for attempt in range(6):
                    price_text = await safe_get_text(page, PRICE_SELECTOR)
                    if price_text:
                        break
                    await asyncio.sleep(0.5 + attempt * 0.2)

                results.append(
                    {
                        "ID": id_value,
                        "Price": price_text,
                        "Status": "ok" if price_text else "no_price",
                    }
                )
                try:
                    await page.fill(normalize_selector(INPUT_SELECTOR), "")
                except Exception:
                    for frame in page.frames:
                        try:
                            await frame.fill(normalize_selector(INPUT_SELECTOR), "")
                        except Exception:
                            continue
                await page.mouse.wheel(0, random.randint(10, SCROLL_VARIANCE))
                await asyncio.sleep(random.uniform(0.15, 0.45))
            except Exception as e:
                logging.exception(f"Unexpected error processing ID {id_value}: {e}")
                results.append({"ID": id_value, "Price": "", "Status": f"error: {e}"})
        await save_results_to_excel(pd.DataFrame(results))

        await browser.close()


if __name__ == "__main__":
    try:
        asyncio.run(run_bot())
    except KeyboardInterrupt:
        logging.info("Interrupted by user")
    except Exception as e:
        logging.exception(f"Fatal error: {e}")
