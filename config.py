import json
import os
import re
import logging
from copy import deepcopy

CONFIG_FILE = "config.json"
logger = logging.getLogger(__name__)

DEFAULT_CONFIG = {
    "check_interval_minutes": 1440,
    "notification_cooldown_hours": 24,
    "email": {
        "sender": "",
        "recipient": "",
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
    },
    "products": [],
    "browser": {
        "headless": False,
        "slow_mo_ms": 120,
        "user_data_dir": "browser_profile",
    },
}


def load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        logger.info("config.json not found — creating default config.")
        save_config(DEFAULT_CONFIG)
        return deepcopy(DEFAULT_CONFIG)
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def add_product(url_or_asin: str, name: str) -> dict:
    asin = _extract_asin(url_or_asin)
    config = load_config()

    existing = {p["asin"] for p in config["products"]}
    if asin in existing:
        logger.warning(f"Product {asin} is already in the list.")
        return config

    config["products"].append({
        "asin": asin,
        "name": name,
        "url": f"https://www.amazon.com/dp/{asin}",
    })
    save_config(config)
    return config


def remove_product(asin: str) -> dict:
    asin = asin.upper().strip()
    config = load_config()
    before = len(config["products"])
    config["products"] = [p for p in config["products"] if p["asin"] != asin]
    if len(config["products"]) == before:
        logger.warning(f"ASIN {asin} not found in product list.")
    save_config(config)
    return config


def _extract_asin(value: str) -> str:
    """Extract a 10-character ASIN from a URL or return the value directly if it looks like an ASIN."""
    value = value.strip()
    patterns = [
        r"/dp/([A-Z0-9]{10})",
        r"/gp/product/([A-Z0-9]{10})",
        r"/product/([A-Z0-9]{10})",
        r"ASIN=([A-Z0-9]{10})",
    ]
    for pattern in patterns:
        match = re.search(pattern, value, re.IGNORECASE)
        if match:
            return match.group(1).upper()

    # If no URL pattern matched, assume it's an ASIN directly
    if re.fullmatch(r"[A-Z0-9]{10}", value, re.IGNORECASE):
        return value.upper()

    raise ValueError(
        f"Could not extract ASIN from: '{value}'\n"
        "Please provide an Amazon product URL (e.g. https://www.amazon.com/dp/B08N5WRWNW) "
        "or a 10-character ASIN (e.g. B08N5WRWNW)."
    )
