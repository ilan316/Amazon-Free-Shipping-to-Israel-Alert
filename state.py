import json
import os
import logging
from datetime import datetime

STATE_FILE = "state.json"
logger = logging.getLogger(__name__)


def load_state() -> dict:
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        logger.warning(f"Could not read state file: {e}")
        return {}


def save_state(state: dict):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2, ensure_ascii=False)
    except IOError as e:
        logger.error(f"Could not write state file: {e}")


def update_product_state(state: dict, asin: str, status: str, notified: bool = False) -> dict:
    now = datetime.now().isoformat(timespec="seconds")
    prev = state.get(asin, {})
    state[asin] = {
        "last_status": status,
        "last_checked": now,
        "last_notified": now if notified else prev.get("last_notified"),
        "consecutive_errors": (
            prev.get("consecutive_errors", 0) + 1
            if status == "ERROR"
            else 0
        ),
    }
    return state


def should_notify(state: dict, asin: str, new_status: str,
                  cooldown_hours: float = 24.0) -> bool:
    """
    Returns True when the product is FREE and enough time has passed
    since the last notification (cooldown_hours, default 24h).

    Logic:
      - If status is not FREE  → never notify
      - If never notified before → notify
      - If last_notified was >= cooldown_hours ago → notify again
      - Otherwise → suppress (still within cooldown window)
    """
    if new_status != "FREE":
        return False

    last_notified = state.get(asin, {}).get("last_notified")
    if last_notified is None:
        return True  # first time FREE detected

    try:
        last_dt = datetime.fromisoformat(last_notified)
        hours_elapsed = (datetime.now() - last_dt).total_seconds() / 3600
        return hours_elapsed >= cooldown_hours
    except Exception:
        return True  # can't parse timestamp → send to be safe


def get_consecutive_errors(state: dict, asin: str) -> int:
    return state.get(asin, {}).get("consecutive_errors", 0)
