"""
Scheduler — runs the Amazon check cycle at the interval defined in config.json.
Uses APScheduler's BlockingScheduler (synchronous); the Playwright check is
run via asyncio.run() inside each job invocation.
"""

import asyncio
import logging

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.interval import IntervalTrigger

from checker import check_all_products, ShippingStatus
from notifier import send_free_shipping_alert
from state import (
    load_state,
    save_state,
    update_product_state,
    should_notify,
    get_consecutive_errors,
)
from config import load_config

logger = logging.getLogger(__name__)

# A product with this many consecutive check errors is temporarily skipped
# and a warning is logged (but it stays in the list).
MAX_CONSECUTIVE_ERRORS = 5


def run_check_cycle():
    """
    Executed by APScheduler on each interval tick.
    Loads the latest config and state on every run so config changes take effect
    without restarting the process.
    """
    config = load_config()
    state = load_state()

    products = config.get("products", [])
    if not products:
        logger.info("No products configured — nothing to check.")
        return

    logger.info(f"=== Check cycle started ({len(products)} product(s)) ===")

    try:
        results = asyncio.run(check_all_products(config, state))
    except Exception as e:
        logger.error(f"Check cycle failed with unexpected error: {e}")
        return

    product_map = {p["asin"]: p for p in products}

    for result in results:
        asin = result.asin
        status_str = result.status.value
        product = product_map.get(asin, {"asin": asin, "name": asin, "url": ""})
        name = product.get("name", asin)

        consecutive = get_consecutive_errors(state, asin)
        if consecutive >= MAX_CONSECUTIVE_ERRORS:
            logger.warning(
                f"[{asin}] Skipped — {consecutive} consecutive errors. "
                "Fix the issue or remove the product to resume."
            )
            continue

        notify = should_notify(state, asin, status_str)
        if notify:
            logger.info(f"[{asin}] FREE shipping detected! Sending email alert.")
            try:
                send_free_shipping_alert(config, product, result.raw_text)
            except RuntimeError as e:
                logger.error(f"[{asin}] Email failed: {e}")

        state = update_product_state(state, asin, status_str, notified=notify)

        status_label = {
            "FREE": "FREE shipping to Israel",
            "PAID": "Ships to Israel (paid)",
            "NO_SHIP": "Does not ship to Israel",
            "UNKNOWN": "Unknown / could not parse",
            "ERROR": f"Error — {result.error_message}",
        }.get(status_str, status_str)

        logger.info(f"  {name}: {status_label}")

    save_state(state)
    logger.info("=== Check cycle complete ===")


def start_scheduler(config: dict):
    """
    Starts the monitoring scheduler.
    Runs one check immediately, then repeats at the configured interval.
    """
    interval = config.get("check_interval_minutes", 60)

    scheduler = BlockingScheduler(timezone="UTC")
    scheduler.add_job(
        run_check_cycle,
        trigger=IntervalTrigger(minutes=interval),
        id="amazon_check",
        name="Amazon Shipping Check",
        misfire_grace_time=300,
    )

    logger.info(f"Monitoring started — checking every {interval} minute(s).")
    logger.info("Running first check now...")
    run_check_cycle()

    try:
        scheduler.start()
    except KeyboardInterrupt:
        logger.info("Monitoring stopped by user (Ctrl+C).")
        scheduler.shutdown(wait=False)
