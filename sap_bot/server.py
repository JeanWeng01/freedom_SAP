"""Web server for Railway deployment.

Endpoints:
    GET /              — health check
    GET /status        — last run results for each tile
    POST /run/tile1    — trigger tile 1 manually
    POST /run/tile2    — trigger tile 2 manually
    POST /run/all      — trigger tiles 1 & 2

The auto-scheduler runs tiles 1 & 2 on the interval from config.yaml.
Manual triggers work anytime (no schedule restriction for tiles 1 & 2).
"""

import os
import sys
import threading
import logging
from datetime import datetime

from flask import Flask, jsonify, request

# Ensure sap_bot dir is on path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yaml
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from bot.driver_setup import create_driver
from bot.login import login
from bot import tile1_confirmation, tile2_reporting, tile3_invoicing, tile4_pod_upload

# ── Logging ─────────────────────────────────────────────────────────────────
LOGS_DIR = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOGS_DIR, exist_ok=True)

log_filename = datetime.now().strftime("sap_bot_%Y%m%d_%H%M%S.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, log_filename), encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("sap_bot.server")

# ── Config ──────────────────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    cfg["login_url"] = os.environ.get("SAP_LOGIN_URL", cfg.get("login_url", ""))
    cfg["launchpad_url"] = os.environ.get("SAP_LAUNCHPAD_URL", cfg.get("launchpad_url", ""))
    cfg["username"] = os.environ.get("SAP_USERNAME", cfg.get("username", ""))
    cfg["password"] = os.environ.get("SAP_PASSWORD", cfg.get("password", ""))
    # Railway runs Chrome on Linux; local runs Edge on Windows
    cfg["browser"] = os.environ.get("BROWSER", cfg.get("browser", "chrome"))
    # dry_run can be overridden by env var (Railway should set DRY_RUN=false)
    dry_run_env = os.environ.get("DRY_RUN")
    if dry_run_env is not None:
        cfg["dry_run"] = dry_run_env.lower() == "true"
    return cfg

CFG = load_config()

# ── Run state ───────────────────────────────────────────────────────────────

run_lock = threading.Lock()
run_status = {
    "tile1": {"last_run": None, "result": None, "running": False},
    "tile2": {"last_run": None, "result": None, "running": False},
    "tile3": {"last_run": None, "result": None, "running": False},
    "tile4": {"last_run": None, "result": None, "running": False},
    "last_auto_run": None,
    "next_auto_run": None,
    "last_invoice_run": None,
}


def run_tile(tile_num: int, dry_run: bool = False) -> dict:
    """Execute a single tile. Returns result dict."""
    tile_key = f"tile{tile_num}"

    if run_status[tile_key]["running"]:
        return {"error": f"Tile {tile_num} is already running"}

    run_status[tile_key]["running"] = True
    start_time = datetime.now().isoformat()

    driver = None
    try:
        driver = create_driver(CFG.get("browser", "chrome"))
        success = login(driver, CFG["login_url"], CFG["launchpad_url"],
                        CFG["username"], CFG["password"])
        if not success:
            result = {"status": "error", "error": "Login failed", "started": start_time}
            run_status[tile_key]["result"] = result
            return result

        if tile_num == 1:
            count = tile1_confirmation.run(driver, dry_run=dry_run)
            result = {
                "status": "completed",
                "tile": 1,
                "orders_confirmed": count,
                "dry_run": dry_run,
                "started": start_time,
                "finished": datetime.now().isoformat(),
            }
        elif tile_num == 2:
            count = tile2_reporting.run(driver, dry_run=dry_run)
            result = {
                "status": "completed",
                "tile": 2,
                "stops_reported": count,
                "dry_run": dry_run,
                "started": start_time,
                "finished": datetime.now().isoformat(),
            }
        elif tile_num == 3:
            counts = tile3_invoicing.run(driver, dry_run=dry_run)
            result = {
                "status": "completed",
                "tile": 3,
                **counts,
                "dry_run": dry_run,
                "started": start_time,
                "finished": datetime.now().isoformat(),
            }
        elif tile_num == 4:
            counts = tile4_pod_upload.run(driver, dry_run=dry_run)
            result = {
                "status": "completed",
                "tile": 4,
                **counts,
                "dry_run": dry_run,
                "started": start_time,
                "finished": datetime.now().isoformat(),
            }
        else:
            result = {"status": "error", "error": f"Unknown tile {tile_num}"}

    except Exception as e:
        log.error("Tile %d failed: %s", tile_num, e, exc_info=True)
        result = {
            "status": "error",
            "tile": tile_num,
            "error": str(e),
            "started": start_time,
            "finished": datetime.now().isoformat(),
        }
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        run_status[tile_key]["running"] = False
        run_status[tile_key]["last_run"] = datetime.now().isoformat()
        run_status[tile_key]["result"] = result

    return result


def cleanup_old_logs(max_age_days: int = 7):
    """Delete log files older than max_age_days."""
    import glob
    cutoff = datetime.now().timestamp() - (max_age_days * 86400)
    for f in glob.glob(os.path.join(LOGS_DIR, "*.log")):
        if os.path.getmtime(f) < cutoff:
            try:
                os.remove(f)
                log.info("Deleted old log: %s", os.path.basename(f))
            except Exception:
                pass


def run_tile_with_retry(tile_num: int, dry_run: bool = False, max_retries: int = 1) -> dict:
    """Run a tile with retry on failure."""
    for attempt in range(1, max_retries + 2):
        result = run_tile(tile_num, dry_run=dry_run)
        if result.get("status") != "error":
            return result
        if attempt <= max_retries:
            log.warning("Tile %d failed (attempt %d) — retrying in 30s...", tile_num, attempt)
            import time
            time.sleep(30)
        else:
            log.error("Tile %d failed after %d attempts", tile_num, attempt)
    return result


def run_auto_cycle_12():
    """Run tiles 1 & 2 in sequence with retry."""
    log.info("═══ Tiles 1 & 2 auto cycle starting ═══")
    cleanup_old_logs(max_age_days=7)
    run_status["last_auto_run"] = datetime.now().isoformat()

    dry_run = CFG.get("dry_run", True)

    for tile_num in [1, 2]:
        result = run_tile_with_retry(tile_num, dry_run=dry_run, max_retries=1)
        log.info("Tile %d result: %s", tile_num, result.get("status"))

    log.info("═══ Tiles 1 & 2 auto cycle complete ═══")


def run_auto_cycle_34():
    """Run tiles 3 & 4 in sequence with retry."""
    log.info("═══ Tiles 3 & 4 auto cycle starting ═══")
    run_status["last_invoice_run"] = datetime.now().isoformat()

    dry_run = CFG.get("dry_run", True)

    for tile_num in [3, 4]:
        result = run_tile_with_retry(tile_num, dry_run=dry_run, max_retries=1)
        log.info("Tile %d result: %s", tile_num, result.get("status"))

    log.info("═══ Tiles 3 & 4 auto cycle complete ═══")


def is_in_invoice_window() -> bool:
    """Check if current time is within the daytime window (9am-9pm) and NOT 4am-9am."""
    hour = datetime.now().hour
    return 9 <= hour < 21


# ── Auto-schedulers (background threads) ────────────────────────────────────

# Tiles 1 & 2 run times (ET): 9am, 12pm, 3pm
TILE12_RUN_HOURS = [int(h) for h in os.environ.get("TILE12_RUN_HOURS", "9,12,15").split(",")]

# Tiles 3 & 4: disabled by default, enable via env var ENABLE_TILES_34=true
TILES_34_ENABLED = os.environ.get("ENABLE_TILES_34", "false").lower() == "true"
TILE34_INTERVAL_MINUTES = int(os.environ.get("TILE34_INTERVAL_MINUTES", "180"))


def scheduler_loop_12():
    """Run tiles 1 & 2 at specific hours (default: 9am, 12pm, 3pm ET)."""
    import time
    log.info("Tiles 1 & 2 scheduler started — runs at hours: %s", TILE12_RUN_HOURS)

    already_ran_this_hour = None

    while True:
        now = datetime.now()
        current_hour = now.hour

        if current_hour in TILE12_RUN_HOURS and already_ran_this_hour != current_hour:
            already_ran_this_hour = current_hour
            try:
                run_auto_cycle_12()
            except Exception as e:
                log.error("Tiles 1 & 2 auto cycle error: %s", e, exc_info=True)

            # Update next run time
            remaining_hours = [h for h in TILE12_RUN_HOURS if h > current_hour]
            if remaining_hours:
                next_hour = remaining_hours[0]
                next_run = now.replace(hour=next_hour, minute=0, second=0).isoformat()
            else:
                next_run = f"tomorrow at {TILE12_RUN_HOURS[0]}:00"
            run_status["next_auto_run"] = next_run
            log.info("Next tiles 1 & 2 run: %s", next_run)

        time.sleep(60)  # check every minute


def scheduler_loop_34():
    """Run tiles 3 & 4 every N hours during daytime (not 4am-9am). Disabled by default."""
    import time

    if not TILES_34_ENABLED:
        log.info("Tiles 3 & 4 scheduler DISABLED (set ENABLE_TILES_34=true to enable)")
        return  # exit thread entirely

    interval = TILE34_INTERVAL_MINUTES * 60
    log.info("Tiles 3 & 4 scheduler started — every %d min during daytime (not 4am-9am)",
             TILE34_INTERVAL_MINUTES)

    while True:
        if is_in_invoice_window():
            try:
                run_auto_cycle_34()
            except Exception as e:
                log.error("Tiles 3 & 4 auto cycle error: %s", e, exc_info=True)
        else:
            log.debug("Outside daytime window — skipping tiles 3 & 4")

        time.sleep(interval)


# ── Flask app ───────────────────────────────────────────────────────────────

app = Flask(__name__)


@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "SAP Freight Workflow Bot",
        "dry_run": CFG.get("dry_run", True),
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify(run_status)


@app.route("/run/tile1", methods=["GET", "POST"])
def trigger_tile1():
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if run_status["tile1"]["running"]:
        return jsonify({"error": "Tile 1 is already running"}), 409

    # Run in background thread so we can return immediately
    def _run():
        run_tile(1, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tile 1 triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


@app.route("/run/tile2", methods=["GET", "POST"])
def trigger_tile2():
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if run_status["tile2"]["running"]:
        return jsonify({"error": "Tile 2 is already running"}), 409

    def _run():
        run_tile(2, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tile 2 triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


@app.route("/run/tile3", methods=["GET", "POST"])
def trigger_tile3():
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if run_status["tile3"]["running"]:
        return jsonify({"error": "Tile 3 is already running"}), 409

    def _run():
        run_tile(3, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tile 3 (invoicing) triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


@app.route("/run/tile4", methods=["GET", "POST"])
def trigger_tile4():
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if run_status["tile4"]["running"]:
        return jsonify({"error": "Tile 4 is already running"}), 409

    def _run():
        run_tile(4, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tile 4 (POD upload) triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


@app.route("/run/all", methods=["GET", "POST"])
def trigger_all():
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if any(run_status[f"tile{n}"]["running"] for n in [1, 2, 3, 4]):
        return jsonify({"error": "A tile is already running"}), 409

    def _run():
        run_tile(1, dry_run=dry_run)
        run_tile(2, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tiles 1 & 2 triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


@app.route("/run/invoices", methods=["GET", "POST"])
def trigger_invoices():
    """Trigger tiles 3 & 4 (invoicing + POD upload) — works anytime, ignores 4am-9am window."""
    dry_run = request.args.get("dry_run", "false").lower() == "true"

    if run_status["tile3"]["running"] or run_status["tile4"]["running"]:
        return jsonify({"error": "Tiles 3 or 4 already running"}), 409

    def _run():
        run_tile(3, dry_run=dry_run)
        run_tile(4, dry_run=dry_run)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({
        "message": f"Tiles 3 & 4 triggered (dry_run={dry_run})",
        "check_status": "/status",
    })


# ── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Start tiles 1 & 2 scheduler (24/7)
    t12 = threading.Thread(target=scheduler_loop_12, daemon=True)
    t12.start()
    log.info("Tiles 1 & 2 scheduler started")

    # Start tiles 3 & 4 scheduler (4am-9am window)
    t34 = threading.Thread(target=scheduler_loop_34, daemon=True)
    t34.start()
    log.info("Tiles 3 & 4 scheduler started (4am-9am window)")

    # Start Flask
    port = int(os.environ.get("PORT", 8080))
    log.info("Starting web server on port %d", port)
    app.run(host="0.0.0.0", port=port)
