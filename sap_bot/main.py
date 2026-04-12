"""SAP Freight Workflow Automation — Entry Point.

Usage:
    python main.py                    # interactive mode selector
    python main.py --tile 1           # run specific tile
    python main.py --tile 1 2         # run tiles 1 and 2
    python main.py --tile 3 --dry-run # dry-run tile 3
    python main.py --autonomous       # run tiles 1 & 2 on a schedule
"""

import os
import sys
import argparse
import logging
from datetime import datetime
from pathlib import Path

import yaml
from dotenv import load_dotenv

# Ensure the sap_bot directory is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Load .env file for secrets
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from bot.driver_setup import create_driver
from bot.login import login
from bot.excel_reader import read_excel
from bot import tile1_confirmation, tile2_reporting, tile3_invoicing, tile4_pod_upload

# ── Logging setup ───────────────────────────────────────────────────────────
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
log = logging.getLogger("sap_bot")


def load_config(path: str = None) -> dict:
    """Load config.yaml and merge in secrets from .env."""
    if path is None:
        path = os.path.join(os.path.dirname(__file__), "config.yaml")
    if not os.path.isfile(path):
        log.error("Config file not found: %s", path)
        raise SystemExit(f"Config file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    # Merge secrets from environment (.env)
    cfg["login_url"] = os.environ.get("SAP_LOGIN_URL", cfg.get("login_url", ""))
    cfg["launchpad_url"] = os.environ.get("SAP_LAUNCHPAD_URL", cfg.get("launchpad_url", ""))
    cfg["username"] = os.environ.get("SAP_USERNAME", cfg.get("username", ""))
    cfg["password"] = os.environ.get("SAP_PASSWORD", cfg.get("password", ""))

    if not cfg["launchpad_url"] or not cfg["username"] or not cfg["password"]:
        raise SystemExit("Missing SAP_LAUNCHPAD_URL, SAP_USERNAME, or SAP_PASSWORD in .env file")

    log.info("Config loaded from %s (secrets from .env)", path)
    return cfg


def resolve_excel_path(cfg: dict) -> str:
    """Resolve the Excel path (may be relative to sap_bot/ dir)."""
    excel_path = cfg.get("excel_path", "SAP_bot.xlsx")
    if not os.path.isabs(excel_path):
        excel_path = os.path.join(os.path.dirname(__file__), excel_path)
    return os.path.abspath(excel_path)


def interactive_menu() -> list[int]:
    """Show a menu and return list of tile numbers to run."""
    print("\n" + "=" * 55)
    print("  SAP Freight Workflow Automation")
    print("=" * 55)
    print("  1. Freight Orders for Confirmation  (autonomous)")
    print("  2. Freight Orders for Reporting      (autonomous)")
    print("  3. Invoice Freight Documents          (human-gated)")
    print("  4. Manage Freight Execution / POD     (human-gated)")
    print("  5. Run Tiles 1 & 2 (autonomous cycle)")
    print("  0. Exit")
    print("=" * 55)

    choice = input("  Select tile(s) to run [1-5, or 0 to exit]: ").strip()
    if choice == "0":
        return []
    if choice == "5":
        return [1, 2]
    try:
        tiles = [int(t.strip()) for t in choice.replace(",", " ").split()]
        valid = [t for t in tiles if t in (1, 2, 3, 4)]
        if not valid:
            print("  Invalid selection.")
            return []
        return valid
    except ValueError:
        print("  Invalid input.")
        return []


def run_tiles(driver, cfg: dict, tiles: list[int]):
    """Run the specified tiles."""
    dry_run = cfg.get("dry_run", True)
    step_through = cfg.get("step_through", False)

    if dry_run:
        print("\n  *** DRY RUN MODE — no destructive actions will be taken ***\n")
        log.info("DRY RUN MODE enabled")

    # Load Excel data if tiles 3 or 4 are requested
    excel_rows = None
    if 3 in tiles or 4 in tiles:
        excel_path = resolve_excel_path(cfg)
        excel_rows = read_excel(excel_path)
        if not excel_rows:
            log.warning("No data rows in Excel — tiles 3 & 4 will be skipped")

    for tile_num in tiles:
        log.info("=" * 40)
        log.info("Starting Tile %d", tile_num)
        log.info("=" * 40)

        try:
            if tile_num == 1:
                tile1_confirmation.run(driver, dry_run=dry_run)

            elif tile_num == 2:
                tile2_reporting.run(driver, dry_run=dry_run)

            elif tile_num == 3:
                if excel_rows:
                    tile3_invoicing.run(driver, excel_rows,
                                        dry_run=dry_run, step_through=step_through)
                else:
                    log.warning("Skipping Tile 3 — no Excel data")

            elif tile_num == 4:
                if excel_rows:
                    tile4_pod_upload.run(driver, excel_rows,
                                         dry_run=dry_run, step_through=step_through)
                else:
                    log.warning("Skipping Tile 4 — no Excel data")

        except KeyboardInterrupt:
            log.info("User interrupted at Tile %d", tile_num)
            print(f"\n  Stopped at Tile {tile_num}.")
            break
        except Exception as e:
            log.error("Tile %d failed: %s", tile_num, e, exc_info=True)
            print(f"\n  Tile {tile_num} error: {e}")
            continue

        # Navigate back to home for next tile
        try:
            driver.get(cfg["launchpad_url"])
        except Exception:
            pass


def run_autonomous(driver, cfg: dict):
    """Run tiles 1 & 2 on a schedule."""
    import schedule
    import time as _time

    interval = cfg.get("autonomous_interval_minutes", 15)
    log.info("Autonomous mode — running tiles 1 & 2 every %d minutes", interval)

    def job():
        log.info("Autonomous cycle starting")
        run_tiles(driver, cfg, [1, 2])
        log.info("Autonomous cycle complete — next run in %d minutes", interval)

    # Run immediately on start
    job()

    schedule.every(interval).minutes.do(job)

    print(f"\n  Autonomous mode active — tiles 1 & 2 every {interval} min. Press Ctrl+C to stop.\n")
    try:
        while True:
            schedule.run_pending()
            _time.sleep(10)
    except KeyboardInterrupt:
        log.info("Autonomous mode stopped by user")
        print("\n  Autonomous mode stopped.")


def main():
    parser = argparse.ArgumentParser(description="SAP Freight Workflow Automation")
    parser.add_argument("--tile", nargs="+", type=int, choices=[1, 2, 3, 4],
                        help="Tile(s) to run")
    parser.add_argument("--autonomous", action="store_true",
                        help="Run tiles 1 & 2 on a schedule")
    parser.add_argument("--dry-run", action="store_true",
                        help="Override: enable dry-run mode")
    parser.add_argument("--step-through", action="store_true",
                        help="Override: enable step-through mode")
    parser.add_argument("--config", type=str, default=None,
                        help="Path to config.yaml")
    args = parser.parse_args()

    # Load config
    cfg = load_config(args.config)

    # CLI overrides
    if args.dry_run:
        cfg["dry_run"] = True
    if args.step_through:
        cfg["step_through"] = True

    # Determine which tiles to run
    if args.autonomous:
        tiles = [1, 2]
    elif args.tile:
        tiles = args.tile
    else:
        tiles = interactive_menu()
        if not tiles:
            print("  Exiting.")
            return

    # Create driver
    driver = create_driver(cfg["browser"])

    try:
        # Login
        success = login(driver, cfg["login_url"], cfg["launchpad_url"],
                        cfg["username"], cfg["password"])
        if not success:
            log.error("Login failed — aborting")
            print("\n  Login failed. Check logs and screenshots for details.")
            return

        # Run
        if args.autonomous:
            run_autonomous(driver, cfg)
        else:
            run_tiles(driver, cfg, tiles)

    finally:
        if not args.autonomous:
            print("\n  Bot finished. Browser will close in 10 seconds...")
            print("  (Press Ctrl+C to keep browser open)")
            try:
                import time as _time
                _time.sleep(10)
            except KeyboardInterrupt:
                print("  Keeping browser open. Close it manually when done.")
                input("  Press Enter to quit...")
        log.info("Closing browser")
        driver.quit()

    log.info("Bot finished")
    print("\n  Done. Check logs/ for details.")


if __name__ == "__main__":
    main()
