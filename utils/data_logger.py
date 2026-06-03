import pandas as pd
from datetime import datetime
import time
from pathlib import Path

LOG_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

def _get_abs_log_path():
    return LOG_FILE.resolve()

def ensure_log_exists():
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    abs_path = _get_abs_log_path()

    if not LOG_FILE.exists():
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Error Log"
        headers = ["timestamp", "hall", "rack_type", "building", "rack",
                   "error_category", "count", "source_file", "processed_by"]
        ws.append(headers)
        wb.save(LOG_FILE)
        wb.close()
        print(f"[DATA_LOGGER] Created new error log at: {abs_path}")
    else:
        # Upgrade schema if 'rack' column missing (for backward compat with old logs)
        try:
            df = pd.read_excel(LOG_FILE)
            if 'rack' not in df.columns:
                # insert after 'building'
                cols = list(df.columns)
                idx = cols.index('building') + 1 if 'building' in cols else len(cols)
                df.insert(idx, 'rack', '')
                df['rack'] = df['rack'].astype('object').fillna('').astype(str)
                df.to_excel(LOG_FILE, index=False)
                print(f"[DATA_LOGGER] Upgraded log schema with 'rack' column at: {abs_path}")
        except Exception as e:
            print(f"[DATA_LOGGER] Schema upgrade check failed: {e}")
        print(f"[DATA_LOGGER] Using existing error log at: {abs_path}")

def log_errors(hall: str, rack_type: str, building: str, rack: str, error_category: str, count: int,
               source_file: str = "", processed_by: str = "system"):
    ensure_log_exists()
    abs_path = _get_abs_log_path()
    print(f"[DATA_LOGGER] Logging: hall={hall}, rack_type={rack_type}, building={building}, rack={rack}, category={error_category}, count={count}")

    from openpyxl import load_workbook, Workbook

    max_retries = 5
    for attempt in range(max_retries):
        wb = None
        try:
            if LOG_FILE.exists():
                wb = load_workbook(LOG_FILE)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = "Error Log"
                headers = ["timestamp", "hall", "rack_type", "building", "rack",
                           "error_category", "count", "source_file", "processed_by"]
                ws.append(headers)

            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                hall, rack_type, building, rack, error_category, count, source_file, processed_by
            ])
            wb.save(LOG_FILE)
            print(f"[DATA_LOGGER] SUCCESS - Logged to: {abs_path}")
            return True

        except PermissionError:
            if wb:
                try: wb.close()
                except: pass
            print(f"[DATA_LOGGER] Log file locked, retrying... ({attempt + 1}/{max_retries})")
            time.sleep(0.8 * (attempt + 1))
        except Exception as e:
            if wb:
                try: wb.close()
                except: pass
            print(f"[DATA_LOGGER] Failed: {e}")
            return False

    print("[DATA_LOGGER] Failed after multiple retries.")
    return False
