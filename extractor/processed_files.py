"""Track which PDF files have been processed and persist extracted results."""
import json
from pathlib import Path

from django.conf import settings

BASE = Path(getattr(settings, "BASE_DIR", Path(__file__).resolve().parent.parent))
PROCESSED_LIST_PATH = BASE / "processed_files.json"
EXTRACTED_RESULTS_PATH = BASE / "extracted_results.json"


def load_processed() -> dict[str, str]:
    """Load processed files: {filename: processed_at_iso}. Merges processed_files + extracted_results."""
    processed = {}
    if PROCESSED_LIST_PATH.exists():
        try:
            with open(PROCESSED_LIST_PATH, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                for k, v in data.items():
                    processed[k] = v if isinstance(v, str) else v.get("processed_at", "")
        except (json.JSONDecodeError, OSError):
            pass
    if EXTRACTED_RESULTS_PATH.exists():
        try:
            with open(EXTRACTED_RESULTS_PATH, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, dict) and "processed_at" in v:
                        processed[k] = v["processed_at"]
        except (json.JSONDecodeError, OSError):
            pass
    return processed


def _migrate_from_processed_files() -> None:
    """One-time migration: convert processed_files.json to extracted_results format."""
    if EXTRACTED_RESULTS_PATH.exists():
        return
    if not PROCESSED_LIST_PATH.exists():
        return
    try:
        with open(PROCESSED_LIST_PATH, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return
        migrated = {}
        for filename, val in data.items():
            ts = val if isinstance(val, str) else (val.get("processed_at") if isinstance(val, dict) else "")
            if ts:
                migrated[filename] = {"processed_at": ts, "data": None}
        if migrated:
            with open(EXTRACTED_RESULTS_PATH, "w", encoding="utf-8") as f:
                json.dump(migrated, f, indent=2)
    except (json.JSONDecodeError, OSError):
        pass


def load_extracted_results() -> list[tuple[str, dict]]:
    """Load all saved extracted results. Returns [(filename, data_dict), ...] sorted by filename."""
    _migrate_from_processed_files()
    if not EXTRACTED_RESULTS_PATH.exists():
        return []
    try:
        with open(EXTRACTED_RESULTS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return []
        result = []
        for filename, entry in sorted(data.items()):
            if isinstance(entry, dict) and "data" in entry and entry["data"]:
                result.append((filename, entry["data"]))
        return result
    except (json.JSONDecodeError, OSError):
        return []


def save_extracted_result(filename: str, data_dict: dict) -> None:
    """Save or update one extracted result. Adds to existing results."""
    from datetime import datetime

    existing = {}
    if EXTRACTED_RESULTS_PATH.exists():
        try:
            with open(EXTRACTED_RESULTS_PATH, encoding="utf-8") as f:
                existing = json.load(f)
            if not isinstance(existing, dict):
                existing = {}
        except (json.JSONDecodeError, OSError):
            existing = {}
    existing[filename] = {
        "processed_at": datetime.utcnow().isoformat(),
        "data": data_dict,
    }
    with open(EXTRACTED_RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2)


def clear_processed() -> None:
    """Clear processed list and all saved extracted results."""
    if PROCESSED_LIST_PATH.exists():
        with open(PROCESSED_LIST_PATH, "w", encoding="utf-8") as f:
            json.dump({}, f)
    if EXTRACTED_RESULTS_PATH.exists():
        with open(EXTRACTED_RESULTS_PATH, "w", encoding="utf-8") as f:
            json.dump({}, f)
