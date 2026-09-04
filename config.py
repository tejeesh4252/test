"""
Configuration and path resolution for the Balboa Mapping Agent.

Design goals (from the architecture review):
- No hardcoded personal machine paths in source code.
- The database location is resolved from an absolute path stored in a
  config file, not inferred from whatever folder the app happens to be
  launched from. That was the bug that silently created a fresh, empty
  database if the app was ever run from a different working directory.
- Fund-specific settings (cashbook paths, sheet mappings, bank-name
  normalization) live in the same config file and are editable from the
  Settings dialog -- no code changes needed to point at a new cashbook.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict


def _base_dir() -> Path:
    """
    Resolve a stable base directory for config/data/logs, independent of
    the working directory the app happens to be launched from.

    Priority:
      1. BALBOA_HOME environment variable, if set (lets you point an
         install at a specific shared folder without touching code).
      2. The folder containing the running executable (frozen/PyInstaller
         build) or this source file (dev run).
    """
    env = os.environ.get("BALBOA_HOME")
    if env:
        return Path(env).expanduser().resolve()

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _base_dir()
CONFIG_PATH = BASE_DIR / "config.json"
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = DATA_DIR / "backups"
LOG_DIR = BASE_DIR / "logs"


DEFAULT_FUND_CONFIGS: Dict[str, Dict[str, Any]] = {
    "Fund I": {
        "property_code": "brf1",
        "fund_name": "Balboa Retail Fund I, L.P.",
        "cashbook": "",
        "sheet_mapping": {
            "CHASE_80003219615": "JPM CHASE - 80003219615",
            "BOFA_1453831338": "BoFA - 1453831338",
            "PNC_1087615896": "PNC - 1087615896",
            "PNC_1087664321": "PNC - 1087664321",
        },
        "bank_normalization": {
            "CHASE": "CHASE", "JPM": "CHASE",
            "BOFA": "BOFA", "BOA": "BOFA", "BANK OF AMERICA": "BOFA",
            "PNC": "PNC",
        },
    },
    "Fund II": {
        "property_code": "brf2f",
        "fund_name": "Balboa Retail Fund II, L.P.",
        "cashbook": "",
        "sheet_mapping": {
            "CHASE_80007962089": "JPM CHASE - 80007962089",
            "CHASE_1453538978": "BoFA - 1453538978",
            "CHASE_1087615909": "PNC - 1087615909",
            "CHASE_1087664436": "PNC - 1087664436",
        },
        "bank_normalization": {
            "CHASE": "CHASE", "JPM": "CHASE",
            "BOFA": "CHASE", "BOA": "CHASE", "BANK OF AMERICA": "CHASE",
            "PNC": "CHASE",
        },
    },
    "Fund III": {
        "property_code": "brf3f",
        "fund_name": "Balboa Retail Fund III, L.P.",
        "cashbook": "",
        "sheet_mapping": {
            "CHASE_689095811": "JPM CHASE - 689095811",
            "PNC_1087615917": "PNC - 1087615917",
            "PNC_1087664444": "PNC - 1087664444",
        },
        "bank_normalization": {
            "CHASE": "CHASE", "JPM": "CHASE",
            "PNC": "PNC",
        },
    },
}

DEFAULT_CONFIG: Dict[str, Any] = {
    "db_path": str(DATA_DIR / "mappings_database.db"),
    "book_num": 1000,
    "header_rows": 9,
    "skip_sheets": ["Summary", "ETL Sample", "COA"],
    "exclude_keywords": [
        "BEGINNING BALANCE", "ENDING BALANCE", "CALCULATION", "ACCRUAL",
        "PAID", "BALANCE", "RUNNING", "CASH BALANCE", "CASH ENTRIES",
    ],
    "high_threshold": 85,
    "medium_threshold": 70,
    "funds": DEFAULT_FUND_CONFIGS,
}


class AppConfig:
    """
    Loads config.json on first use, creates it with defaults if missing.
    Mutate .data in place, then call save() to persist.
    """

    def __init__(self, path: Path = CONFIG_PATH):
        self.path = path
        self.data: Dict[str, Any] = self._load_or_create()

    def _load_or_create(self) -> Dict[str, Any]:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)

        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                merged = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
                merged.update(data)
                # backfill any fund keys added in a newer version
                for fund, defaults in DEFAULT_FUND_CONFIGS.items():
                    merged["funds"].setdefault(fund, defaults)
                return merged
            except Exception:
                # Corrupt config -- keep the broken copy for inspection
                # instead of silently overwriting it.
                try:
                    self.path.replace(self.path.with_suffix(".corrupt.json"))
                except Exception:
                    pass

        data = json.loads(json.dumps(DEFAULT_CONFIG))
        self._write(data)
        return data

    def _write(self, data: Dict[str, Any]) -> None:
        tmp = self.path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        tmp.replace(self.path)

    def save(self) -> None:
        self._write(self.data)

    # -- typed accessors --------------------------------------------
    @property
    def db_path(self) -> Path:
        return Path(self.data["db_path"])

    @db_path.setter
    def db_path(self, value: str) -> None:
        self.data["db_path"] = str(value)

    @property
    def fund_names(self):
        return list(self.data["funds"].keys())

    def fund(self, fund_name: str) -> Dict[str, Any]:
        return self.data["funds"][fund_name]

    def property_code(self, fund_name: str) -> str:
        return self.data["funds"][fund_name]["property_code"]

    def cashbook_path(self, fund_name: str) -> str:
        return self.data["funds"][fund_name].get("cashbook", "")

    def set_cashbook_path(self, fund_name: str, path: str) -> None:
        self.data["funds"][fund_name]["cashbook"] = path
