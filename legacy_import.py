"""
One-time importer for a consolidated legacy training file -- the kind
built by hand from years of prior ETL exports and cash books, used to
seed transaction_history before the app has built up its own history.

Runs entirely locally through the app's own file dialog: nothing about
this leaves the machine it's run on. Always produces a preview before
writing anything -- this puts financial GL data into a live database,
so committing on a first guess at column names would be worse than not
importing at all.

Column matching is alias-based and case-insensitive rather than a fixed
position, because a hand-consolidated file is exactly the kind of thing
that drifts in header naming over time.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

from db import Database
from logging_setup import get_logger

_logger = get_logger("legacy_import")

REQUIRED_ALIASES = {
    "description": ["JE comments", "JE Comment", "Description", "Comments"],
    "gl_code": ["Yardi Account #", "Yardi Account Number", "GL Code", "Account #", "Account"],
    "fund": ["Fund", "Fund Name"],
    "amount": ["Amount", "Transaction Amount"],
}
OPTIONAL_ALIASES = {
    "gl_name": ["Yardi Account name", "Yardi Account Name", "GL Name", "Account Name"],
    "post_date": ["Post Date", "Date", "Transaction Date"],
}


@dataclass
class LegacyImportPreview:
    sheet_name: str
    total_rows: int = 0
    duplicate_rows: int = 0
    unique_rows: int = 0
    fund_counts: Dict[str, int] = field(default_factory=dict)
    unmatched_funds: Dict[str, int] = field(default_factory=dict)
    missing_gl_rows: int = 0
    rows: List[dict] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


def _clean(val) -> str:
    s = str(val).strip() if val is not None else ""
    return "" if s.lower() in ("nan", "none", "nat") else s


def _find_column(columns, aliases) -> Optional[str]:
    lower = {str(c).strip().lower(): c for c in columns}
    for alias in aliases:
        key = alias.strip().lower()
        if key in lower:
            return lower[key]
    return None


def preview_legacy_file(filepath: str, sheet_name: str, known_funds: List[str]) -> LegacyImportPreview:
    preview = LegacyImportPreview(sheet_name=sheet_name)

    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        preview.errors.append(f"Could not open sheet '{sheet_name}': {e}")
        _logger.exception("Legacy preview failed to open %s / %s", filepath, sheet_name)
        return preview

    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    missing_required = []
    for field_name, aliases in REQUIRED_ALIASES.items():
        col = _find_column(df.columns, aliases)
        if col is None:
            missing_required.append(f"{field_name} (tried: {', '.join(aliases)})")
        col_map[field_name] = col
    if missing_required:
        preview.errors.append(
            "Missing required column(s): " + "; ".join(missing_required) +
            f".  Columns actually found: {list(df.columns)}")
        return preview

    for field_name, aliases in OPTIONAL_ALIASES.items():
        col_map[field_name] = _find_column(df.columns, aliases)

    preview.total_rows = len(df)
    seen = set()
    known_fund_lower = {f.lower(): f for f in known_funds}

    for _, row in df.iterrows():
        desc = _clean(row.get(col_map["description"]))
        gl_code = _clean(row.get(col_map["gl_code"]))
        fund_raw = _clean(row.get(col_map["fund"]))
        amount_raw = row.get(col_map["amount"], None)
        gl_name = _clean(row.get(col_map["gl_name"])) if col_map["gl_name"] else ""
        post_date = _clean(row.get(col_map["post_date"])) if col_map["post_date"] else ""

        if not desc or not fund_raw:
            continue
        try:
            amount = float(str(amount_raw).replace(",", ""))
        except (TypeError, ValueError):
            continue

        fund_name = known_fund_lower.get(fund_raw.lower())
        if fund_name is None:
            preview.unmatched_funds[fund_raw] = preview.unmatched_funds.get(fund_raw, 0) + 1
            continue

        if not gl_code:
            preview.missing_gl_rows += 1
            continue

        key = (fund_name, desc.lower(), gl_code, round(amount, 2))
        if key in seen:
            preview.duplicate_rows += 1
            continue
        seen.add(key)

        preview.rows.append({
            "fund_name": fund_name, "bank_desc": desc, "bank_desc2": "",
            "combined_desc": desc, "mapped_gl": gl_code, "mapped_name": gl_name,
            "bank_gl": "", "date": post_date, "sheet": "legacy-import",
            "amount": amount, "status": "MAPPED",
        })
        preview.fund_counts[fund_name] = preview.fund_counts.get(fund_name, 0) + 1

    preview.unique_rows = len(preview.rows)
    _logger.info(
        "Legacy preview: total=%s unique=%s duplicates=%s unmatched_funds=%s missing_gl=%s",
        preview.total_rows, preview.unique_rows, preview.duplicate_rows,
        preview.unmatched_funds, preview.missing_gl_rows)
    return preview


def commit_legacy_import(db: Database, preview: LegacyImportPreview,
                          property_codes: Dict[str, str]) -> Dict[str, tuple]:
    db.backup(tag="legacy_import")
    by_fund: Dict[str, list] = {}
    for row in preview.rows:
        by_fund.setdefault(row["fund_name"], []).append(row)

    results: Dict[str, tuple] = {}
    for fund_name, rows in by_fund.items():
        inserted, skipped = db.insert_transactions(
            fund_name, property_codes.get(fund_name, ""), "legacy-import", rows)
        results[fund_name] = (inserted, skipped)
        _logger.info("Legacy import committed for %s: inserted=%s skipped=%s",
                     fund_name, inserted, skipped)
    return results
