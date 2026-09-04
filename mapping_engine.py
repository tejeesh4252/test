"""
Mapping engine: turns a transaction description into a GL code + confidence.

Match order:
  1. Exact fund-scoped correction (a human already told us the answer
     for this exact description, for this exact fund).
  2. Fuzzy match against this fund's own transaction history, using
     rapidfuzz -- C-accelerated, MIT licensed (unlike the optional
     python-Levenshtein speedup for the old fuzzywuzzy library, which
     is GPLv2 -- worth avoiding given this app already carries PyQt6
     licensing considerations).

process.extractOne runs the whole candidate list in optimized C rather
than a manual Python for-loop, which is the other half of the speed fix
alongside the library swap itself.

Fund scoping matters twice here: corrections never leak across funds,
and the fuzzy corpus itself is filtered to the current fund, so Fund I's
chart of accounts can't get suggested for a Fund III transaction just
because the bank memo text happens to look similar.
"""

from __future__ import annotations

from typing import Optional, Tuple

from rapidfuzz import fuzz, process

from config import AppConfig
from db import Database


class MappingEngine:

    def __init__(self, config: AppConfig, db: Database):
        self.config = config
        self.db = db
        self._corpus_cache: dict = {}
        self._corrections_cache: dict = {}

    def refresh(self, fund_name: str) -> int:
        """Reload this fund's training corpus and corrections straight
        from the database. Call after any archive/import so the very
        next mapping run benefits from what was just learned -- no
        manual 'load training file' step, ever."""
        self._corpus_cache[fund_name] = self.db.get_training_corpus(fund_name)
        self._corrections_cache[fund_name] = self.db.get_corrections(fund_name)
        return len(self._corpus_cache[fund_name])

    def corpus_size(self, fund_name: str) -> int:
        self._ensure_loaded(fund_name)
        return len(self._corpus_cache.get(fund_name, []))

    def _ensure_loaded(self, fund_name: str) -> None:
        if fund_name not in self._corpus_cache:
            self.refresh(fund_name)

    def find_best_match(self, fund_name: str, description: str
                         ) -> Tuple[Optional[str], Optional[str], int, str]:
        description = (description or "").strip()
        if not description:
            return None, None, 0, "LOW"

        self._ensure_loaded(fund_name)
        desc_key = description.lower()

        corrections = self._corrections_cache.get(fund_name, {})
        if desc_key in corrections:
            gl, name = corrections[desc_key]
            return gl, name, 100, "HIGH"

        corpus = self._corpus_cache.get(fund_name, [])
        if not corpus:
            return None, None, 0, "LOW"

        choices = [row[0].lower() for row in corpus]
        match = process.extractOne(desc_key, choices, scorer=fuzz.token_sort_ratio)
        if not match:
            return None, None, 0, "LOW"

        _, score, idx = match
        _, gl, name = corpus[idx]

        high = self.config.data.get("high_threshold", 85)
        med = self.config.data.get("medium_threshold", 70)
        if score >= high:
            conf = "HIGH"
        elif score >= med:
            conf = "MEDIUM"
        else:
            conf = "LOW"
        return gl, name, int(score), conf

    def save_correction(self, fund_name: str, description: str,
                         gl_code: str, gl_name: str) -> None:
        self.db.save_correction(fund_name, description, gl_code, gl_name)
        self._corrections_cache.setdefault(fund_name, {})[
            description.lower().strip()] = (gl_code, gl_name)
