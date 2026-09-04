"""
Background workers. Nothing that touches Excel, SQLite, or rapidfuzz over
a real data set should ever run on the UI thread -- even the fast rapidfuzz
path is still I/O plus compute, and a frozen window reads as broken.
"""

from __future__ import annotations

from PyQt6.QtCore import QThread, pyqtSignal

from logging_setup import get_logger, log_bridge

_import_logger = get_logger("import_worker")
_mapping_logger = get_logger("mapping_worker")


class ImportWorker(QThread):
    log = pyqtSignal(str)
    finished_ok = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, importer, archiver, source_path, cashbook_path,
                 fund_cfg, fund_name, property_code, sheet_mapping):
        super().__init__()
        self.importer = importer
        self.archiver = archiver
        self.source_path = source_path
        self.cashbook_path = cashbook_path
        self.fund_cfg = fund_cfg
        self.fund_name = fund_name
        self.property_code = property_code
        self.sheet_mapping = sheet_mapping

    def run(self):
        log = log_bridge(self.log.emit, _import_logger)
        try:
            _import_logger.info("Import started: %s -> %s", self.source_path, self.cashbook_path)
            result = self.importer.run(
                self.source_path, self.cashbook_path, self.fund_cfg, log=log)

            archive_result = self.archiver.archive_completed_months(
                self.cashbook_path, self.fund_name, self.property_code,
                self.sheet_mapping, result.get("latest_month"), log=log)
            result["archive"] = archive_result
            _import_logger.info(
                "Import finished: added=%s skipped=%s archived_months=%s",
                result.get("total_added"), result.get("total_skipped"),
                len(archive_result.get("archived_months", [])))
            self.finished_ok.emit(result)
        except Exception as e:
            _import_logger.exception("Import failed")
            self.failed.emit(str(e))


class MappingWorker(QThread):
    progress = pyqtSignal(int, int)
    finished_ok = pyqtSignal(list, dict)
    failed = pyqtSignal(str)

    def __init__(self, mapper, fund_name, transactions):
        super().__init__()
        self.mapper = mapper
        self.fund_name = fund_name
        self.transactions = transactions

    def run(self):
        try:
            _mapping_logger.info(
                "Mapping run started: fund=%s transactions=%s",
                self.fund_name, len(self.transactions))
            counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0, "EXISTING": 0}
            total = len(self.transactions)
            for i, txn in enumerate(self.transactions):
                if txn.get("existing_gl"):
                    txn["mapped_gl"] = txn["existing_gl"]
                    txn["mapped_name"] = txn.get("existing_name", "")
                    txn["score"] = 100
                    txn["confidence"] = "HIGH"
                    txn["status"] = "EXISTING"
                    counts["EXISTING"] += 1
                else:
                    gl, name, score, conf = self.mapper.find_best_match(
                        self.fund_name, txn.get("combined_desc", ""))
                    txn["mapped_gl"] = gl or ""
                    txn["mapped_name"] = name or ""
                    txn["score"] = score
                    txn["confidence"] = conf
                    txn["status"] = "MAPPED" if gl else "UNMAPPED"
                    counts[conf] = counts.get(conf, 0) + 1
                if i % 10 == 0 or i == total - 1:
                    self.progress.emit(i + 1, total)
            _mapping_logger.info("Mapping run finished: %s", counts)
            self.finished_ok.emit(self.transactions, counts)
        except Exception as e:
            _mapping_logger.exception("Mapping run failed")
            self.failed.emit(str(e))
