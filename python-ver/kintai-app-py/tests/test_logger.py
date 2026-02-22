"""assets/app_logger.py のユニットテスト"""
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

import assets.app_logger as app_logger_mod
from assets.app_logger import setup_logging, get_logger


# ─────────────────────────── setup_logging ───────────────────────────

class TestSetupLogging:
    def _clear_root_handlers(self):
        root = logging.getLogger()
        for h in list(root.handlers):
            root.removeHandler(h)
            h.close()

    def test_creates_log_dir(self, tmp_path):
        """ログディレクトリが存在しない場合に作成される"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        assert log_dir.exists()
        self._clear_root_handlers()

    def test_adds_rotating_file_handler(self, tmp_path):
        """RotatingFileHandler がルートロガーに追加される"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        root = logging.getLogger()
        rfh_handlers = [h for h in root.handlers if isinstance(h, RotatingFileHandler)]
        assert len(rfh_handlers) == 1
        self._clear_root_handlers()

    def test_no_duplicate_handler(self, tmp_path):
        """複数回呼んでも RotatingFileHandler は1つだけ"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
            setup_logging()
            setup_logging()
        root = logging.getLogger()
        rfh_handlers = [h for h in root.handlers if isinstance(h, RotatingFileHandler)]
        assert len(rfh_handlers) == 1
        self._clear_root_handlers()

    def test_root_logger_level_debug(self, tmp_path):
        """ルートロガーのレベルが DEBUG に設定される"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        assert logging.getLogger().level == logging.DEBUG
        self._clear_root_handlers()

    def test_handler_max_bytes(self, tmp_path):
        """RotatingFileHandler の maxBytes が 250KB"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        root = logging.getLogger()
        rfh = next(h for h in root.handlers if isinstance(h, RotatingFileHandler))
        assert rfh.maxBytes == 250 * 1024
        self._clear_root_handlers()

    def test_handler_backup_count(self, tmp_path):
        """RotatingFileHandler の backupCount が 3"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        self._clear_root_handlers()
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        root = logging.getLogger()
        rfh = next(h for h in root.handlers if isinstance(h, RotatingFileHandler))
        assert rfh.backupCount == 3
        self._clear_root_handlers()

    def test_existing_rotating_handler_not_duplicated(self, tmp_path):
        """既にRFHが存在する場合は追加しない"""
        log_dir = tmp_path / "logs"
        log_file = log_dir / "app.log"
        log_dir.mkdir(parents=True)
        self._clear_root_handlers()
        # 事前に手動で RFH を追加
        existing_rfh = RotatingFileHandler(str(log_file))
        logging.getLogger().addHandler(existing_rfh)
        with (
            patch.object(app_logger_mod, "_LOG_DIR", log_dir),
            patch.object(app_logger_mod, "_LOG_FILE", log_file),
        ):
            setup_logging()
        root = logging.getLogger()
        rfh_handlers = [h for h in root.handlers if isinstance(h, RotatingFileHandler)]
        assert len(rfh_handlers) == 1
        self._clear_root_handlers()


# ─────────────────────────── get_logger ───────────────────────────

class TestGetLogger:
    def test_returns_logger_instance(self):
        """logging.Logger インスタンスが返される"""
        logger = get_logger("test_module")
        assert isinstance(logger, logging.Logger)

    def test_logger_name(self):
        """指定した名前のロガーが返される"""
        logger = get_logger("my.module")
        assert logger.name == "my.module"

    def test_different_names_different_loggers(self):
        """異なる名前では別のロガーインスタンスが返される"""
        l1 = get_logger("module_a")
        l2 = get_logger("module_b")
        assert l1 is not l2
        assert l1.name != l2.name

    def test_same_name_same_logger(self):
        """同じ名前では同一インスタンスが返される（logging.getLogger のキャッシュ動作）"""
        l1 = get_logger("shared_module")
        l2 = get_logger("shared_module")
        assert l1 is l2
