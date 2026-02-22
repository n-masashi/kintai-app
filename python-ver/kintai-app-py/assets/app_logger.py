"""アプリケーションロガー設定"""
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

_LOG_DIR  = Path(__file__).parent / "logs"
_LOG_FILE = _LOG_DIR / "app.log"
_MAX_BYTES    = 250 * 1024  # 250KB
_BACKUP_COUNT = 3
_FMT      = "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s"
_DATE_FMT = "%Y-%m-%d %H:%M:%S"


def setup_logging() -> None:
    """RotatingFileHandler でログ設定を初期化する。アプリ起動時に1回だけ呼ぶ。"""
    _LOG_DIR.mkdir(parents=True, exist_ok=True)

    handler = RotatingFileHandler(
        str(_LOG_FILE),
        maxBytes=_MAX_BYTES,
        backupCount=_BACKUP_COUNT,
        encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter(_FMT, _DATE_FMT))

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    # 多重追加防止
    if not any(isinstance(h, RotatingFileHandler) for h in root.handlers):
        root.addHandler(handler)


def get_logger(name: str) -> logging.Logger:
    """モジュール別ロガーを返す。"""
    return logging.getLogger(name)
