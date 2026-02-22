"""共通フィクスチャ・パス設定"""
import sys
from pathlib import Path

# プロジェクトルートを sys.path に追加
sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest
from assets.config import Config


@pytest.fixture
def base_config():
    """最低限の設定が入った Config オブジェクト"""
    c = Config()
    c.display_name = "山田 太郎"
    c.timesheet_display_name = "山田"
    c.shift_display_name = "山田"
    c.teams_user_id = "yamada@example.com"
    c.timesheet_folder = "/fake/timesheet"
    c.output_folder = ""   # 空 → _BASE_DIR/attendance_data にフォールバック
    c.managers = [
        {"name": "鈴木部長", "teams_id": "suzuki@example.com"},
        {"name": "田中課長", "teams_id": "tanaka@example.com"},
    ]
    return c
