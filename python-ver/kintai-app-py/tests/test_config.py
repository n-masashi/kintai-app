"""assets/config.py のユニットテスト"""
import json
import pytest
from pathlib import Path

from assets.config import Config


class TestConfigDefaults:
    def test_default_theme(self):
        assert Config().theme == "light"

    def test_default_output_folder(self):
        assert Config().output_folder == "attendance_data"

    def test_default_shift_types_empty(self):
        assert Config().shift_types == []

    def test_default_managers_empty(self):
        assert Config().managers == []

    def test_default_strings_empty(self):
        c = Config()
        assert c.ad_name == ""
        assert c.display_name == ""
        assert c.teams_user_id == ""
        assert c.webhook_url == ""
        assert c.timesheet_folder == ""
        assert c.proxy_sh == ""
        assert c.test_date == ""


class TestConfigLoad:
    def test_load_valid_json(self, tmp_path):
        data = {"display_name": "山田 太郎", "theme": "dark"}
        p = tmp_path / "settings.json"
        p.write_text(json.dumps(data), encoding="utf-8")
        c = Config.load(str(p))
        assert c.display_name == "山田 太郎"
        assert c.theme == "dark"

    def test_load_missing_file_returns_defaults(self, tmp_path):
        c = Config.load(str(tmp_path / "nonexistent.json"))
        assert c.theme == "light"

    def test_load_invalid_json_returns_defaults(self, tmp_path):
        p = tmp_path / "settings.json"
        p.write_text("{ invalid json", encoding="utf-8")
        c = Config.load(str(p))
        assert c.theme == "light"

    def test_load_partial_json_fills_defaults(self, tmp_path):
        data = {"theme": "sepia"}
        p = tmp_path / "settings.json"
        p.write_text(json.dumps(data), encoding="utf-8")
        c = Config.load(str(p))
        assert c.theme == "sepia"
        assert c.output_folder == "attendance_data"  # default

    def test_load_managers(self, tmp_path):
        data = {"managers": [{"name": "部長", "teams_id": "bucho@example.com"}]}
        p = tmp_path / "settings.json"
        p.write_text(json.dumps(data), encoding="utf-8")
        c = Config.load(str(p))
        assert len(c.managers) == 1
        assert c.managers[0]["name"] == "部長"

    def test_load_shift_types(self, tmp_path):
        data = {"shift_types": ["日勤", "早番", "深夜"]}
        p = tmp_path / "settings.json"
        p.write_text(json.dumps(data), encoding="utf-8")
        c = Config.load(str(p))
        assert "日勤" in c.shift_types


class TestConfigSave:
    def test_save_creates_file(self, tmp_path):
        p = tmp_path / "configs" / "settings.json"
        c = Config()
        c.display_name = "テスト"
        c.save(str(p))
        assert p.exists()

    def test_save_creates_parent_dirs(self, tmp_path):
        p = tmp_path / "deep" / "nested" / "settings.json"
        Config().save(str(p))
        assert p.exists()

    def test_save_and_reload_roundtrip(self, tmp_path):
        p = tmp_path / "settings.json"
        c = Config()
        c.display_name = "往復テスト"
        c.theme = "green"
        c.shift_types = ["日勤", "早番"]
        c.managers = [{"name": "部長", "teams_id": "x@example.com"}]
        c.save(str(p))
        c2 = Config.load(str(p))
        assert c2.display_name == "往復テスト"
        assert c2.theme == "green"
        assert c2.shift_types == ["日勤", "早番"]
        assert c2.managers[0]["name"] == "部長"

    def test_save_utf8_encoding(self, tmp_path):
        p = tmp_path / "settings.json"
        c = Config()
        c.display_name = "日本語テスト"
        c.save(str(p))
        raw = p.read_text(encoding="utf-8")
        assert "日本語テスト" in raw



class TestConfigToDict:
    def test_to_dict_keys(self):
        d = Config().to_dict()
        expected_keys = {
            "ad_name", "display_name", "teams_user_id", "shift_display_name",
            "timesheet_display_name", "webhook_url", "timesheet_folder",
            "output_folder", "theme", "shift_types", "managers", "proxy_sh", "test_date",
        }
        assert expected_keys == set(d.keys())

    def test_to_dict_values_match(self):
        c = Config()
        c.display_name = "テスト"
        c.theme = "dark"
        d = c.to_dict()
        assert d["display_name"] == "テスト"
        assert d["theme"] == "dark"
