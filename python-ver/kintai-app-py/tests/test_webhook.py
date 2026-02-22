"""assets/teams_webhook.py のユニットテスト"""
import json
import pytest
from datetime import date
from unittest.mock import patch, MagicMock

from assets.teams_webhook import (
    _format_date_short,
    _build_comment_obj,
    _build_column_obj,
    _build_clock_in_payload,
    _build_clock_out_payload,
    _assemble_payload,
    send_teams_post,
)


# ─────────────────────────── _format_date_short ───────────────────────────

class TestFormatDateShort:
    def test_saturday(self):
        assert _format_date_short(date(2026, 2, 21)) == "2/21(土)"

    def test_monday(self):
        assert _format_date_short(date(2026, 2, 23)) == "2/23(月)"

    def test_sunday(self):
        assert _format_date_short(date(2026, 2, 22)) == "2/22(日)"

    def test_single_digit_month_day(self):
        assert _format_date_short(date(2026, 1, 5)) == "1/5(月)"

    def test_all_weekdays(self):
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        base = date(2026, 2, 16)  # 月曜
        for i, name in enumerate(weekdays):
            d = base + __import__("datetime").timedelta(days=i)
            result = _format_date_short(d)
            assert f"({name})" in result


# ─────────────────────────── _build_comment_obj ───────────────────────────

class TestBuildCommentObj:
    def test_empty_comment_returns_empty_dict(self):
        assert _build_comment_obj("", []) == {}

    def test_whitespace_only_returns_empty_dict(self):
        assert _build_comment_obj("   ", []) == {}

    def test_none_like_empty(self):
        assert _build_comment_obj("", ["id1"]) == {}

    def test_comment_with_mention(self):
        r = _build_comment_obj("お願いします", ["id1"])
        assert r["type"] == "TextBlock"
        assert r["spacing"] == "None"
        assert "お願いします" in r["text"]

    def test_comment_without_mention(self):
        r = _build_comment_obj("お願いします", [])
        assert r["type"] == "TextBlock"
        assert r["spacing"] == "Small"
        assert r.get("separator") is True
        assert "お願いします" in r["text"]

    def test_mention_list_empty_vs_none(self):
        """mention_data=[] と mention_data=None相当の区別"""
        r_with = _build_comment_obj("test", ["some_id"])
        r_without = _build_comment_obj("test", [])
        assert r_with["spacing"] == "None"
        assert r_without["spacing"] == "Small"


# ─────────────────────────── _build_column_obj ───────────────────────────

class TestBuildColumnObj:
    def test_clock_in_text(self):
        r = _build_column_obj("山田 太郎", "出勤")
        assert r["type"] == "Column"
        items = r["items"]
        assert len(items) == 1
        assert "山田 太郎が出勤しました" in items[0]["text"]

    def test_clock_out_text(self):
        r = _build_column_obj("鈴木", "退勤")
        assert "鈴木が退勤しました" in r["items"][0]["text"]

    def test_bolder_weight(self):
        r = _build_column_obj("テスト", "出勤")
        assert r["items"][0]["weight"] == "Bolder"


# ─────────────────────────── _build_clock_in_payload ───────────────────────────

class TestBuildClockInPayload:
    def test_basic_structure(self):
        r = _build_clock_in_payload("山田", "yamada@example.com", {
            "shift": "日勤", "work_style": "リモート", "comment": "",
        })
        assert "userId" in r
        assert r["userId"] == "yamada@example.com"
        assert "column" in r
        assert "message" in r
        assert "comment" in r
        assert "mention_data" in r

    def test_message_contains_shift(self):
        r = _build_clock_in_payload("山田", "id", {
            "shift": "早番", "work_style": "出社", "comment": "",
        })
        msg = json.loads(r["message"])
        assert "早番" in msg["text"]

    def test_message_contains_work_style(self):
        r = _build_clock_in_payload("山田", "id", {
            "shift": "日勤", "work_style": "リモート", "comment": "",
        })
        msg = json.loads(r["message"])
        assert "リモート" in msg["text"]

    def test_no_comment_empty_comment_obj(self):
        r = _build_clock_in_payload("山田", "id", {
            "shift": "日勤", "work_style": "リモート", "comment": "",
        })
        comment = json.loads(r["comment"])
        assert comment == {}

    def test_with_comment(self):
        r = _build_clock_in_payload("山田", "id", {
            "shift": "日勤", "work_style": "リモート", "comment": "よろしく",
        })
        comment = json.loads(r["comment"])
        assert "よろしく" in comment.get("text", "")

    def test_mention_data_empty_when_no_mention(self):
        r = _build_clock_in_payload("山田", "id", {
            "shift": "日勤", "work_style": "リモート", "comment": "",
        })
        assert r["mention_data"] == []


# ─────────────────────────── _build_clock_out_payload ───────────────────────────

class TestBuildClockOutPayload:
    def _make_config(self, managers=None):
        from assets.config import Config
        c = Config()
        c.display_name = "山田"
        c.teams_user_id = "yamada@example.com"
        c.managers = managers or []
        return c

    def test_basic_structure(self):
        cfg = self._make_config()
        r = _build_clock_out_payload(cfg, "山田", "yamada@example.com", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "日勤",
            "next_work_mode": "リモート",
            "mention": "",
            "comment": "",
        })
        assert "mention_data" in r
        assert "message" in r

    def test_next_workday_formatted(self):
        cfg = self._make_config()
        r = _build_clock_out_payload(cfg, "山田", "id", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "早番",
            "next_work_mode": "出社",
            "mention": "",
            "comment": "",
        })
        msg = json.loads(r["message"])
        # Container → items[0] に日付テキスト
        text = msg["items"][0]["text"]
        assert "2/22(日)" in text
        assert "早番" in text

    def test_mention_single_manager(self):
        managers = [{"name": "鈴木部長", "teams_id": "suzuki@example.com"}]
        cfg = self._make_config(managers)
        r = _build_clock_out_payload(cfg, "山田", "id", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "日勤",
            "next_work_mode": "リモート",
            "mention": "鈴木部長",
            "comment": "",
        })
        assert "suzuki@example.com" in r["mention_data"]
        assert len(r["mention_data"]) == 1

    def test_mention_all_managers(self):
        managers = [
            {"name": "鈴木部長", "teams_id": "suzuki@example.com"},
            {"name": "田中課長", "teams_id": "tanaka@example.com"},
        ]
        cfg = self._make_config(managers)
        r = _build_clock_out_payload(cfg, "山田", "id", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "日勤",
            "next_work_mode": "リモート",
            "mention": "@All管理職",
            "comment": "",
        })
        assert "suzuki@example.com" in r["mention_data"]
        assert "tanaka@example.com" in r["mention_data"]
        assert len(r["mention_data"]) == 2

    def test_mention_unknown_name_no_mention(self):
        cfg = self._make_config([{"name": "鈴木部長", "teams_id": "suzuki@example.com"}])
        r = _build_clock_out_payload(cfg, "山田", "id", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "日勤",
            "next_work_mode": "リモート",
            "mention": "存在しない人",
            "comment": "",
        })
        assert r["mention_data"] == []

    def test_no_mention(self):
        cfg = self._make_config()
        r = _build_clock_out_payload(cfg, "山田", "id", {
            "next_workday": date(2026, 2, 22),
            "next_shift": "日勤",
            "next_work_mode": "リモート",
            "mention": "",
            "comment": "",
        })
        assert r["mention_data"] == []


# ─────────────────────────── _assemble_payload ───────────────────────────

class TestAssemblePayload:
    def test_keys_present(self):
        col = {"type": "Column", "items": []}
        msg = {"type": "TextBlock", "text": "test"}
        cmt = {}
        r = _assemble_payload("user@example.com", col, msg, cmt, [])
        assert set(r.keys()) == {"mention_data", "userId", "column", "message", "comment"}

    def test_values_are_json_strings(self):
        col = {"type": "Column"}
        msg = {"type": "TextBlock", "text": "hello"}
        cmt = {}
        r = _assemble_payload("id", col, msg, cmt, [])
        # JSON文字列として解析できること
        json.loads(r["column"])
        json.loads(r["message"])
        json.loads(r["comment"])

    def test_mention_data_passed_through(self):
        r = _assemble_payload("id", {}, {}, {}, ["a@b.com", "c@d.com"])
        assert r["mention_data"] == ["a@b.com", "c@d.com"]

    def test_user_id_passed_through(self):
        r = _assemble_payload("yamada@example.com", {}, {}, {}, [])
        assert r["userId"] == "yamada@example.com"


# ─────────────────────────── send_teams_post ───────────────────────────

class TestSendTeamsPost:
    def _make_config(self, url="https://example.com/webhook"):
        from assets.config import Config
        c = Config()
        c.webhook_url = url
        c.display_name = "山田"
        c.teams_user_id = "yamada@example.com"
        c.managers = []
        return c

    def test_no_webhook_url_does_nothing(self):
        cfg = self._make_config(url="")
        with patch("assets.teams_webhook._post") as mock_post:
            send_teams_post(cfg, "clock_in", {"shift": "日勤", "work_style": "リモート", "comment": ""})
            mock_post.assert_not_called()

    def test_none_config_does_nothing(self):
        with patch("assets.teams_webhook._post") as mock_post:
            send_teams_post(None, "clock_in", {})
            mock_post.assert_not_called()

    def test_clock_in_calls_post(self):
        cfg = self._make_config()
        with patch("assets.teams_webhook._post") as mock_post:
            send_teams_post(cfg, "clock_in", {
                "shift": "日勤", "work_style": "リモート", "comment": "",
            })
            mock_post.assert_called_once()

    def test_clock_out_calls_post(self):
        cfg = self._make_config()
        with patch("assets.teams_webhook._post") as mock_post:
            send_teams_post(cfg, "clock_out", {
                "next_workday": date(2026, 2, 22),
                "next_shift": "日勤",
                "next_work_mode": "リモート",
                "mention": "",
                "comment": "",
            })
            mock_post.assert_called_once()

    def test_unknown_message_type_does_nothing(self):
        cfg = self._make_config()
        with patch("assets.teams_webhook._post") as mock_post:
            send_teams_post(cfg, "unknown_type", {})
            mock_post.assert_not_called()
