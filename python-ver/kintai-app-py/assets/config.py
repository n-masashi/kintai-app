"""設定ファイル (settings.json) の読み書きラッパー"""
import json
from pathlib import Path
from typing import List, Dict, Any


class Config:
    """settings.json の読み書きラッパークラス"""

    DEFAULTS = {
        "ad_name": "",
        "display_name": "",
        "teams_user_id": "",
        "shift_display_name": "",
        "timesheet_display_name": "",
        "webhook_url": "",
        "timesheet_folder": "",
        "output_folder": "attendance_data",
        "theme": "light",
        "shift_types": [],
        "managers": [],
        "proxy_sh": "",
        "test_date": "",  # テスト用日付オーバーライド (YYYY-MM-DD)。空文字で無効
    }

    def __init__(self, data: Dict[str, Any] = None):
        d = data or {}
        self.ad_name: str = d.get("ad_name", self.DEFAULTS["ad_name"])
        self.display_name: str = d.get("display_name", self.DEFAULTS["display_name"])
        self.teams_user_id: str = d.get("teams_user_id", self.DEFAULTS["teams_user_id"])
        self.shift_display_name: str = d.get("shift_display_name", self.DEFAULTS["shift_display_name"])
        self.timesheet_display_name: str = d.get("timesheet_display_name", self.DEFAULTS["timesheet_display_name"])
        self.webhook_url: str = d.get("webhook_url", self.DEFAULTS["webhook_url"])
        self.timesheet_folder: str = d.get("timesheet_folder", self.DEFAULTS["timesheet_folder"])
        self.output_folder: str = d.get("output_folder", self.DEFAULTS["output_folder"])
        self.theme: str = d.get("theme", self.DEFAULTS["theme"])
        self.shift_types: List[str] = d.get("shift_types", list(self.DEFAULTS["shift_types"]))
        self.managers: List[Dict[str, str]] = d.get("managers", list(self.DEFAULTS["managers"]))
        self.proxy_sh: str = d.get("proxy_sh", self.DEFAULTS["proxy_sh"])
        self.test_date: str = d.get("test_date", self.DEFAULTS["test_date"])

    @classmethod
    def load(cls, path: str = "settings.json") -> "Config":
        """JSONファイルから設定を読込む。ファイル不在時はデフォルト値を使用。"""
        p = Path(path)
        if p.exists():
            try:
                with open(p, encoding="utf-8") as f:
                    data = json.load(f)
                return cls(data)
            except (json.JSONDecodeError, IOError):
                pass
        return cls()

    def save(self, path: str = "settings.json") -> None:
        """設定をJSONファイルに保存する。"""
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    def to_dict(self) -> Dict[str, Any]:
        """設定をdictに変換する。"""
        return {
            "ad_name": self.ad_name,
            "display_name": self.display_name,
            "teams_user_id": self.teams_user_id,
            "shift_display_name": self.shift_display_name,
            "timesheet_display_name": self.timesheet_display_name,
            "webhook_url": self.webhook_url,
            "timesheet_folder": self.timesheet_folder,
            "output_folder": self.output_folder,
            "theme": self.theme,
            "shift_types": self.shift_types,
            "managers": self.managers,
            "proxy_sh": self.proxy_sh,
            "test_date": self.test_date,
        }
