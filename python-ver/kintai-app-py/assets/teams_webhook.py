"""Teams Webhook 投稿"""
import json
import subprocess
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False


def send_teams_post(config, message_type: str, data: Dict[str, Any]) -> None:
    """
    TeamsにPOST
    message_type: "clock_in" | "clock_out"
    webhook_url 未設定時は何もしない。失敗時はExceptionをraise
    """
    if not config or not config.webhook_url:
        return

    user_name = config.display_name or ""
    user_id   = config.teams_user_id or ""

    if message_type == "clock_in":
        payload = _build_clock_in_payload(user_name, user_id, data)
    elif message_type == "clock_out":
        payload = _build_clock_out_payload(config, user_name, user_id, data)
    else:
        return

    _save_debug_json(payload)
    _post(config, payload)


# ─────────────────────────── ペイロード構築 ───────────────────────────

def _build_clock_in_payload(
    user_name: str,
    user_id: str,
    data: Dict[str, Any],
) -> Dict[str, Any]:
    """出勤用ペイロード"""
    check_type = "出勤"
    work_mode  = data.get("work_style", "")
    comment    = data.get("comment", "") or ""
    mention_data: List = []

    column_obj  = _build_column_obj(user_name, check_type)
    message_obj = {
        "type":    "TextBlock",
        "text":    f"業務を開始します({work_mode})",
        "size":    "Medium",
        "wrap":    True,
        "spacing": "None",
    }
    comment_obj = _build_comment_obj(comment, mention_data)

    # 退勤でないのでmessageへのcomment埋め込みは不要
    mention_arr = [m for m in mention_data if m]

    return _assemble_payload(user_id, column_obj, message_obj, comment_obj, mention_arr)


def _build_clock_out_payload(
    config,
    user_name: str,
    user_id: str,
    data: Dict[str, Any],
) -> Dict[str, Any]:
    """退勤用ペイロード"""
    check_type     = "退勤"
    next_workday   = data.get("next_workday")
    next_shift     = data.get("next_shift", "") or ""
    next_work_mode = data.get("next_work_mode", "") or ""
    mention_name   = data.get("mention", "") or ""
    comment        = data.get("comment", "") or ""

    # 次回出勤日テキスト
    next_date_text = _format_date_short(next_workday) if next_workday else ""

    # メンションデータ（mention_data は必ずリスト）
    mention_data: List = []
    if mention_name == "@All管理職":
        mention_data = [
            m.get("teams_id", "")
            for m in (config.managers or [])
            if m.get("teams_id", "")
        ]
    elif mention_name:
        manager = next(
            (m for m in (config.managers or []) if m.get("name") == mention_name),
            None,
        )
        if manager:
            tid = manager.get("teams_id", "")
            if tid:
                mention_data = [tid]

    column_obj  = _build_column_obj(user_name, check_type)
    message_obj = {
        "type":    "Container",
        "spacing": "None",
        "items": [
            {
                "type":    "TextBlock",
                "text":    f"退勤します。次回は{next_date_text} {next_work_mode}({next_shift})です。",
                "size":    "Medium",
                "wrap":    True,
                "spacing": "None",
            },
            {
                "type":    "TextBlock",
                "text":    "お疲れさまでした。",
                "wrap":    True,
                "spacing": "None",
            },
        ],
    }
    comment_obj = _build_comment_obj(comment, mention_data)
    # 退勤 かつ commentObj が空でない かつ メンションなし
    # → commentObj を messageObj.items に追加
    if comment_obj and not mention_data:
        message_obj["items"].append(comment_obj)

    mention_arr = [m for m in mention_data if m]

    return _assemble_payload(user_id, column_obj, message_obj, comment_obj, mention_arr)


# ─────────────────────────── 部品 ───────────────────────────

def _build_column_obj(user_name: str, check_type: str) -> Dict[str, Any]:
    return {
        "type":  "Column",
        "width": "stretch",
        "items": [
            {
                "type":                    "TextBlock",
                "text":                    f"{user_name}が{check_type}しました",
                "size":                    "Medium",
                "wrap":                    True,
                "weight":                  "Bolder",
                "verticalContentAlignment": "Center",
            }
        ],
    }


def _build_comment_obj(comment: str, mention_data: List) -> Dict[str, Any]:
    """
      - コメント空            → {} (空 dict)
      - コメントあり + メンションあり → spacing="None"
      - コメントあり + メンションなし → spacing="Small", separator=True
    """
    if not comment or not comment.strip():
        return {}
    if mention_data:
        return {
            "type":    "TextBlock",
            "text":    f"コメント: {comment}",
            "wrap":    True,
            "spacing": "None",
        }
    return {
        "type":      "TextBlock",
        "text":      f"コメント: {comment}",
        "wrap":      True,
        "spacing":   "Small",
        "separator": True,
    }


def _assemble_payload(
    user_id: str,
    column_obj: Dict,
    message_obj: Dict,
    comment_obj: Dict,
    mention_arr: List,
) -> Dict[str, Any]:
    """
    column / message / comment はJSON文字列として埋め込む
    """
    return {
        "mention_data": mention_arr,          # 必ずリスト（空でも []）
        "userId":       user_id,
        "column":       json.dumps(column_obj,  ensure_ascii=False, separators=(",", ":")),
        "message":      json.dumps(message_obj, ensure_ascii=False, separators=(",", ":")),
        "comment":      json.dumps(comment_obj, ensure_ascii=False, separators=(",", ":")),
    }


# ─────────────────────────── HTTP 送信 ───────────────────────────

def _post(config, payload: Dict[str, Any]) -> None:
    """ペイロードをJSONとしてPOST。失敗時は Exceptionをraise"""
    proxies = _get_proxies(config)
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json; charset=utf-8"}

    if REQUESTS_AVAILABLE:
        resp = requests.post(
            config.webhook_url,
            data=body,
            headers=headers,
            proxies=proxies,
            timeout=10,
        )
        if resp.status_code not in (200, 202):
            raise Exception(f"HTTP {resp.status_code}: {resp.text[:200]}")
    else:
        req = urllib.request.Request(
            config.webhook_url,
            data=body,
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status not in (200, 202):
                raise Exception(f"HTTP {resp.status}")


# ─────────────────────────── ユーティリティ ───────────────────────────

def _get_proxies(config) -> Optional[Dict[str, str]]:
    """proxy.sh をsourceして環境変数からプロキシ設定を取得する"""
    proxy_sh = getattr(config, "proxy_sh", "") if config else ""
    if not proxy_sh or not Path(proxy_sh).exists():
        return None
    try:
        result = subprocess.run(
            ["bash", "-c", f"source '{proxy_sh}' && env"],
            capture_output=True, text=True, timeout=5,
        )
        env_vars: Dict[str, str] = {}
        for line in result.stdout.splitlines():
            if "=" in line:
                key, _, val = line.partition("=")
                env_vars[key] = val
        http  = env_vars.get("http_proxy")  or env_vars.get("HTTP_PROXY")
        https = env_vars.get("https_proxy") or env_vars.get("HTTPS_PROXY")
        if http or https:
            return {"http": http or "", "https": https or ""}
    except Exception:
        pass
    return None


def _save_debug_json(payload: Dict[str, Any]) -> None:
    """デバッグ用にJSONを保存する"""
    try:
        debug_dir = Path("timesheet")
        debug_dir.mkdir(exist_ok=True)
        with open(debug_dir / "teams_post_debug.json", "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _format_date_short(d: date) -> str:
    """date を 'M月D日(曜)' 形式にフォーマット"""
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    return f"{d.month}月{d.day}日({weekdays[d.weekday()]})"
