# -*- coding: utf-8 -*-
r"""将 cdkey-files 里的 OAuth JSON 转换为 Sub2API 账号文件。

示例：
    python convert_cdkey_json_to_sub2api.py "E:\Users\admin\Desktop\cdkey-files\xxx.json"
    python convert_cdkey_json_to_sub2api.py "E:\Users\admin\Desktop\cdkey-files"
    python convert_cdkey_json_to_sub2api.py "E:\Users\admin\Desktop\cdkey-files\xxx.json" --output sub2api-accounts.json --append
"""
import argparse
import json
import os
from datetime import datetime, timezone
from typing import Any, Dict, List, Tuple


DEFAULT_OUTPUT = "sub2api-accounts.json"


class ConvertError(Exception):
    """转换输入不满足 Sub2API 必要字段时抛出。"""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def load_json_file(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    if not isinstance(payload, dict):
        raise ConvertError("输入 JSON 顶层必须是对象")
    return payload


def load_sub2api_payload(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {"exported_at": utc_now_iso(), "proxies": [], "accounts": []}

    payload = load_json_file(path)
    accounts = payload.get("accounts")
    if not isinstance(accounts, list):
        raise ConvertError(f"输出文件不是有效 Sub2API 格式，缺少 accounts 数组: {path}")

    proxies = payload.get("proxies")
    return {
        "exported_at": payload.get("exported_at") or utc_now_iso(),
        "proxies": proxies if isinstance(proxies, list) else [],
        "accounts": accounts,
    }


def collect_input_files(input_path: str, recursive: bool) -> List[str]:
    if os.path.isfile(input_path):
        return [input_path]
    if not os.path.isdir(input_path):
        raise ConvertError(f"输入路径不存在: {input_path}")

    files: List[str] = []
    if recursive:
        for root, _, names in os.walk(input_path):
            for name in names:
                if name.lower().endswith(".json"):
                    files.append(os.path.join(root, name))
    else:
        for name in os.listdir(input_path):
            path = os.path.join(input_path, name)
            if os.path.isfile(path) and name.lower().endswith(".json"):
                files.append(path)

    files.sort(key=lambda item: item.lower())
    if not files:
        raise ConvertError(f"目录下没有可转换的 .json 文件: {input_path}")
    return files


def unique_name(base_name: str, accounts: List[Dict[str, Any]]) -> str:
    used = {str(item.get("name", "")) for item in accounts if isinstance(item, dict)}
    if base_name not in used:
        return base_name

    idx = 2
    while f"{base_name}-{idx}" in used:
        idx += 1
    return f"{base_name}-{idx}"


def build_credentials(source: Dict[str, Any]) -> Dict[str, Any]:
    if source.get("expires_at") and source.get("expired") and source["expires_at"] != source["expired"]:
        raise ConvertError("输入同时包含 expires_at 与 expired，且值不一致，请先确认有效期字段")

    expires_at = source.get("expires_at") or source.get("expired")
    required = {
        "access_token": source.get("access_token"),
        "refresh_token": source.get("refresh_token"),
        "expires_at": expires_at,
        "email": source.get("email"),
    }
    missing = [key for key, value in required.items() if not value]
    if missing:
        raise ConvertError(f"Sub2API 转换缺少必要字段: {', '.join(missing)}")

    credentials = dict(required)
    optional_mapping = {
        "id_token": "id_token",
        "client_id": "client_id",
        "account_id": "chatgpt_account_id",
        "chatgpt_account_id": "chatgpt_account_id",
        "chatgpt_user_id": "chatgpt_user_id",
        "organization_id": "organization_id",
        "plan_type": "plan_type",
        "subscription_expires_at": "subscription_expires_at",
    }
    for src_key, dst_key in optional_mapping.items():
        value = source.get(src_key)
        if value and dst_key not in credentials:
            credentials[dst_key] = value

    return credentials


def build_account(source: Dict[str, Any], accounts: List[Dict[str, Any]], notes: str) -> Dict[str, Any]:
    email = str(source.get("email") or "").strip()
    name = unique_name(email or f"openai-{int(datetime.now().timestamp())}", accounts)
    account = {
        "name": name,
        "platform": "openai",
        "type": "oauth",
        "credentials": build_credentials(source),
        "concurrency": 10,
        "priority": 1,
        "rate_multiplier": 1.0,
        "auto_pause_on_expired": True,
    }
    if notes:
        account["notes"] = notes
    return account


def write_payload(output_path: str, payload: Dict[str, Any]) -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")


def convert_files(input_files: List[str], output_path: str, append: bool, notes: str, skip_invalid: bool) -> Tuple[str, int, List[str]]:
    if append:
        payload = load_sub2api_payload(output_path)
    else:
        payload = {"exported_at": utc_now_iso(), "proxies": [], "accounts": []}

    errors: List[str] = []
    converted_count = 0
    for input_file in input_files:
        try:
            source = load_json_file(input_file)
            payload["accounts"].append(build_account(source, payload["accounts"], notes))
            converted_count += 1
        except (ConvertError, json.JSONDecodeError, OSError) as e:
            message = f"{input_file}: {e}"
            if not skip_invalid:
                raise ConvertError(message) from e
            errors.append(message)

    if converted_count == 0:
        raise ConvertError("没有成功转换任何账号")

    payload["exported_at"] = utc_now_iso()
    write_payload(output_path, payload)
    return output_path, converted_count, errors


def convert_file(input_path: str, output_path: str, append: bool, notes: str) -> str:
    result, _, _ = convert_files([input_path], output_path, append, notes, skip_invalid=False)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="将 cdkey OAuth JSON 转换为 Sub2API accounts JSON")
    parser.add_argument("input", help="源 JSON 文件或文件夹路径")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"输出文件，默认 {DEFAULT_OUTPUT}")
    parser.add_argument("--append", action="store_true", help="追加到已有 Sub2API 文件；默认覆盖输出文件")
    parser.add_argument("--notes", default="cdkey JSON 转换导出", help="写入账号 notes 字段")
    parser.add_argument("--recursive", action="store_true", help="输入为文件夹时递归扫描子目录")
    parser.add_argument("--skip-invalid", action="store_true", help="批量转换时跳过无效 JSON；默认遇到错误立即停止")
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output)
    input_files = collect_input_files(input_path, args.recursive)
    result, converted_count, errors = convert_files(input_files, output_path, args.append, args.notes, args.skip_invalid)
    print(f"Sub2API 文件已生成: {result}")
    print(f"已转换账号数: {converted_count}")
    if errors:
        print(f"已跳过无效文件数: {len(errors)}")
        for error in errors:
            print(f"- {error}")


if __name__ == "__main__":
    main()
