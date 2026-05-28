# -*- coding: utf-8 -*-
r"""批量重新授权 Sub2API OAuth 账号并原地更新导出文件。

示例：
    python -u reauthorize_sub2api.py --input ".\reauthorize\sub2api-account-20260528103145.json" --dry-run
    python -u reauthorize_sub2api.py --input ".\reauthorize\sub2api-account-20260528103145.json" --limit 1 --proxy "7897" --debug

进度文件只记录账号索引、邮箱和状态，不保存 access_token / refresh_token / OAuth code。
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import gpt_free_core as bot_module
from gpt_free_core import (
    ChatGPTBot,
    DEFAULT_OAUTH_PROXY,
    OUTLOOK_EMAIL_PLUS_API_BASE,
    OUTLOOK_EMAIL_PLUS_API_KEY,
    OUTLOOK_EMAIL_PLUS_CALLER_ID,
    OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
    OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
    OUTLOOK_EMAIL_PLUS_PROVIDER,
    Sub2APIExporter,
    normalize_proxy_url,
    utc_now_iso,
)
from outlook_email_plus_integration import OutlookEmailPlusClient


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
TOKENISH_REPLACEMENTS = [
    re.compile(r'("(?:access_token|refresh_token|id_token|code)"\s*:\s*")[^"]+("?)', re.IGNORECASE),
    re.compile(r"((?:access_token|refresh_token|id_token|code)=)[^&\s]+", re.IGNORECASE),
    re.compile(r"(Bearer\s+)[A-Za-z0-9._\-]+", re.IGNORECASE),
]


class ReauthorizeError(Exception):
    """重授权输入、进度或写入流程不满足要求。"""


class ReauthorizeSkip(Exception):
    """当前账号不适合继续重授权，应跳过并记录。"""


def load_json_file(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as f:
        payload = json.load(f)
    if not isinstance(payload, dict):
        raise ReauthorizeError(f"JSON 顶层必须是对象: {path}")
    return payload


def load_sub2api_payload(path: Path) -> dict[str, Any]:
    payload = load_json_file(path)
    accounts = payload.get("accounts")
    if not isinstance(accounts, list):
        raise ReauthorizeError(f"Sub2API 文件缺少 accounts 数组: {path}")
    return payload


def write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_name(f".{path.name}.tmp-{os.getpid()}")
    try:
        with tmp_path.open("w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def sanitize_error_text(value: object) -> str:
    text = str(value)
    for pattern in TOKENISH_REPLACEMENTS:
        text = pattern.sub(r"\1***", text)
    return text[:500]


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.@+-]+", "_", value.strip())
    return cleaned[:120] or "account"


def mask_email(email: str) -> str:
    email = str(email or "").strip()
    if "@" not in email:
        return email or "<unknown>"
    local, domain = email.split("@", 1)
    if len(local) <= 2:
        return f"{local[:1]}*@{domain}"
    return f"{local[:2]}****@{domain}"


def get_account_email(account: dict[str, Any]) -> str:
    credentials = account.get("credentials") if isinstance(account.get("credentials"), dict) else {}
    extra = account.get("extra") if isinstance(account.get("extra"), dict) else {}
    candidates = [
        account.get("name"),
        account.get("email"),
        extra.get("email"),
        credentials.get("email"),
    ]
    for candidate in candidates:
        email = str(candidate or "").strip()
        if EMAIL_RE.match(email):
            return email
    return ""


def account_key(index: int, account: dict[str, Any]) -> str:
    return str(index)


def load_checkpoint(path: Path, input_path: Path) -> dict[str, Any]:
    if not path.exists():
        return {
            "version": 1,
            "input_path": str(input_path),
            "started_at": utc_now_iso(),
            "updated_at": utc_now_iso(),
            "accounts": {},
        }
    payload = load_json_file(path)
    if not isinstance(payload.get("accounts"), dict):
        raise ReauthorizeError(f"进度文件格式无效，缺少 accounts 对象: {path}")
    return payload


def save_checkpoint(path: Path, checkpoint: dict[str, Any]) -> None:
    checkpoint["updated_at"] = utc_now_iso()
    write_json_atomic(path, checkpoint)


def should_process(status: str, force: bool, retry_failed: bool) -> bool:
    if force:
        return True
    if status == "completed":
        return False
    if status in ("skipped_email_verification", "skipped_phone_verification"):
        return False
    if status == "failed" and not retry_failed:
        return False
    return True


def update_key_set(target: dict[str, Any], canonical: str, aliases: list[str], value: Any) -> None:
    keys = [canonical] + aliases
    existing = [key for key in keys if key in target]
    if existing:
        for key in existing:
            target[key] = value
        return
    target[canonical] = value


def merge_refreshed_credentials(account: dict[str, Any], token_data: dict[str, Any]) -> None:
    credentials = account.get("credentials")
    if not isinstance(credentials, dict):
        raise ReauthorizeError("账号缺少 credentials 对象，无法保形写回 token")

    refreshed = Sub2APIExporter._build_credentials(token_data)
    update_key_set(credentials, "access_token", ["at", "accessToken", "token"], refreshed["access_token"])
    update_key_set(credentials, "refresh_token", ["rt", "refreshToken"], refreshed["refresh_token"])
    update_key_set(credentials, "expires_at", ["time", "token_expires_at", "expired"], refreshed["expires_at"])
    credentials["email"] = refreshed["email"]

    optional_keys = [
        "id_token",
        "client_id",
        "chatgpt_account_id",
        "chatgpt_user_id",
        "organization_id",
        "plan_type",
        "subscription_expires_at",
    ]
    for key in optional_keys:
        if refreshed.get(key):
            credentials[key] = refreshed[key]

    if "updated_at" in credentials:
        credentials["updated_at"] = utc_now_iso()
    if "extra" in account and isinstance(account["extra"], dict) and "email" in account["extra"]:
        account["extra"]["email"] = refreshed["email"]


def select_accounts(
    payload: dict[str, Any],
    checkpoint: dict[str, Any],
    start_index: int,
    limit: int,
    only_emails: set[str],
    force: bool,
    retry_failed: bool,
) -> list[tuple[int, dict[str, Any], str, str]]:
    accounts = payload.get("accounts") or []
    selected: list[tuple[int, dict[str, Any], str, str]] = []
    progress_accounts = checkpoint.get("accounts") if isinstance(checkpoint.get("accounts"), dict) else {}
    normalized_only = {email.lower() for email in only_emails}

    for index, account in enumerate(accounts):
        if index < start_index or not isinstance(account, dict):
            continue
        email = get_account_email(account)
        if normalized_only and email.lower() not in normalized_only:
            continue
        key = account_key(index, account)
        legacy_state = progress_accounts.get(key) or progress_accounts.get(f"{index}:{email}") or {}
        status = str(legacy_state.get("status") or "pending")
        if not should_process(status, force=force, retry_failed=retry_failed):
            continue
        if str(account.get("platform") or "").lower() not in ("", "openai"):
            continue
        if str(account.get("type") or "").lower() not in ("", "oauth"):
            continue
        if not email:
            continue
        selected.append((index, account, key, email))
        if limit > 0 and len(selected) >= limit:
            break
    return selected


def mark_progress(
    checkpoint: dict[str, Any],
    key: str,
    index: int,
    email: str,
    status: str,
    error: str = "",
) -> None:
    accounts = checkpoint.setdefault("accounts", {})
    entry = accounts.setdefault(key, {"index": index, "email": email, "attempts": 0})
    if status == "oauth_started":
        entry["attempts"] = int(entry.get("attempts") or 0) + 1
        entry["started_at"] = utc_now_iso()
    if status == "completed":
        entry["completed_at"] = utc_now_iso()
    if status == "failed":
        entry["failed_at"] = utc_now_iso()
    entry["index"] = index
    entry["email"] = email
    entry["status"] = status
    entry["last_error"] = sanitize_error_text(error) if error else ""
    checkpoint["current"] = {"index": index, "email": email, "stage": status}


async def reauthorize_one_account(
    email: str,
    code_file: Path,
    proxy_url: str,
    headless: bool,
    slow_mo: int,
    code_timeout: int,
) -> dict[str, Any]:
    code_file.parent.mkdir(parents=True, exist_ok=True)
    if code_file.exists():
        code_file.unlink()

    async def poll_and_write_code() -> None:
        client = OutlookEmailPlusClient(
            base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
            api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
            caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
            task_id=f"sub2api-reauth-{int(datetime.now().timestamp())}",
            provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
            project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
            email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
        )
        try:
            await client.open()
            client.use_existing_email(email)
            code_info = await client.poll_for_code(timeout=code_timeout, interval=5)
            code_file.write_text(str(code_info["code"]).strip(), encoding="utf-8")
            print(f"已从 OutlookEmailPlus 获取 OAuth 验证码: source={code_info.get('source')}")
        finally:
            await client.close()

    poll_task: asyncio.Task[None] | None = None
    if OUTLOOK_EMAIL_PLUS_API_KEY:
        poll_task = asyncio.create_task(poll_and_write_code())
    else:
        print("OutlookEmailPlus API Key 未配置，将等待手动写入 OAuth 验证码文件。")

    try:
        async with ChatGPTBot(headless=headless, slow_mo=slow_mo, proxy_url=proxy_url) as bot:
            bot.oauth_code_file = str(code_file)

            async def watch_phone_verification() -> None:
                while True:
                    current_url = str(getattr(bot.page, "url", "") or "") if bot.page else ""
                    if "auth.openai.com/phone-verification" in current_url:
                        raise ReauthorizeSkip("OAuth 邮箱验证后进入 phone-verification，跳过该账号")
                    await asyncio.sleep(0.5)

            oauth_task = asyncio.create_task(bot.export_codex_sub2api_token(email))
            watch_task = asyncio.create_task(watch_phone_verification())
            done, pending = await asyncio.wait({oauth_task, watch_task}, return_when=asyncio.FIRST_COMPLETED)
            for task in pending:
                task.cancel()
            for task in pending:
                try:
                    await task
                except asyncio.CancelledError:
                    pass
            if oauth_task in done:
                return oauth_task.result()
            return watch_task.result()
    finally:
        if poll_task and not poll_task.done():
            poll_task.cancel()
            try:
                await poll_task
            except asyncio.CancelledError:
                pass


async def run_reauthorization(args: argparse.Namespace) -> int:
    input_path = Path(args.input).resolve()
    checkpoint_path = Path(args.checkpoint).resolve() if args.checkpoint else input_path.with_suffix(input_path.suffix + ".checkpoint.json")
    code_dir = Path(args.code_dir).resolve() if args.code_dir else input_path.parent / "oauth-codes"
    payload = load_sub2api_payload(input_path)
    checkpoint = load_checkpoint(checkpoint_path, input_path)
    selected = select_accounts(
        payload,
        checkpoint,
        start_index=max(0, args.start_index),
        limit=max(0, args.limit),
        only_emails=set(args.only_email or []),
        force=args.force,
        retry_failed=args.retry_failed,
    )

    print(f"输入文件: {input_path}")
    print(f"进度文件: {checkpoint_path}")
    print(f"账号总数: {len(payload['accounts'])} | 本次待处理: {len(selected)}")
    if args.dry_run:
        for index, _, _, email in selected:
            print(f"DRY-RUN [{index + 1}/{len(payload['accounts'])}] {mask_email(email)}")
        return 0

    if selected and not args.no_backup:
        backup_path = input_path.with_suffix(input_path.suffix + f".bak-{datetime.now().strftime('%Y%m%d%H%M%S')}")
        shutil.copy2(input_path, backup_path)
        print(f"已创建备份: {backup_path}")

    proxy_url = normalize_proxy_url(args.proxy)
    # 重授权是批量任务：授权成功后必须关闭浏览器，避免多账号处理时残留 Chromium。
    bot_module.DEBUG_MODE = False
    success = 0
    failed = 0
    skipped = 0

    account_retries = max(1, int(args.account_retries or 1))
    for position, (index, account, key, email) in enumerate(selected, start=1):
        masked_email = mask_email(email)
        print(f"[{position}/{len(selected)}] 开始重新授权: index={index} email={masked_email}")
        last_error = ""
        completed = False
        for attempt in range(1, account_retries + 1):
            try:
                logger_stage = f"oauth_started_{attempt}" if attempt > 1 else "oauth_started"
                mark_progress(checkpoint, key, index, email, logger_stage)
                save_checkpoint(checkpoint_path, checkpoint)
                code_file = code_dir / f"oauth-code-{index:04d}.txt"
                print(f"[{position}/{len(selected)}] 第 {attempt}/{account_retries} 次 OAuth，验证码文件: {code_file}")
                token_data = await asyncio.wait_for(
                    reauthorize_one_account(email, code_file, proxy_url, args.headless, args.slow_mo, args.code_timeout),
                    timeout=args.flow_timeout,
                )

                mark_progress(checkpoint, key, index, email, "token_received")
                save_checkpoint(checkpoint_path, checkpoint)
                merge_refreshed_credentials(account, token_data)
                payload["exported_at"] = utc_now_iso()
                write_json_atomic(input_path, payload)

                mark_progress(checkpoint, key, index, email, "completed")
                save_checkpoint(checkpoint_path, checkpoint)
                success += 1
                completed = True
                print(f"[{position}/{len(selected)}] 完成重新授权: {masked_email}")
                break
            except KeyboardInterrupt:
                raise
            except ReauthorizeSkip as exc:
                last_error = sanitize_error_text(exc)
                print(f"[{position}/{len(selected)}] 检测到 phone-verification，跳过账号: {masked_email} | {last_error}")
                mark_progress(checkpoint, key, index, email, "skipped_phone_verification", last_error)
                save_checkpoint(checkpoint_path, checkpoint)
                skipped += 1
                completed = True
                break
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {sanitize_error_text(exc)}"
                print(f"[{position}/{len(selected)}] 第 {attempt}/{account_retries} 次失败，将从头重试该账号: {masked_email} | {last_error}")
                if attempt < account_retries:
                    await asyncio.sleep(2)

        if not completed:
            failed += 1
            mark_progress(checkpoint, key, index, email, "failed", last_error)
            save_checkpoint(checkpoint_path, checkpoint)
            print(f"[{position}/{len(selected)}] 重新授权失败: {masked_email} | {last_error}")
            if args.fail_fast:
                raise ReauthorizeError(last_error)

    checkpoint.pop("current", None)
    save_checkpoint(checkpoint_path, checkpoint)
    print(f"重授权结束: 成功={success} 失败={failed} 检测跳过={skipped} 未处理={len(payload['accounts']) - len(selected)}")
    return 1 if failed else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从 Sub2API 导出文件批量重新授权 Codex OAuth 并原地写回 token")
    parser.add_argument("--input", required=True, help="Sub2API 导出 JSON 文件")
    parser.add_argument("--checkpoint", default="", help="进度文件；默认 <input>.checkpoint.json")
    parser.add_argument("--code-dir", default="", help="OAuth 验证码文件目录；默认输入文件旁 oauth-codes")
    parser.add_argument("--proxy", default=DEFAULT_OAUTH_PROXY, help="浏览器和 token exchange 代理，支持 7897 简写")
    parser.add_argument("--headless", action="store_true", help="无头模式")
    parser.add_argument("--slow-mo", type=int, default=100, help="Playwright 操作延迟(ms)")
    parser.add_argument("--debug", action="store_true", help="保留兼容参数；重授权成功后始终关闭浏览器")
    parser.add_argument("--code-timeout", type=int, default=900, help="OutlookEmailPlus 自动等待 OAuth 邮箱验证码超时秒数")
    parser.add_argument("--account-retries", type=int, default=2, help="单个账号失败后从 OAuth 起点重试次数")
    parser.add_argument("--flow-timeout", type=int, default=180, help="单次 OAuth 授权尝试硬超时秒数；超时后重试或进入下一个账号")
    parser.add_argument("--dry-run", action="store_true", help="只扫描待处理账号，不启动浏览器、不写文件")
    parser.add_argument("--start-index", type=int, default=0, help="从 accounts 的 0-based 索引开始")
    parser.add_argument("--limit", type=int, default=0, help="最多处理账号数；0 表示不限制")
    parser.add_argument("--only-email", action="append", default=[], help="仅处理指定邮箱；可重复传入")
    parser.add_argument("--force", action="store_true", help="忽略 completed 进度，强制重新处理")
    parser.add_argument("--retry-failed", action="store_true", help="重试进度文件中 failed 的账号")
    parser.add_argument("--fail-fast", action="store_true", help="遇到第一个失败立即终止")
    parser.add_argument("--no-backup", action="store_true", help="写回前不创建 .bak 备份")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    try:
        raise SystemExit(asyncio.run(run_reauthorization(args)))
    except KeyboardInterrupt:
        print("用户中断，已保存当前进度。")
        raise SystemExit(130)
    except (ReauthorizeError, json.JSONDecodeError, OSError) as exc:
        print(f"重授权启动失败: {sanitize_error_text(exc)}")
        raise SystemExit(2)


if __name__ == "__main__":
    main()
