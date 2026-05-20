# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import asyncio
import gc
import json
import logging
import os
import random
import re
import sqlite3
import subprocess
import tempfile
import time
import uuid
from datetime import datetime, timedelta
from typing import TYPE_CHECKING, Any, Callable

import gpt_free_core as bot_module
from gpt_free_core import (
    DEFAULT_NAME,
    DEFAULT_OAUTH_PROXY,
    HEROSMS_APIKEY,
    HEROSMS_FINISH_AFTER,
    HEROSMS_INTERVAL,
    OUTLOOK_EMAIL_PLUS_API_BASE,
    OUTLOOK_EMAIL_PLUS_API_KEY,
    OUTLOOK_EMAIL_PLUS_CALLER_ID,
    OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
    OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
    OUTLOOK_EMAIL_PLUS_PROVIDER,
    ChatGPTBot,
    SignupFlowError,
    Sub2APIExporter,
    UserAlreadyExistsError,
    VerificationTimeout,
    generate_birthdate,
    generate_email_prefix,
    generate_real_name,
    generate_strong_password,
    is_valid_email,
    normalize_proxy_url,
    retry_with_backoff,
    wait_for_code_file,
)
from hero_sms_client import HeroSMSActivation, HeroSMSClient, HeroSMSError
from outlook_email_plus_integration import (
    OutlookEmailPlusClient,
    OutlookEmailPlusError,
    OutlookEmailPlusTimeout,
    OutlookEmailPlusUnreadable,
)


logger = logging.getLogger(__name__)
if TYPE_CHECKING:
    from playwright.async_api import Locator, Page

COLOMBIA_ISO = "CO"
COLOMBIA_DIAL_CODE = "57"
COLOMBIA_NAME = "Colombia"
HEROSMS_OPENAI_SERVICE = "dr"
HEROSMS_MAX_PRICE_USD = 0.08
ACCOUNT_RECORD_FILE = os.path.abspath("gpt_free_accounts.json")


class BatchAbortError(Exception):
    """批量运行遇到不可恢复前置条件时终止后续轮次。"""


class EmailPoolExhaustedError(BatchAbortError):
    """邮箱池无可用账号，终止批量运行。"""


class SmsVerificationTimeoutError(SignupFlowError):
    """手机号验证码超时，允许同一轮重新获取手机号重试。"""


class PhoneAlreadyRegisteredError(SignupFlowError):
    """手机号已关联既有账号，允许同一轮取消当前号码并换号重试。"""


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _load_json_list(path: str) -> list[dict[str, Any]]:
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        return []
    return data if isinstance(data, list) else []


def _safe_output_name(value: str, fallback: str) -> str:
    """生成可用于文件名的账号标识。"""
    text = (value or "").strip().lower() or fallback
    safe = re.sub(r"[^a-zA-Z0-9._@+-]+", "_", text)
    safe = safe.strip("._ ")
    return safe or fallback


class WindowsNotificationService:
    """通过 Windows 托盘气泡提示关键人工介入和失败事件。"""

    enabled = os.name == "nt"

    @staticmethod
    def _ps_string(value: str) -> str:
        return "'" + str(value or "").replace("'", "''") + "'"

    @classmethod
    def notify(cls, title: str, message: str, timeout_ms: int = 10000) -> None:
        if not cls.enabled:
            return
        title = " ".join(str(title or "ChatGPT 注册机").split())[:63]
        message = " ".join(str(message or "").split())[:255]
        if not message:
            return
        timeout_ms = max(1000, min(int(timeout_ms or 10000), 30000))
        script = "\n".join(
            [
                "Add-Type -AssemblyName System.Windows.Forms",
                "Add-Type -AssemblyName System.Drawing",
                "$notify = New-Object System.Windows.Forms.NotifyIcon",
                "$notify.Icon = [System.Drawing.SystemIcons]::Information",
                f"$notify.BalloonTipTitle = {cls._ps_string(title)}",
                f"$notify.BalloonTipText = {cls._ps_string(message)}",
                "$notify.Visible = $true",
                f"$notify.ShowBalloonTip({timeout_ms})",
                f"Start-Sleep -Milliseconds {timeout_ms + 1000}",
                "$notify.Dispose()",
            ]
        )
        try:
            subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as exc:
            logger.debug(f"Windows 通知发送失败: {exc}")


def notify_human_intervention(title: str, message: str) -> None:
    WindowsNotificationService.notify(f"需要人工介入: {title}", message, timeout_ms=15000)


def notify_failure(title: str, message: str) -> None:
    WindowsNotificationService.notify(f"流程失败: {title}", message, timeout_ms=15000)


def exit_with_param_error(message: str) -> None:
    notify_failure("参数错误", message)
    raise SystemExit(message)


async def drain_windows_asyncio_shutdown() -> None:
    """在 Windows Proactor 关闭前回收残留 transport，减少退出阶段噪音。"""
    if os.name != "nt":
        return
    await asyncio.sleep(0.1)
    gc.collect()
    await asyncio.sleep(0.1)


def upsert_account_record(record: dict[str, Any], output_file: str = ACCOUNT_RECORD_FILE) -> None:
    """保存账号阶段性状态，失败或中断也保留可恢复信息。"""
    record["updated_at"] = _now_iso()
    records = _load_json_list(output_file)
    record_id = str(record.get("run_id") or record.get("activation_id") or record.get("phone") or "")
    replaced = False
    for index, item in enumerate(records):
        item_id = str(item.get("run_id") or item.get("activation_id") or item.get("phone") or "")
        if record_id and item_id == record_id:
            merged = {**item, **record}
            records[index] = merged
            replaced = True
            break
    if not replaced:
        records.append(record.copy())
    with open(output_file, "w", encoding="utf-8") as file:
        json.dump(records, file, ensure_ascii=False, indent=2)
    logger.info(f"账号阶段记录已写入: {output_file}")


async def _poll_manual_code(code_file: str, timeout: int) -> dict[str, str]:
    if code_file:
        code = await wait_for_code_file(code_file, "邮箱验证码", timeout=timeout)
        return {"code": code, "source": "file"}

    notify_human_intervention("邮箱验证码", "自动读取失败，请在终端手动输入邮箱验证码。")
    print("\n>> 已向邮箱发送验证码，请查看邮件并输入验证码。")
    loop = asyncio.get_running_loop()
    while True:
        code = await loop.run_in_executor(None, lambda: input("请输入邮箱验证码 > ").strip())
        if code and code.isdigit() and 4 <= len(code) <= 8:
            return {"code": code, "source": "manual"}
        print("[!] 验证码应为 4-8 位数字，请重新输入。")


async def _poll_hero_sms_code(
    phone: str,
    api_key: str,
    timeout: int,
    interval: int,
    finish_after: bool,
) -> str | None:
    """通过 HeroSMS 当前激活列表按手机号轮询短信验证码。"""
    phone = phone.strip().lstrip("+")
    api_key = api_key.strip()
    if not phone or not api_key:
        return None

    logger.info(f"HeroSMS 接码启动：phone={phone} timeout={timeout}s interval={interval}s")
    try:
        client = HeroSMSClient(api_key)
        result = await client.poll_code_by_phone(
            phone=phone,
            timeout=timeout,
            interval=interval,
            finish_after=finish_after,
        )
        logger.info(f"HeroSMS 已获取验证码: {result.code[:2]}*** (activation={result.activation_id})")
        return result.code
    except HeroSMSError as exc:
        logger.warning(f"HeroSMS 接码失败: {exc}")
        return None


async def _bridge_outlook_code_to_file(
    outlook_client: OutlookEmailPlusClient,
    code_file: str,
    timeout: int,
) -> None:
    """等待 OutlookEmailPlus 验证码并写入 Codex OAuth 文件入口。"""
    code_info = await outlook_client.poll_for_code(timeout=timeout, interval=5)
    with open(code_file, "w", encoding="utf-8") as file:
        file.write(code_info["code"])
    logger.info(f"已将 OutlookEmailPlus 验证码写入 OAuth code file: {code_file}")


async def _keep_browser_open_on_error(bot: ChatGPTBot | None, reason: str) -> None:
    """异常时保持 Playwright 进程和浏览器窗口，便于人工查看。"""
    notify_human_intervention("浏览器保持打开", f"流程异常: {reason}。请查看浏览器窗口和终端日志。")
    if not bot or not bot.page:
        logger.warning("keep-open 已启用，但没有可保留的 Playwright 页面")
        return
    try:
        state = await _read_page_state(bot.page)
        logger.error(f"keep-open 当前页面: url={state.get('url')} title={state.get('title')}")
        logger.error(f"keep-open 页面文本: {str(state.get('text', ''))[:500]}")
    except Exception as exc:
        logger.warning(f"keep-open 读取页面状态失败: {exc}")
    logger.error("keep-open 已启用：浏览器将保持打开。查看完成后在终端按 Ctrl+C 退出。")
    try:
        while True:
            await asyncio.sleep(10)
    except (KeyboardInterrupt, asyncio.CancelledError):
        logger.info(f"keep-open 结束: {reason}")


def _local_phone_number(phone: str, dial_code: str = COLOMBIA_DIAL_CODE) -> str:
    digits = re.sub(r"\D+", "", phone)
    if digits.startswith(dial_code) and len(digits) > len(dial_code) + 4:
        return digits[len(dial_code):]
    return digits


async def _real_mouse_click_locator(page: "Page", locator: "Locator", description: str, timeout: int = 10000) -> None:
    """通过真实鼠标移动和按键点击目标元素中心。"""
    await locator.wait_for(state="visible", timeout=timeout)
    await locator.scroll_into_view_if_needed(timeout=timeout)
    try:
        if await locator.is_disabled(timeout=1000):
            raise SignupFlowError(f"{description} 当前不可点击：按钮已禁用")
    except SignupFlowError:
        raise
    except Exception:
        pass

    box = await locator.bounding_box(timeout=timeout)
    if not box or box["width"] <= 0 or box["height"] <= 0:
        raise SignupFlowError(f"{description} 无可用点击坐标")

    x = box["x"] + box["width"] / 2
    y = box["y"] + box["height"] / 2
    await page.mouse.move(x, y)
    await page.mouse.down()
    await asyncio.sleep(0.05)
    await page.mouse.up()


async def _click_button_by_text(page: "Page", texts: list[str], timeout: int = 10000) -> None:
    deadline = asyncio.get_running_loop().time() + timeout / 1000
    wanted = [str(text).lower() for text in texts]
    last_error = ""
    while asyncio.get_running_loop().time() < deadline:
        buttons = page.locator('button, [role="button"], a')
        count = await buttons.count()
        for index in range(count):
            locator = buttons.nth(index)
            try:
                if not await locator.is_visible(timeout=500):
                    continue
                text = (await locator.inner_text(timeout=500)).strip().lower()
                if text and any(item in text for item in wanted):
                    await _real_mouse_click_locator(page, locator, f"按钮 {texts}", timeout=1000)
                    return
            except SignupFlowError as exc:
                last_error = str(exc)
            except Exception as exc:
                last_error = str(exc)
        await asyncio.sleep(0.5)
    detail = f"; last_error={last_error}" if last_error else ""
    raise SignupFlowError(f"未找到按钮: {texts}{detail}")


async def _click_submit(page: "Page") -> None:
    texts = {"继续", "Continue", "下一步", "Next"}
    buttons = page.locator('button[type="submit"], button')
    count = await buttons.count()
    last_error = ""
    for index in range(count):
        locator = buttons.nth(index)
        try:
            if not await locator.is_visible(timeout=500):
                continue
            text = (await locator.inner_text(timeout=500)).strip()
            if text in texts:
                await _real_mouse_click_locator(page, locator, "继续/提交按钮", timeout=3000)
                await _random_human_delay(1200, 2500)
                return
        except SignupFlowError as exc:
            last_error = str(exc)
        except Exception as exc:
            last_error = str(exc)
    detail = f"; last_error={last_error}" if last_error else ""
    raise SignupFlowError(f"未找到可点击的继续/提交按钮{detail}")


async def _random_human_delay(min_ms: int = 200, max_ms: int = 800) -> None:
    """人类操作间隔：随机延迟模拟思考/操作停顿。"""
    await asyncio.sleep(random.randint(min_ms, max_ms) / 1000)


async def _human_type(page, text: str, min_delay: int = 40, max_delay: int = 120) -> None:
    """模拟人类逐键输入，加入随机键间延迟。"""
    for char in text:
        await page.keyboard.type(char, delay=random.randint(min_delay, max_delay))
    await _random_human_delay(100, 300)


async def _read_page_state(page) -> dict[str, Any]:
    return await page.evaluate(
        """
        () => ({
            url: location.href,
            title: document.title || '',
            text: (document.body?.innerText || '').slice(0, 1600),
            inputs: Array.from(document.querySelectorAll('input:not([type="hidden"])')).map(i => ({
                type: i.type || '', name: i.name || '', placeholder: i.placeholder || '', id: i.id || '',
                autocomplete: i.autocomplete || '', visible: !!i.offsetParent
            })),
            buttons: Array.from(document.querySelectorAll('button')).map(b => ({
                text: (b.innerText || b.textContent || '').trim(), name: b.name || '', type: b.type || '',
                disabled: !!b.disabled, visible: !!b.offsetParent
            })).filter(b => b.text || b.name)
        })
        """
    )


def _is_cloudflare_challenge(state: dict[str, Any]) -> bool:
    """检测 Cloudflare 人机验证页面（Turnstile / JS Challenge）。"""
    text = str(state.get("text") or "")
    url = str(state.get("url") or "")
    title = str(state.get("title") or "")
    lower = f"{text}\n{url}\n{title}".lower()
    cloudflare_keywords = (
        "cloudflare",
        "ray id:",
        "正在验证您是否是真人",
        "验证您不是自动程序",
        "checking your browser",
        "verify you are human",
        "安全服务防护",
    )
    return any(kw in lower for kw in cloudflare_keywords)


async def _wait_for_cloudflare_clear(page, initial_state: dict[str, Any]) -> dict[str, Any]:
    """检测到 Cloudflare 验证后等待人工解决，解决后返回新页面状态。"""
    ray_id = ""
    match = re.search(r"Ray ID:\s*([a-f0-9]+)", str(initial_state.get("text", "")), re.IGNORECASE)
    if match:
        ray_id = match.group(1)

    logger.warning("=" * 60)
    logger.warning("⚠ Cloudflare 人机验证拦截！")
    if ray_id:
        logger.warning(f"   Ray ID: {ray_id}")
    logger.warning(f"   当前 URL: {str(initial_state.get('url', ''))[:120]}")
    logger.warning(f"   页面标题: {str(initial_state.get('title', ''))[:80]}")
    logger.warning("")
    logger.warning("   请在浏览器窗口中手动完成人机验证")
    logger.warning("   完成后脚本将自动继续...")
    logger.warning("   如验证超时 (10分钟)，将抛出异常")
    logger.warning("=" * 60)
    notify_human_intervention("Cloudflare 人机验证", "请在浏览器窗口中手动完成人机验证，脚本会自动继续。")

    deadline = asyncio.get_running_loop().time() + 600
    notify_interval = 30
    next_notify = asyncio.get_running_loop().time() + notify_interval

    while asyncio.get_running_loop().time() < deadline:
        try:
            state = await _read_page_state(page)
            if not _is_cloudflare_challenge(state):
                logger.info("✅ Cloudflare 验证已通过，继续执行...")
                return state
        except Exception as exc:
            logger.debug(f"Cloudflare 清除检测时读取页面失败: {exc}")

        now = asyncio.get_running_loop().time()
        if now >= next_notify:
            waited = int(now - (deadline - 600))
            logger.warning(f"   仍在等待人工完成 Cloudflare 验证... (已等待 {waited}s)")
            next_notify = now + notify_interval

        await asyncio.sleep(2)

    raise SignupFlowError("Cloudflare 人机验证等待超时（10分钟），请检查网络或代理后重试")


async def _wait_until_state(
    page,
    predicate: Callable[[dict[str, Any]], bool],
    description: str,
    timeout: int = 120,
) -> dict[str, Any]:
    deadline = asyncio.get_running_loop().time() + timeout
    last_state: dict[str, Any] = {}
    while asyncio.get_running_loop().time() < deadline:
        try:
            state = await _read_page_state(page)
            last_state = state
            # Cloudflare 检测：检测到后暂停等待人工解决
            if _is_cloudflare_challenge(state):
                logger.warning(f"等待 '{description}' 时遇到 Cloudflare 验证，暂停等待人工解决...")
                state = await _wait_for_cloudflare_clear(page, state)
                last_state = state
                if predicate(state):
                    logger.info(f"已到达页面状态: {description} ({str(state.get('url', ''))[:120]})")
                    return state
                continue
            if predicate(state):
                logger.info(f"已到达页面状态: {description} ({str(state.get('url', ''))[:120]})")
                return state
        except SignupFlowError:
            raise
        except Exception as exc:
            logger.debug(f"等待 {description} 时读取页面失败: {exc}")
        await asyncio.sleep(1)
    raise SignupFlowError(
        f"等待页面状态超时: {description}; url={str(last_state.get('url', ''))[:160]}; "
        f"text={str(last_state.get('text', ''))[:240]}"
    )


# ── OAuth 多步流程状态检测 ──
def _is_oauth_welcome_back(state: dict[str, Any]) -> bool:
    """OAuth：任意欢迎回来页面（排除 chose-an-account 账号选择页）。"""
    url = str(state.get("url") or "")
    if "choose-an-account" in url:
        return False
    text = str(state.get("text") or "")
    return "欢迎回来" in text


def _is_oauth_email_login(state: dict[str, Any]) -> bool:
    """OAuth：邮箱登录页——需要点击'使用电话号码继续'切换到手机登录。"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    # URL 为 /log-in（不带 ?usernameKind）且有 "使用电话号码继续" 按钮
    return (_is_oauth_welcome_back(state) and "使用电话号码继续" in text
            and "usernameKind=phone_number" not in url)


def _is_oauth_phone_login(state: dict[str, Any]) -> bool:
    """OAuth：手机号登录页——需要选择国家、输入手机号。"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    return (_is_oauth_welcome_back(state)
            and ("log-in?usernameKind=phone_number" in url or "国家号码" in text))


def _is_oauth_enter_password(state: dict[str, Any]) -> bool:
    """OAuth 步骤2：输入密码页"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    return ("输入密码" in text and "log-in/password" in url) or (
        _has_visible_password_input(state) and "log-in" in url and "password" in url
    )


def _is_oauth_add_email(state: dict[str, Any]) -> bool:
    """OAuth 步骤3：要求提供电子邮件地址"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    return "要求提供电子邮件地址" in text or "add-email" in url


def _is_oauth_email_verification(state: dict[str, Any]) -> bool:
    """OAuth 步骤4：检查你的收件箱"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    return "检查你的收件箱" in text or "email-verification" in url.split("?")[0]


def _is_oauth_email_in_use(state: dict[str, Any]) -> bool:
    """OAuth：邮箱已达使用上限或已被绑定。"""
    text = str(state.get("text") or "")
    lower = text.lower()
    return any(
        keyword in lower
        for keyword in (
            "email_in_use",
            "already in use",
            "email already",
            "already been used",
            "已被使用",
            "已在使用",
            "邮箱已绑定",
            "电子邮件地址已",
        )
    )


def _is_oauth_email_code_incorrect(state: dict[str, Any]) -> bool:
    """OAuth：邮箱验证码错误，通常是读到了旧验证码。"""
    text = str(state.get("text") or "")
    lower = text.lower()
    return any(
        keyword in lower
        for keyword in (
            "代码不正确",
            "验证码不正确",
            "incorrect code",
            "invalid code",
            "wrong code",
        )
    )


def _is_oauth_account_selector(state: dict[str, Any]) -> bool:
    """OAuth：已登录账号选择页（已有 session cookie 时的快捷路径）。"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    buttons = state.get("buttons") if isinstance(state, dict) else []
    button_items = buttons if isinstance(buttons, list) else []
    has_session_button = any(
        isinstance(item, dict)
        and item.get("visible")
        and str(item.get("name") or "") == "session_id"
        for item in button_items
    )
    account_selector_text = any(
        phrase in text
        for phrase in (
            "选择一个账户以继续",
            "选择一个帐户以继续",
            "选择一个账号以继续",
            "选择一个账户",
            "选择一个帐户",
            "选择一个账号",
            "选择账户",
            "选择帐户",
            "选择账号",
        )
    )
    return (
        "choose-an-account" in url
        or has_session_button
        or ("欢迎回来" in text and account_selector_text)
    )


def _is_oauth_codex_consent(state: dict[str, Any]) -> bool:
    """OAuth 步骤5：使用 ChatGPT 登录到 Codex - 授权确认页"""
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    return "使用 ChatGPT 登录到 Codex" in text or "sign-in-with-chatgpt/codex/consent" in url


async def _wait_for_oauth_consent_or_email_in_use(page, timeout: int = 60) -> str:
    """等待 Codex 授权页或 email_in_use，避免邮箱上限时等满超时。"""
    state = await _wait_until_state(
        page,
        lambda current: (
            _is_oauth_codex_consent(current)
            or _is_oauth_email_in_use(current)
            or _is_oauth_email_code_incorrect(current)
        ),
        "Codex 授权确认页、邮箱上限错误或验证码错误",
        timeout=timeout,
    )
    if _is_oauth_email_in_use(state):
        return "email_in_use"
    if _is_oauth_email_code_incorrect(state):
        return "code_incorrect"
    return "consent"


def _has_visible_password_input(state: dict[str, Any]) -> bool:
    inputs = state.get("inputs") if isinstance(state, dict) else []
    input_items = inputs if isinstance(inputs, list) else []
    return any(isinstance(item, dict) and item.get("visible") and item.get("type") == "password" for item in input_items)


def _is_login_password_state(state: dict[str, Any]) -> bool:
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    lower_text = text.lower()
    return "log-in/password" in url or "忘记了密码" in text or "forgot password" in lower_text


def _is_create_password_state(state: dict[str, Any]) -> bool:
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    title = str(state.get("title") or "")
    lower = f"{url}\n{text}\n{title}".lower()
    return _has_visible_password_input(state) and (
        "create-account/password" in url
        or "创建密码" in text
        or "create password" in lower
    )


def _is_phone_already_registered_state(state: dict[str, Any]) -> bool:
    """检测当前手机号无法用于新注册的明确页面错误。"""
    text = str(state.get("text") or "")
    lower_text = text.lower()
    return any(
        keyword in lower_text
        for keyword in (
            "phone number is already associated",
            "phone number already associated",
            "phone already associated",
            "already associated with this phone",
            "account already exists with this phone",
            "an account already exists with this phone",
        )
    ) or any(
        keyword in text
        for keyword in (
            "与此电话号码相关联的帐户已存在",
            "与此电话号码关联的帐户已存在",
            "与此手机号相关联的帐户已存在",
            "与此手机号关联的帐户已存在",
            "该手机号已关联账号",
            "此手机号已关联账号",
            "手机号已存在账号",
        )
    )


def _is_sms_verification_state(state: dict[str, Any]) -> bool:
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    title = str(state.get("title") or "")
    lower_text = text.lower()
    inputs = state.get("inputs") if isinstance(state, dict) else []
    has_code_input = any(
        isinstance(item, dict)
        and item.get("visible")
        and item.get("name") != "phoneNumberInput"
        and (item.get("autocomplete") == "one-time-code" or "code" in str(item.get("name", "")).lower())
        for item in (inputs if isinstance(inputs, list) else [])
    )
    return (
        "contact-verification" in url
        or "verify" in url
        or "查看你的手机" in title
        or "查看你的手机" in text
        or "check your phone" in lower_text
        or has_code_input
    )


def _is_about_you_state(state: dict[str, Any]) -> bool:
    url = str(state.get("url") or "")
    text = str(state.get("text") or "")
    lower_text = text.lower()
    return "about-you" in url or "about_you" in url or "tell us about" in lower_text or "告诉我们" in text


def _is_signup_modal_state(state: dict[str, Any]) -> bool:
    text = str(state.get("text") or "")
    lower_text = text.lower()
    return (
        "登录或注册" in text
        or "使用电话号码继续" in text
        or "使用电子邮箱继续" in text
        or "Log in or sign up" in text
        or "continue with phone" in lower_text
        or "continue with email" in lower_text
    )


async def _open_signup_modal(page) -> None:
    await page.wait_for_load_state("domcontentloaded", timeout=30000)
    await _random_human_delay(3000, 6000)

    for attempt in range(1, 4):
        state = await _read_page_state(page)
        if _is_signup_modal_state(state):
            return

        logger.info(f"点击 ChatGPT 首页注册按钮 ({attempt}/3)")
        clicked = False
        try:
            signup_button = page.locator('[data-testid="signup-button"]').first
            await _real_mouse_click_locator(page, signup_button, "signup-button", timeout=10000)
            clicked = True
            logger.info("已用真实鼠标点击 signup-button")
        except Exception as exc:
            logger.warning(f"signup-button 点击失败: {exc}")

        if not clicked:
            try:
                await _click_button_by_text(page, ["免费注册", "Sign up for free", "Sign up"], timeout=5000)
                clicked = True
            except Exception as exc:
                logger.warning(f"注册按钮文本点击失败: {exc}")

        if clicked:
            try:
                await _wait_until_state(page, _is_signup_modal_state, "登录或注册弹窗", timeout=8)
                return
            except SignupFlowError:
                logger.warning("注册按钮点击后未出现弹窗，准备重试")
                await _random_human_delay(1500, 3000)

    state = await _read_page_state(page)
    raise SignupFlowError(f"点击注册后未出现登录或注册弹窗: {state}")


async def _select_phone_country(page, dial_code: str, iso_code: str, country_name: str) -> None:
    selected = await page.evaluate(
        """
        ({dialCode, isoCode, countryName}) => {
            const select = document.querySelector('select');
            if (!select) return false;
            const options = Array.from(select.options);
            const target =
                options.find(opt => opt.value === isoCode) ||
                options.find(opt => opt.text.includes(countryName) || opt.text.includes('哥伦比亚') || opt.text.includes('Colombia')) ||
                options.find(opt => opt.text.includes(`+(${dialCode})`) || opt.text.includes(`(${dialCode})`));
            if (!target) return false;
            const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')?.set;
            if (setter) setter.call(select, target.value);
            else select.value = target.value;
            select.dispatchEvent(new Event('input', { bubbles: true }));
            select.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        """,
        {"dialCode": dial_code, "isoCode": iso_code, "countryName": country_name},
    )
    if selected:
        await _random_human_delay(500, 1500)
        return

    logger.warning("未找到标准国家 select，尝试跳过国家选择；如页面选择器变化，请用 --debug 保留浏览器后提供元素信息")


async def _navigate_to_phone_signup(bot: ChatGPTBot, activation: HeroSMSActivation) -> None:
    if not bot.page:
        raise SignupFlowError("浏览器页面未初始化")
    page = bot.page
    logger.info("第 2 步: 打开 ChatGPT 并进入手机号注册")
    await page.goto("https://chatgpt.com/?temporary-chat=true", wait_until="domcontentloaded", timeout=60000)
    await _open_signup_modal(page)
    try:
        phone_button = page.get_by_role(
            "button",
            name=re.compile(r"(使用电话号码继续|使用手机号继续|使用手机号码继续|手机登录|Continue with phone|phone)", re.IGNORECASE),
        ).first
        await _real_mouse_click_locator(page, phone_button, "手机号继续按钮", timeout=10000)
        logger.info("已用真实鼠标点击手机号继续按钮")
    except SignupFlowError:
        state = await _read_page_state(page)
        logger.error(f"未找到手机号入口，当前页面: {state}")
        raise
    except Exception as exc:
        logger.warning(f"手机号按钮原生点击失败，改用文本点击: {exc}")
        try:
            await _click_button_by_text(
                page,
                [
                    "使用电话号码继续",
                    "使用手机号继续",
                    "使用手机号码继续",
                    "手机登录",
                    "电话号码继续",
                    "Continue with phone",
                    "phone number",
                    "phone",
                ],
                timeout=15000,
            )
        except SignupFlowError:
            state = await _read_page_state(page)
            logger.error(f"未找到手机号入口，当前页面: {state}")
            raise
    phone_input = page.locator('input[name="phoneNumberInput"]').first
    await phone_input.wait_for(state="visible", timeout=20000)
    await _select_phone_country(page, COLOMBIA_DIAL_CODE, COLOMBIA_ISO, COLOMBIA_NAME)
    local_phone = _local_phone_number(activation.phone)
    logger.info(f"输入哥伦比亚手机号: +{COLOMBIA_DIAL_CODE} {local_phone}")
    await _real_mouse_click_locator(page, phone_input, "手机号输入框", timeout=5000)
    await page.keyboard.press("Control+A")
    await page.keyboard.type(local_phone, delay=60)
    await _random_human_delay(400, 1000)
    await _click_submit(page)


async def _complete_phone_registration(
    bot: ChatGPTBot,
    sms_client: HeroSMSClient,
    activation: HeroSMSActivation,
    full_name: str,
    birthdate: str,
    password: str,
    sms_timeout: int,
    sms_interval: int,
    finish_after: bool,
) -> None:
    if not bot.page:
        raise SignupFlowError("浏览器页面未初始化")
    page = bot.page
    logger.info("第 3 步: 等待创建密码页面")
    password_state = await _wait_until_state(
        page,
        lambda state: _is_create_password_state(state) or _is_login_password_state(state),
        "创建密码或登录密码",
        timeout=90,
    )
    if _is_login_password_state(password_state):
        try:
            await sms_client.cancel_activation(activation.activation_id)
        except HeroSMSError as exc:
            logger.warning(f"取消已占用手机号激活失败: {exc}")
        raise PhoneAlreadyRegisteredError("手机号已存在账号，进入登录密码页，无法用于新注册")

    logger.info("第 4 步: 创建随机密码")
    password_input = page.locator('input[type="password"]').first
    await password_input.wait_for(state="visible", timeout=15000)
    await _real_mouse_click_locator(page, password_input, "密码输入框", timeout=5000)
    await _random_human_delay(200, 600)
    await _human_type(page, password)
    await _click_submit(page)

    logger.info("第 5 步: 等待查看手机验证码页面")
    sms_state = await _wait_until_state(
        page,
        lambda state: _is_sms_verification_state(state) or _is_phone_already_registered_state(state),
        "查看你的手机或手机号已存在账号",
        timeout=90,
    )
    if _is_phone_already_registered_state(sms_state):
        try:
            await sms_client.cancel_activation(activation.activation_id)
            logger.info(f"HeroSMS 已取消已存在账号的手机号激活: {activation.activation_id}")
        except HeroSMSError as exc:
            logger.warning(f"取消已存在账号手机号激活失败: {exc}")
        raise PhoneAlreadyRegisteredError("手机号已关联既有账号，无法用于新注册")

    logger.info("第 6 步: 等待 HeroSMS 验证码并提交")
    try:
        sms_result = await sms_client.poll_code_by_activation(
            activation,
            timeout=sms_timeout,
            interval=sms_interval,
            finish_after=finish_after,
            cancel_on_timeout=True,
        )
    except HeroSMSError as exc:
        if "未收到 HeroSMS 验证码" in str(exc):
            raise SmsVerificationTimeoutError(str(exc)) from exc
        raise
    await bot.enter_verification_code(sms_result.code)

    logger.info("第 7 步: 等待 about-you 并填写个人信息")
    await _wait_until_state(page, _is_about_you_state, "about-you", timeout=120)
    await bot.fill_about_you(name=full_name, birthdate=birthdate)


async def _login_existing_phone_account(bot: ChatGPTBot, phone: str, password: str) -> None:
    """使用指定手机号登录已有账号，并停在已登录会话。"""
    if not bot.page:
        raise SignupFlowError("浏览器页面未初始化")
    if not password:
        raise SignupFlowError("恢复模式必须提供 --password")

    activation = HeroSMSActivation(
        phone=phone,
        activation_id="resume-login",
        country=COLOMBIA_ISO,
        service=HEROSMS_OPENAI_SERVICE,
    )
    await _navigate_to_phone_signup(bot, activation)

    page = bot.page
    logger.info("恢复模式: 等待登录密码页")
    await _wait_until_state(page, _is_login_password_state, "登录密码页", timeout=90)
    password_input = page.locator('input[type="password"]').first
    await password_input.wait_for(state="visible", timeout=15000)
    await _real_mouse_click_locator(page, password_input, "登录密码输入框", timeout=5000)
    await page.keyboard.press("Control+A")
    await _random_human_delay(200, 600)
    await _human_type(page, password)
    await _click_submit(page)
    logger.info("恢复模式: 已提交登录密码")

    try:
        await page.wait_for_url(
            re.compile(r"https://chatgpt\.com(?!.*callback).*"),
            timeout=90000,
        )
    except Exception as exc:
        state = await _read_page_state(page)
        raise SignupFlowError(
            f"恢复模式登录后未回到 chatgpt.com: url={str(state.get('url', ''))[:200]}; "
            f"text={str(state.get('text', ''))[:300]}"
        ) from exc

    if not await bot.wait_for_login_complete():
        raise SignupFlowError("恢复模式登录后 session 校验失败")


async def _wait_for_oauth_callback_in_url(page, expected_state: str, timeout: int = 120) -> str:
    """轮询页面 URL 等待 Codex OAuth callback 携带 code。"""
    from urllib.parse import parse_qs, urlparse as _urlparse

    deadline = asyncio.get_running_loop().time() + timeout
    while asyncio.get_running_loop().time() < deadline:
        url = page.url or ""
        if "localhost:1455/auth/callback" in url:
            parsed = _urlparse(url)
            params = parse_qs(parsed.query)
            error = next(iter(params.get("error") or []), "")
            if error:
                desc = next(iter(params.get("error_description") or []), "")
                raise SignupFlowError(f"Codex OAuth 授权失败: {error} {desc}")
            state = next(iter(params.get("state") or []), "")
            if state and state != expected_state:
                logger.warning(
                    "Codex OAuth URL 检测忽略旧回调 state: received=%s expected=%s",
                    state[:8],
                    expected_state[:8],
                )
                await asyncio.sleep(0.5)
                continue
            code = next(iter(params.get("code") or []), "")
            if not code:
                raise SignupFlowError("Codex OAuth callback 缺少 code")
            logger.info(f"URL callback 已获取 code: {code[:4]}...")
            return code
        await asyncio.sleep(0.5)
    raise SignupFlowError(f"{timeout}s 内未在 URL 中检测到 OAuth callback")


async def _execute_oauth_flow(
    bot: ChatGPTBot,
    email_address: str,
    password: str,
    oauth_phone: str = "",
    outlook_client: OutlookEmailPlusClient | None = None,
    email_timeout: int = 180,
) -> dict[str, Any]:
    """执行 OAuth 完整流程：欢迎回来→密码→添加邮箱→验证→Codex授权。

    使用真实鼠标点击 (_real_mouse_click_locator)，不使用 Playwright 虚拟点击。
    """
    if not bot.page:
        raise SignupFlowError("浏览器页面未初始化")
    page = bot.page

    async def retire_current_outlook_claim(reason: str) -> None:
        if not outlook_client:
            return
        try:
            await outlook_client.complete_claim(result="credential_invalid", detail=reason[:500])
        except Exception as claim_exc:
            logger.warning(f"标记 credential_invalid 失败: {claim_exc}")

    async def acquire_next_outlook_email(email_attempt: int, reason: str) -> str:
        if not outlook_client:
            raise SignupFlowError(f"邮箱 {email_address} {reason}，且无 OutlookEmailPlus 客户端无法切换邮箱")
        if email_attempt >= max_email_retries - 1:
            raise SignupFlowError(f"邮箱 {email_address} {reason}，已达最大重试次数 {max_email_retries}")

        logger.info(f"尝试获取新邮箱 (剩余重试: {max_email_retries - email_attempt - 1})")
        new_result = None
        try:
            new_result = await outlook_client.acquire_address()
        except Exception as acquire_exc:
            logger.warning(f"获取新邮箱失败: {acquire_exc}")

        if not new_result or not new_result.get("email"):
            logger.error("❌ OutlookEmailPlus 邮箱池已无可用邮箱！")
            raise EmailPoolExhaustedError(f"邮箱 {email_address} {reason}，邮箱池无可用账号")

        return str(new_result["email"])

    # ── 生成 OAuth 参数并启动回调服务器 ──
    oauth_params = bot._generate_codex_oauth_params()
    auth_url = bot._build_codex_oauth_url(oauth_params)

    callback_server = None
    try:
        try:
            callback_server, _cb_future = await bot._start_codex_oauth_callback_server(oauth_params["state"])
            logger.info("Codex OAuth 本地回调监听已启动")
        except OSError as exc:
            logger.warning(f"Codex OAuth 本地回调监听启动失败，将退回页面 URL 检测: {exc}")
            _cb_future = None

        # ── 导航到 OAuth 授权 URL（触发登录/OAuth流程）──
        logger.info("OAuth: 导航到 Codex OAuth 授权页")
        await page.goto(auth_url, wait_until="domcontentloaded", timeout=60000)
        await _random_human_delay(1500, 3500)

        # ── 分支检测：已登录账号选择 vs 完整登录流程 ──
        # 页面可能还在渲染账号列表，多等一轮确保"选择一个账户"文本出现
        state = await _read_page_state(page)
        if not _is_oauth_account_selector(state) and not _is_oauth_welcome_back(state):
            logger.info("OAuth: 页面初始状态未匹配，等待渲染后重试...")
            await _random_human_delay(1000, 2500)
            state = await _read_page_state(page)

        step_count = 3  # 后三步固定：add-email → verify → consent

        if _is_oauth_account_selector(state):
            # 已有 session cookie，直接展示账号选择
            logger.info("OAuth: 检测到已登录账号选择页，点击已有账号继续")
            clicked_account = False

            # 方案 1: 找 name="session_id" 的 submit 按钮（最可靠）
            session_btn = page.locator('button[name="session_id"]').first
            try:
                if await session_btn.is_visible(timeout=3000):
                    await _real_mouse_click_locator(page, session_btn, "已有账号按钮(session_id)", timeout=8000)
                    clicked_account = True
                    logger.info("OAuth: 已通过 session_id 按钮点击已有账号")
            except Exception as exc:
                logger.debug(f"session_id 按钮不可用: {exc}")

            # 方案 2: 找包含 +57 手机号的按钮
            if not clicked_account:
                phone_btn = page.locator('button, [role="button"]').filter(
                    has_text=re.compile(r"\+57\s?\d{3}")
                ).first
                try:
                    if await phone_btn.is_visible(timeout=3000):
                        await _real_mouse_click_locator(page, phone_btn, "已有账号卡片(+57)", timeout=8000)
                        clicked_account = True
                        logger.info("OAuth: 已通过 +57 手机号点击账号卡片")
                except Exception as exc:
                    logger.debug(f"+57 按钮不可用: {exc}")

            # 方案 3: 找 form[action*="choose-an-account"] 中的 submit 按钮
            if not clicked_account:
                choose_btn = page.locator('form[action*="choose-an-account"] button[type="submit"], '
                                          'form[action*="choose-an-account"] button').first
                try:
                    if await choose_btn.is_visible(timeout=2000):
                        await _real_mouse_click_locator(page, choose_btn, "choose-an-account submit", timeout=5000)
                        clicked_account = True
                        logger.info("OAuth: 已通过 choose-an-account 表单提交点击账号")
                except Exception as exc:
                    logger.debug(f"choose-an-account 按钮不可用: {exc}")

            # 方案 4: 兜底 _click_submit
            if not clicked_account:
                try:
                    submitted = await page.evaluate(
                        """
                        () => {
                            const form = document.querySelector('form[action*="choose-an-account"]');
                            const button = form?.querySelector('button[name="session_id"]');
                            if (!form || !button) return false;
                            if (typeof form.requestSubmit === 'function') {
                                form.requestSubmit(button);
                            } else {
                                button.click();
                            }
                            return true;
                        }
                        """
                    )
                    if submitted:
                        clicked_account = True
                        logger.info("OAuth: 已通过 choose-an-account 表单 requestSubmit 提交账号")
                except Exception as exc:
                    logger.debug(f"choose-an-account 表单 requestSubmit 失败: {exc}")

            # 方案 5: 最后才用通用提交按钮兜底；账号卡片通常没有“继续”文本
            if not clicked_account:
                try:
                    await _click_submit(page)
                    clicked_account = True
                    logger.info("OAuth: 通过 _click_submit 兜底点击账号选择页")
                except Exception:
                    raise SignupFlowError("OAuth 账号选择页：无法点击已有账号继续")

            # ── 等待页面离开 choose-an-account ──
            logger.info("OAuth: 等待账号选择后页面跳转...")
            try:
                await _wait_until_state(
                    page,
                    lambda current: not _is_oauth_account_selector(current),
                    "账号选择后离开选择页",
                    timeout=15,
                )
                logger.info(f"OAuth: 账号选择后已进入下一状态 {page.url[:120]}")
            except Exception:
                logger.warning("OAuth: 点击账号后 15s 未跳转，尝试直接提交表单")
                try:
                    await page.evaluate("""
                        () => {
                            const form = document.querySelector('form[action*=\"choose-an-account\"]');
                            if (form) {
                                const btn = form.querySelector('button[name=\"session_id\"]');
                                if (btn) btn.click();
                            }
                        }
                    """)
                except Exception as e:
                    logger.warning(f"OAuth: 表单直接提交也失败: {e}")
                await _random_human_delay(1500, 3000)
                await _wait_until_state(
                    page,
                    lambda current: not _is_oauth_account_selector(current),
                    "账号选择表单提交后离开选择页",
                    timeout=15,
                )

        elif _is_oauth_welcome_back(state):
            # 未登录，需要完整登录流程（可能分邮箱页→电话页两阶段）
            step_count = 5

            # ── 子分支 1: 邮箱登录页 → 点击"使用电话号码继续"切换 ──
            if _is_oauth_email_login(state):
                logger.info("OAuth 步骤 1a: 邮箱登录页，点击 '使用电话号码继续' 切换到手机登录")
                await _wait_until_state(page, _is_oauth_welcome_back, "欢迎回来(邮箱)", timeout=60)
                # 点击 "使用电话号码继续" 按钮（纯 button，非 submit）
                phone_switch_btn = page.get_by_role("button", name=re.compile(
                    r"(使用电话号码继续|Continue with phone)", re.IGNORECASE
                )).first
                try:
                    if await phone_switch_btn.is_visible(timeout=3000):
                        await _real_mouse_click_locator(page, phone_switch_btn, "使用电话号码继续", timeout=5000)
                        logger.info("OAuth 步骤 1a ✓ 已切换到手机登录页")
                except Exception as exc:
                    logger.warning(f"点击 '使用电话号码继续' 失败，尝试文本匹配: {exc}")
                    await _click_button_by_text(page, [
                        "使用电话号码继续", "Continue with phone",
                        "使用手机号继续", "手机号码继续",
                    ], timeout=10000)
                await _random_human_delay(1000, 2500)

            # ── 子分支 2: 手机号登录页 → 选国家 + 输手机号 + 提交 ──
            logger.info("OAuth 步骤 1b: 手机号登录页，选择国家并输入手机号")
            await _wait_until_state(page, _is_oauth_phone_login, "欢迎回来(手机号)", timeout=60)
            await _random_human_delay(300, 800)

            # 选国家：哥伦比亚 +57
            await _select_phone_country(page, COLOMBIA_DIAL_CODE, COLOMBIA_ISO, COLOMBIA_NAME)

            # 输入本地手机号
            phone_input = page.locator(
                'input[name="__reservedForPhoneNumberInput_tel"], '
                'input[name="phoneNumberInput"], input[type="tel"]'
            ).first
            await phone_input.wait_for(state="visible", timeout=10000)
            local_phone = _local_phone_number(oauth_phone) if oauth_phone else ""
            if local_phone:
                logger.info(f"OAuth 步骤 1b 输入手机号: +{COLOMBIA_DIAL_CODE} {local_phone}")
                await _real_mouse_click_locator(page, phone_input, "手机号输入框", timeout=5000)
                await page.keyboard.press("Control+A")
                await page.keyboard.type(local_phone, delay=60)
            else:
                logger.warning("OAuth 步骤 1b: 无可用手机号，无法自动填写")
            await _random_human_delay(400, 1000)
            await _click_submit(page)
            logger.info("OAuth 步骤 1b ✓ 已提交手机号")

            # ── 步骤 2: 输入密码 ──
            logger.info("OAuth 步骤 2: 等待 '输入密码' 页面")
            await _wait_until_state(page, _is_oauth_enter_password, "输入密码", timeout=60)
            pw_input = page.locator('input[type="password"]').first
            await pw_input.wait_for(state="visible", timeout=15000)
            await _real_mouse_click_locator(page, pw_input, "密码输入框", timeout=5000)
            await _random_human_delay(200, 600)
            await _human_type(page, password)
            logger.info("OAuth 步骤 2 ✓ 已填入密码")
            await _click_submit(page)
        else:
            raise SignupFlowError(
                f"OAuth 导航后未知页面状态: url={str(state.get('url', ''))[:200]}; "
                f"text={str(state.get('text', ''))[:300]}"
            )

        # ── 后续公共步骤：添加邮箱 → 验证邮箱 → Codex 授权 ──
        await _random_human_delay(500, 1500)
        add_email_label = "OAuth 步骤 3/5"
        verify_label = "OAuth 步骤 4/5"
        consent_label = "OAuth 步骤 5/5"
        if step_count == 3:
            add_email_label = "OAuth 步骤 1/3"
            verify_label = "OAuth 步骤 2/3"
            consent_label = "OAuth 步骤 3/3"

        max_email_retries = 3
        for email_attempt in range(max_email_retries):
            logger.info(f"{add_email_label}: 等待 '要求提供电子邮件地址' 页面"
                        f"{' (重试)' if email_attempt > 0 else ''}")
            await _wait_until_state(page, _is_oauth_add_email, "要求提供电子邮件地址", timeout=60)
            em_input = page.locator('input[type="email"], input[name="email"]').first
            await em_input.wait_for(state="visible", timeout=15000)
            await _real_mouse_click_locator(page, em_input, "邮箱输入框", timeout=5000)
            await page.keyboard.press("Control+A")
            await _random_human_delay(150, 400)
            await _human_type(page, email_address, min_delay=30, max_delay=100)
            logger.info(f"{add_email_label} ✓ 已填入邮箱: {email_address}"
                        f" (尝试 {email_attempt + 1}/{max_email_retries})")
            email_code_baseline = int(time.time())
            await _click_submit(page)

            logger.info(f"{verify_label}: 等待 '检查你的收件箱' 页面")
            await _wait_until_state(page, _is_oauth_email_verification, "检查你的收件箱", timeout=60)

            consent_result = ""
            max_code_retries = 3
            for code_attempt in range(1, max_code_retries + 1):
                # 获取验证码：优先 OutlookEmailPlus API 直连，其次交互式输入
                code = ""
                if outlook_client:
                    try:
                        logger.info(
                            f"OAuth: 通过 OutlookEmailPlus API 轮询验证码… "
                            f"(验证码尝试 {code_attempt}/{max_code_retries})"
                        )
                        code_info = await outlook_client.poll_for_code(
                            timeout=email_timeout,
                            interval=5,
                            baseline_timestamp=email_code_baseline,
                        )
                        code = code_info["code"]
                        logger.info(f"OAuth 验证码 API获取: {code[:2]}***")
                    except OutlookEmailPlusUnreadable as exc:
                        reason = f"邮箱读取失败: {exc}"
                        logger.warning(f"邮箱 {email_address} 读取不可用，将从邮箱池剔除: {exc}")
                        await retire_current_outlook_claim(reason)
                        email_address = await acquire_next_outlook_email(email_attempt, "读取不可用")
                        logger.info(f"✅ 已切换到新邮箱: {email_address}")
                        logger.info("跳回 https://auth.openai.com/add-email 重新添加邮箱…")
                        await page.goto("https://auth.openai.com/add-email", wait_until="domcontentloaded", timeout=60000)
                        await _random_human_delay(1000, 2500)
                        consent_result = "retry_email"
                        break
                    except OutlookEmailPlusTimeout:
                        logger.warning("OutlookEmailPlus 轮询验证码超时")
                    except Exception as exc:
                        logger.warning(f"OutlookEmailPlus 获取验证码失败: {exc}")

                if not code:
                    notify_human_intervention("OAuth 邮箱验证码", f"邮箱 {email_address} 自动取码失败，请在终端手动输入验证码。")
                    print("\n>> 已向邮箱发送验证码，请查看邮件并输入验证码。")
                    loop = asyncio.get_running_loop()
                    while True:
                        code = await loop.run_in_executor(None, lambda: input("请输入邮箱验证码 > ").strip())
                        if code and code.isdigit() and 4 <= len(code) <= 8:
                            break
                        print("[!] 验证码应为 4-8 位数字，请重新输入。")

                cd_input = page.locator('input[name="code"], input[autocomplete="one-time-code"]').first
                await cd_input.wait_for(state="visible", timeout=10000)
                await _real_mouse_click_locator(page, cd_input, "验证码输入框", timeout=5000)
                await page.keyboard.press("Control+A")
                await _random_human_delay(150, 500)
                await _human_type(page, code, min_delay=50, max_delay=150)
                logger.info(f"{verify_label} ✓ 已填入验证码: {code[:2]}***")
                await _click_submit(page)

                logger.info(f"{consent_label}: 等待 Codex 授权确认页")
                consent_result = await _wait_for_oauth_consent_or_email_in_use(page, timeout=60)
                if consent_result in ("consent", "email_in_use"):
                    break
                if consent_result == "code_incorrect":
                    if code_attempt >= max_code_retries:
                        raise SignupFlowError("邮箱验证码连续错误，疑似无法获取最新验证码")
                    logger.warning("OAuth 邮箱验证码错误，可能读取到旧验证码，点击重新发送后继续等待新验证码")
                    await _click_button_by_text(page, ["重新发送电子邮件", "Resend email", "Resend"], timeout=10000)
                    await _random_human_delay(1000, 2500)
                    email_code_baseline = int(time.time())
                    await _wait_until_state(page, _is_oauth_email_verification, "检查你的收件箱", timeout=30)
                    continue

            if consent_result == "consent":
                break
            if consent_result == "retry_email":
                continue
            if not _is_oauth_add_email(await _read_page_state(page)) and consent_result not in ("email_in_use", "code_incorrect"):
                continue

            logger.warning(f"邮箱 {email_address} 已达使用上限 (email_in_use)")
            if not outlook_client:
                raise SignupFlowError(
                    f"邮箱 {email_address} email_in_use，且无 OutlookEmailPlus 客户端无法切换邮箱"
                )
            try:
                await outlook_client.complete_claim(result="email_limit", detail="email_in_use")
            except Exception as claim_exc:
                logger.warning(f"标记 email_limit 失败: {claim_exc}")
            email_address = await acquire_next_outlook_email(email_attempt, "email_in_use")
            logger.info(f"✅ 已切换到新邮箱: {email_address}")

            # email_in_use 错误页的"重试"按钮跳转不稳定，直接回到 add-email 重填新邮箱。
            logger.info("跳回 https://auth.openai.com/add-email 重新添加邮箱…")
            await page.goto("https://auth.openai.com/add-email", wait_until="domcontentloaded", timeout=60000)
            await _random_human_delay(1000, 2500)
            await _wait_until_state(page, _is_oauth_add_email, "要求提供电子邮件地址", timeout=30)
            # 继续下一轮循环，从 add-email 页重新开始

        await _click_submit(page)
        logger.info(f"{consent_label} ✓ 已授权 Codex")

        # ── 等待 OAuth callback 获取 code ──
        await _random_human_delay(500, 1500)
        logger.info("OAuth: 等待授权回调...")
        if _cb_future is not None:
            try:
                oauth_code = await asyncio.wait_for(asyncio.shield(_cb_future), timeout=120)
            except asyncio.TimeoutError:
                raise SignupFlowError("120s 内 OAuth callback future 未完成")
        else:
            oauth_code = await _wait_for_oauth_callback_in_url(page, oauth_params["state"], timeout=120)

        # ── 用 code 换 token ──
        logger.info("OAuth: 交换 token...")
        token_data = await bot._exchange_codex_oauth_token(oauth_code, oauth_params["code_verifier"], email_address)
        logger.info("✅ Codex OAuth token 获取成功")
        return token_data

    finally:
        if callback_server:
            callback_server.close()
            await callback_server.wait_closed()


def _release_email_pool(db_path: str) -> int:
    """释放 OutlookEmailPlus 邮箱池中所有 claimed 状态的邮箱。"""
    if not os.path.exists(db_path):
        logger.warning(f"邮箱池数据库不存在: {db_path}")
        return 0
    conn = sqlite3.connect(db_path)
    try:
        conn.row_factory = sqlite3.Row
        accounts = conn.execute(
            "SELECT id, email, claimed_by, claim_token, claimed_project_key "
            "FROM accounts WHERE pool_status = 'claimed'"
        ).fetchall()
        if not accounts:
            logger.info("邮箱池无 claimed 账号，无需释放")
            return 0

        logger.info(f"释放邮箱池 {len(accounts)} 个 claimed 账号")
        now_text = datetime.now().isoformat(timespec="seconds")
        conn.execute("BEGIN IMMEDIATE")
        released = 0
        for acc in accounts:
            cursor = conn.execute(
                "UPDATE accounts SET pool_status='available', claimed_by=NULL, "
                "claimed_at=NULL, lease_expires_at=NULL, claim_token=NULL, "
                "claimed_project_key=NULL, updated_at=? WHERE id=? AND pool_status='claimed'",
                (now_text, acc["id"]),
            )
            if cursor.rowcount == 1:
                released += 1
        conn.execute("COMMIT")
        logger.info(f"邮箱池释放完成: {released} 个已恢复为 available")
        total_released = released

        # ── 2. 清理 pool_status='available' 但 status!='active' 的僵尸账号 ──
        zombie_rows = conn.execute(
            "SELECT id, email, pool_status, status "
            "FROM accounts WHERE pool_status = 'available' AND status != 'active'"
        ).fetchall()
        if zombie_rows:
            logger.warning(
                f"发现 {len(zombie_rows)} 个 pool_status='available' 但 status!='active' 的僵尸账号，"
                f"标记为 retired"
            )
            conn.execute("BEGIN IMMEDIATE")
            cleaned = 0
            for z in zombie_rows:
                cursor = conn.execute(
                    "UPDATE accounts SET pool_status='retired', updated_at=? "
                    "WHERE id=? AND pool_status='available' AND status!='active'",
                    (now_text, z["id"]),
                )
                if cursor.rowcount == 1:
                    cleaned += 1
                    logger.info(f"  僵尸账号已清理: {z['email']} (status={z['status']})")
            conn.execute("COMMIT")
            logger.info(f"僵尸账号清理完成: {cleaned} 个已标记为 retired")
        else:
            logger.info("邮箱池无不一致僵尸账号")

        return total_released
    except Exception as exc:
        conn.execute("ROLLBACK")
        logger.warning(f"邮箱池释放失败: {exc}")
        return 0
    finally:
        conn.close()


def _claim_account_by_email(db_path: str, target_email: str, caller_id: str, task_id: str, project_key: str) -> dict[str, str]:
    """对指定邮箱通过直接 DB 操作创建 claim，直接使用该邮箱地址（不创建别名）。

    返回 {
        "account_id": str,
        "email":        指定邮箱地址,
        "primary_email": 主邮箱,
        "claim_token":  str,
    }。

    若账号不可用（status!='active' / pool_status!='available'）则抛 OutlookEmailPlusError。
    """
    import secrets

    if not os.path.exists(db_path):
        raise OutlookEmailPlusError(f"邮箱池数据库不存在: {db_path}")
    target_email = target_email.strip().lower()
    conn = sqlite3.connect(db_path)
    try:
        conn.row_factory = sqlite3.Row
        account = conn.execute(
            "SELECT id, email, status, pool_status FROM accounts WHERE LOWER(email) = ?",
            (target_email,),
        ).fetchone()
        if not account:
            raise OutlookEmailPlusError(f"账号 {target_email} 在邮箱池中未找到")
        if account["status"] != "active":
            raise OutlookEmailPlusError(
                f"账号 {target_email} status={account['status']}，不可用"
            )
        if account["pool_status"] != "available":
            raise OutlookEmailPlusError(
                f"账号 {target_email} pool_status={account['pool_status']}，不可用，当前状态应为 available"
            )

        account_id = account["id"]
        primary_email = account["email"]
        now_text = datetime.now().isoformat(timespec="seconds")
        lease_expires_at = (datetime.now() + timedelta(hours=1)).isoformat(timespec="seconds") + "Z"
        token = "clm_" + secrets.token_urlsafe(9)

        # 执行 claim（与 claim_atomic 行为一致）
        conn.execute("BEGIN IMMEDIATE")
        conn.execute(
            """UPDATE accounts SET
                pool_status = 'claimed',
                claimed_by = ?,
                claimed_at = ?,
                lease_expires_at = ?,
                claim_token = ?,
                claimed_project_key = ?,
                updated_at = ?
            WHERE id = ? AND pool_status = 'available'""",
            (
                caller_id,
                now_text,
                lease_expires_at,
                token,
                project_key,
                now_text,
                account_id,
            ),
        )
        # 记录 claim 日志
        conn.execute(
            """INSERT INTO account_claim_logs
                (account_id, claim_token, caller_id, task_id, action, result, detail, created_at)
            VALUES (?, ?, ?, ?, 'claim', NULL, ?, ?)""",
            (
                account_id,
                token,
                caller_id,
                task_id,
                f"claim_by_email {target_email}",
                now_text,
            ),
        )
        conn.execute("COMMIT")
        logger.info(
            f"✅ 已 claim 账号 {primary_email} (id={account_id})"
        )
        return {
            "account_id": str(account_id),
            "email": primary_email,
            "primary_email": primary_email,
            "claim_token": token,
        }
    except OutlookEmailPlusError:
        raise
    except Exception as exc:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise OutlookEmailPlusError(f"指定邮箱 claim 失败: {exc}") from exc
    finally:
        conn.close()


async def main() -> None:
    parser = argparse.ArgumentParser(description="ChatGPT 手机号免费注册 + Codex OAuth + Sub2API 导出")
    parser.add_argument("--headless", action="store_true", help="无头模式")
    parser.add_argument("--slow-mo", type=int, default=100, help="Playwright 操作延迟(ms)")
    parser.add_argument("--debug", action="store_true", help="出错时保持浏览器打开")
    parser.add_argument("--keep-open", action="store_true", help="流程结束后保持浏览器打开")
    parser.add_argument("--proxy", type=str, default=DEFAULT_OAUTH_PROXY, help="浏览器与 OAuth token exchange 共用代理")
    parser.add_argument("--email-timeout", type=int, default=180, help="邮箱验证码等待超时(s)")
    parser.add_argument("--email-code-file", type=str, default="", help="等待该文件写入邮箱 / OAuth 验证码")
    parser.add_argument("--email", type=str, default="", help="OAuth/Sub2API 导出时使用的账号标识邮箱；不再用于注册")
    parser.add_argument("--resume-alias-email", type=str, default="", help="复用已有 OutlookEmailPlus 别名作为 OAuth/Sub2API 账号标识")
    parser.add_argument("--resume-phone", type=str, default="", help="恢复模式：使用指定手机号登录已有账号后执行 OAuth/Sub2API 导出")
    parser.add_argument("--claim-email", type=str, default="", help="指定 OutlookEmailPlus 邮箱池中已有邮箱用于注册（不创建别名）")
    parser.add_argument("--herosms-phone", type=str, default="", help="跳过自动买号，直接使用已有 HeroSMS 激活手机号")
    parser.add_argument("--herosms-apikey", type=str, default=HEROSMS_APIKEY, help="HeroSMS APIKEY")
    parser.add_argument("--herosms-service", type=str, default=HEROSMS_OPENAI_SERVICE, help="HeroSMS 服务代码，OpenAI/ChatGPT 默认 dr")
    parser.add_argument("--herosms-country", type=str, default="", help="HeroSMS 国家 ID；默认自动解析 Colombia/CO")
    parser.add_argument("--herosms-max-price", type=float, default=HEROSMS_MAX_PRICE_USD, help="HeroSMS 购号最高价格，默认 0.08 美元")
    parser.add_argument("--herosms-timeout", type=int, default=120, help="HeroSMS 接码等待超时(s)")
    parser.add_argument("--herosms-interval", type=int, default=HEROSMS_INTERVAL, help="HeroSMS 轮询间隔(s)")
    parser.add_argument("--phone-retries", type=int, default=3, help="手机号验证码超时后同轮重新买号重试次数，默认 3")
    parser.add_argument("--herosms-finish", action="store_true", help="读取后调用 HeroSMS 完成激活")
    parser.add_argument("--herosms-only", action="store_true", help="只自动购买 HeroSMS 哥伦比亚号码并等待验证码，不执行注册/OAuth")
    parser.add_argument("--password", type=str, default="", help="手机号注册密码；不传则自动生成")
    parser.add_argument("--name", type=str, default=DEFAULT_NAME, help="账号名称。保持默认时随机生成英文姓名")
    parser.add_argument("--birthdate", type=str, default="", help="生日 YYYY-MM-DD；不传则随机生成")
    parser.add_argument("--output", type=str, default="", help="Sub2API 输出目录；指定时在该目录下按时间创建子目录")
    parser.add_argument("--account-record-file", type=str, default=ACCOUNT_RECORD_FILE, help="阶段性账号记录 JSON，失败也会写入")
    parser.add_argument("--rounds", type=int, default=1, help="批量运行轮数，默认 1 轮")
    args = parser.parse_args()

    email = args.email.strip()
    resume_alias_email = args.resume_alias_email.strip()
    resume_phone = args.resume_phone.strip().lstrip("+")
    claim_email = args.claim_email.strip()
    hero_sms_phone = args.herosms_phone.strip().lstrip("+")
    if email and resume_alias_email:
        exit_with_param_error("[参数错误] --email 不能与 --resume-alias-email 同时使用")
    if claim_email and (email or resume_alias_email):
        exit_with_param_error("[参数错误] --claim-email 不能与 --email / --resume-alias-email 同时使用")
    if resume_phone:
        if not resume_phone.isdigit():
            exit_with_param_error("[参数错误] --resume-phone 必须是手机号数字，可带 +")
        if not args.password.strip():
            exit_with_param_error("[参数错误] --resume-phone 恢复模式必须同时提供 --password")
        if hero_sms_phone:
            exit_with_param_error("[参数错误] --resume-phone 不能与 --herosms-phone 同时使用")
    if email and not is_valid_email(email):
        exit_with_param_error("[参数错误] --email 不是有效邮箱地址")
    if resume_alias_email and not is_valid_email(resume_alias_email):
        exit_with_param_error("[参数错误] --resume-alias-email 不是有效邮箱地址")
    if claim_email and not is_valid_email(claim_email):
        exit_with_param_error("[参数错误] --claim-email 不是有效邮箱地址")

    rounds = max(1, int(args.rounds))
    if rounds > 1 and (claim_email or resume_alias_email or email or hero_sms_phone or resume_phone):
        reused_params = []
        if claim_email:
            reused_params.append("--claim-email")
        if resume_alias_email:
            reused_params.append("--resume-alias-email")
        if email:
            reused_params.append("--email")
        if hero_sms_phone:
            reused_params.append("--herosms-phone")
        if resume_phone:
            reused_params.append("--resume-phone")
        exit_with_param_error(
            f"[参数错误] --rounds={rounds} 不允许复用固定资源: {', '.join(reused_params)}。"
            "批量运行请使用自动领取邮箱和自动购买手机号。"
        )

    # ── 跨轮共享的初始化 ──
    bot_module.DEBUG_MODE = bool(args.debug or args.keep_open)
    proxy_url = normalize_proxy_url(args.proxy)
    hero_sms_api_key = args.herosms_apikey.strip()
    hero_sms_service = args.herosms_service.strip() or HEROSMS_OPENAI_SERVICE
    hero_sms_interval = max(1, int(args.herosms_interval or HEROSMS_INTERVAL))
    hero_sms_timeout = max(1, int(args.herosms_timeout or 120))
    phone_retries = max(1, int(args.phone_retries or 1))
    hero_sms_finish_after = bool(args.herosms_finish or HEROSMS_FINISH_AFTER)
    account_record_file = os.path.abspath(args.account_record_file)
    accounts_root_dir = os.path.abspath(args.output) if args.output.strip() else os.path.abspath(
        os.path.join(os.path.dirname(__file__), "accounts")
    )
    batch_dir_name = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_dir = os.path.join(accounts_root_dir, batch_dir_name)
    batch_output_file = os.path.join(output_dir, f"sub2api-free-batch-{batch_dir_name}.json")
    os.makedirs(output_dir, exist_ok=True)

    sms_client: HeroSMSClient | None = None
    hero_sms_country = args.herosms_country.strip()
    if not resume_phone or args.herosms_only:
        sms_client = HeroSMSClient(hero_sms_api_key)
        hero_sms_country = hero_sms_country or await sms_client.resolve_country_id(COLOMBIA_ISO, COLOMBIA_NAME)
    else:
        hero_sms_country = hero_sms_country or COLOMBIA_ISO

    async def acquire_activation() -> HeroSMSActivation:
        if not sms_client:
            raise HeroSMSError("当前模式不需要 HeroSMS，不能获取新手机号")
        if hero_sms_phone:
            logger.info(f"使用已有 HeroSMS 激活手机号: {hero_sms_phone}")
            activation_data = await sms_client.find_activation_by_phone(hero_sms_phone)
            activation_id = str(
                activation_data.get("activationId")
                or activation_data.get("activation_id")
                or activation_data.get("id")
                or ""
            )
            if not activation_id:
                raise HeroSMSError(f"已匹配手机号但缺少 activation id: {activation_data}")
            return HeroSMSActivation(
                phone=hero_sms_phone,
                activation_id=activation_id,
                country=hero_sms_country,
                service=hero_sms_service,
            )
        return await sms_client.acquire_number(
            service=hero_sms_service,
            country=hero_sms_country,
            max_price=float(args.herosms_max_price),
        )

    if args.herosms_only:
        if not sms_client:
            raise HeroSMSError("HeroSMS 客户端未初始化")
        activation = await acquire_activation()
        code_result = await sms_client.poll_code_by_activation(
            activation,
            timeout=hero_sms_timeout,
            interval=hero_sms_interval,
            finish_after=hero_sms_finish_after,
        )
        print(f"HeroSMS号码: {activation.phone}")
        print(f"HeroSMS验证码: {code_result.code}")
        return

    # ── 批量运行 ──
    successes = 0
    partials = 0
    failures = 0
    batch_aborted = False
    batch_abort_reason = ""
    batch_start = time.time()

    for round_num in range(1, rounds + 1):
        # ── 每轮独立的随机值（如用户指定则复用）──
        round_real_name = args.name if args.name != DEFAULT_NAME else generate_real_name()
        round_birthdate = args.birthdate.strip() or generate_birthdate()
        round_password = args.password.strip() or generate_strong_password()
        round_email_prefix = generate_email_prefix()
        round_token = f"{round_num}-{time.time_ns()}-{uuid.uuid4().hex[:8]}"
        code_file = os.path.abspath(args.email_code_file) if args.email_code_file else os.path.abspath(
            os.path.join(tempfile.gettempdir(), f"gpt_free_oauth_code_{round_token}.txt")
        )
        run_id = f"gpt-free-{round_token}"
        account_record: dict[str, Any] = {
            "run_id": run_id,
            "round": round_num,
            "created_at": _now_iso(),
            "status": "started",
            "stage": "init",
            "mode": "resume_phone" if resume_phone else "register",
            "country": COLOMBIA_ISO,
            "dial_code": COLOMBIA_DIAL_CODE,
            "herosms_service": hero_sms_service,
            "herosms_finish_after": hero_sms_finish_after,
        }
        if resume_phone:
            account_record["phone"] = resume_phone
        upsert_account_record(account_record, account_record_file)

        outlook_client: OutlookEmailPlusClient | None = None
        active_bot: ChatGPTBot | None = None
        address = ""

        async def complete_claim(result: str = "success", detail: str = "") -> None:
            if not outlook_client:
                return
            try:
                await outlook_client.complete_claim(result=result, detail=detail)
            except Exception as exc:
                logger.warning(f"OutlookEmailPlus claim-complete 失败: {exc}")

        async def release_claim(reason: str) -> None:
            if not outlook_client:
                return
            try:
                await outlook_client.release_claim(reason=reason)
            except Exception as exc:
                logger.warning(f"OutlookEmailPlus claim-release 失败: {exc}")

        print(f"\n{'══' * 30}")
        print(f"  第 {round_num}/{rounds} 轮")
        print(f"{'══' * 30}")
        print(f"代理: {proxy_url or '未使用'}")
        print(f"注册国家: Colombia (+57)")
        print(f"HeroSMS service={hero_sms_service} maxPrice=${args.herosms_max_price:.2f}")
        print(f"Sub2API输出目录: {output_dir}")
        if rounds > 1:
            print(f"Sub2API批次总文件: {batch_output_file}")
        print(f"密码: {round_password}")
        print("=" * 60)

        try:
            output_path = ""
            if claim_email:
                logger.info(f"第 1 步: 使用指定邮箱 {claim_email}")
                db_path = os.path.join(
                    os.path.dirname(__file__),
                    "outlookEmailPlus", "data", "outlook_accounts.db",
                )
                task_id = f"gpt-free-claim-{round_email_prefix}-{int(time.time())}"
                claim_info = await asyncio.to_thread(
                    _claim_account_by_email,
                    db_path,
                    claim_email,
                    OUTLOOK_EMAIL_PLUS_CALLER_ID,
                    task_id,
                    OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                )
                address = claim_info["email"]
                outlook_client = OutlookEmailPlusClient(
                    base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
                    api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
                    caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
                    task_id=task_id,
                    provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
                    project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                    email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
                )
                outlook_client.setup_custom_claim(
                    account_id=claim_info["account_id"],
                    email=claim_info["email"],
                    primary_email=claim_info["primary_email"],
                    claim_token=claim_info["claim_token"],
                )
            elif resume_alias_email:
                logger.info("第 1 步: 使用 OutlookEmailPlus 既有别名邮箱恢复登录")
                outlook_client = OutlookEmailPlusClient(
                    base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
                    api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
                    caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
                    task_id=f"gpt-free-resume-{round_email_prefix}-{int(time.time())}",
                    provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
                    project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                    email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
                )
                address = outlook_client.use_existing_email(resume_alias_email)["email"]
            elif email:
                logger.info("第 1 步: 使用自备邮箱")
                address = email
            else:
                logger.info("第 1 步: 创建 OutlookEmailPlus 别名邮箱")
                auto_outlook_client = OutlookEmailPlusClient(
                    base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
                    api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
                    caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
                    task_id=f"gpt-free-{round_email_prefix}-{int(time.time())}",
                    provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
                    project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                    email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
                )
                outlook_client = auto_outlook_client
                try:
                    result = await retry_with_backoff(
                        lambda: auto_outlook_client.acquire_address(),
                        max_retries=3,
                        description="创建 OutlookEmailPlus 地址",
                    )
                    if not result:
                        raise OutlookEmailPlusError("创建 OutlookEmailPlus 地址失败：返回为空")
                    address = result["email"]
                except OutlookEmailPlusError as exc:
                    error_str = str(exc)
                    if "NO_AVAILABLE_ACCOUNT" in error_str:
                        logger.info("OutlookEmailPlus 池空，尝试释放后重试...")
                        try:
                            await asyncio.to_thread(
                                _release_email_pool,
                                os.path.join(
                                    os.path.dirname(__file__),
                                    "outlookEmailPlus", "data", "outlook_accounts.db",
                                ),
                            )
                        except Exception as release_exc:
                            logger.warning(f"释放邮箱池失败: {release_exc}")
                        # 重试一次
                        try:
                            result = await retry_with_backoff(
                                lambda: auto_outlook_client.acquire_address(),
                                max_retries=1,
                                description="创建 OutlookEmailPlus 地址（释放后重试）",
                            )
                            if result:
                                address = result["email"]
                                logger.info(f"✅ 释放后 OutlookEmailPlus 邮箱: {address}")
                        except OutlookEmailPlusError:
                            pass
                    if not address:
                        logger.error(f"❌ OutlookEmailPlus 邮箱池无可用账号，跳过本轮: {exc}")
                        raise EmailPoolExhaustedError(f"OutlookEmailPlus 邮箱池无可用账号: {exc}") from exc

            if address:
                account_record.update({"email": address, "email_status": "ready"})
                upsert_account_record(account_record, account_record_file)

            activation: HeroSMSActivation | None = None
            if resume_phone:
                account_record.update(
                    {
                        "stage": "resume_phone_ready",
                        "status": "in_progress",
                        "phone": resume_phone,
                        "password": round_password,
                    }
                )
                logger.info(f"第 2 步: 恢复模式登录已有手机号账号，手机号={resume_phone}")
            else:
                activation = await acquire_activation()
                account_record.update(
                    {
                        "stage": "phone_acquired",
                        "status": "in_progress",
                        "phone": activation.phone,
                        "activation_id": activation.activation_id,
                        "activation_cost": activation.activation_cost,
                        "phone_attempt": 1,
                        "password": round_password,
                        "name": round_real_name,
                        "birthdate": round_birthdate,
                    }
                )
                logger.info(f"第 2 步: ChatGPT 手机号免费注册，手机号={activation.phone}")
            upsert_account_record(account_record, account_record_file)
            async with ChatGPTBot(headless=args.headless, slow_mo=args.slow_mo, proxy_url=proxy_url) as bot:
                active_bot = bot
                bot.oauth_code_file = code_file
                if resume_phone:
                    await _login_existing_phone_account(bot, resume_phone, round_password)
                    account_record.update({"stage": "resume_login_completed", "status": "logged_in"})
                else:
                    if not sms_client:
                        raise SignupFlowError("注册模式缺少 HeroSMS 激活或客户端")
                    for phone_attempt in range(1, phone_retries + 1):
                        if not activation:
                            activation = await acquire_activation()
                            account_record.update(
                                {
                                    "stage": "phone_acquired",
                                    "status": "in_progress",
                                    "phone": activation.phone,
                                    "activation_id": activation.activation_id,
                                    "activation_cost": activation.activation_cost,
                                    "phone_attempt": phone_attempt,
                                }
                            )
                            upsert_account_record(account_record, account_record_file)
                            logger.info(f"第 2 步: 重新获取 HeroSMS 手机号，手机号={activation.phone} ({phone_attempt}/{phone_retries})")
                        try:
                            await _navigate_to_phone_signup(bot, activation)
                            account_record.update({"stage": "phone_submitted", "phone_attempt": phone_attempt})
                            upsert_account_record(account_record, account_record_file)
                            await _complete_phone_registration(
                                bot=bot,
                                sms_client=sms_client,
                                activation=activation,
                                full_name=round_real_name,
                                birthdate=round_birthdate,
                                password=round_password,
                                sms_timeout=hero_sms_timeout,
                                sms_interval=hero_sms_interval,
                                finish_after=hero_sms_finish_after,
                            )
                            break
                        except (SmsVerificationTimeoutError, PhoneAlreadyRegisteredError) as exc:
                            is_phone_exists = isinstance(exc, PhoneAlreadyRegisteredError)
                            retry_title = "手机号已存在账号" if is_phone_exists else "手机号验证码超时"
                            retry_stage = "phone_already_registered" if is_phone_exists else "phone_sms_timeout"
                            logger.warning(
                                f"{retry_title}，准备回退注册页并换号重试 ({phone_attempt}/{phone_retries})"
                            )
                            notify_failure(
                                retry_title,
                                f"第 {round_num}/{rounds} 轮手机号 {activation.phone} {retry_title}，"
                                f"准备换号重试 ({phone_attempt}/{phone_retries})。",
                            )
                            account_record.update(
                                {
                                    "stage": retry_stage,
                                    "status": "retrying_phone",
                                    "phone": activation.phone,
                                    "activation_id": activation.activation_id,
                                    "phone_attempt": phone_attempt,
                                    "phone_retry_reason": type(exc).__name__,
                                }
                            )
                            upsert_account_record(account_record, account_record_file)
                            if phone_attempt >= phone_retries:
                                raise
                            activation = None
                            continue
                    account_record.update({"stage": "about_you_completed", "status": "registered"})
                upsert_account_record(account_record, account_record_file)

                if not address:
                    account_record.update({"stage": "registered_no_email", "status": "partial", "oauth_status": "skipped_no_email"})
                    upsert_account_record(account_record, account_record_file)
                    logger.warning(f"第 {round_num} 轮: 无可用邮箱，跳过 OAuth/Sub2API")
                    # 跳过 OAuth，但仍完成 claim 记录本轮结果

                if address:
                    # ── 注册完成，导航到 chatgpt.com 建立登录态 ──
                    logger.info("第 6 步: 导航到 chatgpt.com 首页确认登录")
                    if not bot.page:
                        raise SignupFlowError("浏览器页面未初始化")
                    await bot.page.goto("https://chatgpt.com/", wait_until="domcontentloaded", timeout=60000)
                    await _random_human_delay(2000, 4500)

                    # ── OAuth 完整流程：欢迎回来→密码→添加邮箱→验证→Codex授权 ──
                    logger.info("第 7 步: Codex OAuth 5步流程导出 Sub2API")
                    token_data = await _execute_oauth_flow(
                        bot=bot,
                        email_address=address,
                        password=round_password,
                        oauth_phone=resume_phone if resume_phone else (activation.phone if activation else ""),
                        outlook_client=outlook_client,
                        email_timeout=args.email_timeout,
                    )
                    account_file_name = _safe_output_name(address, f"round-{round_num}") + ".json"
                    account_output_file = os.path.join(output_dir, account_file_name)
                    Sub2APIExporter.FILE = os.path.abspath(account_output_file)
                    output_path = Sub2APIExporter.export_account(token_data, notes="GPT free registration 导出")
                    batch_output_path = ""
                    if rounds > 1:
                        Sub2APIExporter.FILE = os.path.abspath(batch_output_file)
                        batch_output_path = Sub2APIExporter.export_account(token_data, notes="GPT free registration 批次导出")
                    account_record.update(
                        {
                            "stage": "sub2api_exported",
                            "status": "completed",
                            "oauth_status": "success",
                            "sub2api_output": output_path,
                            "sub2api_batch_output": batch_output_path,
                        }
                    )
                    upsert_account_record(account_record, account_record_file)

            if address:
                await release_claim("免费注册并导出 Sub2API 成功，释放邮箱以便复用")
                successes += 1
            else:
                await complete_claim("verification_timeout", "注册完成但无可用邮箱，未执行 OAuth/Sub2API")
                partials += 1
            print(f"\n{'=' * 60}")
            print(f"✅ 第 {round_num}/{rounds} 轮 {'成功' if address else '部分完成'}")
            print(f"手机号: {resume_phone if resume_phone else (activation.phone if activation else '')}")
            print(f"邮箱: {address}")
            print(f"密码: {round_password}")
            print(f"Sub2API文件: {output_path}")
            print(f"{'=' * 60}")

        except EmailPoolExhaustedError as exc:
            failures += 1
            batch_aborted = True
            batch_abort_reason = str(exc)
            logger.error(f"第 {round_num} 轮 邮箱池耗尽，终止批量运行: {exc}")
            notify_failure("邮箱池耗尽", f"第 {round_num}/{rounds} 轮失败并终止批量: {exc}")
            account_record.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
            upsert_account_record(account_record, account_record_file)
            await release_claim(str(exc))
            if args.keep_open:
                await _keep_browser_open_on_error(active_bot, type(exc).__name__)
        except (VerificationTimeout, OutlookEmailPlusTimeout) as exc:
            failures += 1
            logger.error(f"第 {round_num} 轮 验证邮件超时: {exc}")
            notify_failure("验证邮件超时", f"第 {round_num}/{rounds} 轮失败: {exc}")
            account_record.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
            upsert_account_record(account_record, account_record_file)
            await complete_claim("verification_timeout", str(exc))
            if args.keep_open:
                await _keep_browser_open_on_error(active_bot, type(exc).__name__)
        except OutlookEmailPlusError as exc:
            failures += 1
            logger.error(f"第 {round_num} 轮 OutlookEmailPlus 失败: {exc}")
            notify_failure("OutlookEmailPlus 失败", f"第 {round_num}/{rounds} 轮失败: {exc}")
            account_record.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
            upsert_account_record(account_record, account_record_file)
            await release_claim(str(exc))
            if args.keep_open:
                await _keep_browser_open_on_error(active_bot, type(exc).__name__)
        except UserAlreadyExistsError as exc:
            failures += 1
            logger.error(f"第 {round_num} 轮 ChatGPT 账号已存在: {exc}")
            notify_failure("ChatGPT 账号已存在", f"第 {round_num}/{rounds} 轮失败: {exc}")
            account_record.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
            upsert_account_record(account_record, account_record_file)
            await complete_claim("credential_invalid", str(exc))
            if args.keep_open:
                await _keep_browser_open_on_error(active_bot, type(exc).__name__)
        except (SignupFlowError, Exception) as exc:
            failures += 1
            logger.error(f"第 {round_num} 轮 注册流程失败: {type(exc).__name__}: {exc}")
            notify_failure("注册流程失败", f"第 {round_num}/{rounds} 轮 {type(exc).__name__}: {exc}")
            account_record.update({"status": "failed", "error_type": type(exc).__name__, "error": str(exc)})
            upsert_account_record(account_record, account_record_file)
            await release_claim(f"{type(exc).__name__}: {exc}")
            if args.keep_open:
                await _keep_browser_open_on_error(active_bot, type(exc).__name__)
        finally:
            if outlook_client:
                await outlook_client.close()

        if batch_aborted:
            break

    # ── 批量运行摘要 ──
    elapsed = time.time() - batch_start
    print(f"\n{'█' * 60}")
    print(f"  批量运行结束  |  总轮数: {rounds}  |  成功: {successes}  |  部分完成: {partials}  |  失败: {failures}")
    if batch_aborted:
        print(f"  已提前终止: {batch_abort_reason}")
    print(f"  总耗时: {elapsed:.0f}s ({elapsed/60:.1f}min)")
    if rounds > 0:
        print(f"  成功率: {successes/rounds*100:.1f}%")
    print(f"{'█' * 60}")
    await drain_windows_asyncio_shutdown()


if __name__ == "__main__":
    asyncio.run(main())
