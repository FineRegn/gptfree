# -*- coding: utf-8 -*-
"""
ChatGPT 免费注册核心模块
================================================================
功能：
    1. 创建临时邮箱（自托管邮箱服务，可替换）
    2. ChatGPT 注册（Playwright 浏览器自动化）
    3. 自动提取邮箱验证码
    4. Codex OAuth 授权
    5. 直接导出 Sub2API OAuth 账号 JSON

依赖：
    pip install playwright aiohttp
    playwright install chromium

⚠ 必读：使用前请阅读下方「用户配置区域」并填写所有 YOUR_xxx 占位符，
否则脚本无法运行。
================================================================
"""
import argparse
import asyncio
import base64
import hashlib
import json
import logging
import os
import random
import re
import secrets
import string
import sys
import time
from datetime import datetime, timezone
from typing import Any, Dict, Optional
from urllib.parse import parse_qs, urlencode, urlparse

import aiohttp
from playwright.async_api import async_playwright, Page, Browser, BrowserContext

from hero_sms_client import HeroSMSClient, HeroSMSError
from outlook_email_plus_integration import OutlookEmailPlusClient, OutlookEmailPlusError, OutlookEmailPlusTimeout

# ===================== 日志 =====================
LOG_FILE = os.path.join(os.path.dirname(__file__), "registration_bot.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "gpt_free_config.json")


def _load_local_config() -> dict:
    """读取本地 JSON 配置；文件不存在时使用安全默认值。"""
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"[配置错误] 读取 {CONFIG_FILE} 失败: {exc}") from exc
    if not isinstance(data, dict):
        raise SystemExit(f"[配置错误] {CONFIG_FILE} 顶层必须是 JSON 对象")
    return data


_LOCAL_CONFIG = _load_local_config()


def _config_str(key: str, default: str = "") -> str:
    value = _LOCAL_CONFIG.get(key, default)
    return str(value).strip() if value is not None else ""


def _config_int(key: str, default: int) -> int:
    value = _LOCAL_CONFIG.get(key, default)
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise SystemExit(f"[配置错误] {key} 必须是整数") from exc

# 全局调试标志
DEBUG_MODE = False
_bot_instance = None  # 保存 ChatGPTBot 引用，用于调试暂停
HEROSMS_APIKEY = _config_str("HEROSMS_APIKEY")
HEROSMS_INTERVAL = 5
HEROSMS_FINISH_AFTER = False

# ============================================================================
# ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼  用 户 配 置 区 域 (USER CONFIG)  ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
# ============================================================================
# 所有需要你自行填写的内容都集中在这一节。把 "YOUR_xxx" 占位符替换为
# 你自己的真实值即可。其他章节通常不需要改动。
# ----------------------------------------------------------------------------

# -------- 1. 临时邮箱服务 ----------------------------------------------------
# 本脚本依赖一个支持「匿名按需创建地址」的临时邮箱后端，用来在 ChatGPT
# 注册时收件、提取 OTP。
#
# 重要：这个后端不需要你设置任何"管理员密码 / API key"！
# 脚本运作方式：
#   1) 调 POST /api/new_address（无 Authorization 头）匿名创建一个新地址
#      响应里会自动给你 { address, password, jwt }，password 是后端随机
#      生成并返回给你的，不是你预先设置的；
#   2) 用响应里的 jwt 当作 Bearer Token 调 GET /api/parsed_mails 拉收件箱；
#   3) 整个流程结束后地址 + password 会被写进 accounts.json / accounts.txt，
#      下次想再登录这个地址就用那个 password。
# 因此你只需要"有一个能匿名创建地址的后端"即可。
#
# 推荐部署方式（开源 + 免费）：
#   ★ dreamhunter2333/cloudflare_temp_email
#     https://github.com/dreamhunter2333/cloudflare_temp_email
#     基于 Cloudflare Workers + D1 + Pages，0 成本、5-10 分钟搭好，
#     原生提供本脚本所需的全部 3 个接口。
#
# 不想自部署 / 想用其他临时邮箱服务（mail.tm 等）？
#   你需要修改下方的 ⌈TempEmailClient⌋ 类，把以下 3 个方法的请求路径
#   和响应字段改成目标服务的格式：
#     - create_address    → 创建地址
#     - poll_for_emails   → 轮询收件箱
#     - address_login     → （可选）复用旧地址
#
# 部署完成后把后端域名填到这里（不要带尾部斜杠）：
TEMP_EMAIL_API = "https://YOUR_TEMP_EMAIL_API_DOMAIN"

# 自动登录链接前缀。脚本会生成 "{TEMP_EMAIL_LOGIN_BASE}?jwt=xxx" 形式的
# 链接保存在 accounts.txt 里，用来事后手动登录查看邮件。
# 如果你的部署 API 和 Web 在同一个域名（最常见情况），填同上即可。
# 仅当你把后端 API 和前端 Web 分别部到不同子域时才需要分开填。
TEMP_EMAIL_LOGIN_BASE = "https://YOUR_TEMP_EMAIL_API_DOMAIN"

# -------- 1.1 OutlookEmailPlus 服务 -----------------------------------------
OUTLOOK_EMAIL_PLUS_API_BASE = _config_str("OUTLOOK_EMAIL_PLUS_API_BASE", "http://127.0.0.1:5001")
OUTLOOK_EMAIL_PLUS_API_KEY = _config_str("OUTLOOK_EMAIL_PLUS_API_KEY")
OUTLOOK_EMAIL_PLUS_PROVIDER = _config_str("OUTLOOK_EMAIL_PLUS_PROVIDER", "outlook")
OUTLOOK_EMAIL_PLUS_CALLER_ID = _config_str("OUTLOOK_EMAIL_PLUS_CALLER_ID", "chatgpt-registration-bot")
OUTLOOK_EMAIL_PLUS_PROJECT_KEY = _config_str("OUTLOOK_EMAIL_PLUS_PROJECT_KEY")
OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN = _config_str("OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN")
OUTLOOK_EMAIL_PLUS_BACKEND = _config_str("OUTLOOK_EMAIL_PLUS_BACKEND", "outlook_email_plus")  # 可选: outlook_email_plus / temp_email

# -------- 3. ChatGPT 注册默认参数 -------------------------------------------
# 注册时填写的"姓"。命令行 --name 可临时覆盖。
# 留默认时会自动生成英文随机姓名（推荐）。
DEFAULT_NAME = "John Doe"

# 临时邮箱前缀的标记，最终生成形如 "<TAG><YYYYMMDD><随机字符>"，例如
# 默认会生成 "oai20260507a"。你可以改成自己的标识便于追踪。
EMAIL_PREFIX_TAG = "oai"

# 自备邮箱母邮箱。启用手动邮箱模式但未通过 --email 指定邮箱时，脚本会自动
# 生成形如 "fineregn+bfh@googlemail.com" 的别名邮箱用于注册。
# 留空则保持原行为：启动后手动输入完整邮箱地址。
MANUAL_EMAIL_BASE = _config_str("MANUAL_EMAIL_BASE")

# 自备邮箱别名长度。例如长度 3 时会生成 bfh 这类随机别名。
MANUAL_EMAIL_ALIAS_LENGTH = _config_int("MANUAL_EMAIL_ALIAS_LENGTH", 5)

# ============================================================================
# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲ 用户配置区域结束 ▲▲▲▲▲▲▲▲▲▲▲▲▲▲
# ============================================================================


# ===================== 系统配置（一般不需要改）=====================
CHATGPT_URL = "https://chatgpt.com"
SCREENSHOTS_DIR = os.path.join(os.path.dirname(__file__), "screenshots")
ACCOUNTS_FILE = os.path.join(os.path.dirname(__file__), "accounts.json")
ACCOUNTS_TXT_FILE = os.path.join(os.path.dirname(__file__), "accounts.txt")
SUB2API_ACCOUNTS_FILE = os.path.join(os.path.dirname(__file__), "sub2api-accounts.json")
CODEX_OAUTH_CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
CODEX_OAUTH_REDIRECT_URI = "http://localhost:1455/auth/callback"
DEFAULT_OAUTH_PROXY = "http://127.0.0.1:7897"

os.makedirs(SCREENSHOTS_DIR, exist_ok=True)


# ===================== 启动校验：用户配置是否完整 =====================
def _validate_user_config(skip_temp_email: bool = False, email_backend: str = "temp_email"):
    """启动时检查 USER CONFIG 区域是否还有未填写的 YOUR_xxx 占位符。
    skip_temp_email: --manual-email 模式可跳过临时邮箱配置检查。
    """
    placeholders = {}
    if not skip_temp_email and email_backend == "temp_email":
        placeholders["TEMP_EMAIL_API"] = TEMP_EMAIL_API
        placeholders["TEMP_EMAIL_LOGIN_BASE"] = TEMP_EMAIL_LOGIN_BASE
    if not skip_temp_email and email_backend == "outlook_email_plus":
        placeholders["OUTLOOK_EMAIL_PLUS_API_BASE"] = OUTLOOK_EMAIL_PLUS_API_BASE
        placeholders["OUTLOOK_EMAIL_PLUS_API_KEY"] = OUTLOOK_EMAIL_PLUS_API_KEY
        placeholders["OUTLOOK_EMAIL_PLUS_CALLER_ID"] = OUTLOOK_EMAIL_PLUS_CALLER_ID
    missing = [k for k, v in placeholders.items() if "YOUR_" in str(v).upper()]
    if missing:
        raise SystemExit(
            "\n[配置缺失] 请先在脚本顶部「用户配置区域」填写以下变量：\n  - "
            + "\n  - ".join(missing)
            + "\n详见脚本头部 USER CONFIG 注释。\n"
        )
# ===================== 工具函数 =====================
def generate_strong_password(length: int = 16) -> str:
    alphabet = string.ascii_letters + string.digits + "!@#$%^&*"
    return "".join(secrets.choice(alphabet) for _ in range(length))


def generate_birthdate() -> str:
    """生成随机出生日期（18-35岁）"""
    import datetime as dt
    today = dt.date.today()
    years_ago = random.randint(18, 35)
    days_ago = random.randint(0, 365)
    bd = today - dt.timedelta(days=365 * years_ago + days_ago)
    return bd.strftime("%Y-%m-%d")


# 真人姓名库
REAL_FIRST_NAMES = [
    "James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph",
    "Thomas", "Charles", "Mary", "Patricia", "Jennifer", "Linda", "Barbara",
    "Elizabeth", "Susan", "Jessica", "Sarah", "Karen", "Lisa", "Nancy", "Betty",
    "Margaret", "Sandra", "Ashley", "Dorothy", "Kimberly", "Emily", "Donna",
    "Michelle", "Carol", "Amanda", "Melissa", "Deborah", "Stephanie", "Rebecca",
    "Sharon", "Laura", "Cynthia", "Kathleen", "Amy", "Angela", "Shirley",
    "Anna", "Brenda", "Pamela", "Emma", "Nicole", "Helen", "Samantha",
]
REAL_LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
    "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson",
    "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson",
    "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson",
    "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen",
    "Hill", "Flores", "Green", "Adams", "Nelson", "Baker", "Hall", "Rivera",
    "Campbell", "Mitchell", "Carter", "Roberts",
]


def generate_real_name() -> str:
    return f"{random.choice(REAL_FIRST_NAMES)} {random.choice(REAL_LAST_NAMES)}"


def generate_email_prefix() -> str:
    """生成 <TAG><YYYYMMDD><随机字符> 格式的前缀。TAG 只有 USER CONFIG 里的 EMAIL_PREFIX_TAG 控制。"""
    today = datetime.now().strftime("%Y%m%d")
    suffix = random.choice("abcdefghjkmnpqrstuvwxyz")
    return f"{EMAIL_PREFIX_TAG}{today}{suffix}"


def is_valid_email(email: str) -> bool:
    text = (email or "").strip()
    return bool(text and "@" in text and "." in text.split("@", 1)[-1])


def generate_manual_email_alias(base_email: str, alias_length: int = MANUAL_EMAIL_ALIAS_LENGTH) -> str:
    """基于母邮箱生成 plus alias 注册邮箱。"""
    text = (base_email or "").strip()
    if not is_valid_email(text):
        raise ValueError("母邮箱格式不正确")
    local, domain = text.split("@", 1)
    local = local.split("+", 1)[0]
    length = max(1, int(alias_length or 3))
    alias = "".join(secrets.choice(string.ascii_lowercase) for _ in range(length))
    return f"{local}+{alias}@{domain}"


def normalize_proxy_url(proxy: str) -> str:
    """规范化代理地址，支持直接传端口号。"""
    text = (proxy or "").strip()
    if not text:
        return ""
    if text.isdigit():
        return f"http://127.0.0.1:{text}"
    if "://" not in text:
        return f"http://{text}"
    return text


async def wait_for_code_file(path: str, label: str, timeout: int = 600) -> str:
    """等待外部写入验证码文件，用于非交互式测试流程。"""
    start_time = time.time()
    deadline = start_time + timeout
    logger.info(f"等待{label}写入文件: {path}")
    while time.time() < deadline:
        try:
            if os.path.exists(path) and os.path.getmtime(path) >= start_time:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                match = re.search(r"\b(\d{4,8})\b", content)
                if match:
                    code = match.group(1)
                    logger.info(f"已从文件读取{label}: {code[:2]}***")
                    return code
        except Exception as e:
            logger.warning(f"读取{label}文件失败: {e}")
        await asyncio.sleep(2)
    raise VerificationTimeout(f"{timeout}s 内未等到{label}文件验证码: {path}")


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def decode_jwt_payload(token: str) -> dict:
    if not token or token.count(".") < 2:
        return {}
    payload = token.split(".", 2)[1]
    payload += "=" * (-len(payload) % 4)
    try:
        return json.loads(base64.urlsafe_b64decode(payload.encode("ascii")).decode("utf-8"))
    except Exception:
        return {}


def first_non_empty(*values) -> str:
    for value in values:
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def first_audience_value(audience) -> str:
    if isinstance(audience, str):
        return audience.strip()
    if isinstance(audience, list):
        for item in audience:
            if isinstance(item, str) and item.strip():
                return item.strip()
    return ""


def extract_organization_id(raw_organizations) -> str:
    if not isinstance(raw_organizations, list):
        return ""
    first_org = ""
    for org in raw_organizations:
        if not isinstance(org, dict):
            continue
        org_id = first_non_empty(org.get("id"))
        if not org_id:
            continue
        if not first_org:
            first_org = org_id
        if org.get("is_default") is True:
            return org_id
    return first_org


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


async def retry_with_backoff(coro_fn, max_retries=3, base_delay=2.0, description="operation"):
    for attempt in range(max_retries):
        try:
            return await coro_fn()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            delay = base_delay * (2**attempt)
            logger.warning(f"{description} 失败 (重试 {attempt+1}/{max_retries}): {e}")
            await asyncio.sleep(delay)


async def screenshot(page: Page, name: str):
    # 截图功能已禁用（调试用途）。调用点遍布全文未删改；
    # 如需恢复截图，取消下面被注释的代码块并删掉这行 return 即可。
    return
    # path = os.path.join(SCREENSHOTS_DIR, f"{time.strftime('%H%M%S')}_{name}.png")
    # try:
    #     await page.screenshot(path=path, full_page=True)
    #     logger.info(f"截图: {path}")
    # except Exception as e:
    #     logger.warning(f"截图失败: {e}")


async def debug_pause(page: Page, reason: str):
    """调试模式：打印当前状态并保持浏览器打开，等待手动检查。"""
    global _bot_instance
    logger.error("=" * 60)
    logger.error(f"调试暂停: {reason}")
    logger.error(f"当前 URL: {page.url[:150]}")
    try:
        text = await page.evaluate("() => document.body?.innerText?.substring(0, 800) || '(no body text)'")
        logger.error(f"页面文本:\n{text}")
    except Exception:
        pass
    await screenshot(page, f"DEBUG_{reason.replace(' ', '_')[:40]}")
    logger.error(f"截图已保存到 screenshots/ 目录")
    logger.error("浏览器保持打开。调试完成后按 Ctrl+C 退出。")
    logger.error("=" * 60)
    try:
        while True:
            await asyncio.sleep(10)
    except asyncio.CancelledError:
        pass


async def log_page_state(page: Page, step_name: str):
    """输出当前页面状态用于调试"""
    url = page.url
    title = await page.title()
    try:
        body = await page.evaluate("() => document.body?.innerText?.substring(0, 300) || '(empty)'")
    except Exception:
        body = "(无法读取)"
    logger.info(f"[{step_name}] URL={url[:120]}")
    logger.info(f"[{step_name}] Title={title}")
    logger.info(f"[{step_name}] Body={body[:200]}")


# ===================== 自定义异常 =====================
class TempEmailError(Exception):
    pass


class SignupFlowError(Exception):
    pass


class UserAlreadyExistsError(SignupFlowError):
    """ChatGPT 提示该邮箱账号已存在。"""

    def __init__(self, email: str = ""):
        self.email = email
        message = "ChatGPT 提示该邮箱账号已存在"
        if email:
            message = f"{message}: {email}"
        super().__init__(message)


class VerificationTimeout(Exception):
    pass


# ===================== 1. 验证码提取器 =====================
class VerificationCodeExtractor:
    PATTERNS = [
        (r"\b(\d{6})\b", "6-digit code"),
        (r">(\d{6})<", "6-digit inside tags"),
        (r"\b(\d{5})\b", "5-digit code"),
        (r"\b(\d{4})\b", "4-digit code"),
        (r"\b(\d{8})\b", "8-digit code"),
        (r"code[:\s]*(\d{4,8})", "'code: XXXX' (case insensitive)"),
        (r"OTP[:\s]*(\d{4,8})", "'OTP: XXXX' (case insensitive)"),
        (r"verification[^0-9]*(\d{4,8})", "'verification code: XXXX'"),
        (r"confirm[^0-9]*(\d{4,8})", "'confirm: XXXX'"),
        (r"(\d{4,8})\s*is your", "'XXXX is your code'"),
        (r"验证码[:\s]*(\d{4,8})", "中文验证码"),
        (r"kode[:\s]*(\d{4,8})", "'kode: XXXX' (Bahasa)"),
    ]

    @staticmethod
    def _strip_html(text: str) -> str:
        return re.sub(r"<[^>]+>", " ", text)

    @classmethod
    def extract_from_text(cls, text: str) -> Optional[str]:
        if not text:
            return None
        clean = cls._strip_html(text)
        for pattern, _desc in cls.PATTERNS:
            match = re.search(pattern, clean, re.IGNORECASE)
            if match:
                return match.group(1)
        return None

    @classmethod
    def extract_from_subject(cls, subject: str) -> Optional[str]:
        if not subject:
            return None
        return cls.extract_from_text(subject)

    @classmethod
    def extract_from_html(cls, html: str) -> Optional[str]:
        if not html:
            return None
        # 从 <b>/<strong> 标签或数字串中提取
        bold_match = re.search(r"<b[^>]*>(\d{4,8})</b>", html, re.IGNORECASE)
        if bold_match:
            return bold_match.group(1)
        strong_match = re.search(r"<strong[^>]*>(\d{4,8})</strong>", html, re.IGNORECASE)
        if strong_match:
            return strong_match.group(1)
        return cls.extract_from_text(html)

    @classmethod
    def find_verification_link(cls, text: str, html: str) -> Optional[str]:
        content = (html or "") + (text or "")
        urls = re.findall(r"https?://[^\s<>\"'\)]+", content)
        for url in urls:
            if any(k in url.lower() for k in ("verify", "confirm", "activate", "email-verification")):
                return url
        return None

    @classmethod
    def comprehensive_extract(cls, mail: dict) -> Optional[dict]:
        subject = mail.get("subject", "")
        text_body = mail.get("text", "")
        html_body = mail.get("html", "")

        for source, content in [("subject", subject), ("text", text_body), ("html", html_body)]:
            code = cls.extract_from_text(content)
            if code:
                return {"code": code, "source": source}

        link = cls.find_verification_link(text_body, html_body)
        if link:
            code_from_link = cls.extract_from_text(link)
            if code_from_link:
                return {"code": code_from_link, "source": "link"}

        return None


# ===================== 2. 临时邮箱客户端 =====================
class TempEmailClient:
    BASE = TEMP_EMAIL_API
    _session: Optional[aiohttp.ClientSession] = None
    latest_jwt: str = ""

    async def __aenter__(self):
        self._session = aiohttp.ClientSession(
            timeout=aiohttp.ClientTimeout(total=30),
            connector=aiohttp.TCPConnector(limit=5),
        )
        self.latest_jwt = ""
        return self

    async def __aexit__(self, *args):
        if self._session:
            await self._session.close()

    def _assert_session(self) -> aiohttp.ClientSession:
        if not self._session:
            raise TempEmailError("Session 未初始化，使用 async with TempEmailClient() as client:")
        return self._session

    @staticmethod
    def sha256(text: str) -> str:
        return hashlib.sha256(text.encode()).hexdigest()

    async def create_address(self, name: str = "") -> dict[str, Any]:
        s = self._assert_session()
        payload = {"name": name} if name else {}
        async with s.post(f"{self.BASE}/api/new_address", json=payload) as resp:
            if resp.status == 429:
                raise TempEmailError("API 限流 (429)")
            body = await resp.json()
            if resp.status != 200:
                raise TempEmailError(f"创建地址失败: {resp.status} {body}")
            self.latest_jwt = self._extract_jwt(body)
            logger.info(f"✅ 临时邮箱: {body['address']}  (password: {body.get('password', 'N/A')})")
            return body

    @staticmethod
    def _extract_jwt(data: Dict[str, Any]) -> str:
        for key in ("jwt", "token", "access_token"):
            value = data.get(key) if isinstance(data, dict) else None
            if isinstance(value, str) and value.strip():
                return value.strip()
        return ""

    async def _refresh_jwt(self, email_addr: str, plaintext_password: str) -> str:
        if not email_addr or not plaintext_password:
            return ""
        logger.warning("邮箱 JWT 可能已失效，尝试使用邮箱密码重新登录刷新 token")
        refreshed = await self.address_login(email_addr, plaintext_password)
        new_jwt = self._extract_jwt(refreshed)
        if new_jwt:
            self.latest_jwt = new_jwt
            logger.info("✅ 邮箱 JWT 已刷新，继续轮询收件箱")
        return new_jwt

    async def poll_for_emails(
        self,
        jwt: str,
        timeout: int = 120,
        interval: int = 5,
        email_addr: str = "",
        plaintext_password: str = "",
    ) -> list[dict[str, Any]]:
        s = self._assert_session()
        current_jwt = jwt
        headers = {"Authorization": f"Bearer {current_jwt}"}
        seen_ids = set()
        deadline = time.time() + timeout
        refresh_attempted = False

        logger.info(f"等待验证邮件... (最长 {timeout}s)")

        while time.time() < deadline:
            try:
                async with s.get(
                    f"{self.BASE}/api/parsed_mails", params={"offset": "0", "limit": "10"}, headers=headers
                ) as resp:
                    if resp.status in (401, 403) and not refresh_attempted:
                        refresh_attempted = True
                        new_jwt = await self._refresh_jwt(email_addr, plaintext_password)
                        if new_jwt:
                            current_jwt = new_jwt
                            headers = {"Authorization": f"Bearer {current_jwt}"}
                            await asyncio.sleep(1)
                            continue
                    if resp.status != 200:
                        logger.warning(f"查询邮件失败: {resp.status}")
                        await asyncio.sleep(interval)
                        continue
                    data = await resp.json()
                    results = data.get("results", [])
                    new_mails = [m for m in results if m.get("id") not in seen_ids]

                    if new_mails:
                        for m in new_mails:
                            seen_ids.add(m["id"])
                            logger.info(
                                f"📧 新邮件 #{m['id']}: from={m.get('source','?')[:50]} subj={m.get('subject','?')[:60]}"
                            )
                        if results:
                            return results

            except Exception as e:
                logger.warning(f"轮询异常: {e}")

            await asyncio.sleep(interval)

        raise VerificationTimeout(f"{timeout}s 内未收到邮件")

    async def get_parsed_mail(
        self,
        jwt: str,
        mail_id: int,
        email_addr: str = "",
        plaintext_password: str = "",
    ) -> dict[str, Any]:
        s = self._assert_session()
        headers = {"Authorization": f"Bearer {jwt}"}
        async with s.get(f"{self.BASE}/api/parsed_mail/{mail_id}", headers=headers) as resp:
            if resp.status in (401, 403):
                new_jwt = await self._refresh_jwt(email_addr, plaintext_password)
                if new_jwt:
                    headers = {"Authorization": f"Bearer {new_jwt}"}
                    async with s.get(f"{self.BASE}/api/parsed_mail/{mail_id}", headers=headers) as retry_resp:
                        if retry_resp.status != 200:
                            raise TempEmailError(f"获取邮件失败: {retry_resp.status}")
                        return await retry_resp.json()
            if resp.status != 200:
                raise TempEmailError(f"获取邮件失败: {resp.status}")
            return await resp.json()

    async def address_login(self, email_addr: str, plaintext_password: str) -> dict[str, Any]:
        s = self._assert_session()
        hashed = self.sha256(plaintext_password)
        payload = {"email": email_addr, "password": hashed}
        async with s.post(f"{self.BASE}/api/address_login", json=payload) as resp:
            if resp.status != 200:
                body = await resp.text()
                raise TempEmailError(f"登录失败: {resp.status} {body}")
            body = await resp.json()
            self.latest_jwt = self._extract_jwt(body)
            return body


# ===================== 3. ChatGPT 浏览器自动化 =====================
class ChatGPTBot:
    CHATGPT_URL = "https://chatgpt.com"
    USER_ALREADY_EXISTS_RE = re.compile(
        r"user_already_exists|user already exists|account already exists|账号已存在|帳戶已存在|账户已存在|用户已存在|已存在.*账号|已存在.*账户",
        re.IGNORECASE,
    )

    def __init__(self, headless: bool = False, slow_mo: int = 100, proxy_url: str = ""):
        self.headless = headless
        self.slow_mo = slow_mo
        self.proxy_url = normalize_proxy_url(proxy_url)
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.playwright = None
        # 注册时填的姓名，支付时用于 Stripe 账单字段
        self.registration_name: str = ""
        self.current_plan_type: str = ""
        self.oauth_code_file: str = ""
        self.registration_email: str = ""
        self.allow_existing_account_login: bool = False

    async def __aenter__(self):
        await self.launch()
        return self

    async def __aexit__(self, *args):
        if not DEBUG_MODE:
            await self.close()
        else:
            logger.info("调试模式 — 浏览器保持打开")

    async def launch(self):
        global _bot_instance
        _bot_instance = self
        logger.info("启动浏览器...")
        self.playwright = await async_playwright().__aenter__()
        launch_args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-features=ChromeWhatsNewUI",
            "--no-first-run",
            "--no-default-browser-check",
        ]
        if self.proxy_url:
            logger.info(f"浏览器代理: {self.proxy_url}")
            self.browser = await self.playwright.chromium.launch(
                headless=self.headless,
                slow_mo=self.slow_mo,
                args=launch_args,
                ignore_default_args=["--enable-automation"],
                proxy={"server": self.proxy_url},
            )
        else:
            self.browser = await self.playwright.chromium.launch(
                headless=self.headless,
                slow_mo=self.slow_mo,
                args=launch_args,
                ignore_default_args=["--enable-automation"],
            )
        self.context = await self.browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            permissions=["geolocation"],
        )
        await self.context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            window.chrome = window.chrome || { runtime: {} };
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (p) =>
                p.name === 'notifications'
                    ? Promise.resolve({ state: 'prompt', onchange: null })
                    : originalQuery(p);
        """)
        self.page = await self.context.new_page()
        self.page.set_default_timeout(30000)
        logger.info("✅ 浏览器启动完成")

    async def close(self):
        try:
            if self.page:
                await self.page.close()
        except Exception:
            pass
        try:
            if self.context:
                await self.context.close()
        except Exception:
            pass
        try:
            if self.browser:
                await self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                await self.playwright.stop()
        except Exception:
            pass
        logger.info("浏览器已关闭")

    async def _random_delay(self, ms_min=500, ms_max=2000):
        await asyncio.sleep(random.randint(ms_min, ms_max) / 1000)

    async def _type_human_like(self, locator, text: str):
        for ch in text:
            await locator.press(ch, delay=random.randint(50, 150))
            await asyncio.sleep(random.randint(10, 30) / 1000)

    async def _fill_controlled_input(self, locator, value: str) -> str:
        await locator.scroll_into_view_if_needed(timeout=3000)
        # 先尝试 force click（绕过 label 遮挡），失败则用 JS focus
        try:
            await locator.click(timeout=3000, force=True)
        except Exception:
            try:
                await locator.evaluate("(el) => { el.focus(); el.dispatchEvent(new Event('focus', {bubbles:true})); }")
            except Exception:
                pass
        try:
            await self.page.keyboard.press("Control+A")
            await self.page.keyboard.press("Backspace")
        except Exception:
            pass
        try:
            await locator.fill("")
        except Exception:
            pass
        await self.page.keyboard.type(value, delay=random.randint(45, 95))
        await locator.evaluate("""
            (el) => {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }
        """)
        try:
            actual = await locator.input_value(timeout=1000)
        except Exception:
            actual = ""
        if actual != value:
            await locator.evaluate("""
                (el, value) => {
                    const proto = el instanceof HTMLTextAreaElement
                        ? window.HTMLTextAreaElement.prototype
                        : window.HTMLInputElement.prototype;
                    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                    if (setter) setter.call(el, value);
                    else el.value = value;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                }
            """, value)
            try:
                actual = await locator.input_value(timeout=1000)
            except Exception:
                actual = ""
        return actual

    async def _fill_age_only_input(self, age_value: str) -> bool:
        candidates = [
            self.page.locator('input[name="age"]').first,
            self.page.get_by_placeholder(re.compile(r"^(年龄|age)$", re.IGNORECASE)).first,
            self.page.get_by_label(re.compile(r"^(年龄|age)$", re.IGNORECASE)).first,
            self.page.get_by_role("spinbutton", name=re.compile(r"(年龄|age)", re.IGNORECASE)).first,
            self.page.locator('input[type="number"]').first,
        ]
        for loc in candidates:
            try:
                if not await loc.is_visible(timeout=1500):
                    continue
                meta = await loc.evaluate("""
                    (el) => ({
                        type: el.getAttribute('type') || '',
                        name: el.getAttribute('name') || '',
                        id: el.getAttribute('id') || '',
                        placeholder: el.getAttribute('placeholder') || '',
                        ariaLabel: el.getAttribute('aria-label') || '',
                        role: el.getAttribute('role') || ''
                    })
                """)
                combined = "".join(str(v) for v in meta.values())
                if re.search(r"(full.?name|姓名|全名)", combined, re.IGNORECASE):
                    continue
                actual = await self._fill_controlled_input(loc, age_value)
                if actual == age_value:
                    logger.info(f"已填写年龄: {actual}")
                    return True
                logger.warning(f"年龄输入校验失败: expected={age_value}, actual={actual}, meta={meta}")
            except Exception as e:
                logger.warning(f"年龄候选输入失败: {e}")
        return False

    async def _detect_user_already_exists(self, step_name: str, email: str = "") -> bool:
        """检测 ChatGPT 注册页是否提示当前邮箱账号已存在。"""
        if not self.page:
            return False
        page = self.page
        try:
            body_text = await page.evaluate("() => document.body?.innerText || ''")
        except Exception:
            body_text = ""
        if not body_text or not self.USER_ALREADY_EXISTS_RE.search(body_text):
            return False
        target_email = email or self.registration_email
        if self.allow_existing_account_login:
            logger.info(f"检测到账号已存在提示，恢复模式继续登录流程: {target_email}")
            await screenshot(page, f"user_already_exists_resume_{step_name}")
            return True
        logger.error(f"检测到 ChatGPT 账号已存在提示 ({step_name}): {target_email}")
        await screenshot(page, f"user_already_exists_{step_name}")
        await log_page_state(page, f"账号已存在 ({step_name})")
        raise UserAlreadyExistsError(target_email)

    async def navigate_to_signup(self, email: str) -> bool:
        logger.info(f"导航到 ChatGPT 注册页, 邮箱={email}")
        self.registration_email = email
        if not self.page:
            raise SignupFlowError("浏览器页面未初始化")
        page = self.page

        await page.goto(self.CHATGPT_URL, wait_until="domcontentloaded")
        await self._random_delay(1500, 2500)
        await screenshot(page, "step_01_chatgpt_home")

        signup_clicked = False
        for sel in [
            page.locator('[data-testid="signup-button"]').first,
            page.get_by_role("link", name=re.compile(r"^(免费注册|注册|Sign up|Create account|Get started)$", re.IGNORECASE)).first,
            page.get_by_role("button", name=re.compile(r"^(免费注册|注册|Sign up|Create account|Get started)$", re.IGNORECASE)).first,
            page.get_by_text(re.compile(r"^(免费注册|注册|Sign up|Create account|Get started)$", re.IGNORECASE)).first,
        ]:
            try:
                if await sel.is_visible(timeout=5000):
                    await sel.click()
                    signup_clicked = True
                    logger.info("已从 chatgpt.com 首页点击注册入口")
                    break
            except Exception:
                continue

        if not signup_clicked:
            logger.warning("未找到 chatgpt.com 首页注册按钮，继续等待邮箱输入框")

        try:
            await page.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception:
            pass
        await self._random_delay(1500, 2500)
        await screenshot(page, "step_01_after_signup_click")

        # 找邮箱输入框（auth.openai.com 整页 OR chatgpt.com dialog）
        email_input = None
        deadline = asyncio.get_event_loop().time() + 30  # 最多等 30 秒
        attempt = 0
        while asyncio.get_event_loop().time() < deadline:
            attempt += 1
            for sel in [
                page.get_by_role("textbox", name=re.compile(r"(电子邮件|email)", re.IGNORECASE)).first,
                page.locator('input[type="email"]').first,
                page.locator('input[name="email"]').first,
                page.locator('input[autocomplete="email"]').first,
                page.locator('input[id*="email" i]').first,
                page.get_by_placeholder(re.compile(r"(电子邮件|邮箱|email)", re.IGNORECASE)).first,
            ]:
                try:
                    if await sel.is_visible(timeout=1200):
                        email_input = sel
                        break
                except Exception:
                    continue
            if email_input:
                break
            # 偶尔可能页面跳转到了 auth.openai.com 但 DOM 还在加载，多等一下
            await asyncio.sleep(1)

        if not email_input:
            page_text = await page.evaluate("() => document.body.innerText.substring(0, 1500)")
            logger.error(f"找不到邮箱输入框，当前 URL: {page.url}")
            logger.error(f"页面文本（前 1500 字）: {page_text}")
            await screenshot(page, "step_02_email_input_missing")
            raise SignupFlowError(f"找不到邮箱输入框（attempts={attempt}, url={page.url[:100]}）")

        actual_email = await self._fill_controlled_input(email_input, email)
        if actual_email != email:
            raise SignupFlowError(f"注册邮箱填写失败: expected={email}, actual={actual_email}")
        await self._random_delay()
        await screenshot(page, "step_02_email_in_dialog")

        # 4. 点击 "继续"
        continue_btn = None
        for sel in [
            page.get_by_role("button", name=re.compile(r"^(继续|Continue)$", re.IGNORECASE)),
        ]:
            try:
                if await sel.is_visible(timeout=5000):
                    continue_btn = sel
                    break
            except Exception:
                continue

        if not continue_btn:
            raise SignupFlowError("找不到继续按钮")

        await continue_btn.click()
        await self._random_delay(800, 1500)
        await self._detect_user_already_exists("email_continue", email)

        # 5. 等待跳转：可能进 password 中间页，也可能直接跳 email-verification
        # （ChatGPT 对部分邮箱会跳过 password 中间页）
        # 用 predicate 匹配任意一个目标 URL，谁先到谁返回，避免死等不存在的中间页
        logger.info("等待跳转到 password 中间页 或 email-verification 页...")
        try:
            await page.wait_for_url(
                lambda url: ("create-account/password" in url) or ("email-verification" in url) or ("log-in/password" in url),
                timeout=30000,
            )
        except Exception:
            logger.warning(f"等待跳转超时，当前 URL: {page.url[:120]}")

        await self._random_delay(800, 1500)
        await screenshot(page, "step_03_password_page")
        current_url = page.url
        await self._detect_user_already_exists("post_email_continue", email)

        if "log-in/password" in current_url:
            if self.allow_existing_account_login:
                logger.info(f"恢复模式检测到已有账号登录页，继续使用邮箱验证码登录: {email}")
            else:
                logger.error(f"注册邮箱已跳转到登录密码页，判定账号已存在: {email}")
                await screenshot(page, "step_03_user_already_exists_login_password")
                await log_page_state(page, "账号已存在 (log-in/password)")
                raise UserAlreadyExistsError(email)

        if "log-in/password" in current_url and self.allow_existing_account_login:
            otp_btn = None
            for sel in [
                page.get_by_role("button", name=re.compile(r"(一次性验证码|one.time|use a code|email.*code|code instead)", re.IGNORECASE)),
                page.get_by_text(re.compile(r"(一次性验证码|use a one.time code|email.*code|code instead)", re.IGNORECASE)),
            ]:
                try:
                    if await sel.first.is_visible(timeout=3000):
                        otp_btn = sel.first
                        break
                except Exception:
                    continue

            if otp_btn:
                await otp_btn.click()
                logger.info("已在已有账号登录页点击邮箱验证码登录")
                await self._random_delay()
            else:
                logger.info("已有账号登录页未找到验证码按钮，继续等待邮箱验证页")

        # 5b. 直接跳到了 email-verification → 跳过 password / OTP 按钮步骤
        if "email-verification" in current_url:
            logger.info("✅ 已直接跳到 email-verification（无中间 password 页）")
            await screenshot(page, "step_04_verification_page")
            return True

        # 6. 处于 password 中间页 → 点 "使用一次性验证码注册"
        if "create-account/password" in current_url:
            otp_btn = None
            for sel in [
                page.get_by_role("button", name=re.compile(r"(一次性验证码|one.time|use a code)", re.IGNORECASE)),
                page.get_by_text(re.compile(r"(一次性验证码|use a one.time code)", re.IGNORECASE)),
            ]:
                try:
                    if await sel.first.is_visible(timeout=3000):
                        otp_btn = sel.first
                        break
                except Exception:
                    continue

            if otp_btn:
                await otp_btn.click()
                await self._random_delay()
            else:
                logger.info("未找到 OTP 按钮，尝试直接等 email-verification 页...")
        elif "log-in/password" not in current_url:
            logger.warning(f"既不在 password 也不在 email-verification: {current_url[:120]}")

        # 7. 等待 email-verification 页面
        try:
            await page.wait_for_url("**/email-verification**", timeout=30000)
            logger.info("✅ 已到达 email-verification 页面")
        except Exception:
            logger.warning("等待 email-verification 超时")
            current_url = page.url
            logger.info(f"当前 URL: {current_url[:120]}")

        await screenshot(page, "step_04_verification_page")
        return True

    async def wait_for_verification_page(self) -> bool:
        try:
            await self.page.wait_for_url("**/email-verification**", timeout=60000)
            logger.info("到达邮箱验证页面")
            await screenshot(self.page, "step_04_verification_page")
            return True
        except Exception:
            logger.warning("未检测到邮箱验证 URL，检查页面内容...")
            url = self.page.url
            if "email-verification" in url or "auth.openai.com" in url:
                return True
            return False

    async def enter_verification_code(self, code: str) -> bool:
        logger.info(f"输入验证码: {code[:2]}...{code[-1]}")
        await screenshot(self.page, "step_05_before_code")

        # email-verification 页面: textbox "验证码"
        code_input = None
        for sel in [
            self.page.get_by_role("textbox", name=re.compile(r"(验证码|verification code)", re.IGNORECASE)),
            self.page.get_by_placeholder(re.compile(r"(code|验证码|verification)", re.IGNORECASE)),
            self.page.locator("form input").first,
        ]:
            try:
                if await sel.is_visible(timeout=5000):
                    code_input = sel
                    break
            except Exception:
                continue

        if not code_input:
            # 尝试 6 个独立数字输入框
            single_inputs = self.page.locator('input[maxlength="1"]')
            count = await single_inputs.count()
            if count >= 4:
                logger.info(f"检测到 {count} 个独立数字输入框")
                for i in range(min(len(code), count)):
                    await single_inputs.nth(i).fill(code[i])
                    await asyncio.sleep(0.1)
            else:
                raise SignupFlowError("找不到验证码输入框")
        else:
            actual = await self._fill_controlled_input(code_input, code)
            if actual != code:
                raise SignupFlowError(f"验证码输入校验失败: expected={code}, actual={actual}")

        await screenshot(self.page, "step_06_code_entered")

        # 点击 "继续" 按钮
        for sel in [
            self.page.get_by_role("button", name=re.compile(r"^(继续|Continue)$", re.IGNORECASE)),
            self.page.locator("button").last,
        ]:
            try:
                if await sel.is_visible(timeout=2000):
                    await sel.click()
                    break
            except Exception:
                continue

        logger.info("验证码已提交，等待跳转...")
        return True

    async def detect_post_verification_destination(self, timeout: int = 30) -> str:
        """验证码提交后判断下一步：about-you 新号流程，或直接进入登录完成等待。"""
        deadline = time.time() + timeout
        last_url = ""

        while time.time() < deadline:
            current_url = self.page.url or ""
            if current_url != last_url:
                logger.info(f"验证码后跳转检测: {current_url[:150]}")
                last_url = current_url

            if "about-you" in current_url or "about_you" in current_url:
                logger.info("检测到 about-you 页面，继续执行新号资料填写")
                await screenshot(self.page, "step_06b_detect_about_you")
                return "about_you"

            if CHATGPT_URL in current_url and "callback" not in current_url and "auth" not in current_url:
                logger.info("验证码后已回到 chatgpt.com，判定跳过 about-you")
                return "logged_in"

            await asyncio.sleep(1)

        logger.warning(f"验证码后 {timeout}s 未进入 about-you，交给登录完成检测处理")
        await log_page_state(self.page, "验证码后未检测到 about-you")
        return "skip_about_you"

    async def fill_about_you(self, name=None, birthdate=None):
        if name is None:
            name = DEFAULT_NAME
        if birthdate is None:
            birthdate = generate_birthdate()
        # 记下注册姓名，用于 Stripe 账单
        self.registration_name = name
        parts = birthdate.split("-")
        year, month, day = parts[0], parts[1], parts[2]

        for _ in range(30):
            url = self.page.url
            if "about-you" in url or "about_you" in url:
                break
            await asyncio.sleep(2)

        logger.info(f"到达 about-you, name={name}, birthdate={birthdate}")
        await screenshot(self.page, "step_07_about_you")
        await self._random_delay()

        name_filled = False
        for name_input in [
            self.page.locator('input[name="name"]').first,
            self.page.get_by_placeholder(re.compile(r"(全名|full name|name|姓名)", re.IGNORECASE)).first,
            self.page.get_by_label(re.compile(r"(全名|full name|name|姓名)", re.IGNORECASE)).first,
            self.page.get_by_role("textbox", name=re.compile(r"(全名|full name|name|姓名)", re.IGNORECASE)).first,
        ]:
            try:
                if await name_input.is_visible(timeout=1500):
                    actual_name = await self._fill_controlled_input(name_input, name)
                    if actual_name == name:
                        name_filled = True
                        logger.info(f"已填写姓名: {name}")
                        break
            except Exception:
                continue
        if not name_filled:
            logger.warning("未找到姓名输入框")

        await self._random_delay(300, 800)

        # 检测页面是要求"年龄"还是"生日日期"
        page_body = await self.page.evaluate("() => document.body?.innerText || ''")
        logger.info(f"about-you 页面关键词检测: body前150字={page_body[:150]}")

        # 先处理 "成年人" 复选框 (常见于年龄验证)
        if any(k in page_body for k in ("成年人", "adult", "18+")):
            try:
                adult_cb = self.page.locator('input[type="checkbox"]').first
                if await adult_cb.is_visible(timeout=2000):
                    if not await adult_cb.is_checked():
                        await adult_cb.click()
                        logger.info("已勾选「成年人」复选框")
            except Exception:
                pass

        has_birthdate = any(k in page_body for k in ("生日", "出生", "birth", "YYYY", "MM", "DD"))
        has_age_only = any(k in page_body for k in ("年龄", "age")) and not has_birthdate

        # 计算一个合理的年龄
        age_value = str(random.randint(22, 30))

        if has_age_only:
            logger.info(f"检测到年龄输入（非日期），填入: {age_value}")
            if not await self._fill_age_only_input(age_value):
                raise SignupFlowError("找不到可填写的年龄输入框")
            try:
                current_name = await self.page.locator('input[name="name"]').first.input_value(timeout=1000)
                if current_name != name:
                    logger.warning(f"姓名字段被改写，重新填写: {current_name!r} -> {name!r}")
                    await self._fill_controlled_input(self.page.locator('input[name="name"]').first, name)
            except Exception:
                pass
        else:
            # ── 尝试 contenteditable 日期分段（React Aria DateField）──
            filled_segments = await self.page.evaluate(f"""
                (() => {{
                    const year = '{year}', month = '{month}', day = '{day}';
                    const segments = document.querySelectorAll(
                        'div[contenteditable="true"][role="spinbutton"]'
                    );
                    const visible = Array.from(segments).filter(el => {{
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    }});
                    const log = [];

                    if (visible.length >= 3) {{
                        // 按 data-type 匹配年/月/日
                        const map = {{}};
                        for (const el of visible) {{
                            const type = el.getAttribute('data-type');
                            if (type === 'year' || type === 'month' || type === 'day') {{
                                map[type] = el;
                            }}
                        }}
                        const values = {{ year, month, day }};
                        for (const [type, el] of Object.entries(map)) {{
                            if (values[type]) {{
                                el.focus();
                                // 清除并填入新值
                                const range = document.createRange();
                                range.selectNodeContents(el);
                                const sel = window.getSelection();
                                sel.removeAllRanges();
                                sel.addRange(range);
                                document.execCommand('delete', false);
                                document.execCommand('insertText', false, values[type]);
                                el.dispatchEvent(new InputEvent('input', {{ bubbles: true, inputType: 'insertText' }}));
                                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                log.push(type + '=' + values[type]);
                            }}
                        }}
                        if (Object.keys(map).length >= 3) {{
                            return JSON.stringify({{ mode: 'contenteditable', count: Object.keys(map).length, log }});
                        }}
                    }}

                    // 兜底：旧 input/role=spinbutton 方式
                    const allInputs = document.querySelectorAll(
                        'input[type="number"], [role="spinbutton"]'
                    );
                    const inputVisible = Array.from(allInputs).filter(el => {{
                        const r = el.getBoundingClientRect();
                        const combined = (el.name || '') + (el.id || '') + (el.getAttribute('aria-label') || '');
                        return r.width > 0 && r.height > 0 && !/name|姓名|full/i.test(combined);
                    }});
                    inputVisible.sort((a, b) => {{
                        const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
                        return (ra.top - rb.top) || (ra.left - rb.left);
                    }});

                    const s = Object.getOwnPropertyDescriptor(
                        window.HTMLInputElement.prototype, 'value'
                    ).set;
                    const set = (el, val) => {{
                        s.call(el, val);
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    }};
                    const values = [year, month, day];
                    for (let i = 0; i < inputVisible.length && i < 3; i++) {{
                        set(inputVisible[i], values[i]);
                        log.push('[' + i + ']=' + values[i]);
                    }}
                    return JSON.stringify({{ mode: 'legacy', count: inputVisible.length, log }});
                }})()
            """)
            logger.info(f"年龄填写 (y/m/d): {filled[:300]}")

        await self._random_delay(300, 500)

        await screenshot(self.page, "step_08_about_filled")
        await log_page_state(self.page, "about-you 填完表单后")

        # 提交 — 按钮文本是 "完成帐户创建"
        clicked = False
        for btn_text in [
            re.compile(r"(完成|Complete|create account|agree)", re.IGNORECASE),
            re.compile(r"(Continue|继续|Save|保存|Next|下一步)", re.IGNORECASE),
        ]:
            try:
                btn = self.page.get_by_role("button", name=btn_text)
                if await btn.is_visible(timeout=2000):
                    text = (await btn.text_content() or "").strip()
                    logger.info(f"点击提交按钮: {text[:40]}")
                    await btn.click()
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            buttons = self.page.locator("button")
            cnt = await buttons.count()
            logger.warning(f"未找到标准提交按钮，页面有 {cnt} 个按钮")
            if cnt > 0:
                await buttons.last.click()
                clicked = True
                logger.info("已点击最后一个按钮作为保底")

        logger.info(f"about-you 提交 {'成功' if clicked else '未点击'}")
        await self._detect_user_already_exists("about_you_submit")

        # 主动检测是否有跳转，卡住则立即截断
        age_retry_done = False
        for i in range(15):  # 最多等 30 秒
            await asyncio.sleep(2)
            current = self.page.url
            if "about-you" not in current and "about_you" not in current:
                await self._detect_user_already_exists("about_you_after_redirect")
                logger.info(f"about-you 已跳转: {current[:100]}")
                return  # 已离开 about-you
            await self._detect_user_already_exists("about_you_wait")
            if has_age_only and not age_retry_done:
                try:
                    body_text = await self.page.evaluate("() => document.body?.innerText || ''")
                except Exception:
                    body_text = ""
                if "请输入有效年龄" in body_text or "valid age" in body_text.lower():
                    age_retry_done = True
                    logger.warning("检测到年龄校验失败，重新填写年龄并再次提交")
                    if await self._fill_age_only_input(age_value):
                        try:
                            retry_btn = self.page.get_by_role(
                                "button",
                                name=re.compile(r"(完成|Complete|create account|agree|Continue|继续|Save|保存|Next|下一步)", re.IGNORECASE),
                            ).first
                            if await retry_btn.is_visible(timeout=2000):
                                await retry_btn.click()
                                logger.info("年龄重填后已再次提交")
                                continue
                        except Exception as e:
                            logger.warning(f"年龄重填后再次提交失败: {e}")
            if i == 5:
                await log_page_state(self.page, f"about-you 等待跳转中 ({i*2}s)")
            if i == 14:
                # 卡住了，立刻截断留给 MCP 处理
                logger.error(f"about-you 提交后 30s 无跳转，当前仍在: {current[:100]}")
                await log_page_state(self.page, "about-you 卡住")
                if DEBUG_MODE:
                    await debug_pause(self.page, "about-you 提交后卡住，无跳转")
                return

    async def wait_for_login_complete(self) -> bool:
        logger.info("等待 OAuth 回调返回 chatgpt.com...")
        # 先检查是否已经回到了 chatgpt.com
        url = self.page.url
        if CHATGPT_URL in url and "callback" not in url and "auth" not in url:
            logger.info(f"已在 chatgpt.com: {url[:100]}")
            await self._random_delay(2000, 3000)
            return await self._check_session()

        # 等待 60s 跳转，超时就截断
        try:
            await self.page.wait_for_url(
                re.compile(r"https://chatgpt\.com(?!.*callback).*"),
                timeout=60000
            )
            logger.info(f"已回到: {self.page.url[:100]}")
        except Exception:
            url = self.page.url
            logger.error(f"60s 内未跳转回 chatgpt.com，当前: {url[:120]}")
            await log_page_state(self.page, "wait_for_login 超时")
            if DEBUG_MODE:
                await debug_pause(self.page, "wait_for_login 超时 60s")
            return False

        await self._random_delay(2000, 3000)
        return await self._check_session()

    async def _check_session(self) -> bool:
        """检查 /api/auth/session 确认登录状态"""
        try:
            session_data = await self.page.evaluate("""
                async () => {
                    const resp = await fetch('/api/auth/session', { credentials: 'include' });
                    const text = await resp.text();
                    try { return JSON.parse(text); } catch(e) { return null; }
                }
            """)
            if session_data and session_data.get("accessToken"):
                user_info = session_data.get("user", {})
                plan = session_data.get("account", {}).get("planType", "?")
                self.current_plan_type = str(plan or "")
                logger.info(f"✅ 登录成功: {user_info.get('email', '?')} plan={plan}")
                return True
            else:
                logger.warning(f"Session 无 accessToken: {str(session_data)[:300]}")
        except Exception as e:
            logger.warning(f"Session 检查失败: {e}")
        return False

        try:
            session_data = await self.page.evaluate("""
                async () => {
                    const resp = await fetch('/api/auth/session', { credentials: 'include' });
                    const text = await resp.text();
                    try { return JSON.parse(text); } catch(e) { return null; }
                }
            """)
            if session_data and session_data.get("accessToken"):
                user_info = session_data.get("user", {})
                plan = session_data.get("account", {}).get("planType", "?")
                logger.info(f"✅ 登录成功: {user_info.get('email', '?')} plan={plan}")
                return True
            else:
                logger.warning(f"Session 内容: {str(session_data)[:300]}")
        except Exception as e:
            logger.warning(f"Session 检查失败: {e}")

        # 保底：URL 确认
        if CHATGPT_URL in self.page.url:
            logger.info("保底确认：已在 chatgpt.com")
            return True
        return False

    async def get_access_token(self) -> str:
        token = await self.page.evaluate("""
            async () => {
                const resp = await fetch('/api/auth/session', { credentials: 'include' });
                const data = await resp.json();
                return data?.accessToken || null;
            }
        """)
        if not token:
            raise SignupFlowError("获取 accessToken 失败")
        return token

    def _generate_codex_oauth_params(self) -> dict:
        code_verifier = secrets.token_urlsafe(48)
        challenge = _b64url_encode(hashlib.sha256(code_verifier.encode("ascii")).digest())
        state = secrets.token_hex(16)
        return {"code_verifier": code_verifier, "code_challenge": challenge, "state": state}

    def _build_codex_oauth_url(self, oauth_params: dict) -> str:
        query = urlencode({
            "client_id": CODEX_OAUTH_CLIENT_ID,
            "code_challenge": oauth_params["code_challenge"],
            "code_challenge_method": "S256",
            "codex_cli_simplified_flow": "true",
            "id_token_add_organizations": "true",
            "prompt": "login",
            "redirect_uri": CODEX_OAUTH_REDIRECT_URI,
            "response_type": "code",
            "scope": "openid email profile offline_access",
            "state": oauth_params["state"],
        })
        return f"https://auth.openai.com/oauth/authorize?{query}"

    async def _submit_codex_oauth_email(self, email: str) -> bool:
        """在 Codex OAuth 登录页填写邮箱并提交。"""
        email_input = None
        for locator in [
            self.page.locator('input[type="email"][name="email"]').first,
            self.page.locator('input[name="email"]').first,
            self.page.locator('input[autocomplete="email"]').first,
            self.page.get_by_role("textbox", name=re.compile(r"(电子邮件地址|电子邮件|邮箱|email)", re.IGNORECASE)).first,
            self.page.get_by_placeholder(re.compile(r"(电子邮件地址|电子邮件|邮箱|email)", re.IGNORECASE)).first,
        ]:
            try:
                if await locator.is_visible(timeout=700):
                    email_input = locator
                    break
            except Exception:
                continue

        if not email_input:
            return False

        actual = await self._fill_controlled_input(email_input, email)
        if actual != email:
            raise SignupFlowError(f"Codex OAuth 邮箱填写失败: expected={email}, actual={actual}")

        for submit in [
            self.page.locator('form[aria-label="选择登录选项"] button[type="submit"]').first,
            self.page.locator('button[type="submit"][name="intent"]').first,
            self.page.get_by_role("button", name=re.compile(r"^(继续|Continue)$", re.IGNORECASE)).first,
        ]:
            try:
                if await submit.is_visible(timeout=1000):
                    await submit.click(timeout=5000)
                    logger.info(f"已提交 Codex OAuth 登录邮箱: {email}")
                    await self._random_delay(1500, 2500)
                    return True
            except Exception:
                continue

        raise SignupFlowError("Codex OAuth 邮箱页找不到继续按钮")

    async def _click_codex_oauth_one_time_code(self) -> bool:
        """在 Codex OAuth 密码页切换到一次性验证码登录。"""
        password_input_visible = False
        for password_locator in [
            self.page.locator('input[type="password"]').first,
            self.page.get_by_role("textbox", name=re.compile(r"^(密码|password)$", re.IGNORECASE)).first,
        ]:
            try:
                if await password_locator.is_visible(timeout=500):
                    password_input_visible = True
                    break
            except Exception:
                continue
        if not password_input_visible:
            return False

        for otp_button in [
            self.page.get_by_role("button", name=re.compile(r"(一次性验证码|one.?time code|verification code)", re.IGNORECASE)).first,
            self.page.get_by_text(re.compile(r"(使用一次性验证码登录|one.?time code)", re.IGNORECASE)).first,
        ]:
            try:
                if await otp_button.is_visible(timeout=1000):
                    await otp_button.click(timeout=5000)
                    logger.info("已切换 Codex OAuth 为一次性验证码登录")
                    await self._random_delay(1500, 2500)
                    return True
            except Exception:
                continue

        return False

    async def _submit_codex_oauth_code(self) -> bool:
        """在 Codex OAuth 验证码页读取文件验证码并提交。"""
        code_input = None
        for locator in [
            self.page.locator('input[autocomplete="one-time-code"]').first,
            self.page.locator('input[name*="code" i]').first,
            self.page.get_by_role("textbox", name=re.compile(r"(验证码|verification code|code)", re.IGNORECASE)).first,
            self.page.get_by_placeholder(re.compile(r"(验证码|verification code|code)", re.IGNORECASE)).first,
        ]:
            try:
                if await locator.is_visible(timeout=500):
                    code_input = locator
                    break
            except Exception:
                continue

        if not code_input:
            return False
        if not self.oauth_code_file:
            logger.info("Codex OAuth 验证码页需要人工输入，未配置验证码文件，等待页面后续跳转")
            return True

        code = await wait_for_code_file(self.oauth_code_file, "Codex OAuth 验证码", timeout=900)
        actual = await self._fill_controlled_input(code_input, code)
        if actual != code:
            raise SignupFlowError(f"Codex OAuth 验证码输入失败: expected={code}, actual={actual}")

        for submit in [
            self.page.get_by_role("button", name=re.compile(r"^(继续|Continue)$", re.IGNORECASE)).first,
            self.page.locator('button[type="submit"]').first,
        ]:
            try:
                if await submit.is_visible(timeout=1000):
                    await submit.click(timeout=5000)
                    logger.info("已提交 Codex OAuth 验证码")
                    await self._random_delay(1500, 2500)
                    return True
            except Exception:
                continue

        raise SignupFlowError("Codex OAuth 验证码页找不到继续按钮")

    async def _start_codex_oauth_callback_server(self, expected_state: str):
        """监听 Codex OAuth localhost 回调，避免浏览器跳到 chrome-error 丢失 code。"""
        loop = asyncio.get_running_loop()
        callback_future = loop.create_future()

        async def handle_callback(reader, writer) -> None:
            try:
                raw_request = await reader.read(8192)
                request_text = raw_request.decode("utf-8", errors="ignore")
                request_line = request_text.splitlines()[0] if request_text else ""
                parts = request_line.split()
                target = parts[1] if len(parts) >= 2 else ""
                parsed = urlparse(target)
                params = parse_qs(parsed.query)

                error = first_non_empty(*(params.get("error") or []))
                state = first_non_empty(*(params.get("state") or []))
                code = first_non_empty(*(params.get("code") or []))
                if error:
                    desc = first_non_empty(*(params.get("error_description") or []))
                    if not callback_future.done():
                        callback_future.set_exception(SignupFlowError(f"Codex OAuth 授权失败: {error} {desc}"))
                    body = "Codex OAuth authorization failed. You can close this window."
                elif state and state != expected_state:
                    logger.warning(
                        "Codex OAuth 忽略旧回调 state: received=%s expected=%s",
                        state[:8],
                        expected_state[:8],
                    )
                    body = "Codex OAuth stale callback ignored. You can close this window."
                elif code:
                    if not callback_future.done():
                        callback_future.set_result(code)
                    body = "Codex OAuth authorization complete. You can close this window."
                else:
                    body = "Codex OAuth callback received without code."

                response = (
                    "HTTP/1.1 200 OK\r\n"
                    "Content-Type: text/html; charset=utf-8\r\n"
                    "Connection: close\r\n\r\n"
                    f"<html><body><h3>{body}</h3></body></html>"
                )
                writer.write(response.encode("utf-8"))
                await writer.drain()
            except Exception as e:
                if not callback_future.done():
                    callback_future.set_exception(e)
            finally:
                try:
                    writer.close()
                    await writer.wait_closed()
                except Exception:
                    pass

        server = await asyncio.start_server(handle_callback, host=None, port=1455)
        logger.info("Codex OAuth 本地回调监听已启动: http://localhost:1455/auth/callback")
        return server, callback_future

    async def _wait_for_codex_oauth_callback(self, expected_state: str, email: str, callback_future=None, timeout: int = 120) -> str:
        deadline = time.time() + timeout
        authorize_pattern = re.compile(r"(Authorize|Allow|Sign in|Log in|允许|授权|Continue|继续|确认|同意|登录)", re.IGNORECASE)
        last_url = ""
        login_email_submitted = False
        oauth_code_submitted = False
        last_login_wait_log = 0.0
        last_blocking_input_log = 0.0

        while time.time() < deadline:
            if callback_future and callback_future.done():
                return callback_future.result()

            current_url = self.page.url or ""
            if current_url != last_url:
                logger.info(f"Codex OAuth 跳转检测: {current_url[:160]}")
                last_url = current_url
                if "auth.openai.com/log-in" not in current_url:
                    login_email_submitted = False

            if "localhost:1455/auth/callback" in current_url:
                parsed = urlparse(current_url)
                params = parse_qs(parsed.query)
                error = first_non_empty(*(params.get("error") or []))
                if error:
                    desc = first_non_empty(*(params.get("error_description") or []))
                    raise SignupFlowError(f"Codex OAuth 授权失败: {error} {desc}")
                state = first_non_empty(*(params.get("state") or []))
                if state and state != expected_state:
                    logger.warning(
                        "Codex OAuth URL 检测忽略旧回调 state: received=%s expected=%s",
                        state[:8],
                        expected_state[:8],
                    )
                    await asyncio.sleep(0.5)
                    continue
                code = first_non_empty(*(params.get("code") or []))
                if not code:
                    raise SignupFlowError("Codex OAuth callback 缺少 code")
                return code

            email_input_visible = False
            for email_locator in [
                self.page.locator('input[type="email"][name="email"]').first,
                self.page.locator('input[name="email"]').first,
                self.page.locator('input[autocomplete="email"]').first,
            ]:
                try:
                    if await email_locator.is_visible(timeout=300):
                        email_input_visible = True
                        break
                except Exception:
                    continue

            if email_input_visible:
                if not login_email_submitted:
                    login_email_submitted = await self._submit_codex_oauth_email(email)
                    deadline = time.time() + timeout
                    last_login_wait_log = time.time()
                elif time.time() - last_login_wait_log >= 10:
                    logger.info("Codex OAuth 仍在邮箱登录页，等待跳转或下一步验证，不重复点击继续")
                    last_login_wait_log = time.time()
                await asyncio.sleep(1)
                continue

            if await self._click_codex_oauth_one_time_code():
                deadline = time.time() + timeout
                await asyncio.sleep(1)
                continue

            code_input_visible = False
            for code_locator in [
                self.page.locator('input[autocomplete="one-time-code"]').first,
                self.page.locator('input[name*="code" i]').first,
            ]:
                try:
                    if await code_locator.is_visible(timeout=300):
                        code_input_visible = True
                        break
                except Exception:
                    continue

            if code_input_visible:
                if not oauth_code_submitted:
                    oauth_code_submitted = await self._submit_codex_oauth_code()
                    deadline = time.time() + timeout
                elif time.time() - last_blocking_input_log >= 10:
                    logger.info("Codex OAuth 仍在验证码页，等待跳转，不重复提交")
                    last_blocking_input_log = time.time()
                await asyncio.sleep(1)
                continue

            blocking_input_visible = False
            for blocking_locator in [
                self.page.locator('input[type="password"]').first,
                self.page.locator('input[autocomplete="one-time-code"]').first,
                self.page.locator('input[name*="code" i]').first,
                self.page.get_by_role("textbox", name=re.compile(r"(验证码|verification code|code|password|密码)", re.IGNORECASE)).first,
            ]:
                try:
                    if await blocking_locator.is_visible(timeout=300):
                        blocking_input_visible = True
                        break
                except Exception:
                    continue

            if blocking_input_visible:
                if time.time() - last_blocking_input_log >= 10:
                    logger.info("Codex OAuth 检测到验证码/密码输入页，等待人工或页面后续跳转，不自动点击继续")
                    last_blocking_input_log = time.time()
                await asyncio.sleep(1)
                continue

            for locator in [
                self.page.get_by_role("button", name=authorize_pattern).first,
                self.page.get_by_text(authorize_pattern).first,
            ]:
                try:
                    if await locator.is_visible(timeout=700):
                        await locator.click(timeout=3000)
                        logger.info("已点击 Codex OAuth 授权/继续按钮")
                        await self._random_delay(800, 1500)
                        break
                except Exception:
                    continue

            await asyncio.sleep(1)

        await log_page_state(self.page, "Codex OAuth callback 超时")
        raise SignupFlowError(f"{timeout}s 内未获取 Codex OAuth callback")

    async def _exchange_codex_oauth_token(self, code: str, code_verifier: str, email: str) -> dict:
        payload = {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": CODEX_OAUTH_REDIRECT_URI,
            "client_id": CODEX_OAUTH_CLIENT_ID,
            "code_verifier": code_verifier,
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        last_error = None

        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
            for attempt in range(1, 6):
                try:
                    if self.proxy_url:
                        logger.info(f"Codex OAuth token exchange 使用代理: {self.proxy_url}")
                        request_ctx = session.post(
                            "https://auth.openai.com/oauth/token",
                            data=payload,
                            headers=headers,
                            proxy=self.proxy_url,
                        )
                    else:
                        request_ctx = session.post(
                            "https://auth.openai.com/oauth/token",
                            data=payload,
                            headers=headers,
                        )

                    async with request_ctx as resp:
                        body_text = await resp.text()
                        try:
                            body = json.loads(body_text)
                        except json.JSONDecodeError:
                            body = {"raw": body_text[:500]}
                        if resp.status == 200:
                            return self._build_codex_token_data(body, email)
                        retryable = resp.status == 429 or 500 <= resp.status <= 599
                        last_error = SignupFlowError(f"Codex OAuth token HTTP {resp.status}: {body}")
                        if not retryable:
                            raise last_error
                except Exception as e:
                    last_error = e
                    if attempt == 5:
                        raise
                    if isinstance(e, SignupFlowError) and not str(e).startswith("Codex OAuth token HTTP 429") and "HTTP 5" not in str(e):
                        raise
                wait_ms = attempt * 3
                logger.warning(f"Codex OAuth 换 token 第 {attempt} 次失败，{wait_ms}s 后重试: {last_error}")
                await asyncio.sleep(wait_ms)

        raise SignupFlowError(f"Codex OAuth 换 token 失败: {last_error}")

    def _build_codex_token_data(self, tokens: dict, email: str) -> dict:
        access_token = tokens.get("access_token") or ""
        refresh_token = tokens.get("refresh_token") or ""
        if not access_token or not refresh_token:
            raise SignupFlowError("Codex OAuth token 响应缺少 access_token/refresh_token")

        id_token = tokens.get("id_token") or ""
        access_claims = decode_jwt_payload(access_token)
        id_claims = decode_jwt_payload(id_token)
        access_auth = access_claims.get("https://api.openai.com/auth") or {}
        id_auth = id_claims.get("https://api.openai.com/auth") or {}
        access_profile = access_claims.get("https://api.openai.com/profile") or {}
        expires_in = int(tokens.get("expires_in") or 0)
        access_exp = access_claims.get("exp") if isinstance(access_claims.get("exp"), int) else 0
        expires_at_ts = time.time() + expires_in if expires_in > 0 else access_exp
        expires_at = datetime.fromtimestamp(expires_at_ts, timezone.utc).isoformat().replace("+00:00", "Z")

        token_data = {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": expires_at,
            "email": first_non_empty(email, access_profile.get("email"), id_claims.get("email")),
            "id_token": id_token,
            "client_id": first_non_empty(access_claims.get("client_id"), first_audience_value(id_claims.get("aud"))),
            "chatgpt_account_id": first_non_empty(access_auth.get("chatgpt_account_id"), id_auth.get("chatgpt_account_id")),
            "chatgpt_user_id": first_non_empty(
                access_auth.get("chatgpt_user_id"),
                id_auth.get("chatgpt_user_id"),
                access_auth.get("user_id"),
                id_auth.get("user_id"),
            ),
            "organization_id": extract_organization_id(id_auth.get("organizations")),
            "plan_type": first_non_empty(access_auth.get("chatgpt_plan_type"), id_auth.get("chatgpt_plan_type")),
            "subscription_expires_at": first_non_empty(
                access_auth.get("chatgpt_subscription_active_until"),
                id_auth.get("chatgpt_subscription_active_until"),
            ),
        }
        return {k: v for k, v in token_data.items() if v}

    async def export_codex_sub2api_token(self, email: str) -> dict:
        """使用当前登录态完成 Codex OAuth，并返回可导出为 Sub2API 的 token 数据。"""
        logger.info("开始 Codex OAuth 授权换取 Sub2API token")
        oauth_params = self._generate_codex_oauth_params()
        auth_url = self._build_codex_oauth_url(oauth_params)
        callback_server = None
        callback_future = None
        try:
            try:
                callback_server, callback_future = await self._start_codex_oauth_callback_server(oauth_params["state"])
            except OSError as e:
                logger.warning(f"Codex OAuth 本地回调监听启动失败，将退回页面 URL 检测: {e}")

            await self.page.goto(auth_url, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            logger.warning(f"打开 Codex OAuth 授权页异常，继续等待 callback: {e}")
        try:
            code = await self._wait_for_codex_oauth_callback(oauth_params["state"], email, callback_future=callback_future)
            token_data = await self._exchange_codex_oauth_token(code, oauth_params["code_verifier"], email)
            logger.info("✅ Codex OAuth token 获取成功")
            return token_data
        finally:
            if callback_server:
                callback_server.close()
                await callback_server.wait_closed()

    async def _click_confirm_button(self) -> bool:
        for sel in [
            self.page.get_by_role("button", name=re.compile(r"(Confirm|确认|Submit|Pay|Verify|验证)", re.IGNORECASE)),
            self.page.get_by_role("button", name=re.compile(r"(Save|保存|Done|完成|Continue|继续)", re.IGNORECASE)),
            self.page.get_by_text(re.compile(r"(Confirm|确认|Save|保存|Done|完成)", re.IGNORECASE)),
        ]:
            try:
                btn = sel.last
                if await btn.is_visible(timeout=2000):
                    await btn.click()
                    return True
            except Exception:
                continue
        return False

    async def _dump_password_settings_page(self, reason: str) -> None:
        """输出设置页结构，便于补充安全/密码入口选择器。"""
        logger.warning(f"设置密码页诊断: {reason}")
        logger.warning(f"当前 URL: {self.page.url[:200]}")
        try:
            page_info = await self.page.evaluate(
                """
                () => {
                    const visible = (el) => {
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    };
                    const items = Array.from(document.querySelectorAll('button, [role="button"], a, input'))
                        .filter(visible)
                        .slice(0, 80)
                        .map((el) => ({
                            tag: el.tagName,
                            role: el.getAttribute('role') || '',
                            type: el.getAttribute('type') || '',
                            text: (el.innerText || el.value || el.getAttribute('aria-label') || el.getAttribute('placeholder') || '').trim().slice(0, 100),
                            testid: el.getAttribute('data-testid') || '',
                        }));
                    return {
                        title: document.title || '',
                        body: (document.body?.innerText || '').trim().slice(0, 1200),
                        items,
                    };
                }
                """
            )
            logger.warning(f"设置页标题: {page_info.get('title', '')!r}")
            logger.warning(f"设置页可见元素: {page_info.get('items', [])}")
            logger.warning(f"设置页正文: {page_info.get('body', '')!r}")
        except Exception as e:
            logger.warning(f"设置页诊断失败: {e}")

        await screenshot(self.page, f"password_settings_debug_{reason[:30]}")

    async def _open_password_settings_entry(self) -> bool:
        """进入设置/安全/密码中的设置密码入口。"""
        try:
            await self.page.goto(f"{self.CHATGPT_URL}/#settings/Security")
            await self._random_delay(1500, 2500)
        except Exception as e:
            logger.warning(f"打开安全设置页失败: {e}")

        # 兼容 hash 未直接切到安全页的情况，先尝试点击“安全”标签。
        for security_tab in [
            self.page.get_by_role("tab", name=re.compile(r"^(Security|安全)$", re.IGNORECASE)).first,
            self.page.get_by_role("button", name=re.compile(r"^(Security|安全)$", re.IGNORECASE)).first,
            self.page.get_by_text(re.compile(r"^(Security|安全)$", re.IGNORECASE)).first,
        ]:
            try:
                if await security_tab.is_visible(timeout=1500):
                    await security_tab.click(timeout=3000)
                    logger.info("已进入设置/安全")
                    await self._random_delay(800, 1500)
                    break
            except Exception:
                continue

        # 安全页里的“密码”行通常需要先点开，再出现设置/新增/更改密码表单。
        password_entry_patterns = [
            r"(Set password|Add password|Create password|Change password|Manage password|Password|设置密码|新增密码|创建密码|更改密码|管理密码|密码)",
        ]
        for pattern in password_entry_patterns:
            for entry in [
                self.page.get_by_test_id("password-setting").first,
                self.page.get_by_role("button", name=re.compile(pattern, re.IGNORECASE)).first,
                self.page.get_by_text(re.compile(pattern, re.IGNORECASE)).first,
            ]:
                try:
                    if await entry.is_visible(timeout=2500):
                        await entry.click(timeout=5000)
                        logger.info("已点击设置/安全/密码入口")
                        await self._random_delay(1000, 1800)
                        try:
                            await self.page.wait_for_url(re.compile(r"auth\.openai\.com/.*/new-password"), timeout=10000)
                        except Exception:
                            pass
                        try:
                            await self.page.locator('input[name="new-password"], input[type="password"]').first.wait_for(
                                state="visible",
                                timeout=10000,
                            )
                        except Exception:
                            pass
                        return True
                except Exception:
                    continue

        return False

    async def _fill_visible_password_form(self, password: str) -> bool:
        """填写当前可见的密码表单，并点击确认。"""
        new_password = self.page.locator('input[name="new-password"]').first
        confirm_password = self.page.locator('input[name="confirm-password"]').first

        try:
            if await new_password.is_visible(timeout=1000) and await confirm_password.is_visible(timeout=1000):
                await new_password.fill(password)
                await confirm_password.fill(password)
                logger.info("已填写 auth.openai.com 新密码表单")
                await screenshot(self.page, "step_15_password_set")
                if not await self._click_confirm_button():
                    logger.warning("没找到设置密码确认按钮")
                    await self._dump_password_settings_page("confirm_button_missing")
                    return False
                await self._random_delay(1500, 2500)
                logger.info("✅ 密码已设置")
                return True
        except Exception as e:
            logger.warning(f"填写 auth.openai.com 新密码表单失败: {e}")

        pw_inputs = self.page.locator('input[type="password"]')
        count = await pw_inputs.count()
        if count <= 0:
            return False

        logger.info(f"检测到 {count} 个密码输入框")
        for i in range(count):
            try:
                inp = pw_inputs.nth(i)
                if await inp.is_visible(timeout=1000):
                    await inp.fill(password)
            except Exception as e:
                logger.warning(f"填写第 {i + 1} 个密码输入框失败: {e}")

        await screenshot(self.page, "step_15_password_set")
        if not await self._click_confirm_button():
            logger.warning("没找到设置密码确认按钮")
            await self._dump_password_settings_page("confirm_button_missing")
            return False

        await self._random_delay(1500, 2500)
        logger.info("✅ 密码已设置")
        return True

    async def add_password_login(self, password: str) -> bool:
        logger.info("设置账号密码...")
        try:
            await self._open_password_settings_entry()
            await screenshot(self.page, "step_14_settings")

            if await self._fill_visible_password_form(password):
                return True

            # 查找设置密码的入口
            set_pw_btn = None
            for sel in [
                self.page.get_by_role(
                    "button",
                    name=re.compile(r"(set password|设置密码|add password|新增密码|create password|创建密码|change password|更改密码)", re.IGNORECASE),
                ),
                self.page.get_by_text(
                    re.compile(r"(set password|设置密码|add password|新增密码|create password|创建密码|change password|更改密码)", re.IGNORECASE)
                ),
            ]:
                try:
                    if await sel.first.is_visible(timeout=3000):
                        set_pw_btn = sel.first
                        break
                except Exception:
                    continue

            if set_pw_btn:
                await set_pw_btn.click()
                await self._random_delay()
                if await self._fill_visible_password_form(password):
                    return True
                logger.warning("已点击设置密码入口，但未出现密码输入框")
                await self._dump_password_settings_page("password_inputs_missing")
                return False

            logger.warning("没找到设置密码的入口")
            await self._dump_password_settings_page("password_entry_missing")
            return False

        except Exception as e:
            logger.error(f"设置密码失败: {e}")
            return False

    async def cancel_subscription(self) -> bool:
        """通过 ChatGPT 后端 API 直接取消续订（MCP 抓包验证）：
            POST https://chatgpt.com/backend-api/subscriptions/cancel
            Headers: Authorization: Bearer <accessToken>
            Body: {}
            Response: 200 {} （成功后 GET subscriptions 会显示 will_renew=false）

        在浏览器上下文 fetch 可自动带上 cookies + access token。比 UI 流程稳定得多。
        """
        logger.info("取消订阅（API 调用）...")
        try:
            # 必须先在 chatgpt.com 同源页面下，才能正确读到 NextAuth session 与 cookies
            if "chatgpt.com" not in (self.page.url or ""):
                await self.page.goto(self.CHATGPT_URL)
                await self._random_delay(2000, 3000)

            await screenshot(self.page, "step_16_before_cancel")

            result = await self.page.evaluate(
                """
                async () => {
                    // 1. 取 accessToken
                    let token = null;
                    try {
                        const sess = await fetch('/api/auth/session', { credentials: 'include' }).then(r => r.json());
                        token = sess && sess.accessToken;
                    } catch (e) {}
                    if (!token) return { ok: false, stage: 'session', error: 'no accessToken' };

                    // 2. 解析 JWT 拿到 chatgpt_account_id（cancel API 必填字段）
                    let acctId = null;
                    try {
                        const parts = token.split('.');
                        if (parts.length === 3) {
                            const payload = JSON.parse(atob(parts[1].replace(/-/g,'+').replace(/_/g,'/')));
                            const cgpt = payload && payload['https://api.openai.com/auth'];
                            if (cgpt && cgpt.chatgpt_account_id) acctId = cgpt.chatgpt_account_id;
                        }
                    } catch (e) {}
                    // JWT 解析失败时退回到 /backend-api/me
                    if (!acctId) {
                        try {
                            const me = await fetch('/backend-api/me', {
                                credentials: 'include',
                                headers: { Authorization: 'Bearer ' + token },
                            }).then(r => r.json());
                            acctId = me && me.chatgpt_account_id;
                            if (!acctId && me && me.orgs && me.orgs.data) {
                                // org id 仅作最后兜底
                                acctId = me.orgs.data[0] && me.orgs.data[0].id;
                            }
                        } catch (e) {}
                    }
                    if (!acctId) return { ok: false, stage: 'acct_id', error: 'no chatgpt_account_id' };

                    // 3. POST 取消（必须带 account_id）
                    let cancelStatus = 0, cancelBody = '';
                    try {
                        const r = await fetch('/backend-api/subscriptions/cancel', {
                            method: 'POST',
                            credentials: 'include',
                            headers: {
                                'Content-Type': 'application/json',
                                'Authorization': 'Bearer ' + token,
                            },
                            body: JSON.stringify({ account_id: acctId }),
                        });
                        cancelStatus = r.status;
                        cancelBody = await r.text();
                    } catch (e) {
                        return { ok: false, stage: 'cancel', error: String(e), acctId };
                    }
                    if (cancelStatus !== 200) {
                        return { ok: false, stage: 'cancel', status: cancelStatus, body: cancelBody, acctId };
                    }

                    // 4. 验证 will_renew=false
                    let willRenew = null, plan = null;
                    try {
                        const sub = await fetch('/backend-api/subscriptions?account_id=' + acctId, {
                            credentials: 'include',
                            headers: { Authorization: 'Bearer ' + token },
                        }).then(r => r.json());
                        willRenew = sub && sub.will_renew;
                        plan = sub && sub.plan_type;
                    } catch (e) {}

                    return { ok: true, willRenew, plan, acctId };
                }
                """
            )

            logger.info(f"取消订阅结果: {result}")
            if not result.get("ok"):
                logger.error(f"取消订阅失败: {result}")
                return False

            # 校验 will_renew
            wr = result.get("willRenew")
            if wr is False:
                logger.info(f"✅ 订阅已取消续订（plan={result.get('plan')}）")
                await screenshot(self.page, "step_17_cancelled")
                return True
            if wr is None:
                # 校验信息缺失但 cancel 接口返回 200，认为成功
                logger.info("✅ 取消请求已被接受（will_renew 未能校验，但 200 OK）")
                return True
            logger.warning(f"取消接口 200 但 will_renew 仍为 {wr}")
            return False

        except Exception as e:
            logger.error(f"取消订阅失败: {e}")
            return False


# ===================== 4. 账号管理器 =====================
class AccountManager:
    FILE = ACCOUNTS_FILE
    TXT_FILE = ACCOUNTS_TXT_FILE

    @classmethod
    def load_accounts(cls) -> list:
        if not os.path.exists(cls.FILE):
            return []
        try:
            with open(cls.FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []

    @staticmethod
    def _disp_width(s: str) -> int:
        """计算字符串在等宽终端下的显示宽度。CJK 全角字符算 2 列，其他算 1 列。"""
        w = 0
        for ch in s:
            cp = ord(ch)
            if (
                0x1100 <= cp <= 0x115F  # Hangul Jamo
                or 0x2E80 <= cp <= 0x9FFF  # CJK
                or 0xA000 <= cp <= 0xA4CF
                or 0xAC00 <= cp <= 0xD7A3  # Hangul Syllables
                or 0xF900 <= cp <= 0xFAFF
                or 0xFE30 <= cp <= 0xFE4F
                or 0xFF00 <= cp <= 0xFF60
                or 0xFFE0 <= cp <= 0xFFE6
            ):
                w += 2
            else:
                w += 1
        return w

    @classmethod
    def _pad(cls, s: str, width: int) -> str:
        """左对齐，按显示宽度补空格。"""
        diff = width - cls._disp_width(s)
        return s + (" " * diff if diff > 0 else "")

    @classmethod
    def _fmt_time(cls, iso: str) -> str:
        """ISO UTC 时间 → 本地时间 'YYYY-MM-DD HH:MM:SS'。"""
        try:
            dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
            return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return iso[:19] if iso else ""

    @classmethod
    def write_txt_dump(cls, accounts: list) -> None:
        """根据全部账号重写 accounts.txt。

        每个账号 4 行：邮箱:、密码:、时间:、登录地址:；账号之间用空行分隔。
        4 个标签按显示宽度补齐，使后面的冒号 + 值竖直对齐。
        """
        labels = ["邮箱", "密码", "时间", "登录地址"]
        label_w = max(cls._disp_width(lb) for lb in labels)

        def line(label: str, value: str) -> str:
            # "<label>:<空格补齐>  <value>"  —— 冒号紧贴 label，再补空格后接 value
            return cls._pad(label + ":", label_w + 1) + "  " + value

        chunks = [
            f"# ChatGPT 注册账号清单（共 {len(accounts)} 个，最近一次更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）",
            "",
        ]
        for a in accounts:
            chunks.append(line("邮箱", a.get("email", "")))
            chunks.append(line("密码", a.get("password", "")))
            chunks.append(line("时间", cls._fmt_time(a.get("created_at", ""))))
            chunks.append(line("登录地址", a.get("auto_login_url", "")))
            chunks.append("")  # 账号之间空行

        content = "\n".join(chunks)
        if not content.endswith("\n"):
            content += "\n"
        with open(cls.TXT_FILE, "w", encoding="utf-8") as f:
            f.write(content)

    @classmethod
    def save_account(cls, account: dict) -> None:
        accounts = cls.load_accounts()
        accounts.append(account)
        with open(cls.FILE, "w", encoding="utf-8") as f:
            json.dump(accounts, f, ensure_ascii=False, indent=2)
        try:
            cls.write_txt_dump(accounts)
            logger.info(f"✅ 账号已保存到 {cls.FILE} 与 {cls.TXT_FILE} ({len(accounts)} 个账号)")
        except Exception as e:
            logger.warning(f"写入 TXT 失败（JSON 已保存）: {e}")
            logger.info(f"✅ 账号已保存到 {cls.FILE} ({len(accounts)} 个账号)")

    @staticmethod
    def generate_auto_login_url(address: str, jwt: str) -> str:
        if not jwt:
            return ""
        return f"{TEMP_EMAIL_LOGIN_BASE}?jwt={jwt}"

    @classmethod
    def create_account(
        cls,
        email: str,
        chatgpt_password: str,
        temp_email_password: str,
        jwt: str,
    ) -> dict:
        account = {
            "email": email,
            "password": chatgpt_password,
            "temp_email_password": temp_email_password,
            "jwt": jwt,
            "auto_login_url": cls.generate_auto_login_url(email, jwt),
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
        cls.save_account(account)
        return account


class Sub2APIExporter:
    FILE = SUB2API_ACCOUNTS_FILE

    @classmethod
    def load_payload(cls) -> dict:
        if not os.path.exists(cls.FILE):
            return {"exported_at": utc_now_iso(), "proxies": [], "accounts": []}
        try:
            with open(cls.FILE, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (json.JSONDecodeError, IOError):
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
        proxies = payload.get("proxies") if isinstance(payload.get("proxies"), list) else []
        accounts = payload.get("accounts") if isinstance(payload.get("accounts"), list) else []
        return {"exported_at": payload.get("exported_at") or utc_now_iso(), "proxies": proxies, "accounts": accounts}

    @staticmethod
    def _derive_name(email: str) -> str:
        return (email or "").strip() or f"openai-{int(time.time())}"

    @staticmethod
    def _unique_name(base_name: str, accounts: list) -> str:
        used = {str(item.get("name", "")) for item in accounts if isinstance(item, dict)}
        if base_name not in used:
            return base_name
        idx = 2
        while f"{base_name}-{idx}" in used:
            idx += 1
        return f"{base_name}-{idx}"

    @staticmethod
    def _build_credentials(token_data: dict) -> dict:
        required = ["access_token", "refresh_token", "expires_at", "email"]
        missing = [key for key in required if not token_data.get(key)]
        if missing:
            raise SignupFlowError(f"Sub2API 导出缺少必要字段: {', '.join(missing)}")
        allowed = [
            "access_token",
            "refresh_token",
            "expires_at",
            "email",
            "id_token",
            "client_id",
            "chatgpt_account_id",
            "chatgpt_user_id",
            "organization_id",
            "plan_type",
            "subscription_expires_at",
        ]
        return {key: token_data[key] for key in allowed if token_data.get(key)}

    @classmethod
    def export_account(cls, token_data: dict, notes: str = "") -> str:
        payload = cls.load_payload()
        accounts = payload["accounts"]
        email = token_data.get("email", "")
        name = cls._unique_name(cls._derive_name(email), accounts)
        account = {
            "name": name,
            "platform": "openai",
            "type": "oauth",
            "credentials": cls._build_credentials(token_data),
            "concurrency": 10,
            "priority": 1,
            "rate_multiplier": 1.0,
            "auto_pause_on_expired": True,
        }
        if notes:
            account["notes"] = notes
        accounts.append(account)
        payload["exported_at"] = utc_now_iso()
        with open(cls.FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        logger.info(f"✅ Sub2API 账号已导出到 {cls.FILE} ({len(accounts)} 个账号)")
        return cls.FILE


# ===================== 5. 主流程 =====================
def _interactive_startup_menu() -> dict:
    """无参数启动时的交互式菜单，仅保留邮箱模式选择。"""
    print("=" * 60)
    print("ChatGPT 注册机 — 启动选项")
    print("=" * 60)
    print("请输入选项编号：")
    print("  1 — 使用自动邮箱后端（自动收件 + 自动提取验证码）")
    print("  2 — 手动输入自己的邮箱（验证码自己看邮箱后粘贴）")
    print()
    print("规则: 必须选 1 或 2 之一（邮箱模式互斥）")
    print("默认: 1（自动邮箱后端 + 免费账号）")
    print("=" * 60)

    while True:
        raw = input("请选择 [回车默认 1] > ").strip()
        if not raw:
            raw = "1"
        digits = set(c for c in raw if c.isdigit())
        invalid = digits - {"1", "2"}
        if invalid:
            print(f"[!] 无效选项: {','.join(sorted(invalid))}，仅支持 1/2")
            continue
        if "1" in digits and "2" in digits:
            print("[!] 不能同时选 1 和 2（邮箱模式互斥）")
            continue
        if "1" not in digits and "2" not in digits:
            print("[!] 必须选 1 或 2 中的一个（邮箱模式必填）")
            continue
        result = {"manual_email": "2" in digits}
        mode_email = "手动邮箱" if result["manual_email"] else "自动邮箱后端"
        print(f"\n>> 已选择: 邮箱={mode_email} | 模式=免费账号\n")
        return result


async def main():
    global DEBUG_MODE, HEROSMS_APIKEY
    global HEROSMS_INTERVAL, HEROSMS_FINISH_AFTER

    parser = argparse.ArgumentParser(description="ChatGPT 注册机")
    parser.add_argument("--headless", action="store_true", help="无头模式")
    parser.add_argument("--slow-mo", type=int, default=100, help="操作延迟(ms)")
    parser.add_argument("--password", type=str, default=None, help="保留兼容参数：当前主流程不再设置 ChatGPT 密码")
    parser.add_argument("--name", type=str, default=DEFAULT_NAME, help="账号名称。保持默认时会随机生成英文姓名")
    parser.add_argument("--birthdate", type=str, default=None, help="生日 YYYY-MM-DD")
    parser.add_argument("--email-timeout", type=int, default=180, help="邮件等待超时(s)")
    parser.add_argument("--debug", action="store_true", help="调试模式：出错时保持浏览器打开不关闭")
    parser.add_argument("--email", type=str, default=None, help="指定注册邮箱；配合 --jwt 时复用已有临时邮箱")
    parser.add_argument("--resume-alias-email", type=str, default="", help="使用已有 OutlookEmailPlus 别名邮箱登录恢复")
    parser.add_argument("--manual-email-base", type=str, default=None, help="自备邮箱母邮箱；未传 --email 时自动生成 +alias 注册邮箱")
    parser.add_argument("--email-code-file", type=str, default="", help="非交互式测试：等待该文件写入邮箱验证码")
    parser.add_argument("--herosms-apikey", type=str, default=HEROSMS_APIKEY, help="HeroSMS API Key；提供后按手机号匹配当前激活并自动读取 OTP")
    parser.add_argument("--herosms-interval", type=int, default=HEROSMS_INTERVAL, help="HeroSMS 轮询间隔秒数")
    parser.add_argument("--herosms-finish", action="store_true", help="读取 HeroSMS 后调用 setStatus=6 完成激活；默认不调用，避免号码从激活列表消失")
    parser.add_argument("--no-herosms-finish", action="store_true", help="兼容旧参数：读取 HeroSMS 后不调用 setStatus=6 完成激活")
    parser.add_argument("--proxy", type=str, default="", help="浏览器与 Codex OAuth token exchange 使用的代理，例如 7897 或 http://127.0.0.1:7897")
    parser.add_argument("--email-backend", type=str, default=None, choices=["temp_email", "outlook_email_plus", "manual"], help="选择邮箱后端：temp_email / outlook_email_plus / manual")
    parser.add_argument("--jwt", type=str, default=None, help="已有邮箱的JWT")
    parser.add_argument("--temp-password", type=str, default="", help="已有邮箱临时密码")
    parser.add_argument(
        "--manual-email",
        action="store_true",
        help="使用自己的邮箱注册：可配合 --email 指定邮箱；未传 --email 时启动后手动输入。"
        "验证码也手动从邮箱复制粘贴，启用后不再调用临时邮箱 API（即使 USER CONFIG 没填也能跑）",
    )
    args = parser.parse_args()

    # 无任何 CLI 参数 → 进入交互菜单收集邮箱模式 + 是否付款
    if len(sys.argv) == 1:
        choices = _interactive_startup_menu()
        args.manual_email = choices["manual_email"]

    resume_alias_email = args.resume_alias_email.strip()
    if resume_alias_email:
        if not is_valid_email(resume_alias_email):
            raise SystemExit("[参数错误] --resume-alias-email 不是有效邮箱地址")
        if args.email or args.manual_email or args.jwt:
            raise SystemExit("[参数错误] --resume-alias-email 不能与 --email / --manual-email / --jwt 同时使用")

    if args.email and not args.jwt:
        args.manual_email = True

    selected_backend = args.email_backend or OUTLOOK_EMAIL_PLUS_BACKEND
    if selected_backend not in ("temp_email", "outlook_email_plus", "manual"):
        raise SystemExit("[配置错误] OUTLOOK_EMAIL_PLUS_BACKEND 仅支持 temp_email / outlook_email_plus / manual")
    if resume_alias_email:
        selected_backend = "outlook_email_plus"
    if selected_backend == "manual":
        args.manual_email = True
    if args.manual_email:
        selected_backend = "manual"
    elif args.email and args.jwt:
        selected_backend = "manual"

    DEBUG_MODE = args.debug
    HEROSMS_APIKEY = args.herosms_apikey.strip()
    HEROSMS_INTERVAL = max(1, int(args.herosms_interval or 5))
    HEROSMS_FINISH_AFTER = bool(args.herosms_finish and not args.no_herosms_finish)

    # 启动前校验 USER CONFIG 是否完整（依据用户最终选择）
    _validate_user_config(skip_temp_email=args.manual_email, email_backend=selected_backend)

    chatgpt_password = args.password or generate_strong_password(16)
    birthdate = args.birthdate or generate_birthdate()
    real_name = args.name if args.name != DEFAULT_NAME else generate_real_name()
    email_prefix = generate_email_prefix()
    manual_email_base = (args.manual_email_base if args.manual_email_base is not None else MANUAL_EMAIL_BASE).strip()

    print("=" * 60)
    print("ChatGPT 注册机")
    print("=" * 60)
    if args.manual_email:
        print("邮箱模式: 手动（用户自备邮箱）")
    elif resume_alias_email:
        print(f"邮箱模式: 恢复登录（OutlookEmailPlus 既有别名: {resume_alias_email}）")
    elif selected_backend == "outlook_email_plus":
        print(f"邮箱 API: {OUTLOOK_EMAIL_PLUS_API_BASE}")
    else:
        print(f"邮箱 API: {TEMP_EMAIL_API}")
        print(f"邮箱前缀: {email_prefix}")
    print("ChatGPT密码: 已跳过设置")
    print(f"姓名: {real_name}")
    print(f"生日: {birthdate}")
    print("模式: 免费账号")
    print(f"HeroSMS自动OTP: {bool(HEROSMS_APIKEY)} (interval={HEROSMS_INTERVAL}s)")
    print("=" * 60)

    # 存储各步骤结果
    address = None
    jwt = None
    temp_password = None
    sub2api_output_path = ""
    outlook_email_client = None

    async def complete_outlook_claim(result: str = "success", detail: str = "") -> None:
        if not outlook_email_client:
            return
        try:
            await outlook_email_client.complete_claim(result=result, detail=detail)
        except Exception as e:
            logger.warning(f"OutlookEmailPlus claim-complete 失败: {e}")

    async def release_outlook_claim(reason: str) -> None:
        if not outlook_email_client:
            return
        try:
            await outlook_email_client.release_claim(reason=reason)
        except Exception as e:
            logger.warning(f"OutlookEmailPlus claim-release 失败: {e}")

    try:
        # Step 1: 获取邮箱地址（手动输入 / 复用 / 临时邮箱 API / outlookEmailPlus）
        if resume_alias_email:
            logger.info("第 1 步: 使用 OutlookEmailPlus 既有别名邮箱恢复登录")
            outlook_email_client = OutlookEmailPlusClient(
                base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
                api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
                caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
                task_id=f"chatgpt-resume-{email_prefix}-{int(time.time())}",
                provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
                project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
            )
            result = outlook_email_client.use_existing_email(resume_alias_email)
            address = result["email"]
            jwt = ""
            temp_password = ""
        elif args.manual_email:
            logger.info("第 1 步: 使用自有邮箱")
            if args.email:
                entered = args.email.strip()
                if not is_valid_email(entered):
                    raise SystemExit("[参数错误] --email 不是有效邮箱地址")
                address = entered
            elif manual_email_base:
                try:
                    address = generate_manual_email_alias(manual_email_base)
                except ValueError as e:
                    raise SystemExit(f"[配置错误] MANUAL_EMAIL_BASE / --manual-email-base: {e}")
                logger.info(f"母邮箱: {manual_email_base} -> 注册别名邮箱: {address}")
            else:
                loop = asyncio.get_running_loop()
                while True:
                    entered = await loop.run_in_executor(
                        None, lambda: input("请输入用于注册的邮箱地址 > ").strip()
                    )
                    if is_valid_email(entered):
                        address = entered
                        break
                    print("[!] 邮箱格式不正确，请重新输入。")
            jwt = ""
            temp_password = ""
            logger.info(f"使用自有邮箱: {address}")
        elif args.email and args.jwt:
            address = args.email
            jwt = args.jwt
            temp_password = args.temp_password or ""
            logger.info(f"复用已有邮箱: {address}")
        elif selected_backend == "outlook_email_plus":
            logger.info("第 1 步: 创建 OutlookEmailPlus 邮箱")
            outlook_email_client = OutlookEmailPlusClient(
                base_url=OUTLOOK_EMAIL_PLUS_API_BASE,
                api_key=OUTLOOK_EMAIL_PLUS_API_KEY,
                caller_id=OUTLOOK_EMAIL_PLUS_CALLER_ID,
                task_id=f"chatgpt-{email_prefix}-{int(time.time())}",
                provider=OUTLOOK_EMAIL_PLUS_PROVIDER,
                project_key=OUTLOOK_EMAIL_PLUS_PROJECT_KEY,
                email_domain=OUTLOOK_EMAIL_PLUS_EMAIL_DOMAIN,
            )
            result = await retry_with_backoff(
                lambda: outlook_email_client.acquire_address(),
                max_retries=3,
                description="创建 OutlookEmailPlus 地址",
            )
            address = result["email"]
            temp_password = ""
        else:
            logger.info("第 1 步: 创建临时邮箱")
            async with TempEmailClient() as email_client:
                result = await retry_with_backoff(
                    lambda: email_client.create_address(name=email_prefix),
                    max_retries=3, description="创建地址"
                )
                address = result["address"]
                jwt = email_client.latest_jwt or result["jwt"]
                temp_password = result.get("password") or ""

        # Step 2-7: ChatGPT 注册
        logger.info("第 2 步: ChatGPT 注册")
        async with ChatGPTBot(headless=args.headless, slow_mo=args.slow_mo, proxy_url=args.proxy) as bot:
            bot.oauth_code_file = args.email_code_file.strip()
            bot.allow_existing_account_login = bool(resume_alias_email)
            await bot.navigate_to_signup(address)
            await bot.wait_for_verification_page()

            # Step 3-4: 获取验证码
            code_info = None
            if args.manual_email:
                logger.info("第 3 步: 等待用户从自有邮箱复制验证码")
                if args.email_code_file:
                    manual_code = await wait_for_code_file(args.email_code_file, "邮箱验证码", timeout=args.email_timeout)
                    code_info = {"code": manual_code, "source": "file"}
                else:
                    print("\n>> 已向你的邮箱发送验证码，请去查看邮件并把 6 位验证码粘贴回来。")
                    loop = asyncio.get_running_loop()
                    while True:
                        manual_code = await loop.run_in_executor(
                            None, lambda: input("请输入邮箱验证码 > ").strip()
                        )
                        if manual_code and manual_code.isdigit() and 4 <= len(manual_code) <= 8:
                            code_info = {"code": manual_code, "source": "manual"}
                            break
                        print("[!] 验证码应为 4-8 位数字，请重新输入。")
            elif selected_backend == "outlook_email_plus":
                logger.info("第 3 步: 等待验证邮件（OutlookEmailPlus）")
                if not outlook_email_client:
                    raise OutlookEmailPlusError("OutlookEmailPlus 客户端未初始化")
                code_info = await outlook_email_client.poll_for_code(timeout=args.email_timeout, interval=5)
                logger.info(f"✅ 验证码: {code_info['code'][:3]}... (source={code_info['source']})")
            else:
                logger.info("第 3 步: 等待验证邮件...")
                async with TempEmailClient() as email_client:
                    mails = await email_client.poll_for_emails(
                        jwt,
                        timeout=args.email_timeout,
                        interval=5,
                        email_addr=address or "",
                        plaintext_password=temp_password or "",
                    )
                    if email_client.latest_jwt:
                        jwt = email_client.latest_jwt

                logger.info("第 4 步: 提取验证码")
                extractor = VerificationCodeExtractor()
                sorted_mails = sorted(mails, key=lambda m: m.get("created_at", ""), reverse=True)
                for mail in sorted_mails:
                    code_info = extractor.comprehensive_extract(mail)
                    if code_info:
                        logger.info(f"✅ 验证码: {code_info['code'][:3]}... (source={code_info['source']})")
                        break

                if not code_info:
                    logger.warning("自动提取失败，等待手动输入...")
                    loop = asyncio.get_running_loop()
                    manual_code = await loop.run_in_executor(
                        None, lambda: input("请输入邮箱验证码 (手动) > ").strip()
                    )
                    code_info = {"code": manual_code, "source": "manual"}

            # Step 5: 输入验证码
            logger.info("第 5 步: 输入验证码")
            await bot.enter_verification_code(code_info["code"])

            post_verify_state = await bot.detect_post_verification_destination()
            if post_verify_state == "about_you":
                # Step 6: 填写 about-you
                logger.info("第 6 步: about-you")
                await bot.fill_about_you(name=real_name, birthdate=birthdate)
            else:
                logger.info(f"跳过第 6 步 about-you，当前状态: {post_verify_state}")

            # Step 7: 等待登录完成
            logger.info("第 7 步: 等待登录完成")
            logged_in = await bot.wait_for_login_complete()
            if not logged_in:
                logger.error("登录验证失败，无法导出 Sub2API OAuth 账号")
                await release_outlook_claim("登录验证失败")
                return

            # Step 8: Codex OAuth → Sub2API 导出
            logger.info("第 8 步: Codex OAuth 导出 Sub2API")
            token_data = await bot.export_codex_sub2api_token(address)
            sub2api_output_path = Sub2APIExporter.export_account(token_data, notes="ChatGPT 免费注册机导出")

        print("\n" + "=" * 60)
        print("🎉 注册完成！")
        print(f"📧 邮箱: {address}")
        print("🔑 密码: 已跳过设置")
        print(f"📁 Sub2API文件: {sub2api_output_path}")
        print("=" * 60)
        await complete_outlook_claim("success", "注册成功")

    except VerificationTimeout as e:
        logger.error(f"验证邮件超时: {e}")
        logger.error("未导出 Sub2API：邮箱验证未完成")
        await complete_outlook_claim("verification_timeout", str(e))

    except OutlookEmailPlusTimeout as e:
        logger.error(f"验证邮件超时: {e}")
        logger.error("未导出 Sub2API：邮箱验证未完成")
        await complete_outlook_claim("verification_timeout", str(e))

    except OutlookEmailPlusError as e:
        logger.error(f"OutlookEmailPlus 失败: {e}")
        logger.error("未导出 Sub2API：邮箱后端失败")
        await release_outlook_claim(str(e))

    except UserAlreadyExistsError as e:
        logger.error(f"ChatGPT 账号已存在: {e}")
        logger.error("未导出 Sub2API：当前邮箱已被 OpenAI 占用")
        await complete_outlook_claim("credential_invalid", str(e))

    except (TempEmailError, SignupFlowError) as e:
        logger.error(f"注册流程失败: {type(e).__name__}: {e}")
        logger.error("未导出 Sub2API：流程未完成或 OAuth token 获取失败")
        await release_outlook_claim(f"{type(e).__name__}: {e}")

    except KeyboardInterrupt:
        logger.info("用户中断，未导出 Sub2API")
        await release_outlook_claim("用户中断")

    except Exception as e:
        logger.exception(f"未预期的错误: {e}")
        logger.error("未导出 Sub2API：出现未预期错误")
        await release_outlook_claim(f"未预期错误: {e}")

    finally:
        if outlook_email_client:
            await outlook_email_client.close()


if __name__ == "__main__":
    asyncio.run(main())
