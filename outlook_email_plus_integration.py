# -*- coding: utf-8 -*-
"""outlookEmailPlus 对外邮箱池 API 客户端。"""
from __future__ import annotations

import asyncio
import json
import logging
import re
import time
import uuid
from dataclasses import dataclass
from typing import Any, Optional

import aiohttp


logger = logging.getLogger(__name__)


class OutlookEmailPlusError(Exception):
    """outlookEmailPlus API 操作失败。"""


class OutlookEmailPlusTimeout(OutlookEmailPlusError):
    """等待邮箱验证码超时。"""


class OutlookEmailPlusUnreadable(OutlookEmailPlusError):
    """邮箱读取链路不可用，应从邮箱池剔除。"""


@dataclass
class ClaimResult:
    account_id: str
    email: str
    primary_email: str
    alias_email: str
    claim_token: str
    caller_id: str
    task_id: str


def _first_text(*values: Any) -> str:
    for value in values:
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, (int, float)):
            return str(int(value))
    return ""


def _walk(value: Any):
    if isinstance(value, dict):
        for child in value.values():
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)
    else:
        yield value


def extract_verification_code(data: Any) -> Optional[str]:
    """从未完全固定 schema 的响应中提取 4-8 位验证码。"""
    preferred_keys = {"code", "verification_code", "verificationCode", "otp", "pin"}
    if isinstance(data, dict):
        for key, value in data.items():
            if key in preferred_keys:
                text = _first_text(value)
                match = re.fullmatch(r"\d{4,8}", text)
                if match:
                    return text

    texts = [item for item in _walk(data) if isinstance(item, str)]
    contextual = re.compile(
        r"(?:code|otp|验证码|verification|confirm|kode)[^0-9]{0,40}(\d{4,8})"
        r"|(\d{4,8})[^0-9]{0,40}(?:code|otp|验证码|verification|confirm|kode)",
        re.IGNORECASE,
    )
    for text in texts:
        match = contextual.search(text)
        if match:
            return next(group for group in match.groups() if group)

    for text in texts:
        match = re.search(r"(?<!\d)(\d{6})(?!\d)", text)
        if match:
            return match.group(1)
    return None


def _is_unreadable_mailbox_error(message: str) -> bool:
    upper = (message or "").upper()
    markers = (
        "UPSTREAM_READ_FAILED",
        "ACCOUNT_AUTH_EXPIRED",
        "NO_MAIL_PERMISSION",
        "IMAP_AUTH_FAILED",
        "IMAP_CONNECT_FAILED",
        "GRAPH/IMAP",
        "GRAPH/IMAP 均读取失败",
    )
    return any(marker in upper for marker in markers)


class OutlookEmailPlusClient:
    def __init__(
        self,
        base_url: str,
        api_key: str,
        caller_id: str,
        task_id: str = "",
        provider: str = "outlook",
        project_key: str = "",
        email_domain: str = "",
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key.strip()
        self.caller_id = caller_id.strip()
        self.task_id = task_id.strip() or f"task-{uuid.uuid4().hex[:12]}"
        self.provider = provider.strip()
        self.project_key = project_key.strip()
        self.email_domain = email_domain.strip()
        self._session: Optional[aiohttp.ClientSession] = None
        self._claim: Optional[ClaimResult] = None

    async def __aenter__(self):
        await self.open()
        return self

    async def __aexit__(self, *args: object) -> None:
        await self.close()

    async def open(self) -> None:
        if self._session and not self._session.closed:
            return
        self._session = aiohttp.ClientSession(
            timeout=aiohttp.ClientTimeout(total=30),
            connector=aiohttp.TCPConnector(limit=5),
        )

    async def close(self) -> None:
        if self._session and not self._session.closed:
            await self._session.close()

    async def _session_or_open(self) -> aiohttp.ClientSession:
        await self.open()
        if not self._session:
            raise OutlookEmailPlusError("OutlookEmailPlus session 初始化失败")
        return self._session

    def _headers(self) -> dict[str, str]:
        return {"X-API-Key": self.api_key}

    async def _read_json(self, response: aiohttp.ClientResponse) -> Any:
        text = await response.text()
        try:
            return json.loads(text) if text else {}
        except json.JSONDecodeError as exc:
            raise OutlookEmailPlusError(f"API 返回非 JSON: HTTP {response.status} {text[:300]}") from exc

    @staticmethod
    def _is_success(body: Any) -> bool:
        return isinstance(body, dict) and body.get("success") is True

    @staticmethod
    def _error_message(body: Any) -> str:
        if isinstance(body, dict):
            return f"{body.get('code') or 'ERROR'}: {body.get('message') or body}"
        return str(body)

    def _claim_payload(self) -> dict[str, object]:
        payload: dict[str, object] = {"caller_id": self.caller_id, "task_id": self.task_id, "use_alias": True}
        if self.provider:
            payload["provider"] = self.provider
        if self.project_key:
            payload["project_key"] = self.project_key
        if self.email_domain:
            payload["email_domain"] = self.email_domain
        return payload

    def _normalize_claim(self, data: Any) -> ClaimResult:
        if not isinstance(data, dict):
            data = {}
        return ClaimResult(
            account_id=_first_text(data.get("account_id"), data.get("accountId"), data.get("id")),
            email=_first_text(data.get("email"), data.get("address")),
            primary_email=_first_text(data.get("primary_email"), data.get("primaryEmail")),
            alias_email=_first_text(data.get("alias_email"), data.get("aliasEmail")),
            claim_token=_first_text(data.get("claim_token"), data.get("claimToken"), data.get("token")),
            caller_id=self.caller_id,
            task_id=self.task_id,
        )

    async def acquire_address(self) -> dict[str, str]:
        session = await self._session_or_open()
        async with session.post(
            f"{self.base_url}/api/external/pool/claim-random",
            json=self._claim_payload(),
            headers=self._headers(),
        ) as response:
            body = await self._read_json(response)
        if response.status != 200 or not self._is_success(body):
            raise OutlookEmailPlusError(f"claim-random 失败: HTTP {response.status} {self._error_message(body)}")
        claim = self._normalize_claim(body.get("data") if isinstance(body, dict) else {})
        if not claim.account_id or not claim.email or not claim.claim_token:
            raise OutlookEmailPlusError(f"claim-random 返回缺少 account_id/email/claim_token: {body}")
        self._claim = claim
        logger.info(f"✅ OutlookEmailPlus 邮箱: {claim.email} (account_id={claim.account_id})")
        return {
            "account_id": claim.account_id,
            "email": claim.email,
            "primary_email": claim.primary_email,
            "alias_email": claim.alias_email,
            "claim_token": claim.claim_token,
            "task_id": claim.task_id,
        }

    async def create_address(self, name: str = "") -> dict[str, str]:
        _ = name
        return await self.acquire_address()

    def use_existing_email(self, email: str) -> dict[str, str]:
        """使用已知邮箱地址轮询验证码，不占用邮箱池 claim。"""
        address = (email or "").strip()
        if not address:
            raise OutlookEmailPlusError("恢复流程邮箱不能为空")
        claim = ClaimResult(
            account_id="",
            email=address,
            primary_email=address.split("+", 1)[0] + "@" + address.split("@", 1)[1] if "+" in address and "@" in address else address,
            alias_email=address,
            claim_token="",
            caller_id=self.caller_id,
            task_id=self.task_id,
        )
        self._claim = claim
        logger.info(f"✅ OutlookEmailPlus 使用已有邮箱: {claim.email}")
        return {
            "account_id": claim.account_id,
            "email": claim.email,
            "primary_email": claim.primary_email,
            "alias_email": claim.alias_email,
            "claim_token": claim.claim_token,
            "task_id": claim.task_id,
        }

    def setup_custom_claim(
        self,
        account_id: str,
        email: str,
        primary_email: str,
        claim_token: str,
    ) -> dict[str, str]:
        """使用外部预先分配好的 claim（如直接 DB 操作创建的别名），
        使 complete_claim / release_claim 仍可正常工作。"""
        self._claim = ClaimResult(
            account_id=str(account_id),
            email=str(email),
            primary_email=str(primary_email),
            alias_email=str(email),
            claim_token=str(claim_token),
            caller_id=self.caller_id,
            task_id=self.task_id,
        )
        logger.info(f"✅ OutlookEmailPlus 自定义 claim: {self._claim.email} (account_id={self._claim.account_id})")
        return {
            "account_id": self._claim.account_id,
            "email": self._claim.email,
            "primary_email": self._claim.primary_email,
            "alias_email": self._claim.alias_email,
            "claim_token": self._claim.claim_token,
            "task_id": self._claim.task_id,
        }

    def _require_claim(self) -> ClaimResult:
        if not self._claim:
            raise OutlookEmailPlusError("尚未领取 OutlookEmailPlus 邮箱")
        return self._claim

    def _mail_params(self, timeout_seconds: int = 0, folder: str = "", baseline_timestamp: int = 0) -> dict[str, str]:
        claim = self._require_claim()
        params: dict[str, str] = {"email": claim.email, "since_minutes": "30", "top": "20"}
        if baseline_timestamp > 0:
            params["baseline_timestamp"] = str(baseline_timestamp)
        if timeout_seconds > 0:
            params["timeout_seconds"] = str(max(1, min(timeout_seconds, 120)))
            params["poll_interval"] = str(max(1, min(5, timeout_seconds)))
            params["mode"] = "sync"
        if folder:
            params["folder"] = folder
        return params

    async def _get_external(self, path: str, params: dict[str, str]) -> Any:
        session = await self._session_or_open()
        try:
            async with session.get(f"{self.base_url}{path}", params=params, headers=self._headers()) as response:
                body = await self._read_json(response)
        except asyncio.TimeoutError as exc:
            raise OutlookEmailPlusError(f"{path} 请求超时") from exc
        except aiohttp.ClientError as exc:
            raise OutlookEmailPlusError(f"{path} 网络请求失败: {exc}") from exc
        if response.status not in (200, 202):
            raise OutlookEmailPlusError(f"{path} 失败: HTTP {response.status} {self._error_message(body)}")
        return body

    async def _fetch_code_once(self, wait_seconds: int, baseline_timestamp: int = 0) -> Optional[dict[str, str]]:
        paths = [
            ("/api/external/verification-code", self._mail_params(baseline_timestamp=baseline_timestamp)),
            ("/api/external/verification-code", self._mail_params(folder="junkemail", baseline_timestamp=baseline_timestamp)),
            ("/api/external/wait-message", self._mail_params(wait_seconds, baseline_timestamp=baseline_timestamp)),
            ("/api/external/wait-message", self._mail_params(wait_seconds, folder="junkemail", baseline_timestamp=baseline_timestamp)),
            ("/api/external/messages/latest", self._mail_params(baseline_timestamp=baseline_timestamp)),
            ("/api/external/messages/latest", self._mail_params(folder="junkemail", baseline_timestamp=baseline_timestamp)),
            ("/api/external/messages", self._mail_params(baseline_timestamp=baseline_timestamp)),
            ("/api/external/messages", self._mail_params(folder="junkemail", baseline_timestamp=baseline_timestamp)),
        ]
        last_error = ""
        unreadable_errors: list[str] = []
        for path, params in paths:
            try:
                body = await self._get_external(path, params)
            except OutlookEmailPlusError as exc:
                last_error = str(exc)
                if _is_unreadable_mailbox_error(last_error):
                    unreadable_errors.append(last_error)
                logger.warning(f"OutlookEmailPlus 本轮接口失败，继续尝试下一个接口: {last_error}")
                continue
            code = extract_verification_code(body)
            if code:
                return {"code": code, "source": path}
            if isinstance(body, dict) and body.get("success") is False:
                last_error = self._error_message(body)
                if _is_unreadable_mailbox_error(last_error):
                    unreadable_errors.append(last_error)
        if len(unreadable_errors) >= 3:
            raise OutlookEmailPlusUnreadable(unreadable_errors[-1])
        if last_error:
            logger.debug(f"OutlookEmailPlus 本轮未取得验证码: {last_error}")
        return None

    async def poll_for_code(self, timeout: int = 120, interval: int = 5, baseline_timestamp: int = 0) -> dict[str, str]:
        if baseline_timestamp <= 0:
            baseline_timestamp = int(time.time())
        deadline = asyncio.get_event_loop().time() + timeout
        while True:
            remaining = int(deadline - asyncio.get_event_loop().time())
            if remaining <= 0:
                break
            wait_seconds = max(1, min(interval, remaining, 120))
            code_info = await self._fetch_code_once(wait_seconds, baseline_timestamp=baseline_timestamp)
            if code_info:
                return code_info
            await asyncio.sleep(max(1, min(interval, remaining)))
        raise OutlookEmailPlusTimeout(f"{timeout}s 内未获取到 OutlookEmailPlus 邮箱验证码")

    async def complete_claim(self, result: str = "success", detail: str = "") -> None:
        claim = self._require_claim()
        if not claim.account_id or not claim.claim_token:
            logger.info("OutlookEmailPlus 既有邮箱模式，跳过 claim-complete")
            return
        payload = {
            "account_id": claim.account_id,
            "claim_token": claim.claim_token,
            "caller_id": claim.caller_id,
            "task_id": claim.task_id,
            "result": result,
        }
        if detail:
            payload["detail"] = detail
        session = await self._session_or_open()
        async with session.post(
            f"{self.base_url}/api/external/pool/claim-complete",
            json=payload,
            headers=self._headers(),
        ) as response:
            body = await self._read_json(response)
        if response.status != 200 or not self._is_success(body):
            raise OutlookEmailPlusError(f"claim-complete 失败: HTTP {response.status} {self._error_message(body)}")

    async def release_claim(self, reason: str = "abandoned") -> None:
        claim = self._require_claim()
        if not claim.account_id or not claim.claim_token:
            logger.info("OutlookEmailPlus 既有邮箱模式，跳过 claim-release")
            return
        payload = {
            "account_id": claim.account_id,
            "claim_token": claim.claim_token,
            "caller_id": claim.caller_id,
            "task_id": claim.task_id,
            "reason": reason,
        }
        session = await self._session_or_open()
        async with session.post(
            f"{self.base_url}/api/external/pool/claim-release",
            json=payload,
            headers=self._headers(),
        ) as response:
            body = await self._read_json(response)
        if response.status != 200 or not self._is_success(body):
            raise OutlookEmailPlusError(f"claim-release 失败: HTTP {response.status} {self._error_message(body)}")
