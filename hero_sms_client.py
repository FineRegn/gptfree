# -*- coding: utf-8 -*-
"""HeroSMS 短信查询客户端。

用于已知手机号场景：通过 HeroSMS 当前激活列表按手机号匹配 activation，
再轮询短信验证码。
"""
from __future__ import annotations

import asyncio
import logging
import re
from dataclasses import dataclass
from typing import Any, ClassVar, cast

import aiohttp


logger = logging.getLogger(__name__)

# OTP 通常为 6 位。这里拒绝 4 位码，避免误取 HeroSMS 激活对象里的旧短码。
CODE_RE = re.compile(r"(?<!\d)(\d{5,8})(?!\d)")
PHONE_FIELD_KEYS = (
    "phoneNumber",
    "phone_number",
    "phone",
    "number",
    "tel",
    "msisdn",
)
ACTIVATION_ID_FIELD_KEYS = (
    "activationId",
    "activation_id",
    "id",
)
ACTIVE_LIST_FIELD_KEYS = (
    "activeActivations",
    "activations",
    "data",
    "items",
    "result",
    "list",
)


class HeroSMSError(RuntimeError):
    """HeroSMS API 操作失败。"""


@dataclass(frozen=True)
class HeroSMSResult:
    code: str
    message: str
    phone: str
    activation_id: str


@dataclass(frozen=True)
class HeroSMSActivation:
    phone: str
    activation_id: str
    activation_cost: float | None = None
    country: str = ""
    service: str = ""


class HeroSMSClient:
    BASE_URL: ClassVar[str] = "https://hero-sms.com/stubs/handler_api.php"

    def __init__(self, api_key: str, timeout: int = 30):
        self.api_key = (api_key or "").strip()
        self.timeout = timeout
        if not self.api_key:
            raise HeroSMSError("HeroSMS api_key 为空")

    async def _get(self, action: str, params: dict[str, str] | None = None) -> Any:
        request_params = {"api_key": self.api_key, "action": action, **(params or {})}
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=self.timeout)) as session:
            async with session.get(self.BASE_URL, params=request_params) as resp:
                text = await resp.text()
                if resp.status not in (200, 204):
                    raise HeroSMSError(f"HeroSMS {action} HTTP {resp.status}: {text[:200]}")
                if resp.status == 204 or not text.strip():
                    return "OK"
                try:
                    return await resp.json(content_type=None)
                except Exception:
                    return text.strip()

    async def resolve_country_id(self, iso_code: str = "CO", country_name: str = "Colombia") -> str:
        """从 HeroSMS 国家列表解析国家 ID，避免硬编码供应商 ID。"""
        iso_code = iso_code.strip().upper()
        name_candidates = {country_name.strip().lower(), "colombia", "哥伦比亚"}
        last_error = ""
        for action in ("getCountries", "getCountriesList"):
            try:
                data = await self._get(action)
            except HeroSMSError as exc:
                last_error = str(exc)
                continue
            country_id = extract_country_id(data, iso_code, name_candidates)
            if country_id:
                logger.info(f"HeroSMS 已解析国家 {iso_code}/{country_name}: country={country_id}")
                return country_id
        detail = f"，最后错误: {last_error}" if last_error else ""
        raise HeroSMSError(f"HeroSMS 国家列表中未找到 {iso_code}/{country_name}{detail}")

    async def acquire_number(
        self,
        service: str,
        country: str,
        max_price: float,
        operator: str = "",
        max_retries: int = 5,
    ) -> HeroSMSActivation:
        """调用 getNumberV2 自动购买号码。"""
        params = {"service": service, "country": country, "maxPrice": f"{max_price:.2f}"}
        if operator:
            params["operator"] = operator

        last_response = ""
        for attempt in range(1, max_retries + 1):
            data = await self._get("getNumberV2", params)
            if isinstance(data, str):
                last_response = data
                if data == "NO_BALANCE":
                    raise HeroSMSError("HeroSMS 余额不足")
                if data == "BAD_KEY":
                    raise HeroSMSError("HeroSMS API Key 无效")
                if data == "NO_NUMBERS" and attempt < max_retries:
                    logger.warning(f"HeroSMS 暂无可用号码，3 秒后重试 ({attempt}/{max_retries})")
                    await asyncio.sleep(3)
                    continue
                raise HeroSMSError(f"HeroSMS 获取号码失败: {data}")

            if not isinstance(data, dict):
                raise HeroSMSError(f"HeroSMS getNumberV2 返回格式异常: {str(data)[:200]}")

            activation_id = _first_text(data.get("activationId"), data.get("activation_id"), data.get("id"))
            phone = _first_text(data.get("phoneNumber"), data.get("phone_number"), data.get("phone"), data.get("number"))
            if not activation_id or not phone:
                raise HeroSMSError(f"HeroSMS getNumberV2 缺少 activationId/phoneNumber: {sanitize_payload(data)}")
            if not phone.startswith("+"):
                phone = f"+{phone}"
            cost = parse_float(data.get("activationCost") or data.get("cost") or data.get("price"))
            logger.info(f"HeroSMS 已获取号码: {phone} (activation={activation_id}, cost={cost})")
            return HeroSMSActivation(
                phone=phone,
                activation_id=activation_id,
                activation_cost=cost,
                country=country,
                service=service,
            )

        raise HeroSMSError(f"HeroSMS 获取号码失败，最后响应: {last_response}")

    async def mark_ready(self, activation_id: str) -> None:
        await self._get("setStatus", {"id": activation_id, "status": "1"})

    async def cancel_activation(self, activation_id: str) -> None:
        await self._get("setStatus", {"id": activation_id, "status": "8"})

    async def poll_code_by_activation(
        self,
        activation: HeroSMSActivation,
        timeout: int = 180,
        interval: int = 5,
        finish_after: bool = True,
        cancel_on_timeout: bool = True,
    ) -> HeroSMSResult:
        """按 activation id 轮询验证码，用于自动购号后的短信接收。"""
        try:
            await self.mark_ready(activation.activation_id)
        except HeroSMSError as exc:
            logger.warning(f"HeroSMS 标记准备接码失败，继续轮询: {exc}")

        deadline = asyncio.get_running_loop().time() + timeout
        last_message = ""
        while asyncio.get_running_loop().time() < deadline:
            code, message = await self._get_status_code(activation.activation_id)
            if message:
                last_message = message
            if code:
                logger.info(f"HeroSMS 已获取验证码: {code[:2]}***")
                if finish_after:
                    await self.finish_activation(activation.activation_id)
                return HeroSMSResult(
                    code=code,
                    message=message,
                    phone=activation.phone,
                    activation_id=activation.activation_id,
                )
            await asyncio.sleep(interval)

        if cancel_on_timeout:
            try:
                await self.cancel_activation(activation.activation_id)
                logger.info(f"HeroSMS 已取消超时激活: {activation.activation_id}")
            except HeroSMSError as exc:
                logger.warning(f"HeroSMS 取消超时激活失败: {exc}")
        detail = f"，最后响应: {last_message[:120]}" if last_message else ""
        raise HeroSMSError(f"{timeout}s 内未收到 HeroSMS 验证码{detail}")

    async def find_activation_by_phone(self, phone: str) -> dict[str, Any]:
        normalized_phone = normalize_phone(phone)
        data = await self._get("getActiveActivations", {"start": "0", "limit": "100"})
        activations = extract_activation_items(data)
        target_variants = phone_variants(normalized_phone)
        seen_phones: list[str] = []

        for item in activations:
            candidate_phone = extract_activation_phone(item)
            if not candidate_phone:
                continue
            seen_phones.append(candidate_phone)
            candidate_variants = phone_variants(candidate_phone)
            if phones_match(target_variants, candidate_variants):
                activation_id = extract_activation_id(item)
                if not activation_id:
                    logger.warning(f"HeroSMS 已匹配手机号但缺少 activation id: {sanitize_payload(item)}")
                    continue
                logger.info(f"HeroSMS 已匹配手机号: {candidate_phone} (activation={activation_id})")
                return cast(dict[str, Any], item)

        sample = ", ".join(seen_phones[:10]) or "无可识别手机号"
        raise HeroSMSError(
            f"HeroSMS 当前激活列表中未找到手机号: {normalized_phone} "
            f"(activations={len(activations)}, seen={sample})"
        )

    async def _get_status_code(self, activation_id: str) -> tuple[str, str]:
        data = await self._get("getStatusV2", {"id": activation_id})
        code, message = extract_code_from_payload(data)
        if code:
            return code, message

        data = await self._get("getStatus", {"id": activation_id})
        if isinstance(data, str):
            if data.startswith("STATUS_OK:"):
                code = data.split(":", 1)[1].strip()
                return code, data
            if data in ("STATUS_CANCEL", "STATUS_CANCELLED"):
                raise HeroSMSError("HeroSMS 激活已取消")
        return "", str(data)[:200]

    async def _get_all_sms_code(self, activation_id: str) -> tuple[str, str]:
        data = await self._get("getAllSms", {"id": activation_id, "size": "10", "page": "1"})
        return extract_code_from_payload(data)

    async def finish_activation(self, activation_id: str) -> None:
        try:
            await self._get("setStatus", {"id": activation_id, "status": "6"})
        except HeroSMSError as e:
            logger.warning(f"HeroSMS setStatus 完成激活失败，尝试 finishActivation: {e}")
            await self._get("finishActivation", {"id": activation_id})
        logger.info(f"HeroSMS 已完成激活: {activation_id}")

    async def poll_code_by_phone(
        self,
        phone: str,
        timeout: int = 120,
        interval: int = 5,
        finish_after: bool = False,
    ) -> HeroSMSResult:
        activation = await self.find_activation_by_phone(phone)
        activation_id = extract_activation_id(activation)
        matched_phone = extract_activation_phone(activation) or str(phone).strip()
        initial_code, initial_message = extract_code_from_payload(activation)
        if initial_code:
            logger.info(f"HeroSMS 激活列表已有旧验证码，忽略并继续等待新 OTP: {initial_code[:2]}***")

        deadline = asyncio.get_running_loop().time() + timeout
        last_message = initial_message
        while asyncio.get_running_loop().time() < deadline:
            code, message = await self._get_status_code(activation_id)
            if not code:
                code, message = await self._get_all_sms_code(activation_id)
            if message:
                last_message = message
            if code:
                logger.info(f"HeroSMS 已获取 OTP: {code[:2]}***")
                if finish_after:
                    await self.finish_activation(activation_id)
                else:
                    logger.info("HeroSMS 保留激活状态：未调用 setStatus=6/finishActivation")
                return HeroSMSResult(code=code, message=message, phone=matched_phone, activation_id=activation_id)
            await asyncio.sleep(interval)

        detail = f"，最后响应: {last_message[:120]}" if last_message else ""
        raise HeroSMSError(f"{timeout}s 内未收到 HeroSMS 验证码{detail}")


def normalize_phone(phone: str, strict: bool = True) -> str:
    normalized = re.sub(r"\D+", "", (phone or "").strip())
    if strict and not normalized:
        raise HeroSMSError("手机号必须包含数字")
    return normalized


def _first_text(*values: Any) -> str:
    for value in values:
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, (int, float)):
            return str(int(value))
    return ""


def parse_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(re.sub(r"[^0-9.]+", "", str(value)))
    except ValueError:
        return None


def extract_country_id(data: Any, iso_code: str, name_candidates: set[str]) -> str:
    if isinstance(data, str):
        try:
            import json

            return extract_country_id(json.loads(data), iso_code, name_candidates)
        except Exception:
            return ""
    if isinstance(data, list):
        for item in data:
            country_id = extract_country_id(item, iso_code, name_candidates)
            if country_id:
                return country_id
        return ""
    if not isinstance(data, dict):
        return ""

    direct_id = _first_text(data.get("id"), data.get("countryId"), data.get("country_id"))
    direct_iso = _first_text(data.get("isoCode"), data.get("iso"), data.get("code"), data.get("iso2")).upper()
    direct_names = {
        _first_text(data.get("eng"), data.get("en"), data.get("name"), data.get("country"), data.get("title")).lower(),
        _first_text(data.get("chn"), data.get("zh"), data.get("label")).lower(),
    }
    direct_names.discard("")
    if direct_id and (direct_iso == iso_code or direct_names.intersection(name_candidates)):
        return direct_id

    for key, value in data.items():
        if isinstance(value, str):
            if key.isdigit() and value.strip().lower() in name_candidates:
                return key
            continue
        if isinstance(value, dict):
            country_id = extract_country_id({"id": key, **value} if key.isdigit() else value, iso_code, name_candidates)
            if country_id:
                return country_id
        elif isinstance(value, list):
            country_id = extract_country_id(value, iso_code, name_candidates)
            if country_id:
                return country_id
    return ""


def phone_variants(phone: str) -> set[str]:
    """生成常见手机号格式变体，用于匹配 HeroSMS 不同返回格式。"""
    normalized = normalize_phone(phone, strict=False)
    if not normalized:
        return set()

    variants = {normalized}
    if normalized.startswith("00") and len(normalized) > 2:
        variants.add(normalized[2:])
    if normalized.startswith("+"):
        variants.add(normalize_phone(normalized[1:], strict=False))
    if normalized.startswith("0") and len(normalized) > 1:
        variants.add(normalized[1:])
    if normalized.startswith("62") and len(normalized) > 2:
        local = normalized[2:]
        variants.add(local)
        variants.add("0" + local)
    if normalized.startswith("8"):
        variants.add("62" + normalized)
        variants.add("0" + normalized)
    return {item for item in variants if item}


def phones_match(target_variants: set[str], candidate_variants: set[str]) -> bool:
    for target in target_variants:
        for candidate in candidate_variants:
            if target == candidate or target.endswith(candidate) or candidate.endswith(target):
                return True
    return False


def extract_activation_items(data: Any) -> list[dict[str, Any]]:
    """从 HeroSMS 不同响应结构中递归提取 activation 列表。"""
    items: list[dict[str, Any]] = []
    if isinstance(data, list):
        for value in data:
            if isinstance(value, dict):
                if extract_activation_id(value) or extract_activation_phone(value):
                    items.append(value)
                else:
                    items.extend(extract_activation_items(value))
        return items

    if isinstance(data, dict):
        if extract_activation_id(data) or extract_activation_phone(data):
            return [data]
        for key in ACTIVE_LIST_FIELD_KEYS:
            value = data.get(key)
            if value is not None:
                items.extend(extract_activation_items(value))
        if items:
            return items
        for value in data.values():
            items.extend(extract_activation_items(value))
    return items


def extract_activation_phone(item: dict[str, Any]) -> str:
    for key in PHONE_FIELD_KEYS:
        value = item.get(key)
        phone = normalize_phone(str(value or ""), strict=False)
        if phone:
            return phone
    return ""


def extract_activation_id(item: dict[str, Any]) -> str:
    for key in ACTIVATION_ID_FIELD_KEYS:
        value = str(item.get(key) or "").strip()
        if value:
            return value
    return ""


def sanitize_payload(item: dict[str, Any]) -> str:
    preview = {key: item.get(key) for key in (*ACTIVATION_ID_FIELD_KEYS, *PHONE_FIELD_KEYS) if key in item}
    return str(preview)[:200]


def extract_sms_code(message: str) -> str | None:
    if not message:
        return None
    context_patterns = [
        r"(?:OTP|kode|code|verification|验证码|驗證碼)[^0-9]{0,40}(\d{5,8})",
        r"(\d{5,8})[^0-9]{0,40}(?:OTP|kode|code|verification|验证码|驗證碼)",
    ]
    for pattern in context_patterns:
        match = re.search(pattern, message, re.IGNORECASE)
        if match:
            return match.group(1)
    match = CODE_RE.search(message)
    return match.group(1) if match else None


def extract_code_from_payload(data: Any) -> tuple[str, str]:
    if isinstance(data, str):
        return extract_sms_code(data) or "", data
    if isinstance(data, dict):
        direct_code = data.get("smsCode") or data.get("code")
        direct_text = str(data.get("smsText") or data.get("text") or data.get("message") or "")
        if direct_code:
            return str(direct_code).strip(), direct_text
        for key in ("sms", "call"):
            node = data.get(key)
            if isinstance(node, dict):
                code = str(node.get("code") or "").strip()
                text = str(node.get("text") or "").strip()
                if code:
                    return code, text
                parsed = extract_sms_code(text)
                if parsed:
                    return parsed, text
        for key in ("data", "activeActivations", "items", "result"):
            value = data.get(key)
            code, text = extract_code_from_payload(value)
            if code:
                return code, text
    if isinstance(data, list):
        for item in data:
            code, text = extract_code_from_payload(item)
            if code:
                return code, text
    return "", str(data)[:200]
