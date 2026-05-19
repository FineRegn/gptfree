# 临时邮箱：已配置 CF 域名但前端域名下拉无变化（Options 口径不一致）(BUG)

**创建日期**: 2026-04-09  
**修复日期**: 2026-04-09  
**关联功能**: CF Worker 临时邮箱 / 临时邮箱 UI (Temp Emails) / Settings v0.3 Tab 重构  
**分析人员**: AI 代码分析助手  
**状态**: ✅ 已修复（方案 A 已实施 + v0.3.1 自动同步 + 人工验收通过）

## 概述

在「设置」页面完成 **Cloudflare Worker 域名同步/配置**后，用户切换到「⚡ 临时邮箱」页面，选择 Provider 为 **Cloudflare Temp Mail**（`cloudflare_temp_mail`），期望能在“域名下拉框”中看到刚同步的域名列表。

但实际表现为：**域名下拉仍只有“自动分配域名”**，看起来像“设置了 CF 域名但前端没有反应”。

本质原因是：**临时邮箱页依赖的 `/api/temp-emails/options` 返回的 domains 数据源与 CF 域名同步写入的 settings key 不一致**，并且 options 接口当前 **未按前端选择的 provider 返回**。

---

## 影响范围

- **用户侧 UI**：临时邮箱页面无法选择 CF 域名创建邮箱（只能走“自动分配”或空域名）。
- **功能侧**：即使 CF Worker 已配置多域名，也无法通过 UI 指定域名创建临时邮箱。
- **可观测性**：用户误以为“同步失败/保存失败”，实际是 options 读取口径错误。

---

## 复现步骤（最小）

1. 启动 Web 应用并登录。
2. 进入「设置」页面，在 CF Worker 相关区域配置：
   - `cf_worker_base_url`
   - `cf_worker_admin_key`
3. 点击“从 CF Worker 同步域名”（对应接口 `POST /api/settings/cf-worker-sync-domains`），确认设置页只读域名字段出现域名列表。
4. 切换到「⚡ 临时邮箱」页面（Temp Emails）。
5. 在 Provider 下拉框选择 `cloudflare_temp_mail`。
6. 观察域名下拉框（`#tempEmailDomainSelect`）：
   - **实际**：下拉仍只有“自动分配域名”，不出现已同步的域名。
   - **预期**：下拉展示 CF 域名列表，可手动选择。

---

## 预期行为 vs 实际行为

### 预期

- 临时邮箱页面选择 Provider=CF 后，域名下拉应展示 CF Worker 同步的域名列表。
- 用户选择域名并点击“创建”，后端应按指定域名创建邮箱。

### 实际

- 临时邮箱页面域名下拉不更新；仍为“自动分配域名”。
- 用户无法通过 UI 指定 CF 域名。

---

## 关键代码定位

### 前端（临时邮箱页面）

文件：`static/js/features/temp_emails.js`

- `loadTempEmailOptions()`
  - 调用：`GET /api/temp-emails/options`
  - 使用：`data.options.domains` 渲染 `#tempEmailDomainSelect`
- `onTempEmailProviderChange(selectedProvider)`
  - 当 provider 为 `cloudflare_temp_mail` 时，会 `loadTempEmailOptions(true)` 强制刷新域名列表

> 当前前端请求 options **没有携带 provider_name 参数**，后端只能按默认 provider 返回 options。

### 后端（临时邮箱 options API）

文件：`outlook_web/controllers/temp_emails.py`

- `api_get_temp_email_options()`
  - 当前实现：`options = temp_mail_service.get_options()`
  - 问题：**无法按前端选择的 provider 返回**（只能按全局 runtime provider 返回）。

### 后端（CF Provider options 数据源）

文件：`outlook_web/services/temp_mail_provider_cf.py`

- `CloudflareTempMailProvider.get_options()`
  - 当前读取：
    - `temp_mail_domains`
    - `temp_mail_default_domain`
    - `temp_mail_prefix_rules`
  - 但设置页 CF 域名同步写入的是 **独立 key**：
    - `cf_worker_domains`
    - `cf_worker_default_domain`
    - `cf_worker_prefix_rules`
  - 结果：即使 CF 域名已同步，CF provider 仍可能返回空 domains。

### 后端（设置页 CF 域名同步写入点）

文件：`outlook_web/controllers/settings.py`

- `api_sync_cf_worker_domains()`
  - 写入：`cf_worker_domains`、`cf_worker_default_domain`
  - 设计意图：v0.3 之后 CF Worker 配置与 GPTMail 配置 **完全隔离**

---

## 根因总结（Root Cause）

该 BUG 由两类“口径不一致”叠加导致：

1) **options API 不按 provider 返回**

- `/api/temp-emails/options` 固定返回 `temp_mail_service.get_options()`，取的是“全局 runtime provider”的 options。
- 但临时邮箱页面允许用户在 UI 内临时选择 provider（`legacy_bridge` / `cloudflare_temp_mail`），两者不一致时会出现“UI 选择 CF，但 options 仍返回 legacy 的 domains”。

2) **CF provider options 读取的 settings key 与 CF 同步写入 key 不一致**

- CF 域名同步写入：`cf_worker_*`
- CF provider 读取：`temp_mail_*`

因此即使 runtime provider 恰好是 CF，也可能拿不到已同步的域名列表。

---

## 修复实现（已落地）

> 最终采用并落地：**方案 A（推荐）** + **v0.3.1 自动同步兜底**

1. 后端：`/api/temp-emails/options` 支持 `provider_name` 参数
   - 例如：`GET /api/temp-emails/options?provider_name=cloudflare_temp_mail`
2. Service：`TempMailService.get_options(provider_name=...)`
   - 调用 provider factory 根据指定 provider 返回 options
3. CF provider：`get_options()` 改为读取 `cf_worker_domains/cf_worker_default_domain/cf_worker_prefix_rules`
4. 前端：`loadTempEmailOptions()` 调用时带上当前 provider
5. v0.3.1：当 `cf_worker_domains` 为空且 `cf_worker_base_url` 已配置时，自动调用 `GET {base_url}/open_api/settings` 拉取 domains 并写回：
   - `cf_worker_domains`
   - `cf_worker_default_domain`
   - 同步失败非阻塞（仅 warning，不影响 options 返回）

---

## 验收结果（2026-04-09）

### 1) 接口验收

- `GET /api/temp-emails/options?provider_name=cloudflare_temp_mail` 返回 200
- 返回 `domains` 包含：
  - `zerodotsix.top`（默认）
  - `outlookmailplus.tech`
- DB 已写回：
  - `cf_worker_domains=[{"name":"zerodotsix.top","enabled":true},{"name":"outlookmailplus.tech","enabled":true}]`
  - `cf_worker_default_domain=zerodotsix.top`

### 2) UI 验收

- 临时邮箱页切换 provider=CF 后，域名下拉可见已同步域名
- 指定域名创建邮箱成功（前端提示 success）

### 3) 一次现场误报说明（非代码缺陷）

- 现象：前端创建时报 `UNAUTHORIZED` / 502
- 根因：`cf_worker_admin_key` 配置值错误（当时写入了 `admin123`，而该 Worker 实际 admin 密码是 `1234567890-=`）
- 结论：保存/加解密链路正常，属于配置值不一致导致的鉴权失败
- 处置：更新 `cf_worker_admin_key` 为正确值后，创建立即恢复正常

---

## 验收标准（修复后）

1. ✅ 设置页同步 CF 域名成功后，临时邮箱页选择 CF provider，域名下拉展示同步的域名列表。
2. ✅ 选择某个域名创建临时邮箱，后端实际创建的邮箱域名与选择一致。
3. ✅ 切换 provider 为 legacy 时，域名下拉禁用并显示“自动分配域名”。
4. ✅ 回归测试通过（重点用例：`tests.test_temp_mail_provider_cf`、`tests.test_temp_emails_api_regression`）。
