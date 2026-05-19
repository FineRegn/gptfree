# ChatGPT 免费手机号注册自动化

基于 HeroSMS + OutlookEmailPlus 的 ChatGPT 免费账号自动化注册工具。

## 目录结构

```
gptfree/
├── gpt_free_ancientmethod_reg.py    # ★ 主入口：免费手机号注册 + OAuth + Sub2API
├── gpt_free_core.py                 # 核心 Bot 模块（公共库）
├── outlook_email_plus_integration.py  # Outlook 邮箱池客户端
├── hero_sms_client.py               # HeroSMS 接码客户端
├── codex_oauth_sub2api_once.py      # 单步 OAuth Sub2API 导出
├── release_outlook_email_pool.py/bat/cmd  # 邮箱池释放工具
├── outlookEmailPlus/                # 邮箱池 API 服务（Flask）
├── gpt_free_config.example.json     # 本地配置示例
├── screenshots/                     # 调试截图目录
└── gpt_free_accounts.json           # 默认账号记录输出
```

## 1. 配置步骤

### 1.1 安装依赖

```powershell
# 进入项目目录
cd "E:\Users\admin\Desktop\gptfree"

# 安装 Playwright 浏览器
python -m playwright install chromium
```

### 1.2 配置 `gpt_free_config.json`

复制配置示例并填写本地密钥；`gpt_free_config.json` 已被 `.gitignore` 排除，不要提交到仓库：

```powershell
Copy-Item .\gpt_free_config.example.json .\gpt_free_config.json
notepad .\gpt_free_config.json
```

常用配置项：

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `OUTLOOK_EMAIL_PLUS_API_BASE` | 邮箱池 API 地址 | `http://127.0.0.1:5001` |
| `OUTLOOK_EMAIL_PLUS_API_KEY` | 邮箱池 API Key | 本地 JSON 配置 |
| `OUTLOOK_EMAIL_PLUS_CALLER_ID` | 调用方标识 | `chatgpt-registration-bot` |
| `OUTLOOK_EMAIL_PLUS_PROJECT_KEY` | 项目 Key（防重复领取） | 留空 |
| `HEROSMS_APIKEY` | HeroSMS API Key | 本地 JSON 配置，也可用参数覆盖 |
| `MANUAL_EMAIL_BASE` | 自备邮箱母邮箱 | 留空 |
| `DEFAULT_OAUTH_PROXY` | 代理地址 | `http://127.0.0.1:7897` |

HeroSMS API Key 默认从 `gpt_free_config.json` 读取，也可用参数临时覆盖：`--herosms-apikey "your_key"`。

### 1.3 启动邮箱池服务

注册脚本依赖 OutlookEmailPlus 邮箱池 API，运行前需先启动：

```powershell
cd "E:\Users\admin\Desktop\gptfree\outlookEmailPlus"
$env:PYTHONIOENCODING="utf-8"
.\.venv\Scripts\python.exe start.py
```

保持该窗口运行，API 监听在 `http://127.0.0.1:5001`。

### 1.4 邮箱池维护

如果注册脚本异常中断，邮箱池中的账号可能停留在 `claimed` 状态，需释放：

```powershell
cd "E:\Users\admin\Desktop\gptfree"
.\release_outlook_email_pool.bat
```

预览（不实际修改）：

```powershell
python release_outlook_email_pool.py --dry-run
```

释放脚本会自动：
1. 将所有 `claimed` 状态的账号恢复为 `available`
2. 清理 `pool_status='available'` 但 `status!='active'` 的僵尸账号 → `retired`

## 2. 常用命令

> 以下命令直接使用当前环境中的 `python` 执行。

### 2.1 基础注册（自动领取邮箱 + HeroSMS 手机验证）

```powershell
cd "E:\Users\admin\Desktop\gptfree"

python -u gpt_free_ancientmethod_reg.py `
  --proxy "7897" `
  --debug
```

### 2.2 指定已有邮箱（直接从邮箱池 claim 指定邮箱）

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --claim-email "VetaCusenza25@hotmail.com" `
  --proxy "7897" `
  --debug
```

> `--claim-email`：直接 claim 指定邮箱，`complete_claim("email_limit")` 可正常退休该账号。

### 2.3 使用既有别名邮箱（不 claim 邮箱池）

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --resume-alias-email "user+tag@example.com" `
  --proxy "7897" `
  --debug
```

> 说明：`--resume-alias-email` 不占用邮箱池 claim，仅用于轮询验证码。

### 2.4 自备邮箱

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --email "your@email.com" `
  --proxy "7897" `
  --debug
```

### 2.5 复用已有 HeroSMS 号码

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --herosms-phone "57xxxxxxxxxx" `
  --proxy "7897" `
  --debug
```

### 2.6 恢复已有手机号账号并导出 Sub2API

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --resume-phone "57xxxxxxxxxx" `
  --password "账号登录密码" `
  --proxy "7897" `
  --debug
```

> `--resume-phone`：不购号、不注册，使用指定手机号走同一个手机号入口登录已有账号；登录成功后继续执行 Codex OAuth、添加邮箱验证并导出 Sub2API。未指定 `--email` / `--claim-email` / `--resume-alias-email` 时，会自动从 OutlookEmailPlus 领取邮箱。

### 2.7 仅测试 HeroSMS 接码（不注册）

```powershell
python -u gpt_free_ancientmethod_reg.py --herosms-only
```

### 2.8 指定输出目录

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --output ".\accounts" `
  --proxy "7897" `
  --debug
```

### 2.9 流程结束保持浏览器打开

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --keep-open `
  --proxy "7897" `
  --debug
```

### 2.10 自定义密码/姓名/生日

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --password "MyP@ssw0rd!" `
  --name "Alice" `
  --birthdate "1995-06-15" `
  --proxy "7897" `
  --debug
```

> `--output` 指定 Sub2API 输出根目录；不指定时默认使用 `accounts/`。每次运行都会在根目录下创建时间子目录，并按邮箱生成单账号 JSON。

### 2.11 批量顺序运行

```powershell
python -u gpt_free_ancientmethod_reg.py `
  --rounds 5 `
  --proxy "7897" `
  --debug
```

> `--rounds` 会按轮次顺序执行，不是并发。批量模式每轮自动领取 OutlookEmailPlus 邮箱并自动完成 HeroSMS 手机验证；邮箱池不足时直接终止后续轮次。

同一次运行会在 `accounts/<时间戳>/` 下按邮箱生成单账号 JSON。批量模式还会额外生成 `sub2api-free-batch-<时间戳>.json`，汇总本批次所有成功账号，便于一次性导入。

批量模式不允许复用固定资源，不能同时传：`--claim-email`、`--resume-alias-email`、`--email`、`--herosms-phone`、`--resume-phone`。

## 3. 完整参数速查

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--claim-email <email>` | 指定邮箱池中已有邮箱（直接 claim） | — |
| `--resume-alias-email <email>` | 使用既有别名邮箱（不 claim） | — |
| `--resume-phone <phone>` | 恢复模式：指定手机号登录已有账号后执行 OAuth/Sub2API 导出，必须配合 `--password` | — |
| `--email <email>` | 自备邮箱 | — |
| `--herosms-phone <phone>` | 跳过购号，复用已有号码 | — |
| `--herosms-only` | 仅测试 HeroSMS 接码 | — |
| `--herosms-max-price <n>` | HeroSMS 购号最高价(USD) | `0.08` |
| `--herosms-apikey <key>` | 覆盖 HeroSMS API Key | — |
| `--herosms-country <id>` | 手动指定国家 ID | 自动解析 Colombia |
| `--herosms-timeout <s>` | HeroSMS 接码超时 | `120` |
| `--herosms-interval <s>` | HeroSMS 轮询间隔 | `5` |
| `--phone-retries <n>` | 手机验证码超时后同轮换号重试次数 | `3` |
| `--herosms-finish` | 读取后完成 HeroSMS 激活 | `False` |
| `--proxy <port/url>` | 代理 | `http://127.0.0.1:7897` |
| `--headless` | 无头模式 | `False` |
| `--debug` | 出错时保持浏览器 | `False` |
| `--keep-open` | 流程结束后保持浏览器 | `False` |
| `--slow-mo <ms>` | Playwright 操作延迟 | `100` |
| `--email-timeout <s>` | 邮箱验证码超时 | `180` |
| `--email-code-file <path>` | 验证码文件桥接 | 系统临时文件 |
| `--password <pw>` | 密码 | 自动生成 |
| `--name <name>` | 账号名称 | 随机英文姓名 |
| `--birthdate <YYYY-MM-DD>` | 生日 | 随机生成 |
| `--output <path>` | Sub2API 输出根目录；每次运行按时间创建子目录，单账号按邮箱生成 JSON，批量额外生成总文件 | `accounts/` |
| `--rounds <n>` | 批量顺序运行轮数；邮箱池不足时直接终止 | `1` |

### 参数互斥规则

- `--claim-email` / `--resume-alias-email` / `--email` 三者互斥，不能同时使用
- `--rounds > 1` 时不能与 `--claim-email` / `--resume-alias-email` / `--email` / `--herosms-phone` / `--resume-phone` 同时使用

## 4. 注册流程

```
第 1 步: 获取邮箱
  ├── --claim-email  → claim 指定邮箱（DB 直接操作，含有效 claim_token）
  ├── --resume-alias-email → 使用既有别名（不 claim，无 claim_token）
  ├── --email → 自备邮箱
  └── 默认 → OutlookEmailPlus 自动随机领取（含 +alias）
第 2 步: HeroSMS 自动手机号验证
第 3 步: ChatGPT 手机号注册
第 4 步: 邮箱验证码
第 5 步: 个人信息（about-you）
第 6 步: 导航到 chatgpt.com 建立登录态
第 7 步: Codex OAuth（欢迎回来→密码→添加邮箱→验证→Codex授权）
第 8 步: 导出 Sub2API
```

手机号验证码超时处理：

```
当前 HeroSMS activation 超时 → 自动取消该 activation → 回到注册入口 → 重新获取手机号重试
```

默认同一轮最多重试 `--phone-retries 3` 次。重试期间当前 OutlookEmailPlus 邮箱 claim 不释放，避免同一轮反复领取邮箱。

完整流程成功后，OutlookEmailPlus 邮箱会执行 `release_claim` 释放回邮箱池，便于后续复用；只有 `email_in_use`、邮箱不可读等不可复用场景才会标记完成并退休。

恢复模式流程：

```
第 1 步: 获取 OAuth 添加邮箱（同上，默认自动领取 OutlookEmailPlus 邮箱）
第 2 步: 使用 --resume-phone 进入手机号登录
第 3 步: 跳转到 https://auth.openai.com/log-in/password 后输入 --password
第 4 步: 登录成功后进入 Codex OAuth
第 5 步: 添加邮箱 → 邮箱验证码 → Codex 授权 → 导出 Sub2API
```

Windows 通知：

- 在 Windows 下运行时，脚本会通过系统托盘气泡通知提示关键事件。
- 需要人工介入时会通知，例如 Cloudflare 人机验证、邮箱验证码自动读取失败转手动输入、`--keep-open` 保留浏览器排查。
- 当前轮失败时会通知，例如邮箱池耗尽、验证邮件超时、OutlookEmailPlus 失败、账号已存在、注册流程异常。
- 通知使用 PowerShell + Windows Forms 系统组件实现，无需额外安装 Python 依赖；非 Windows 环境会自动跳过。

## 5. email_in_use 自动重试

当邮箱达到注册上限（email_in_use）时，脚本会自动：

1. 调用 `complete_claim("email_limit")` 将该邮箱标记为 `retired`
2. 从邮箱池领取新邮箱
3. 若池空 → 直接终止当前运行；批量模式下不再继续后续轮次
4. 最多切换 3 个可用邮箱
5. 若 `--resume-alias-email` 模式（无 OutlookEmailPlus 客户端），直接终止

## 6. 单步 Sub2API 导出

已注册成功的账号，仅需补跑 OAuth 导出：

```powershell
cd "E:\Users\admin\Desktop\gptfree"
Remove-Item ".\oauth_code.txt" -ErrorAction SilentlyContinue

python -u codex_oauth_sub2api_once.py `
  --email "your@email.com" `
  --code-file ".\oauth_code.txt" `
  --output ".\sub2api-accounts.json" `
  --proxy "7897" `
  --debug
```

## 7. 相关文件

| 文件 | 作用 |
|------|------|
| `gpt_free_ancientmethod_reg.py` | 主入口：免费注册 + OAuth + Sub2API |
| `gpt_free_core.py` | 核心 Bot 模块（Playwright + HTTP） |
| `outlook_email_plus_integration.py` | 邮箱池客户端（领取/读信/验证码轮询） |
| `hero_sms_client.py` | HeroSMS 查询激活与 OTP 抽取 |
| `codex_oauth_sub2api_once.py` | 单步 Sub2API 导出 |
| `release_outlook_email_pool.bat` | 一键释放邮箱池 claimed 账号 |
| `registration_bot.log` | 运行日志 |
| `screenshots/` | 自动截图 |
