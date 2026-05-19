# Codex 注册流程脚本（当前版本）

本项目当前实现不是 DDG + Browserbase 方案，而是：

- 本地 `puppeteer-real-browser` 自动化浏览器
- HeroSMS 接码（手机号）
- cloud-mail / 兼容 API 邮箱服务（默认按 cloud-mail 协议）
- OpenAI OAuth 授权换取 Token

如果你是按旧 README 配的 `ddgToken/mailInboxUrl`，会直接跑不通。

## 运行前提

- Node.js 18+
- 可用的 HeroSMS API Key
- 可用的邮箱服务（默认 cloud-mail 协议）
- 能启动图形浏览器的环境

注意：
- Linux 下禁止通过 `xvfb-run` 启动，脚本会主动报错退出
- `config.json` 必须是严格 JSON，不能写 `//` 注释

## 安装

```bash
npm install
```

## 配置

编辑 `config.json`（示例）：

```json
{
  "heroSmsApiKey": "YOUR_HEROSMS_API_KEY",
  "heroSmsService": "dr",
  "heroSmsCountry": 16,
  "heroSmsPromptCountrySelection": true,
  "heroSmsCountryTopN": 10,
  "phoneCountryCode": "GB",
  "mailBaseUrl": "https://your-mail-site.example.com",
  "mailProvider": "cloud-mail",
  "mailAdminEmail": "admin@example.com",
  "mailAdminPassword": "YOUR_ADMIN_PASSWORD",
  "mailAdminToken": "",
  "mailUserType": 1,
  "mailSitePassword": "",
  "mailDomain": "your-domain.example.com",
  "mailDomains": ["your-domain.example.com", "your-other-domain.example.com"],
  "proxyHost": "",
  "proxyPort": 0,
  "proxyUsername": "",
  "proxyPassword": "",
  "tokenOutputDir": "tokens",
  "tokenOutputDirs": ["tokens"]
}
```

另外现在支持环境覆盖文件：

- `config.json`：共用基础配置
- `config.local.json`：本地 macOS 覆盖配置
- `config.server.json`：Linux / 服务器覆盖配置

加载规则：

- macOS 默认叠加 `config.local.json`
- Linux 默认叠加 `config.server.json`
- 也可以手动指定：

```bash
CONFIG_PROFILE=local node index.js 1
CONFIG_PROFILE=server node index.js 1
CONFIG_FILE=./config.local.json node index.js 1
```

### 配置字段说明

| 字段 | 说明 | 必填 |
|---|---|---|
| `heroSmsApiKey` | HeroSMS 的 API Key | 是 |
| `heroSmsService` | HeroSMS 服务代码，默认 `dr` | 否 |
| `heroSmsCountry` | 默认 HeroSMS 国家 ID，价格查询失败时的兜底值 | 否 |
| `heroSmsPromptCountrySelection` | 启动时是否先展示最便宜国家前 N 名并交互选择 | 否 |
| `heroSmsCountryTopN` | 价格列表展示前几名，默认 `10` | 否 |
| `phoneCountryCode` | 默认手机国家 ISO 代码，例如 `GB` / `US` | 否 |
| `phoneCountries` | 可参与比价与匹配的国家清单；未填时使用内置常见国家列表 | 否 |
| `mailBaseUrl` | 邮箱服务根地址 | 是 |
| `mailProvider` | 邮箱协议，默认 `cloud-mail` | 否 |
| `mailAdminEmail` | cloud-mail 管理员登录邮箱（可不填：会尝试用 `mailAdminPassword` 作为邮箱） | 建议填 |
| `mailAdminPassword` | cloud-mail 管理员登录密码（或 legacy 的 admin key） | 是 |
| `mailAdminToken` | cloud-mail 管理员 token，填了就不再走 `/api/login` | 否 |
| `mailUserType` | cloud-mail 新建邮箱用户角色 ID，默认 `1` | 否 |
| `mailSitePassword` | legacy 协议二次鉴权（`x-custom-auth`） | 否 |
| `mailDomain` | 单个邮箱域名；作为兼容旧配置的兜底值 | 否 |
| `mailDomains` | 多域名池；每一轮创建邮箱前会随机选一个域名 | 建议填 |
| `proxyHost/proxyPort/proxyUsername/proxyPassword` | 浏览器和 OAuth 请求代理（可选） | 否 |
| `tokenOutputDir` | 单目录输出 token | 否 |
| `tokenOutputDirs` | 多目录输出 token，配置后优先于 `tokenOutputDir` | 否 |

## 启动方式

### 1) 完整流程（默认）

```bash
node index.js 1
```

含义：跑一轮完整链路，直到产出 1 个 token 文件。  
你也可以批量：

```bash
node index.js 5
```

完整流程启动前会先尝试：

1. 从 HeroSMS 拉取国家列表与服务价格
2. 按价格从低到高列出前 `heroSmsCountryTopN` 个国家
3. 等你输入序号 / ISO / HeroSMS 国家 ID
4. 如果该国家聚合库存较低（当前阈值 `< 20`），再列出运营商 / 聚合库存供你二次选择
5. 再用你选中的国家与运营商继续后面的注册流程

也可以跳过交互，直接指定国家：

```bash
node index.js 1 --country=US
```

如果当前不是交互终端，脚本会自动回退到 `phoneCountryCode` 对应国家。

### 2) 只跑已注册账号的 OAuth（`--phase2`）

```bash
node index.js --phase2
```

前提：`accounts.json` 内存在可恢复账号。  
启动后会按创建时间倒序列出候选账号（手机号 / 时间 / 状态 / 国家 / 姓名），让你交互式选择要继续绑定邮箱的那一条。

### 3) 只补最后一步 Token（`--phase3`）

```bash
node index.js --phase3
```

作用：只执行“邮箱登录 OAuth 并换取 token”这最后一步。  
适合前面的手机号注册、邮箱绑定都已经成功，但最后 token 没拿到时单独补跑。

启动后会按时间倒序列出 `username.json` 里的候选记录（邮箱 / 手机号 / 时间 / 状态），让你交互式选择要补 token 的那一条。

它会：

1. 读取 `username.json` 最新一条记录
2. 用该记录里的邮箱重新走邮箱 OTP / OAuth
3. 换取并保存 token

它不会：

- 重新注册手机号
- 重新绑定邮箱
- 重跑完整三阶段流程

### 4) 按 `username.json` 批量补 Token（`--phase8`）

```bash
node index.js --phase8
```

用于对已有邮箱账号进行邮箱 OTP 登录并批量换 token。失败记录会追加到 `shibai.json`。

## 实际执行流程（默认模式）

1. Phase 1：手机号注册 ChatGPT（含 SMS 验证）
2. Phase 1.5：首次登录补全 about-you
3. Phase 2：绑定临时邮箱
4. Phase 3：邮箱登录 OAuth，换取并保存 token

## 产物文件

| 文件 | 作用 |
|---|---|
| `accounts.json` | 手机号注册账号池，含状态（`registered` / `oauth_done`） |
| `username.json` | 已绑定邮箱账号池（供 phase8 使用） |
| `tokens/codex-<email>-free.json` | 最终 token 文件 |
| `shibai.json` | phase8 失败账号记录 |

## 常见问题

### 1. 启动就报配置解析失败

`config.json` 含注释或格式错误。请用严格 JSON。

### 2. 提示未配置 `heroSmsApiKey`

`config.json` 未填写或读取失败。

### 3. 代理失败

脚本会检测代理连接错误并自动回退直连重试当前轮。

### 4. `--phase2` 报没有可用账号

先跑 `node index.js 1` 生成 `accounts.json` 记录。

### 5. cloud-mail 提示认证失效 / 401

优先检查：
1. `mailAdminEmail` / `mailAdminPassword` 是否正确
2. `mailAdminToken` 是否过期（若填写了）
3. `mailDomain` 是否为可用域名（例如 `raven97.xyz`，不要写错）

## 备注

当前代码中 `oauthClientId/oauthRedirectPort/useChrome/chromePath` 虽有配置导出，但主流程未使用这些配置值，不建议依赖它们。
