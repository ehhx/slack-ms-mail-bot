# Slack Outlook Mail Bot (Deno Deploy)

一个部署在 **Deno Deploy** 的 Slack 机器人：连接多个 Outlook / Microsoft 365 账号，把这些账号 Inbox 的新邮件推送到 Slack 指定频道。

当前支持两种同步后端：

- `graph_native`：直接使用 Microsoft Graph Webhook + Delta Query
- `ms_oauth2api`：由本机器人服务端调用 [msOauth2api](https://github.com/HChaoHui/msOauth2api) 接口做轮询同步

## 功能

- 集中托管多个 Outlook 账号
- 每个邮箱绑定一个 Slack 频道
- 支持 **Microsoft Graph 原生同步** 或 **msOauth2api 轮询同步**
- 使用 Slack Slash Command 管理邮箱、路由、状态、手动同步
- 使用 Deno KV 存储 OAuth 状态、邮箱路由、订阅 lease、delta link、去重记录

## 目录结构

- `src/local.ts`：本地开发入口
- `src/deploy.ts`：Deno Deploy 入口 + 定时维护任务
- `src/handler.ts`：统一 HTTP 路由
- `src/slack/*`：Slash command、交互、签名校验、Slack API
- `src/microsoft/*`：Microsoft OAuth + Graph 封装
- `src/mail/*`：邮箱领域模型、同步编排、通知格式
- `src/store/*`：Deno KV 仓储

## 环境变量

必填：

- `SLACK_SIGNING_SECRET`
- `SLACK_BOT_TOKEN`
- `APP_BASE_URL`
- `MICROSOFT_CLIENT_ID`
- `MICROSOFT_CLIENT_SECRET`
- `MICROSOFT_REDIRECT_URI`
- `TOKEN_ENCRYPTION_KEY`

推荐：

- `MICROSOFT_AUTH_TENANT=common`
- `MAIL_PREVIEW_MAX_CHARS=220`
- `GRAPH_SUBSCRIPTION_RENEWAL_WINDOW_MINUTES=180`
- `GRAPH_SUBSCRIPTION_MAX_MINUTES=4230`
- `MAIL_SYNC_POLL_INTERVAL_MINUTES=15`
- `MAIL_PROVIDER_DEFAULT=graph_native`
- `MSOAUTH2API_BASE_URL=https://your-ms-oauth2api.example.com`
- `MSOAUTH2API_PASSWORD=<optional shared password>`
- `MSOAUTH2API_MAILBOX=INBOX`
- `SLACK_API_TIMEOUT_MS=15000`
- `GRAPH_WEBHOOK_CLIENT_STATE=<custom secret>`
- `KV_PATH=<local kv sqlite path>`

示例：

```bash
export SLACK_SIGNING_SECRET="..."
export SLACK_BOT_TOKEN="xoxb-..."
export APP_BASE_URL="https://your-app.example.com"
export MICROSOFT_CLIENT_ID="..."
export MICROSOFT_CLIENT_SECRET="..."
export MICROSOFT_REDIRECT_URI="https://your-app.example.com/oauth/microsoft/callback"
export TOKEN_ENCRYPTION_KEY="replace-me"
export MAIL_PROVIDER_DEFAULT="graph_native"
export MSOAUTH2API_BASE_URL="https://your-ms-oauth2api.example.com"
export MSOAUTH2API_PASSWORD=""
export MSOAUTH2API_MAILBOX="INBOX"
```

## Slack App 配置

在 [Slack API](https://api.slack.com/apps) 创建 App 后：

### Slash Command

- Command: `/mail`
- Request URL:
  - `https://<your-domain>/slack/commands`

### Interactivity

开启 **Interactivity & Shortcuts**：

- Request URL:
  - `https://<your-domain>/slack/interactivity`

### OAuth Scopes

- `commands`
- `chat:write`
- `chat:write.public`（如果要向机器人未加入的公共频道投递，按需开启）

## Microsoft App 配置

在 Azure Portal / Microsoft Entra ID 创建应用：

### Redirect URI

- `https://<your-domain>/oauth/microsoft/callback`

### Delegated permissions

- `offline_access`
- `Mail.Read`
- `User.Read`

### Webhook endpoint（仅 `graph_native` provider 使用）

机器人会使用：

- `POST https://<your-domain>/graph/webhook`

作为 Microsoft Graph 通知地址与生命周期通知地址。

## HTTP 路由

- `GET /healthz`
- `POST /slack/commands`
- `POST /slack/interactivity`
- `POST /graph/webhook`
- `GET /oauth/microsoft/callback`

## Slack 命令

- `/mail help`
- `/mail connect [graph|msoauth2api]`
- `/mail list`
- `/mail status`
- `/mail provider <mailbox> <graph|msoauth2api>`
- `/mail route <mailbox> <#channel>`
- `/mail test <mailbox>`
- `/mail sync <mailbox>`
- `/mail disconnect <mailbox>`

其中 `<mailbox>` 可以用：

- 邮箱地址，如 `ops@example.com`
- 邮箱 ID 前缀

## provider 行为说明

### `graph_native`

- OAuth 完成后立即建立 Graph delta 基线
- 自动创建 Graph subscription
- 通过 `/graph/webhook` 接收变更通知
- 仍会由维护任务做补偿同步与续租

### `ms_oauth2api`

- OAuth 完成后只保存 refresh token，并通过服务端调用 `msOauth2api /api/mail-all`
- 首次连接/切换到该 provider 时会先建立“历史邮件基线”，避免把旧邮件一次性推到 Slack
- 不依赖 `/graph/webhook`
- 主要依赖维护轮询与 `/mail sync <mailbox>` 手动补偿

## msOauth2api 部署要求

如果你选择 `ms_oauth2api`，需要额外部署一个 `msOauth2api` 服务，并把它的公网地址配置到：

- `MSOAUTH2API_BASE_URL`

如果 `msOauth2api` 开启了共享密码校验，再配置：

- `MSOAUTH2API_PASSWORD`

## 本地开发

1. 安装 Deno
2. 配置环境变量
3. 启动：

```bash
deno task dev
```

默认监听 `http://localhost:8000`。

建议用 ngrok / Cloudflare Tunnel 暴露公网地址，分别配置到 Slack 与 Microsoft Graph。

## 部署到 Deno Deploy

- 入口文件：`src/deploy.ts`
- 配置好所有环境变量
- 确保公网 URL 与 `APP_BASE_URL`、`MICROSOFT_REDIRECT_URI` 一致

## 测试

```bash
deno task test
```

> 当前仓库环境里未检测到 `deno` 命令，因此本项目已提供测试文件，但尚未在本地运行验证。
