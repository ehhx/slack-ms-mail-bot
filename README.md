# Lark Outlook Mail Bot

部署在 Deno Deploy 的多账号 Outlook 邮件机器人。每个 Outlook
邮箱可以绑定一个独立 Lark 群，同时监控 **Inbox + Junk**，并提供一个只读、多账号
Web 邮件工作台。

## 主要能力

- 使用 Lark 自建应用机器人接收管理命令和发送邮件卡片
- 集中连接多个个人 Outlook / Microsoft 365 账号
- 每个邮箱独立绑定一个 Lark 群
- Microsoft Graph Webhook + Delta Query 实时通知和补偿同步
- 可选使用 [msOauth2api](https://github.com/HChaoHui/msOauth2api) 轮询后端
- 支持从 Deno KV 迁移到共享 Supabase PostgreSQL，保存加密 refresh
  token、邮箱路由、订阅 lease、delta cursor 和去重记录
- Web 工作台支持多邮箱切换、Inbox/Junk、验证码提取、附件信息和 HTML 正文

## 项目结构

- `src/deploy.ts`：Deno Deploy 入口和定时维护任务
- `src/local.ts`：本地开发入口
- `src/handler.ts`：统一 HTTP 路由
- `src/lark/*`：Lark 事件、命令、卡片和 OpenAPI 客户端
- `src/microsoft/*`：Microsoft OAuth 与 Graph 客户端
- `src/mail/*`：邮箱同步、路由、通知和去重逻辑
- `src/store/*`：Deno KV 兼容仓储、Supabase 适配器和迁移工具
- `supabase/migrations/*`：Supabase 持久化状态表与 RPC
- `src/web/*`：只读 Web 邮件工作台

## HTTP 路由

- `GET /healthz`
- `POST /lark/events`
- `POST /graph/webhook`
- `GET /oauth/microsoft/callback`
- `GET /app`
- `GET|POST /app/login`
- `POST /app/logout`
- `GET /api/mailboxes`
- `GET /api/mailboxes/:mailboxId/messages?folder=inbox|junk`
- `GET /api/mailboxes/:mailboxId/messages/:messageId?folder=inbox|junk`
- `POST /admin/migrate-deno-kv`（迁移期间临时启用，需 `x-kv-migration-token`）

## 环境变量

### 必填

```text
LARK_APP_ID
LARK_APP_SECRET
LARK_VERIFICATION_TOKEN
APP_BASE_URL
MICROSOFT_CLIENT_ID
MICROSOFT_CLIENT_SECRET
MICROSOFT_REDIRECT_URI
TOKEN_ENCRYPTION_KEY
```

### 推荐配置

```text
LARK_API_BASE_URL=https://open.larksuite.com/open-apis
LARK_API_TIMEOUT_MS=15000
LARK_ADMIN_OPEN_IDS=ou_xxx,ou_yyy
MICROSOFT_AUTH_TENANT=common
MAIL_PREVIEW_MAX_CHARS=800
GRAPH_SUBSCRIPTION_RENEWAL_WINDOW_MINUTES=180
GRAPH_SUBSCRIPTION_MAX_MINUTES=4230
MAIL_SYNC_POLL_INTERVAL_MINUTES=15
MAIL_PROVIDER_DEFAULT=graph_native
GRAPH_WEBHOOK_CLIENT_STATE=<随机密钥>
WEB_ADMIN_PASSWORD=<Web 工作台密码>
WEB_SESSION_SECRET=<独立的 Cookie 签名密钥>
```

### Supabase 持久化配置

```text
PERSISTENCE_BACKEND=deno_kv|dual_write|supabase
SUPABASE_URL=https://<共享 Supabase 项目 REST API 地址>
SUPABASE_SERVICE_ROLE_KEY=<仅服务端保存的密钥>
SUPABASE_REQUEST_TIMEOUT_MS=10000
LEGACY_DENO_KV_MIGRATION_TOKEN=<一次性高强度随机值>
LEGACY_DENO_KV_MIGRATION_BATCH_SIZE=100
```

默认值 `PERSISTENCE_BACKEND=deno_kv` 保留旧行为。迁移时先使用
`dual_write`，应用会从旧 Deno KV 读取，并按“先 Supabase、后 Deno KV”的顺序双写；
迁移完成并验证后再切换为 `supabase`。`SUPABASE_SERVICE_ROLE_KEY` 和迁移 token
不得放入前端、仓库或聊天记录。

如果使用中国版飞书，将 API 地址改为：

```text
LARK_API_BASE_URL=https://open.feishu.cn/open-apis
```

如果使用 `ms_oauth2api`，再增加：

```text
MSOAUTH2API_BASE_URL=https://your-ms-oauth2api.example.com
MSOAUTH2API_PASSWORD=<可选共享密码>
MSOAUTH2API_MAILBOX=INBOX,Junk
```

生成安全密钥示例：

```bash
openssl rand -base64 32
```

## 一、创建 Lark 自建应用

1. 打开
   [Lark Developer Console](https://open.larksuite.com/app)，创建企业自建应用。
2. 在应用能力中添加“机器人”。
3. 为应用开启机器人发送消息、接收群聊中提及机器人的消息以及接收单聊消息所需权限。
4. 在“事件订阅”中选择“将事件发送至开发者服务器”。
5. 添加事件：

   ```text
   im.message.receive_v1
   ```

6. 请求地址填写：

   ```text
   https://<你的 Deno 域名>/lark/events
   ```

7. 将事件订阅页面的 Verification Token 保存为 `LARK_VERIFICATION_TOKEN`。
8. 当前版本不处理加密事件，**Encrypt Key 必须留空**。
9. 发布应用版本，并把应用可用范围限制为实际管理员或使用成员。

> Lark 后台必须能成功访问 `/lark/events` 后才会通过 URL 校验。因此可以先在 Deno
> Deploy 配置 App ID、App Secret 和 Verification Token，再回来保存事件地址。

## 二、配置 Microsoft 应用

在 Azure Portal / Microsoft Entra ID 中创建应用注册。

### 支持的账号类型

个人 Outlook 账号需要选择：

```text
Accounts in any organizational directory and personal Microsoft accounts
```

### Redirect URI

平台选择 Web，填写：

```text
https://<你的 Deno 域名>/oauth/microsoft/callback
```

### Delegated permissions

```text
offline_access
Mail.Read
User.Read
```

在“证书和密码”中创建 Client Secret，并保存其 **Value**。环境变量对应关系：

```text
Application (client) ID -> MICROSOFT_CLIENT_ID
Client secret Value     -> MICROSOFT_CLIENT_SECRET
Redirect URI            -> MICROSOFT_REDIRECT_URI
```

如果之前 Slack 版本已经使用同一个 Microsoft App，Microsoft Redirect URI
和权限不需要因为迁移到 Lark 而修改；只有 Deno 域名变化时才需要更新 Redirect
URI。

## 三、部署到 Deno Deploy

1. 将 GitHub 仓库连接到 Deno Deploy。
2. 入口文件设置为：

   ```text
   src/deploy.ts
   ```

3. 创建 Deno KV
   数据库并绑定到当前生产部署。迁移完成前不要解绑，它是旧数据源和回滚依据。
4. 在项目环境变量中填写前述必填配置；首次部署仍使用
   `PERSISTENCE_BACKEND=deno_kv`。
5. `APP_BASE_URL` 必须是生产域名，不带末尾斜杠：

   ```text
   APP_BASE_URL=https://your-project.deno.dev
   ```

6. `MICROSOFT_REDIRECT_URI` 必须和 Azure 中完全一致：

   ```text
   MICROSOFT_REDIRECT_URI=https://your-project.deno.dev/oauth/microsoft/callback
   ```

7. 部署完成后访问：

   ```text
   https://your-project.deno.dev/healthz
   ```

   正常响应应为 `ok`。

8. 返回 Lark 后台配置 `/lark/events`，完成事件地址校验并发布应用。

## 四、迁移到共享 Supabase

### 1. 目标项目和区域

当前优先共用现有阿里云 AnalyticDB Supabase 新加坡项目：

```text
Project Ref: spb-krh6cml28cs1hvc9
Region: ap-southeast-1
```

选择理由：`mail_persistent_state` / `mail_state_*` 与现有 bot 项目的表和 RPC
不冲突； 共用同一新加坡项目可以减少新增项目成本，
也避免因为另选区域引入额外跨区延迟。只有当共享项目的权限隔离或容量策略不能接受时，
再选择其他区域的新项目。

共享项目的主要限制是 `service_role` 是项目级高权限密钥。两个服务都使用
`service_role` 时，服务端理论上可以访问同一项目内的其他表。当前迁移通过表名/RPC
命名隔离和 RLS
限制公开访问，但这不是强安全边界；如果后续需要严格隔离，应改用独立 Supabase
项目，或进一步设计专用数据库角色/JWT 权限。

### 2. 同步共享项目 migration history

共享 Supabase 的 migration history 是项目级的。这个项目已经存在
`20260720101645_create_persistent_state_store.sql`，本仓库当前只有
`20260723110555_create_mail_persistent_state.sql`。因此不能直接
`db push`，必须先把远端 history 同步到本仓库，再推送 `mail` 的新 migration。

先在本机 shell 中设置凭据，不要写入仓库或聊天记录：

```bash
export ALIYUN_ACCESS_KEY_ID=<你的阿里云 AccessKeyId>
export ALIYUN_ACCESS_KEY_SECRET=<你的阿里云 AccessKeySecret>
read -s SUPABASE_DB_PASSWORD
export SUPABASE_DB_PASSWORD
```

然后连接共享项目。阿里云 CLI 默认地域是
`cn-hangzhou`，这个项目在新加坡，所以远程命令必须带
`--aliyun-region ap-southeast-1`：

```bash
supabase link \
  --project-ref spb-krh6cml28cs1hvc9 \
  --aliyun-region ap-southeast-1 \
  -p "$SUPABASE_DB_PASSWORD"

supabase migration list \
  --project-ref spb-krh6cml28cs1hvc9 \
  --aliyun-region ap-southeast-1 \
  -p "$SUPABASE_DB_PASSWORD"

supabase migration fetch --linked
```

如果 `migration fetch --linked` 不能从阿里云 linked project
读取历史，则改为通过数据库连接串执行
`supabase migration fetch --db-url <percent-encoded DSN>`。不要从其他项目手工复制
migration 文件，除非已经确认远端 history 只有对应版本且 SQL 完全一致。

同步完成后再次执行：

```bash
supabase migration list \
  --project-ref spb-krh6cml28cs1hvc9 \
  --aliyun-region ap-southeast-1 \
  -p "$SUPABASE_DB_PASSWORD"
```

确认本地和远端已共同包含旧版本
`20260720101645_create_persistent_state_store.sql`，且只有
`20260723110555_create_mail_persistent_state.sql` 处于待推送状态。

### 3. 应用 mail Schema

当前仓库的 `supabase/migrations/20260723110555_create_mail_persistent_state.sql`
会创建 `mail_persistent_state`、RLS、仅 `service_role` 可用的 RPC
和过期清理索引。先 dry-run，确认只会应用 `mail` 这一条 migration：

如果 Deno 服务调用 RPC 时返回 404，先在 Supabase Dashboard 的 Data API
设置里确认 `public` schema 已暴露；不要为了排查 404 给 `anon` 或 `authenticated`
授权。

```bash
supabase db push \
  --project-ref spb-krh6cml28cs1hvc9 \
  --aliyun-region ap-southeast-1 \
  -p "$SUPABASE_DB_PASSWORD" \
  --dry-run
```

确认无误后正式推送：

```bash
supabase db push \
  --project-ref spb-krh6cml28cs1hvc9 \
  --aliyun-region ap-southeast-1 \
  -p "$SUPABASE_DB_PASSWORD"
```

如使用 Supabase.com 或其他非阿里云项目，不要添加
`--aliyun-region`；命令应改用对应平台的 project ref 和 CLI profile。

### 4. 部署双写版本

在 Deno Deploy 配置：

```text
PERSISTENCE_BACKEND=dual_write
SUPABASE_URL=https://<共享 Supabase 项目 REST API 地址>
SUPABASE_SERVICE_ROLE_KEY=<服务端密钥>
LEGACY_DENO_KV_MIGRATION_TOKEN=<随机值>
```

重新部署后，先确认：

```bash
curl -fsS https://<你的 Deno 域名>/healthz
```

应返回 `{"ok":true,"persistence":"dual_write"}`。

### 5. 按前缀分页导入旧 KV

以下 10 个前缀都必须执行，且每个前缀都要持续提交上一次响应的
`nextCursor`，直到返回 `null`。批次刚好填满时，最后仍可能返回一个游标，
需要再请求一次空页才能结束。

```text
oauth_state
mailbox_connection
mailbox_email
team_mailbox
mailbox_route
mailbox_sync
mailbox_lease
subscription_mailbox
delivered_mail
sync_queue
```

请求示例（不要把 token 写入仓库）：

```bash
curl -fsS -X POST https://<你的 Deno 域名>/admin/migrate-deno-kv \
  -H "x-kv-migration-token: $LEGACY_DENO_KV_MIGRATION_TOKEN" \
  -H 'content-type: application/json' \
  -d '{"kind":"mailbox_connection"}'
```

接口使用 `setIfAbsent`，可安全重试且不会覆盖双写期间已写入 Supabase 的新状态。
`oauth_state` 会保留原 `expiresAt`，`delivered_mail` 按 `deliveredAt + 90 天`
设置过期时间； 旧 Deno KV 不会被删除。

### 6. 切换并验证

所有前缀导入完成后，将 Deno Deploy 的 `PERSISTENCE_BACKEND` 改为 `supabase`
并重新部署：

```bash
curl -fsS https://<你的 Deno 域名>/healthz
```

应返回 `{"ok":true,"persistence":"supabase"}`。随后至少验证一次 Lark 事件、Graph
Webhook、邮箱租约、同步队列和邮件去重；确认无误后删除
`LEGACY_DENO_KV_MIGRATION_TOKEN`。

回滚时保留旧 Deno KV，不要直接删除数据库。切回 `deno_kv`
只能恢复切换前已同步到旧 KV 的状态；切换到 `supabase`
后产生的新写入需要先评估，不能宣称跨存储零丢失回滚。

## 五、初始化 Lark 管理员

第一次部署时可以暂时不填写
`LARK_ADMIN_OPEN_IDS`，但必须通过应用可用范围限制使用者。

1. 把机器人加入一个测试群。
2. 在群中发送：

   ```text
   @机器人 mail whoami
   ```

3. 复制机器人返回的 `open_id`。
4. 写入 Deno Deploy：

   ```text
   LARK_ADMIN_OPEN_IDS=ou_xxxxxxxxx
   ```

5. 重新部署。此后只有名单中的用户可以执行管理命令。

## 六、连接和迁移邮箱

### 连接新邮箱

在希望接收该邮箱通知的 Lark 群里发送：

```text
@机器人 mail connect graph
```

点击卡片中的 Microsoft 授权按钮。看到 `Mailbox connected` 后返回
Lark，该邮箱已经绑定当前群。

### 迁移 Slack 版本中的已有邮箱

旧 Outlook refresh token、Graph subscription 和 delta cursor 都保存在原 Deno KV
中，不需要重新授权。在目标 Lark 群发送：

```text
@机器人 mail claim your-account@outlook.com
```

每个邮箱都需要在对应目标群执行一次 `claim`。完成前，旧 Slack
路由不会继续发送通知。

## Lark 命令

```text
mail help
mail whoami
mail connect [graph|msoauth2api]
mail claim <mailbox>
mail list
mail status
mail provider <mailbox> <graph|msoauth2api>
mail route <mailbox>
mail test <mailbox>
mail sync <mailbox>
mail disconnect <mailbox>
```

`<mailbox>` 可以使用完整邮箱地址或邮箱 ID 前缀。`connect`、`claim` 和 `route`
必须在群聊中执行，它们会把邮箱绑定到当前群。

## Web 邮件工作台

配置以下变量后启用：

```text
WEB_ADMIN_PASSWORD=<登录密码>
WEB_SESSION_SECRET=<Cookie 签名密钥>
```

访问：

```text
https://<你的域名>/app
```

当前 Web 工作台只读，并且只支持 `graph_native` 邮箱。页面采用 Deno 原生轻量
SPA：

- 首屏不再等待 Microsoft Graph，页面壳层会立即显示
- CSS 与客户端脚本单独缓存
- 列表、正文和 access token 都有短时内存缓存
- 快速切换时会取消旧请求，并在空闲时预取下一封邮件
- 桌面端列表和正文拥有独立滚动条
- 移动端使用列表/正文单列切换，不压缩双栏
- HTML 邮件在无脚本、无同源权限的 iframe 中打开

当前不支持回复、转发、删除、标记已读以及 Web 端修改邮箱路由。

## 本地开发

配置环境变量后运行：

```bash
deno task dev
```

默认地址为 `http://localhost:8000`。如需验证 Lark 和 Graph 回调，使用 Cloudflare
Tunnel 或 ngrok 暴露 HTTPS 地址，并同步修改 `APP_BASE_URL` 和 Microsoft Redirect
URI。

## 验证

```bash
deno fmt --check
deno lint
deno check src/deploy.ts
deno test --unstable-kv --allow-env --allow-read --allow-write src
```

本机没有全局 Deno 时，可以使用：

```bash
env npm_config_cache=/tmp/ehhx-npm-cache npx --yes deno check src/deploy.ts
env npm_config_cache=/tmp/ehhx-npm-cache npx --yes deno test --unstable-kv --allow-env --allow-read --allow-write src
```
