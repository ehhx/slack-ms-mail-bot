import type { MailMessageSummary, MailboxBundle } from "../mail/types.ts";
import { formatMailboxRef, mailboxStatusLine, toPreviewText } from "../mail/message.ts";

function fmtTime(iso: string | undefined): string {
  if (!iso) return "-";
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return iso;
  return parsed.toLocaleString("zh-CN", { hour12: false });
}

function providerLabel(bundle: MailboxBundle): string {
  return bundle.connection.providerType === "ms_oauth2api" ? "msOauth2api" : "Graph Native";
}

export function buildHelpBlocks(): unknown[] {
  const lines = [
    "*用法*",
    "`/mail connect [graph|msoauth2api]` 连接一个 Outlook 账号",
    "`/mail list` 查看已连接邮箱与路由",
    "`/mail status` 查看授权/订阅/同步状态",
    "`/mail provider <mailbox> <graph|msoauth2api>` 切换同步 provider",
    "`/mail route <mailbox> <#channel>` 修改默认 Slack 频道",
    "`/mail test <mailbox>` 发送测试通知",
    "`/mail sync <mailbox>` 手动触发补偿同步",
    "`/mail disconnect <mailbox>` 断开邮箱连接",
  ];

  return [
    { type: "section", text: { type: "mrkdwn", text: lines.join("\n") } },
  ];
}

export function buildConnectBlocks(authorizeUrl: string, providerText: string): unknown[] {
  return [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `点击下面的按钮，使用 Microsoft 账号授权这个机器人访问你的 Outlook Inbox。当前将使用 *${providerText}* 作为邮件同步后端。授权成功后，默认会把该邮箱的新邮件投递到当前频道。`,
      },
      accessory: {
        type: "button",
        text: { type: "plain_text", text: "Connect Outlook" },
        url: authorizeUrl,
        action_id: "oauth_connect",
      },
    },
  ];
}

function buildMailboxActions(mailboxId: string): unknown[] {
  return [
    {
      type: "button",
      text: { type: "plain_text", text: "Sync" },
      action_id: "mail_sync",
      value: mailboxId,
    },
    {
      type: "button",
      text: { type: "plain_text", text: "Test" },
      action_id: "mail_test",
      value: mailboxId,
    },
    {
      type: "button",
      text: { type: "plain_text", text: "Disconnect" },
      style: "danger",
      action_id: "mail_disconnect",
      value: mailboxId,
      confirm: {
        title: { type: "plain_text", text: "Disconnect mailbox?" },
        text: { type: "mrkdwn", text: "断开后将停止同步此邮箱的新邮件。" },
        confirm: { type: "plain_text", text: "Disconnect" },
        deny: { type: "plain_text", text: "Cancel" },
      },
    },
  ];
}

export function buildMailboxListBlocks(bundles: MailboxBundle[]): unknown[] {
  const blocks: unknown[] = [
    { type: "header", text: { type: "plain_text", text: "已连接邮箱" } },
  ];

  if (bundles.length === 0) {
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: "当前还没有连接任何 Outlook 邮箱。先运行 `/mail connect [graph|msoauth2api]`。",
      },
    });
    return blocks;
  }

  for (const bundle of bundles) {
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text:
          `*${bundle.connection.displayName || bundle.connection.emailAddress}*\n邮箱：\`${bundle.connection.emailAddress}\`\nID：\`${formatMailboxRef(bundle.connection.mailboxId)}\`\nProvider：\`${providerLabel(bundle)}\`\n${mailboxStatusLine(bundle)}`,
      },
    });
    blocks.push({ type: "actions", elements: buildMailboxActions(bundle.connection.mailboxId) });
    blocks.push({ type: "divider" });
  }

  blocks.push({
    type: "actions",
    elements: [{
      type: "button",
      text: { type: "plain_text", text: "Refresh" },
      action_id: "mail_refresh",
      value: "refresh",
    }],
  });

  return blocks;
}

export function buildStatusBlocks(bundles: MailboxBundle[]): unknown[] {
  const blocks: unknown[] = [
    { type: "header", text: { type: "plain_text", text: "邮箱状态" } },
  ];

  if (bundles.length === 0) {
    blocks.push({
      type: "section",
      text: { type: "mrkdwn", text: "没有可显示的邮箱状态。" },
    });
    return blocks;
  }

  for (const bundle of bundles) {
    const pollingOnly = bundle.connection.providerType === "ms_oauth2api";
    blocks.push({
      type: "section",
      fields: [
        { type: "mrkdwn", text: `*Mailbox*\n${bundle.connection.emailAddress}` },
        { type: "mrkdwn", text: `*Provider*\n${providerLabel(bundle)}` },
        { type: "mrkdwn", text: `*Route*\n${bundle.route ? `<#${bundle.route.slackChannelId}>` : "未配置"}` },
        { type: "mrkdwn", text: `*Connection*\n${bundle.connection.status}` },
        {
          type: "mrkdwn",
          text: `*Subscription*\n${pollingOnly ? "polling-only" : (bundle.lease?.status ?? "missing")}`,
        },
        {
          type: "mrkdwn",
          text: `*Lease expires*\n${pollingOnly ? "不适用" : fmtTime(bundle.lease?.expiresAt)}`,
        },
        { type: "mrkdwn", text: `*Last sync*\n${fmtTime(bundle.syncState?.lastSyncAt)}` },
        { type: "mrkdwn", text: `*Last notification*\n${fmtTime(bundle.syncState?.lastNotificationAt)}` },
      ],
    });
    if (bundle.connection.lastError || bundle.syncState?.lastError || bundle.lease?.lastError) {
      blocks.push({
        type: "context",
        elements: [{
          type: "mrkdwn",
          text: `最近错误：${bundle.connection.lastError ?? bundle.syncState?.lastError ?? bundle.lease?.lastError}`,
        }],
      });
    }
    blocks.push({ type: "divider" });
  }

  return blocks;
}

export function buildMailNotificationBlocks(
  mailbox: MailboxBundle,
  message: MailMessageSummary,
  maxPreviewChars: number,
): unknown[] {
  const preview = toPreviewText(message.bodyPreview, maxPreviewChars);
  const sender = message.fromName || message.fromAddress || "Unknown sender";
  const elements: unknown[] = [];
  if (message.webLink) {
    elements.push({
      type: "button",
      text: { type: "plain_text", text: "Open in Outlook" },
      url: message.webLink,
      action_id: "open_outlook",
    });
  }

  return [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*📬 新邮件：${message.subject || "(no subject)"}*\n*发件人*：${sender}\n*邮箱*：${mailbox.connection.emailAddress}\n*Provider*：${providerLabel(mailbox)}\n*时间*：${fmtTime(message.receivedDateTime)}`,
      },
    },
    {
      type: "section",
      text: { type: "mrkdwn", text: preview },
    },
    ...(elements.length ? [{ type: "actions", elements }] : []),
  ];
}
