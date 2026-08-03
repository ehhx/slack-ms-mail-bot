import type { MailboxBundle, MailMessageSummary } from "../mail/types.ts";
import {
  attachmentSummaryText,
  formatFolderLabel,
  formatMailboxRef,
  monitoredFoldersText,
  notificationBodyText,
  toPreviewText,
} from "../mail/message.ts";

type LarkCard = Record<string, unknown>;

function escapeMarkdown(input: string | undefined): string {
  return (input ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/([`*_{}\[\]<>])/g, "\\$1")
    .replace(/\r/g, "")
    .trim();
}

function fmtTime(iso: string | undefined): string {
  if (!iso) return "-";
  const parsed = new Date(iso);
  return Number.isNaN(parsed.getTime())
    ? iso
    : parsed.toLocaleString("zh-CN", { hour12: false });
}

function providerLabel(bundle: MailboxBundle): string {
  return bundle.connection.providerType === "ms_oauth2api"
    ? "msOauth2api"
    : "Graph Native";
}

function routeLabel(bundle: MailboxBundle): string {
  if (!bundle.route) return "未配置";
  if (bundle.route.platform !== "lark") return "等待迁移到 Lark";
  return bundle.route.chatName || bundle.route.chatId;
}

export function buildLarkCard(input: {
  title: string;
  content: string;
  template?: "blue" | "green" | "orange" | "red" | "grey";
  actions?: Array<Record<string, unknown>>;
}): LarkCard {
  const elements: Array<Record<string, unknown>> = [
    {
      tag: "div",
      text: { tag: "lark_md", content: input.content },
    },
  ];
  if (input.actions?.length) {
    elements.push({ tag: "hr" });
    elements.push({ tag: "action", actions: input.actions });
  }
  return {
    config: { wide_screen_mode: true },
    header: {
      template: input.template ?? "blue",
      title: { tag: "plain_text", content: input.title },
    },
    elements,
  };
}

export function buildLarkHelpCard(): LarkCard {
  return buildLarkCard({
    title: "Outlook 邮件机器人",
    content: [
      "在群里发送以下命令：",
      "`mail connect [graph|msoauth2api]` 连接新的 Outlook 账号并绑定当前群",
      "`mail claim <mailbox>` 认领原 Slack 连接的邮箱并绑定当前群",
      "`mail list` 查看已连接邮箱",
      "`mail status` 查看授权、订阅和同步状态",
      "`mail provider <mailbox> <graph|msoauth2api>` 切换同步后端",
      "`mail route <mailbox>` 将邮箱投递到当前群",
      "`mail test <mailbox>` 发送测试通知",
      "`mail sync <mailbox>` 手动补偿同步",
      "`mail disconnect <mailbox>` 断开邮箱",
      "`mail whoami` 查看当前群、租户和 open_id，用于管理员配置",
    ].join("\n"),
  });
}

export function buildLarkConnectCard(
  authorizeUrl: string,
  providerLabelText: string,
): LarkCard {
  return buildLarkCard({
    title: "连接 Outlook 邮箱",
    content:
      `使用 Microsoft 账号授权机器人读取 **Inbox + Junk**。当前同步后端：**${
        escapeMarkdown(providerLabelText)
      }**。授权完成后，新邮件会投递到当前 Lark 群。`,
    actions: [{
      tag: "button",
      type: "primary",
      text: { tag: "plain_text", content: "连接 Outlook" },
      url: authorizeUrl,
    }],
  });
}

export function buildLarkMailboxListCard(bundles: MailboxBundle[]): LarkCard {
  if (bundles.length === 0) {
    return buildLarkCard({
      title: "已连接邮箱",
      content:
        "当前租户还没有已连接邮箱。先发送 `mail connect graph`，或对旧邮箱使用 `mail claim <mailbox>`。",
      template: "grey",
    });
  }

  const content = bundles.map((bundle) =>
    [
      `**${
        escapeMarkdown(
          bundle.connection.displayName || bundle.connection.emailAddress,
        )
      }**`,
      `邮箱：\`${escapeMarkdown(bundle.connection.emailAddress)}\``,
      `ID：\`${formatMailboxRef(bundle.connection.mailboxId)}\``,
      `同步：${providerLabel(bundle)} · ${monitoredFoldersText(bundle)}`,
      `投递：${escapeMarkdown(routeLabel(bundle))}`,
    ].join("\n")
  ).join("\n\n---\n\n");
  return buildLarkCard({ title: "已连接邮箱", content });
}

export function buildLarkStatusCard(bundles: MailboxBundle[]): LarkCard {
  if (bundles.length === 0) {
    return buildLarkCard({
      title: "邮箱状态",
      content: "没有可显示的邮箱状态。",
      template: "grey",
    });
  }

  const content = bundles.map((bundle) => {
    const pollingOnly = bundle.connection.providerType === "ms_oauth2api";
    return [
      `**${escapeMarkdown(bundle.connection.emailAddress)}**`,
      `路由：${escapeMarkdown(routeLabel(bundle))}`,
      `连接：${escapeMarkdown(bundle.connection.status)} · 同步：${
        fmtTime(bundle.syncState?.lastSyncAt)
      }`,
      `订阅：${
        pollingOnly
          ? "轮询模式"
          : escapeMarkdown(bundle.lease?.status ?? "missing")
      }`,
      `最近通知：${fmtTime(bundle.syncState?.lastNotificationAt)}`,
      ...(bundle.connection.lastError || bundle.syncState?.lastError ||
          bundle.lease?.lastError
        ? [
          `最近错误：${
            escapeMarkdown(
              bundle.connection.lastError ?? bundle.syncState?.lastError ??
                bundle.lease?.lastError,
            )
          }`,
        ]
        : []),
    ].join("\n");
  }).join("\n\n---\n\n");
  return buildLarkCard({ title: "邮箱状态", content });
}

export function buildLarkMailNotificationCard(
  mailbox: MailboxBundle,
  message: MailMessageSummary,
  maxPreviewChars: number,
): LarkCard {
  const subject = escapeMarkdown(message.subject || "(无主题)");
  const sender = escapeMarkdown(
    message.fromName || message.fromAddress || "未知发件人",
  );
  const preview = escapeMarkdown(toPreviewText(
    notificationBodyText(message),
    Math.max(maxPreviewChars, 800),
  ));
  const attachmentText = attachmentSummaryText(message.attachments);
  const content = [
    `**发件人**：${sender}`,
    `**邮箱**：${escapeMarkdown(mailbox.connection.emailAddress)}`,
    `**文件夹**：${
      escapeMarkdown(formatFolderLabel(message.folderKind, message.folderName))
    }`,
    `**时间**：${fmtTime(message.receivedDateTime)}`,
    "",
    preview || "(无可用正文)",
    ...(attachmentText ? ["", escapeMarkdown(attachmentText)] : []),
  ].join("\n");

  return buildLarkCard({
    title: `📬 ${subject}`,
    content,
    actions: message.webLink
      ? [{
        tag: "button",
        type: "primary",
        text: { tag: "plain_text", content: "在 Outlook 中打开" },
        url: message.webLink,
      }]
      : undefined,
  });
}

export function buildLarkWhoAmICard(input: {
  tenantId: string;
  senderOpenId: string;
  chatId: string;
  chatType: string;
}): LarkCard {
  return buildLarkCard({
    title: "Lark 当前上下文",
    content: [
      `租户：\`${escapeMarkdown(input.tenantId)}\``,
      `当前群：\`${escapeMarkdown(input.chatId)}\` (${
        escapeMarkdown(input.chatType)
      })`,
      `你的 open_id：\`${escapeMarkdown(input.senderOpenId)}\``,
    ].join("\n"),
    template: "grey",
  });
}
