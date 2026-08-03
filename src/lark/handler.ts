import { getConfigAsync } from "../config.ts";
import {
  claimMailboxForLark,
  createConnectUrl,
  disconnectMailbox,
  listMailboxes,
  processQueuedSyncs,
  queueMailboxSyncByMailboxRef,
  sendTestNotification,
  updateMailboxProvider,
  updateMailboxRoute,
} from "../mail/service.ts";
import type { WaitUntilLike } from "../runtime.ts";
import { runBackground } from "../runtime.ts";
import { postLarkCard } from "./api.ts";
import {
  getLarkUrlVerificationChallenge,
  getLarkVerificationToken,
  type LarkTextMessageEvent,
  parseLarkMailCommand,
  parseLarkPayload,
  parseLarkTextMessageEvent,
} from "./parse.ts";
import {
  buildLarkCard,
  buildLarkConnectCard,
  buildLarkHelpCard,
  buildLarkMailboxListCard,
  buildLarkStatusCard,
  buildLarkWhoAmICard,
} from "./ui.ts";

function jsonResponse(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });
}

function providerLabel(providerType: "graph_native" | "ms_oauth2api"): string {
  return providerType === "ms_oauth2api" ? "msOauth2api" : "Graph Native";
}

function isLarkAdmin(openId: string, configuredAdmins: string[]): boolean {
  // 个人部署刚接入时允许先运行 `mail whoami` 取得 open_id；未配置管理员名单前，
  // 依赖飞书应用的可用范围控制。部署完成后应设置 LARK_ADMIN_OPEN_IDS 收紧权限。
  return configuredAdmins.length === 0 || configuredAdmins.includes(openId);
}

function commandRequiresGroup(kind: string): boolean {
  return ["connect", "claim", "route"].includes(kind);
}

async function replyError(chatId: string, error: unknown): Promise<void> {
  const message = error instanceof Error ? error.message : String(error);
  await postLarkCard({
    chatId,
    card: buildLarkCard({
      title: "邮件机器人操作失败",
      content: `错误：${message}`,
      template: "red",
    }),
  });
}

async function handleLarkCommand(event: LarkTextMessageEvent): Promise<void> {
  const command = parseLarkMailCommand(event.text);
  if (!command) return;

  const config = await getConfigAsync();
  if (command.kind === "whoami") {
    await postLarkCard({
      chatId: event.chatId,
      card: buildLarkWhoAmICard(event),
    });
    return;
  }

  if (!isLarkAdmin(event.senderOpenId, config.larkAdminOpenIds)) {
    await postLarkCard({
      chatId: event.chatId,
      card: buildLarkCard({
        title: "未授权的管理操作",
        content:
          "当前 open_id 不在 `LARK_ADMIN_OPEN_IDS` 中。先运行 `mail whoami`，再更新部署环境变量。",
        template: "orange",
      }),
    });
    return;
  }

  if (commandRequiresGroup(command.kind) && event.chatType !== "group") {
    await postLarkCard({
      chatId: event.chatId,
      card: buildLarkCard({
        title: "请在目标群中执行此命令",
        content:
          "为了让每个邮箱绑定独立 Lark 群，`connect`、`claim` 和 `route` 只能在群聊中执行。",
        template: "orange",
      }),
    });
    return;
  }

  try {
    switch (command.kind) {
      case "help":
        await postLarkCard({ chatId: event.chatId, card: buildLarkHelpCard() });
        return;
      case "connect": {
        const { authorizeUrl, providerType } = await createConnectUrl({
          teamId: event.tenantId,
          userId: event.senderOpenId,
          channelId: event.chatId,
          providerType: command.providerType,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkConnectCard(authorizeUrl, providerLabel(providerType)),
        });
        return;
      }
      case "claim": {
        const bundle = await claimMailboxForLark({
          teamId: event.tenantId,
          userId: event.senderOpenId,
          chatId: event.chatId,
          mailbox: command.mailbox,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "邮箱已迁移到 Lark",
            content:
              `**${bundle.connection.emailAddress}** 已认领并绑定到当前群。Microsoft OAuth 和同步进度已保留。`,
            template: "green",
          }),
        });
        return;
      }
      case "list": {
        const bundles = await listMailboxes(event.tenantId);
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkMailboxListCard(bundles),
        });
        return;
      }
      case "status": {
        const bundles = await listMailboxes(event.tenantId);
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkStatusCard(bundles),
        });
        return;
      }
      case "provider": {
        const bundle = await updateMailboxProvider({
          teamId: event.tenantId,
          mailbox: command.mailbox,
          providerType: command.providerType,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "同步后端已更新",
            content: `**${bundle.connection.emailAddress}** 已切换到 **${
              providerLabel(bundle.connection.providerType)
            }**。`,
            template: "green",
          }),
        });
        return;
      }
      case "route": {
        const bundle = await updateMailboxRoute({
          teamId: event.tenantId,
          mailbox: command.mailbox,
          chatId: event.chatId,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "Lark 路由已更新",
            content:
              `**${bundle.connection.emailAddress}** 的新邮件现在会投递到当前群。`,
            template: "green",
          }),
        });
        return;
      }
      case "sync": {
        const bundle = await queueMailboxSyncByMailboxRef({
          teamId: event.tenantId,
          mailbox: command.mailbox,
          reason: "lark_sync",
          requestedByUserId: event.senderOpenId,
        });
        await processQueuedSyncs();
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "同步完成",
            content:
              `已完成 **${bundle.connection.emailAddress}** 的同步队列处理。`,
            template: "green",
          }),
        });
        return;
      }
      case "test": {
        const bundle = await sendTestNotification({
          teamId: event.tenantId,
          mailbox: command.mailbox,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "测试通知已发送",
            content: `已向 **${
              bundle.route?.chatName || bundle.route?.chatId || "目标群"
            }** 发送 ${bundle.connection.emailAddress} 的测试邮件通知。`,
            template: "green",
          }),
        });
        return;
      }
      case "disconnect": {
        const bundle = await disconnectMailbox({
          teamId: event.tenantId,
          mailbox: command.mailbox,
        });
        await postLarkCard({
          chatId: event.chatId,
          card: buildLarkCard({
            title: "邮箱已断开",
            content:
              `已断开 **${bundle.connection.emailAddress}**，后续不再同步新邮件。`,
            template: "green",
          }),
        });
        return;
      }
    }
  } catch (error) {
    await replyError(event.chatId, error);
  }
}

export async function handleLarkEvent(
  request: Request,
  ctx?: WaitUntilLike,
): Promise<Response> {
  const bodyText = await request.text();
  let payload;
  try {
    payload = parseLarkPayload(bodyText || "{}");
  } catch (error) {
    return jsonResponse(
      {
        code: 400,
        msg: error instanceof Error ? error.message : "Invalid Lark payload",
      },
      400,
    );
  }

  const config = await getConfigAsync();
  const receivedToken = getLarkVerificationToken(payload);
  if (receivedToken !== config.larkVerificationToken) {
    return jsonResponse(
      { code: 401, msg: "Lark verification token mismatch" },
      401,
    );
  }

  const challenge = getLarkUrlVerificationChallenge(payload);
  if (challenge) return jsonResponse({ challenge });

  const event = parseLarkTextMessageEvent(payload);
  if (event) {
    runBackground(ctx, handleLarkCommand(event));
  }

  return jsonResponse({ code: 0 });
}
