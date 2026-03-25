import type { WaitUntilLike } from "../../runtime.ts";
import { runBackground } from "../../runtime.ts";
import {
  disconnectMailbox,
  enqueueMailboxSync,
  listMailboxes,
  processQueuedSyncs,
  sendTestNotification,
} from "../../mail/service.ts";
import { parseInteractivityPayload } from "../parse.ts";
import { jsonResponse, slackEphemeral } from "../respond.ts";
import { buildMailboxListBlocks } from "../ui.ts";
import { postEphemeralMessage } from "../api.ts";

export async function handleSlackInteractivity(
  bodyText: string,
  ctx?: WaitUntilLike,
): Promise<Response> {
  const payload = parseInteractivityPayload(bodyText);
  const action = payload.actions?.[0];
  if (!action) {
    return jsonResponse(slackEphemeral("未收到可执行动作。"));
  }

  switch (action.action_id) {
    case "mail_refresh": {
      const mailboxes = await listMailboxes(payload.team.id);
      return jsonResponse(slackEphemeral("已连接邮箱", buildMailboxListBlocks(mailboxes)));
    }
    case "mail_sync": {
      runBackground(ctx, (async () => {
        try {
          const mailboxId = action.value ?? "";
          if (!mailboxId) throw new Error("Missing mailbox id");
          await enqueueMailboxSync({
            mailboxId,
            reason: "interactive_sync",
            requestedByUserId: payload.user.id,
          });
          await processQueuedSyncs();
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              "✅ 已完成该邮箱的同步处理。",
            );
          }
        } catch (error) {
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              `❌ 同步失败：${error instanceof Error ? error.message : String(error)}`,
            );
          }
        }
      })());
      return jsonResponse(slackEphemeral("已开始同步。"));
    }
    case "mail_test": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await sendTestNotification({
            teamId: payload.team.id,
            mailbox: action.value ?? "",
          });
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              `✅ 已向 <#${bundle.route?.slackChannelId}> 发送测试通知。`,
            );
          }
        } catch (error) {
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              `❌ 测试通知失败：${error instanceof Error ? error.message : String(error)}`,
            );
          }
        }
      })());
      return jsonResponse(slackEphemeral("正在发送测试通知。"));
    }
    case "mail_disconnect": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await disconnectMailbox({
            teamId: payload.team.id,
            mailbox: action.value ?? "",
          });
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              `✅ 已断开邮箱 ${bundle.connection.emailAddress}`,
            );
          }
        } catch (error) {
          if (payload.channel?.id) {
            await postEphemeralMessage(
              payload.channel.id,
              payload.user.id,
              `❌ 断开失败：${error instanceof Error ? error.message : String(error)}`,
            );
          }
        }
      })());
      return jsonResponse(slackEphemeral("正在断开邮箱连接。"));
    }
    default:
      return jsonResponse(slackEphemeral(`未知动作：${action.action_id}`));
  }
}
