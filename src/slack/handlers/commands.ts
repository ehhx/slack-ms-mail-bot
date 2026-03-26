import { parseMailCommand, parseSlashCommand } from "../parse.ts";
import { buildConnectBlocks, buildHelpBlocks, buildMailboxListBlocks, buildStatusBlocks } from "../ui.ts";
import { jsonResponse, postToResponseUrl, slackEphemeral } from "../respond.ts";
import type { WaitUntilLike } from "../../runtime.ts";
import { runBackground } from "../../runtime.ts";
import {
  createConnectUrl,
  disconnectMailbox,
  listMailboxes,
  processQueuedSyncs,
  queueMailboxSyncByMailboxRef,
  sendTestNotification,
  updateMailboxProvider,
  updateMailboxRoute,
} from "../../mail/service.ts";

function providerLabel(providerType: "graph_native" | "ms_oauth2api"): string {
  return providerType === "ms_oauth2api" ? "msOauth2api" : "Graph Native";
}

export async function handleSlackCommands(
  bodyText: string,
  ctx?: WaitUntilLike,
): Promise<Response> {
  const slash = parseSlashCommand(bodyText);
  const cmd = parseMailCommand(slash.text);

  switch (cmd.kind) {
    case "help":
      return jsonResponse(slackEphemeral("Slack Outlook 邮件机器人", buildHelpBlocks()));
    case "connect": {
      runBackground(ctx, (async () => {
        try {
          const { authorizeUrl, providerType } = await createConnectUrl({
            teamId: slash.team_id,
            userId: slash.user_id,
            channelId: slash.channel_id,
            channelName: slash.channel_name,
            providerType: cmd.providerType,
          });
          await postToResponseUrl(
            slash.response_url,
            {
              ...slackEphemeral(
                `点击按钮完成 Outlook 授权（provider：${providerLabel(providerType)}）`,
                buildConnectBlocks(authorizeUrl, providerLabel(providerType)),
              ),
              replace_original: true,
            },
          );
        } catch (error) {
          await postToResponseUrl(
            slash.response_url,
            {
              ...slackEphemeral(
                `❌ 生成授权链接失败：${error instanceof Error ? error.message : String(error)}`,
              ),
              replace_original: true,
            },
          );
        }
      })());
      return jsonResponse(slackEphemeral("正在生成 Outlook 授权链接，请稍候…"));
    }
    case "list": {
      const mailboxes = await listMailboxes(slash.team_id);
      return jsonResponse(slackEphemeral("已连接邮箱", buildMailboxListBlocks(mailboxes)));
    }
    case "status": {
      const mailboxes = await listMailboxes(slash.team_id);
      return jsonResponse(slackEphemeral("邮箱状态", buildStatusBlocks(mailboxes)));
    }
    case "provider": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await updateMailboxProvider({
            teamId: slash.team_id,
            mailbox: cmd.mailbox,
            providerType: cmd.providerType,
          });
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(
              `✅ ${bundle.connection.emailAddress} 已切换到 ${providerLabel(bundle.connection.providerType)}。`,
            ),
          );
        } catch (error) {
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(
              `❌ provider 切换失败：${error instanceof Error ? error.message : String(error)}`,
            ),
          );
        }
      })());
      return jsonResponse(
        slackEphemeral(
          `正在切换 ${cmd.mailbox} 的 provider 为 ${providerLabel(cmd.providerType)}。`,
        ),
      );
    }
    case "route": {
      const updated = await updateMailboxRoute({
        teamId: slash.team_id,
        mailbox: cmd.mailbox,
        channelId: cmd.channelId,
        channelName: cmd.channelName,
      });
      return jsonResponse(
        slackEphemeral(
          `✅ ${updated.connection.emailAddress} 现在会投递到 <#${cmd.channelId}>。`,
        ),
      );
    }
    case "sync": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await queueMailboxSyncByMailboxRef({
            teamId: slash.team_id,
            mailbox: cmd.mailbox,
            reason: "slash_sync",
            requestedByUserId: slash.user_id,
          });
          await processQueuedSyncs();
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`✅ 已完成同步队列处理：${bundle.connection.emailAddress}`),
          );
        } catch (error) {
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`❌ 同步失败：${error instanceof Error ? error.message : String(error)}`),
          );
        }
      })());
      return jsonResponse(slackEphemeral(`已开始排队同步 ${cmd.mailbox}。`));
    }
    case "test": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await sendTestNotification({
            teamId: slash.team_id,
            mailbox: cmd.mailbox,
          });
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`✅ 已向 <#${bundle.route?.slackChannelId}> 发送测试通知。`),
          );
        } catch (error) {
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`❌ 测试通知失败：${error instanceof Error ? error.message : String(error)}`),
          );
        }
      })());
      return jsonResponse(slackEphemeral(`正在发送测试通知：${cmd.mailbox}`));
    }
    case "disconnect": {
      runBackground(ctx, (async () => {
        try {
          const bundle = await disconnectMailbox({
            teamId: slash.team_id,
            mailbox: cmd.mailbox,
          });
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`✅ 已断开邮箱 ${bundle.connection.emailAddress}`),
          );
        } catch (error) {
          await postToResponseUrl(
            slash.response_url,
            slackEphemeral(`❌ 断开失败：${error instanceof Error ? error.message : String(error)}`),
          );
        }
      })());
      return jsonResponse(slackEphemeral(`正在断开邮箱：${cmd.mailbox}`));
    }
  }
}
