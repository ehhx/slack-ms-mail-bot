import { getConfigAsync } from "./config.ts";
import { handleRequest } from "./handler.ts";
import { runMaintenance } from "./mail/service.ts";

await getConfigAsync();

const maybeCron = (Deno as typeof Deno & {
  cron?: (name: string, schedule: string, callback: () => void | Promise<void>) => void;
}).cron;

if (typeof maybeCron === "function") {
  maybeCron("mail-bot-maintenance", "*/10 * * * *", async () => {
    await runMaintenance();
  });
}

Deno.serve((request) => handleRequest(request));
