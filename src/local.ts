import { getConfigAsync } from "./config.ts";
import { handleRequest } from "./handler.ts";

const port = Number.parseInt(Deno.env.get("PORT") ?? "8000", 10);

await getConfigAsync();
console.log(`Listening on http://localhost:${port}`);

Deno.serve({ port }, (request) => handleRequest(request));
