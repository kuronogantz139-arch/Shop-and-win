/**
 * Creates the Neon tables without running a full sync.
 *
 *   npm run migrate
 */

import "dotenv/config";
import { loadConfig } from "../lib/config.js";
import { createSqlClient, ensureSchema } from "../lib/db.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const sql = createSqlClient(config);
  await ensureSchema(sql);
  console.log("Schema ensured: sharepoint_entries, sync_state");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
