/**
 * Runs the full sync locally against the configured SharePoint list and Neon
 * database. Loads .env automatically.
 *
 *   npm run sync:local
 */

import "dotenv/config";
import { runSync } from "../lib/sync.js";

async function main(): Promise<void> {
  const result = await runSync();
  console.log(JSON.stringify(result, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
