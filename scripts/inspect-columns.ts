/**
 * Prints the SharePoint list's column INTERNAL names alongside their display
 * names, plus the `fields` payload of the first item. Use this once to fill in
 * FIELD_MAP in lib/transform.ts correctly.
 *
 *   npm run inspect:columns
 */

import "dotenv/config";
import { loadConfig } from "../lib/config.js";
import { createGraphClient } from "../lib/graph.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const graph = createGraphClient(config);

  const cols = await graph
    .api(`/sites/${config.sharepointSiteId}/lists/${config.sharepointListId}/columns`)
    .get();

  console.log("\n=== Columns (displayName -> internal name) ===");
  for (const c of cols.value ?? []) {
    if (c.readOnly && c.hidden) continue;
    console.log(`  ${JSON.stringify(c.displayName)}\t->\t${c.name}`);
  }

  const items = await graph
    .api(`/sites/${config.sharepointSiteId}/lists/${config.sharepointListId}/items`)
    .expand("fields")
    .top(1)
    .get();

  console.log("\n=== First item `fields` payload ===");
  console.log(JSON.stringify(items.value?.[0]?.fields ?? {}, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
