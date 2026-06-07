/**
 * The pipeline: authenticate → pull new SharePoint items → transform →
 * upsert into Neon → advance the watermark. Kept framework-agnostic so it can
 * be driven by the Vercel function (`api/sync.ts`) or the local CLI runner.
 */

import { loadConfig } from "./config.js";
import { createGraphClient, fetchListItems } from "./graph.js";
import {
  createSqlClient,
  ensureSchema,
  getWatermark,
  setWatermark,
  upsertEntry,
  deleteRemovedEntries,
} from "./db.js";
import { mapItemToEntry, type GraphListItem } from "./transform.js";

export interface SyncResult {
  fetched: number;
  upserted: number;
  deleted: number;
  watermarkBefore: string | null;
  watermarkAfter: string | null;
  durationMs: number;
}

export async function runSync(): Promise<SyncResult> {
  const start = Date.now();
  const config = loadConfig();

  const sql = createSqlClient(config);
  await ensureSchema(sql);

  const watermarkBefore = await getWatermark(sql);

  const graph = createGraphClient(config);
  const items: GraphListItem[] = await fetchListItems(graph, config);

  let upserted = 0;
  let maxCreated = watermarkBefore;
  const currentIds: string[] = [];

  for (const item of items) {
    const entry = mapItemToEntry(item);
    await upsertEntry(sql, entry);
    upserted += 1;
    currentIds.push(entry.sharepointId);

    const created = (item as { createdDateTime?: string }).createdDateTime;
    if (created && (!maxCreated || created > maxCreated)) {
      maxCreated = created;
    }
  }

  // Delete any DB rows that no longer exist in SharePoint
  const deleted = await deleteRemovedEntries(sql, currentIds);

  if (maxCreated && maxCreated !== watermarkBefore) {
    await setWatermark(sql, maxCreated);
  }

  return {
    fetched: items.length,
    upserted,
    deleted,
    watermarkBefore,
    watermarkAfter: maxCreated,
    durationMs: Date.now() - start,
  };
}
