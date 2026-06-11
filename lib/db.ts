/**
 * Neon Postgres access via the serverless HTTP driver.
 *
 * `@neondatabase/serverless` issues queries over HTTP, which suits Vercel
 * functions: there is no long-lived TCP connection pool to exhaust across
 * cold/warm invocations. The `neon()` tagged-template client safely
 * parameterises interpolated values (no manual escaping).
 */

import { neon } from "@neondatabase/serverless";
import type { AppConfig } from "./config.js";
import type { SyncEntry } from "./transform.js";

export type Sql = ReturnType<typeof neon>;

export function createSqlClient(config: AppConfig): Sql {
  return neon(config.databaseUrl);
}

/**
 * Creates the destination table and the single-row sync watermark table.
 * Safe to run on every deploy/invocation (idempotent).
 */
export async function ensureSchema(sql: Sql): Promise<void> {
  await sql`
    CREATE TABLE IF NOT EXISTS sharepoint_entries (
      sharepoint_id     TEXT PRIMARY KEY,
      full_name         TEXT,
      phone             TEXT,
      age               INTEGER,
      gender            TEXT,
      email             TEXT,
      id_number         TEXT,
      address           TEXT,
      shop_name         TEXT,
      price             INTEGER,
      notes             TEXT,
      number_of_coupons INTEGER,
      receipt_number    TEXT,
      coupon_code       TEXT,
      submission_date   TIMESTAMPTZ,
      synced_at         TIMESTAMPTZ NOT NULL DEFAULT now()
    )
  `;

  await sql`
    CREATE TABLE IF NOT EXISTS sync_state (
      id              BOOLEAN PRIMARY KEY DEFAULT true,
      last_watermark  TIMESTAMPTZ,
      updated_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
      CONSTRAINT sync_state_singleton CHECK (id)
    )
  `;
}

/** Returns the high-water mark (max createdDateTime processed) or null. */
export async function getWatermark(sql: Sql): Promise<string | null> {
  const rows = (await sql`
    SELECT last_watermark FROM sync_state WHERE id = true
  `) as Array<{ last_watermark: string | null }>;

  const value = rows[0]?.last_watermark ?? null;
  return value ? new Date(value).toISOString() : null;
}

export async function setWatermark(sql: Sql, iso: string): Promise<void> {
  await sql`
    INSERT INTO sync_state (id, last_watermark, updated_at)
    VALUES (true, ${iso}, now())
    ON CONFLICT (id)
    DO UPDATE SET last_watermark = EXCLUDED.last_watermark, updated_at = now()
  `;
}

/**
 * Upserts one entry keyed on the SharePoint item id. Re-running the sync over
 * the same item is therefore a no-op data-wise (idempotent), which makes the
 * whole pipeline safe to retry.
 */
/** Deletes any rows whose sharepoint_id is not in the current fetch.
 *  Safety guard: if the fetch returned fewer than 80% of the rows currently
 *  in the DB, skip deletion — it likely means SharePoint returned an
 *  incomplete result set, and we must not mass-delete valid records.
 */
export async function deleteRemovedEntries(
  sql: Sql,
  currentIds: string[],
): Promise<number> {
  if (currentIds.length === 0) return 0;

  const countRows = (await sql`
    SELECT COUNT(*)::int AS total FROM sharepoint_entries
  `) as Array<{ total: number }>;
  const dbTotal = countRows[0]?.total ?? 0;

  // If the fetch looks incomplete (< 80% of what we have), abort deletion.
  if (dbTotal > 0 && currentIds.length < dbTotal * 0.8) {
    console.warn(
      `[sync] Skipping deletion: fetched ${currentIds.length} items but DB has ${dbTotal} rows — looks like a partial fetch.`,
    );
    return 0;
  }

  const rows = (await sql`
    DELETE FROM sharepoint_entries
    WHERE sharepoint_id != ALL(${currentIds})
    RETURNING sharepoint_id
  `) as Array<{ sharepoint_id: string }>;
  return rows.length;
}

export async function upsertEntry(sql: Sql, e: SyncEntry): Promise<void> {
  await sql`
    INSERT INTO sharepoint_entries (
      sharepoint_id, full_name, phone, age, gender, email, id_number,
      address, shop_name, price, notes, number_of_coupons, receipt_number,
      coupon_code, submission_date, synced_at
    )
    VALUES (
      ${e.sharepointId}, ${e.fullName}, ${e.phone}, ${e.age}, ${e.gender},
      ${e.email}, ${e.idNumber}, ${e.address}, ${e.shopName}, ${e.price},
      ${e.notes}, ${e.numberOfCoupons}, ${e.receiptNumber}, ${e.couponCode},
      ${e.submissionDate}, now()
    )
    ON CONFLICT (sharepoint_id) DO UPDATE SET
      full_name         = EXCLUDED.full_name,
      phone             = EXCLUDED.phone,
      age               = EXCLUDED.age,
      gender            = EXCLUDED.gender,
      email             = EXCLUDED.email,
      id_number         = EXCLUDED.id_number,
      address           = EXCLUDED.address,
      shop_name         = EXCLUDED.shop_name,
      price             = EXCLUDED.price,
      notes             = EXCLUDED.notes,
      number_of_coupons = EXCLUDED.number_of_coupons,
      receipt_number    = EXCLUDED.receipt_number,
      coupon_code       = EXCLUDED.coupon_code,
      submission_date   = EXCLUDED.submission_date,
      synced_at         = now()
  `;
}
