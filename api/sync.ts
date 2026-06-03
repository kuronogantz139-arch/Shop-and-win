/**
 * Vercel Serverless Function: POST/GET /api/sync
 *
 * Invoked by Vercel Cron (see vercel.json). Vercel automatically attaches an
 * `Authorization: Bearer <CRON_SECRET>` header to scheduled invocations when
 * the CRON_SECRET env var is set; we reject anything that doesn't match so the
 * endpoint can't be triggered by the public internet.
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import { runSync } from "../lib/sync.js";

export default async function handler(
  req: VercelRequest,
  res: VercelResponse,
): Promise<void> {
  const expected = process.env.CRON_SECRET;
  if (!expected) {
    res.status(500).json({ ok: false, error: "CRON_SECRET is not configured" });
    return;
  }

  const auth = req.headers["authorization"];
  if (auth !== `Bearer ${expected}`) {
    res.status(401).json({ ok: false, error: "Unauthorized" });
    return;
  }

  try {
    const result = await runSync();
    res.status(200).json({ ok: true, ...result });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[sync] failed:", err);
    res.status(500).json({ ok: false, error: message });
  }
}
