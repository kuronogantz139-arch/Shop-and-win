# SharePoint → Neon Sync

A scheduled sync service for Vercel. On a cron schedule it authenticates to
Microsoft Graph (app-only), pulls new items from a SharePoint list, transforms
them (including parsing the formatted `Price` string into an integer), and
upserts them into a Neon Postgres database.

## Architecture

| Concern | Choice | Why |
|---|---|---|
| Host | Vercel Serverless Function (`/api/sync`) | Headless job; no UI needed. |
| Schedule | Vercel Cron (`vercel.json`) | Native, no extra infra. Default: hourly. |
| Graph auth | `@azure/identity` `ClientSecretCredential` + Graph SDK | Correct unattended (daemon) flow — no signed-in user. |
| Database | `@neondatabase/serverless` (HTTP driver) | No TCP pool to exhaust across serverless invocations. |
| De-dup | `ON CONFLICT (sharepoint_id)` upsert + watermark | Idempotent and incremental — safe to retry. |

```
api/sync.ts        Vercel function entry (+ CRON_SECRET auth)
lib/config.ts      Validated env var loading
lib/graph.ts       Graph client + paginated list-item fetch
lib/db.ts          Neon client, schema, upsert, watermark
lib/transform.ts   Field mapping + Price → integer parsing
lib/sync.ts        Orchestration
scripts/           inspect-columns, migrate, run-local
```

## Data flow & the Price transform

`Price` arrives as a formatted string like `"200,750 IQD"`. `parsePriceToInteger`
in [lib/transform.ts](lib/transform.ts) strips all non-digit characters
(currency label, commas, and dots used as thousands separators) and parses the
remainder, yielding `200750`. Empty/invalid values become `NULL`.

## ⚠️ Before first run: fix the column names

Graph returns columns under `fields` keyed by their **internal** name, which
often differs from the SharePoint display name (`"Full Name"` may be
`FullName`, `Full_x0020_Name`, or even `field_1`). Discover the real names:

```bash
npm install
cp .env.example .env   # fill in values
npm run inspect:columns
```

Then update `FIELD_MAP` in [lib/transform.ts](lib/transform.ts) to match.

## Local development

```bash
npm install
cp .env.example .env        # fill in all values
npm run inspect:columns     # discover internal column names → update FIELD_MAP
npm run migrate             # create tables in Neon
npm run sync:local          # run the full pipeline once
npm run typecheck
```

## Deploy to Vercel

```bash
npm i -g vercel
vercel link
# set the env vars (see below), then:
vercel --prod
```

Cron is defined in `vercel.json` (hourly, `0 * * * *`). Adjust the schedule
there. Note: Vercel Cron requires a Pro plan for sub-daily schedules.

## Required Vercel environment variables

Set these in **Vercel → Project → Settings → Environment Variables** (mark them
for Production, and Preview if you deploy previews). All are secrets — none
should be `NEXT_PUBLIC_`.

| Variable | Description | Where to get it |
|---|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID | App registration → Overview |
| `AZURE_CLIENT_ID` | Application (client) ID | App registration → Overview |
| `AZURE_CLIENT_SECRET` | Client secret **value** | App registration → Certificates & secrets |
| `SHAREPOINT_SITE_ID` | `host,siteGuid,webGuid` | `GET /sites/{host}:/sites/{name}` |
| `SHAREPOINT_LIST_ID` | List GUID or name | `GET /sites/{site-id}/lists` |
| `DATABASE_URL` | Neon **pooled** connection string (`sslmode=require`) | Neon dashboard → Connection string |
| `CRON_SECRET` | Long random token guarding `/api/sync` | `openssl rand -hex 32` |
| `GRAPH_PAGE_SIZE` | *(optional)* page size, default `200` | — |

## Azure AD setup (one time)

1. Azure Portal → **App registrations** → New registration.
2. **API permissions** → Microsoft Graph → **Application permissions** →
   `Sites.Read.All` (or `Sites.Selected` scoped to the site) → **Grant admin
   consent**.
3. **Certificates & secrets** → New client secret → copy the **value**.

## Security notes

- `/api/sync` rejects any request whose `Authorization` header isn't
  `Bearer <CRON_SECRET>`, so only Vercel Cron (or you, with the secret) can run it.
- The Neon driver parameterises all interpolated values — no SQL string concatenation.
- Secrets live only in Vercel env vars / local `.env` (gitignored).
