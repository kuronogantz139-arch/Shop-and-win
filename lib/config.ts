/**
 * Centralised, validated access to environment variables.
 *
 * Throwing early (with a clear message naming the missing variable) is far
 * easier to debug in Vercel logs than a downstream `undefined` blowing up
 * inside the Graph or Postgres client.
 */

function required(name: string): string {
  const value = process.env[name];
  if (!value || value.trim() === "") {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value.trim();
}

function optional(name: string, fallback: string): string {
  const value = process.env[name];
  return value && value.trim() !== "" ? value.trim() : fallback;
}

export interface AppConfig {
  // Azure AD app registration (Microsoft Graph, client-credentials flow)
  azureTenantId: string;
  azureClientId: string;
  azureClientSecret: string;

  // SharePoint targeting
  sharepointSiteId: string;
  sharepointListId: string;

  // Neon Postgres
  databaseUrl: string;

  // Operational
  cronSecret: string;
  pageSize: number;
}

export function loadConfig(): AppConfig {
  return {
    azureTenantId: required("AZURE_TENANT_ID"),
    azureClientId: required("AZURE_CLIENT_ID"),
    azureClientSecret: required("AZURE_CLIENT_SECRET"),

    sharepointSiteId: required("SHAREPOINT_SITE_ID"),
    sharepointListId: required("SHAREPOINT_LIST_ID"),

    databaseUrl: required("DATABASE_URL"),

    cronSecret: required("CRON_SECRET"),
    pageSize: Number(optional("GRAPH_PAGE_SIZE", "200")),
  };
}
