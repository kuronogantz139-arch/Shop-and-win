/**
 * Microsoft Graph access using the app-only (client credentials) flow.
 *
 * We authenticate with a confidential client (`ClientSecretCredential`) and let
 * the official Graph SDK handle token acquisition/caching via
 * `TokenCredentialAuthenticationProvider`. This is the correct pattern for an
 * unattended daemon — there is no signed-in user.
 *
 * Required Azure AD application permission (admin-consented):
 *   - Sites.Read.All  (or Sites.Selected scoped to the target site)
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client, type PageCollection } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import type { AppConfig } from "./config.js";
import { type GraphListItem } from "./transform.js";

export function createGraphClient(config: AppConfig): Client {
  const credential = new ClientSecretCredential(
    config.azureTenantId,
    config.azureClientId,
    config.azureClientSecret,
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return Client.initWithMiddleware({ authProvider });
}

/**
 * Fetches list items from the target SharePoint list, newest entries since
 * `sinceIso` (a watermark on the Submission Date column).
 *
 * Notes on the query:
 *  - `expand=fields` pulls the column values inline so we avoid an extra round
 *    trip per item.
 *  - We filter and order on the Created system field rather than the custom
 *    Submission Date column, because custom columns must be explicitly indexed
 *    in SharePoint before Graph will allow `$filter`/`$orderby` on them
 *    (otherwise Graph returns a "field is not indexed" error). `Created` is
 *    always queryable. If your business "Submission Date" can differ from the
 *    row creation time, index that column and swap the field name here.
 *  - Graph paginates with `@odata.nextLink`; we follow it to completion.
 */
export async function fetchListItems(
  client: Client,
  config: AppConfig,
): Promise<GraphListItem[]> {
  const base = `/sites/${config.sharepointSiteId}/lists/${config.sharepointListId}/items`;

  // Plain fetch — no $filter, no $orderby, no Prefer header.
  // The "HonorNonIndexedQueries" Prefer header was causing SharePoint to
  // return partial result sets, which made the deletion logic incorrectly
  // remove rows that still exist in SharePoint.
  const request = client
    .api(base)
    .expand("fields")
    .top(config.pageSize);

  const items: GraphListItem[] = [];
  let page: PageCollection = await request.get();

  while (true) {
    const value = (page.value ?? []) as GraphListItem[];
    items.push(...value);

    const nextLink = page["@odata.nextLink"] as string | undefined;
    if (!nextLink) break;

    page = await client.api(nextLink).get();
  }

  return items;
}
