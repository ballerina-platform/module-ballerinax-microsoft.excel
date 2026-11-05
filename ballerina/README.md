## Overview

[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel) is a spreadsheet application in the Microsoft 365 suite, used to organise, format, and calculate data. The Microsoft Excel connector provides programmatic access to workbooks stored in OneDrive and SharePoint through the [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview) REST API.

This connector supports the Microsoft Graph **v1.0** API and covers the workbook surface: worksheets, ranges, tables, table rows and columns, charts, PivotTables, named items, comments, and workbook sessions. All operations act on behalf of the signed-in user.

### Key features

- Create, read, update, and delete worksheets, tables, table rows, and table columns
- Read and write cell ranges by address, and apply formatting, sorting, merging, and clearing
- Create and manage charts, including chart series, axes, titles, legends, and chart image export
- Manage PivotTables, named items, and workbook comments with threaded replies
- Run a batch of calls inside a persistent or non-persistent workbook session for consistency and performance

## Setup guide

To use the Microsoft Excel connector, you need access to a Microsoft 365 business account and an application registered in Microsoft Entra ID. Workbooks stored on the OneDrive consumer platform are not supported by the Excel REST API — only files stored on the business platform.

> **Note:** The screenshots in this guide are for illustration only. The Microsoft Entra admin center changes over time, so treat them as a visual reference rather than an exact match — follow the described actions and choose the values (application name, redirect URI, permissions, and so on) that fit your own scenario.

### Step 1: Register an application in Microsoft Entra ID

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com).
2. Navigate to **Entra ID** > **App registrations** and click **New registration**.

   ![App registrations](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-1.jpg)

3. Enter a name of your choice and select the account types appropriate for your organization. Add a **Redirect URI** under the **Web** platform — this is where the sign-in flow returns the authorization code (a Microsoft-hosted page such as `https://jwt.ms` is a convenient choice for Step 4). The **Web** platform is required because the connector redeems the code using a client secret; the public "Mobile and desktop applications" platform does not accept one. Click **Register**.

   ![Register an application](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-2.jpeg)

4. On the app's **Overview** page, note the **Application (client) ID** and the **Directory (tenant) ID** — you need the client ID for `Config.toml`, and the tenant ID to build the token endpoint in Step 4.

   ![Overview of application](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-3.jpeg)

### Step 2: Add Microsoft Graph permissions

The Excel API is **delegated-only** — it supports no application (app-only) permissions, so every call acts on behalf of a signed-in user.

1. In the registered application, go to **API permissions** > **Add a permission** > **Microsoft Graph**.

   ![Select Microsoft Graph](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-5.jpeg)

2. Choose **Delegated permissions** and add the scopes your integration needs:

   | Permission | Purpose |
   |---|---|
   | `Files.Read` | Read the signed-in user's files — sufficient for read-only integrations |
   | `Files.ReadWrite` | Read and write the signed-in user's files |
   | `offline_access` | Return a refresh token alongside the access token |

   `Files.ReadWrite` is the least-privileged permission Microsoft documents for the Excel API, and there is no higher-privileged alternative — request only what your integration needs. Click **Add permissions**.

   ![Request API permissions](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-4.jpeg)

3. Back on the **API permissions** page, click **Grant admin consent for \<your tenant\>** and confirm **Yes**. Each permission should then show **Granted for \<your tenant\>** in the **Status** column.

   ![Grant admin consent](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-6.jpeg)

### Step 3: Create a client secret

1. Go to **Certificates & secrets** > **New client secret**, add a description and an expiry, and click **Add**.

   ![Add a client secret](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-8.jpeg)

2. Copy the secret **Value** immediately — it is shown only once. (The **Secret ID** next to it is just a label, not a credential.)

   ![Certificates & secrets](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/main/docs/resources/setup-guide-7.jpeg)

### Step 4: Obtain a refresh token

Replace `<TENANT_ID>`, `<CLIENT_ID>`, and `<CLIENT_SECRET>` below with the values from Steps 1 and 3. Use your directory (tenant) ID for a single-tenant application, or `common` for a multi-tenant one.

**1. Get an authorization code.** This is an interactive sign-in and consent step, so it cannot be curled — paste this URL into a browser instead:

```text
https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/authorize?client_id=<CLIENT_ID>&response_type=code&redirect_uri=https%3A%2F%2Fjwt.ms&response_mode=query&scope=https%3A%2F%2Fgraph.microsoft.com%2FFiles.ReadWrite%20offline_access&state=12345
```

Sign in and accept the consent prompt. You land on `https://jwt.ms/?code=<AUTH_CODE>&state=12345` — copy the `code` value from that URL. It is single-use and short-lived, so use it within a few minutes.

**2. Exchange the code for a refresh token.**

macOS / Linux:

```bash
curl -X POST "https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token" \
  --data-urlencode "client_id=<CLIENT_ID>" \
  --data-urlencode "client_secret=<CLIENT_SECRET>" \
  --data-urlencode "scope=https://graph.microsoft.com/Files.ReadWrite offline_access" \
  --data-urlencode "code=<AUTH_CODE>" \
  --data-urlencode "redirect_uri=https://jwt.ms" \
  --data-urlencode "grant_type=authorization_code"
```

Windows (PowerShell — call `curl.exe` explicitly; plain `curl` is aliased to `Invoke-WebRequest` and takes different flags):

```powershell
curl.exe -X POST "https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token" `
  --data-urlencode "client_id=<CLIENT_ID>" `
  --data-urlencode "client_secret=<CLIENT_SECRET>" `
  --data-urlencode "scope=https://graph.microsoft.com/Files.ReadWrite offline_access" `
  --data-urlencode "code=<AUTH_CODE>" `
  --data-urlencode "redirect_uri=https://jwt.ms" `
  --data-urlencode "grant_type=authorization_code"
```

The JSON response includes a `refresh_token` field — that is the `refreshToken` value for `Config.toml`. The `refreshUrl` is the same token endpoint you just called: `https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token`.

> The `access_token` in the same response is short-lived (about an hour) and is not used directly; the connector uses the refresh token to mint new access tokens automatically.

## Quickstart

To use the Microsoft Excel connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the connector

Import the `ballerinax/microsoft.excel` module into the Ballerina project.

```ballerina
import ballerinax/microsoft.excel;
```

### Step 2: Create a new connector instance

Create an `excel:ConnectionConfig` with the OAuth2 tokens obtained and initialize the connector with it.

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;
configurable string refreshUrl = ?;
configurable string driveId = ?;
configurable string driveItemId = ?;

final excel:Client excel = check new ({
    auth: {clientId, clientSecret, refreshToken, refreshUrl}
});
```

Provide the values in a `Config.toml` file in the project root. Use your tenant id in `refreshUrl` for a single-tenant application, or `common` for a multi-tenant one.

```toml
clientId = "<CLIENT_ID>"
clientSecret = "<CLIENT_SECRET>"
refreshToken = "<REFRESH_TOKEN>"
refreshUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
driveId = "<DRIVE_ID>"
driveItemId = "<DRIVE_ITEM_ID>"
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations. The following example lists the worksheets in a workbook, identified by the drive that holds it and the workbook file itself.

```ballerina
public function main() returns error? {
    excel:WorksheetCollectionResponse _ = check excel->listWorksheets(driveId, driveItemId);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples

The `Microsoft Excel` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/module-ballerinax-microsoft.excel/tree/main/examples/), covering the following use cases:

- [workbook_session_summary](../examples/workbook_session_summary/workbook_session_summary.md) — List the worksheets in a workbook and report each one's used range, inside a single workbook session.
- [expense_table_ingestion](../examples/expense_table_ingestion/expense_table_ingestion.md) — Add a worksheet, define a table over its header range, and append expense rows to it.
- [revenue_chart_export](../examples/revenue_chart_export/revenue_chart_export.md) — Create a chart from a range of revenue figures, position it, and export it as a base64-encoded image.
- [inventory_table_audit](../examples/inventory_table_audit/inventory_table_audit.md) — List the tables on a worksheet with their columns and header/total row ranges.
