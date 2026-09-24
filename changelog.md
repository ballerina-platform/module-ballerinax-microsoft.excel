# Change Log
This file contains all the notable changes done to the Ballerina Microsoft Excel package through the releases.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

This release regenerates the connector from the [Microsoft Graph v1.0 OpenAPI description](https://github.com/microsoftgraph/msgraph-metadata/blob/master/openapi/v1.0/openapi.yaml)
instead of maintaining the client by hand. The operation surface grows from 30 hand-written
operations to 191 generated ones, covering the whole Excel workbook API.

It contains breaking changes. See the "Migrating from 2.x" section below.

### Added
- 161 new operations, bringing the total to 191. The connector now covers worksheets, ranges,
  tables, table rows and columns, charts (including series, axes, legend, title, data labels and
  image export), PivotTables, named items, comments and replies, workbook operations, and
  workbook sessions.
- Workbook session support on every applicable operation. 190 of the 191 operations accept an
  optional `workbookSessionId` header — every operation except `createSession`, which mints the
  ID. Microsoft recommends passing a session ID whenever more than one call is made in a short
  period; previously the session endpoints returned an ID the client had no typed way to use.
- A mock-server test suite covering 30 operations, runnable without credentials via `bal test`,
  and against the live API with `IS_LIVE_SERVER=true`.
- Four runnable example packages under `examples/`: `workbook_session_summary`,
  `expense_table_ingestion`, `revenue_chart_export` and `inventory_table_audit`.
- `docs/resources/script.py`, which rebuilds the Excel-scoped specification from the upstream
  Graph description so the starting point is reproducible rather than a checked-in copy of
  unknown provenance.
- A setup guide with screenshots covering Microsoft Entra ID application registration, Graph
  permissions, client secret creation, and obtaining a refresh token.

### Changed
- **[Breaking]** Operations are addressed by `driveId` and `driveItemId` rather than a single
  `workbookIdOrPath` string. These identify the drive holding the workbook and the workbook file
  itself, and can be obtained from the Graph [drive](https://learn.microsoft.com/en-us/graph/api/resources/drive)
  and [driveItem](https://learn.microsoft.com/en-us/graph/api/resources/driveitem) resources.
- **[Breaking]** Collection operations return a collection response record rather than a bare
  array — for example `listWorksheets` returns `WorksheetCollectionResponse` (with an optional
  `value` field) instead of `Worksheet[]`.
- **[Breaking]** Several operations were renamed to match their resource and scope:
  `getWorkbookApplication` → `getApplication`, `calculateWorkbookApplication` →
  `calculateApplication`, `getTable` → `getWorksheetTable`, `addTable` → `addWorksheetTable`,
  `deleteTable` → `deleteWorksheetTable`.
- The service URL remains the canonical Microsoft Graph root,
  `https://graph.microsoft.com/v1.0`, matching the other Microsoft Graph connectors.
- `@odata.type` is optional on every generated record. It was previously required on 51 records,
  which caused response binding to fail whenever Microsoft Graph omitted it — Graph returns it
  only for polymorphic instances.
- Generated documentation no longer carries the OData "navigation property" boilerplate.
  149 doc lines and 72 generic `Success` return descriptions were replaced with resource-specific
  wording, in the specification, so regeneration reproduces them.
- Package documentation moved to `README.md`. `Package.md` and `Module.md` were removed, as
  `bal pack` packages only `docs/README.md` for Ballerina Central.

### Removed
- **[Breaking]** `resetChartData` — the operation does not exist in the Microsoft Graph v1.0
  Excel API.
- **[Breaking]** The module-private `constants.bal` file. Its 48 URL-fragment constants were used
  by the hand-written client to assemble request paths and are unreferenced by the generated one.

### Not changed
- Authentication. `ConnectionConfig.auth` remains
  `http:BearerTokenConfig|OAuth2RefreshTokenGrantConfig`, so existing credential configuration
  continues to work unchanged. Note that the Excel API is delegated-only; it supports no
  application (app-only) permissions.

### Migrating from 2.x

Identify the workbook by drive and item rather than by path, and read the collection response's
`value` field:

```ballerina
// 2.x
excel:Worksheet[] response = check excelClient->listWorksheets(workbookIdOrPath);
foreach excel:Worksheet sheet in response {
    io:println(sheet.name);
}
```

```ballerina
// 3.0.0
excel:WorksheetCollectionResponse response =
    check excelClient->listWorksheets(driveId, driveItemId);
foreach excel:Worksheet sheet in response.value ?: [] {
    io:println(sheet?.name);
}
```

Operations that took positional arguments now take a typed payload record:

```ballerina
// 2.x
excel:Worksheet response = check excelClient->addWorksheet(workbookIdOrPath, "Sheet1");
```

```ballerina
// 3.0.0
excel:AddWorksheetResponse response =
    check excelClient->addWorksheet(driveId, driveItemId, {name: "Sheet1"});
```

To run a batch of calls inside a single workbook session, create the session and pass its ID on
each subsequent request:

```ballerina
// 3.0.0
excel:SessionInfoResponse session =
    check excelClient->createSession(driveId, driveItemId, {persistChanges: true});

if session is excel:SessionInfo {
    excel:Worksheet sheet = check excelClient->getWorksheet(
        driveId, driveItemId, worksheetId,
        headers = {workbookSessionId: session?.id}
    );
}
```
