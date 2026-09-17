# Tests

The test suite covers 30 of the connector's 191 operations, chosen for CRUD coverage across each workbook resource type plus the session lifecycle: workbook and application retrieval, `createSession`/`refreshSession`/`closeSession`, worksheet list/get/add/update/delete, table list/get/add/update/delete, table row and column list/add/delete, chart list/get/add/update/delete plus chart image export, and range retrieval via `getRange` and `getUsedRange`.

Every test carries both the `mock_tests` and `live_tests` groups and runs against a mock server by default.

## Running Tests

```bash
bal test
```

The test suite uses a mock server (`tests/mock_service.bal`) that intercepts HTTP calls so no real credentials are required.

### Running against the live API

Set `IS_LIVE_SERVER=true` and supply credentials and workbook identifiers as environment variables:

| Variable | Description |
|---|---|
| `MS_EXCEL_ACCESS_TOKEN` | Bearer token for the Microsoft Graph API |
| `MS_EXCEL_DRIVE_ID` | Identifier of the drive holding the workbook |
| `MS_EXCEL_ITEM_ID` | Identifier of the workbook `driveItem` |
| `MS_EXCEL_WORKSHEET_ID` | Identifier of a worksheet in that workbook |
| `MS_EXCEL_TABLE_ID` | Identifier of a table in that worksheet |
| `MS_EXCEL_COLUMN_ID` | Identifier of a column in that table |
| `MS_EXCEL_ROW_ID` | Index of a row in that table |
| `MS_EXCEL_CHART_ID` | Identifier of a chart in that worksheet |

```bash
export IS_LIVE_SERVER=true
bal test --groups live_tests
```

Note that the live tests create and delete worksheets, tables, rows, columns, and charts in the target workbook. Point them at a scratch workbook, not one holding real data.
