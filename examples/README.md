# Examples

The `ballerinax/microsoft.excel` connector provides practical examples illustrating usage in various scenarios.

| Example | Description |
|---------|-------------|
| [`workbook_session_summary`](./workbook_session_summary/workbook_session_summary.md) | List the worksheets in a workbook and report each one's used range, inside a single workbook session. |
| [`expense_table_ingestion`](./expense_table_ingestion/expense_table_ingestion.md) | Add a worksheet, define a table over its header range, and append expense rows to it. |
| [`revenue_chart_export`](./revenue_chart_export/revenue_chart_export.md) | Create a chart from a range of revenue figures, position it, and export it as a base64-encoded image. |
| [`inventory_table_audit`](./inventory_table_audit/inventory_table_audit.md) | List the tables on a worksheet with their columns and header/total row ranges. |

## Prerequisites

1. Follow the [Setup guide](../ballerina/README.md#setup-guide) to register an application with the Microsoft identity platform and obtain a refresh token.

2. For each example, create a `Config.toml` file in its directory with your credentials and the identifiers of the workbook to operate on:

    ```toml
    clientId = "<CLIENT_ID>"
    clientSecret = "<CLIENT_SECRET>"
    refreshToken = "<REFRESH_TOKEN>"
    driveId = "<DRIVE_ID>"
    driveItemId = "<DRIVE_ITEM_ID>"
    ```

## Running an example

Execute the following commands to build an example from the source:

* To build an example:

    ```bash
    bal build
    ```

* To run an example:

    ```bash
    bal run
    ```

## Building the examples with the local module

**Warning**: Due to the absence of support for reading local repositories for single Ballerina files, the Bala of the module is manually written to the central repository as a workaround. Consequently, the bash script may modify your local Ballerina repositories.

Execute the following commands to build all the examples against the changes you have made to the module locally:

* To build all the examples:

    ```bash
    ./build.sh build
    ```

* To run all the examples:

    ```bash
    ./build.sh run
    ```
