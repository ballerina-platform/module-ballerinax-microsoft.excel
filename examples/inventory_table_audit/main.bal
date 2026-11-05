// Audits the tables on an inventory worksheet: lists each table, its columns, and the
// address of its header and total rows, so a data-quality check can confirm the sheet
// still has the structure a downstream job expects.

import ballerina/io;
import ballerinax/microsoft.excel;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;
configurable string refreshUrl = ?;
configurable string driveId = ?;
configurable string driveItemId = ?;
configurable string worksheetId = ?;

public function main() returns error? {
    excel:Client excel = check new ({
        auth: {clientId, clientSecret, refreshToken, refreshUrl}
    });

    // Step 1: List the tables defined on the worksheet.
    excel:TableCollectionResponse tables =
        check excel->listWorksheetTables(driveId, driveItemId, worksheetId);

    foreach excel:Table tbl in tables.value ?: [] {
        string tableId = tbl.id ?: "";
        if tableId == "" {
            continue;
        }
        io:println("Table: ", tbl?.name);

        // Step 2: List the columns of the table.
        excel:TableColumnCollectionResponse columns =
            check excel->listColumns(driveId, driveItemId, worksheetId, tableId);

        foreach excel:TableColumn column in columns.value ?: [] {
            io:println("  column: ", column?.name);
        }

        // Step 3: Report the header row range.
        excel:RangeResponse header =
            check excel->getTableHeaderRowRange(driveId, driveItemId, worksheetId, tableId);
        if header is excel:Range {
            io:println("  header row: ", header?.address);
        }

        // Step 4: Report the total row range, when the table shows one.
        excel:RangeResponse total =
            check excel->getTableTotalRowRange(driveId, driveItemId, worksheetId, tableId);
        if total is excel:Range {
            io:println("  total row: ", total?.address);
        }
    }
}
