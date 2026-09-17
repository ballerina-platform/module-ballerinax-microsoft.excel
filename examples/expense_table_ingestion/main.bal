// Sets up a fresh expense sheet: adds a worksheet, defines a table over its header
// range, and appends expense rows to it. This is the shape of an ingestion job that
// writes periodic records into a workbook.

import ballerina/io;
import ballerinax/microsoft.excel;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;
configurable string driveId = ?;
configurable string driveItemId = ?;

public function main() returns error? {
    excel:Client excel = check new ({
        auth: {
            clientId,
            clientSecret,
            refreshToken,
            refreshUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        }
    });

    // Step 1: Add a worksheet to hold the expense records.
    excel:AddWorksheetResponse added =
        check excel->addWorksheet(driveId, driveItemId, {name: "Expenses"});

    if added !is excel:Worksheet {
        return error("addWorksheet did not return the created worksheet");
    }
    string worksheetId = added.id ?: "";
    io:println("Created worksheet: ", added?.name);

    // Step 2: Define a table over the header row so rows can be appended to it.
    excel:AddTableResponse tableAdded = check excel->addWorksheetTable(
            driveId, driveItemId, worksheetId, {address: "A1:C1", hasHeaders: true});

    if tableAdded !is excel:Table {
        return error("addWorksheetTable did not return the created table");
    }
    string tableId = tableAdded.id ?: "";
    io:println("Created table: ", tableAdded?.name);

    // Step 3: Append expense rows to the table.
    excel:TableRowOperationResponse row = check excel->addRow(
            driveId, driveItemId, worksheetId, tableId,
            {values: [["2026-09-01", "Travel", 420.50], ["2026-09-03", "Hardware", 1299.00]]});

    if row is excel:TableRow {
        io:println("Appended rows at index: ", row?.index);
    }
}
