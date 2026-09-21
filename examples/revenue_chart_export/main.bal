// Builds a revenue chart from a range of cells, positions it on the worksheet, and
// exports it as a base64-encoded image that can be embedded in a report or email.

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

    // Step 1: Create a clustered column chart over the revenue range.
    excel:AddChartResponse created = check excel->addChart(
            driveId, driveItemId, worksheetId,
            {'type: "ColumnClustered", sourceData: "A1:C5", seriesBy: "Auto"});

    if created !is excel:Chart {
        return error("addChart did not return the created chart");
    }
    string chartId = created.id ?: "";
    io:println("Created chart: ", created?.name);

    // Step 2: Anchor the chart to a cell range on the worksheet.
    check excel->setChartPosition(driveId, driveItemId, worksheetId, chartId,
            {startCell: "E2", endCell: "L20"});

    // Step 3: Export the chart as a base64-encoded PNG.
    excel:ChartImageResponse image =
        check excel->getChartImage(driveId, driveItemId, worksheetId, chartId);

    string encoded = image?.value ?: "";
    io:println("Chart image bytes (base64 length): ", encoded.length());
}
