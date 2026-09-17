// Opens a workbook session, lists the worksheets in the workbook, and reports the
// used range of each one, so a reporting job can see how much data each sheet holds
// before deciding what to read in full.

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

    // Step 1: Open a session. Microsoft recommends passing a session id with every
    // request when making more than one call in a short period.
    excel:SessionInfoResponse session =
        check excel->createSession(driveId, driveItemId, {persistChanges: false});

    string sessionId = session is excel:SessionInfo ? (session?.id ?: "") : "";

    // Step 2: List every worksheet in the workbook.
    excel:WorksheetCollectionResponse worksheets =
        check excel->listWorksheets(driveId, driveItemId, {workbookSessionId: sessionId});

    // Step 3: Report the used range of each worksheet.
    foreach excel:Worksheet sheet in worksheets.value ?: [] {
        string sheetId = sheet.id ?: "";
        if sheetId == "" {
            continue;
        }
        excel:RangeResponse usedRange = check excel->getUsedRange(
                driveId, driveItemId, sheetId, {workbookSessionId: sessionId});

        if usedRange is excel:Range {
            io:println(string `${sheet?.name ?: "(unnamed)"}: ${usedRange?.address ?: "empty"}`);
        }
    }

    // Step 4: Close the session.
    check excel->closeSession(driveId, driveItemId, {workbookSessionId: sessionId});
}
