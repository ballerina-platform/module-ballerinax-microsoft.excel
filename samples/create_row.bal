import ballerina/log;
import ballerinax/microsoft.excel;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;
configurable string refreshUrl = ?;
configurable string workbookIdOrPath = ?;

excel:ExcelConfiguration configuration = {
    authConfig: {
        clientId: clientId,
        clientSecret: clientSecret,
        refreshToken: refreshToken,
        refreshUrl: refreshUrl
    }
};

excel:Client excelClient = check new (configuration);

public function main() {
    excel:Row|error response = excelClient->createRow(workbookIdOrPath, "sheetName", "tableName", [[1, 2, 3]], 4);
    if (response is excel:Row) {
        log:printInfo(response.toString());
    }
}
