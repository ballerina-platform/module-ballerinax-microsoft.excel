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
    excel:Column column = {
        index: 3,
        values: [["a3"], ["c3"], ["aa"]]
    };
    
    excel:Column|error response = excelClient->createRow(workbookIdOrPath, "sheetName", "tableName", column);
    if (response is excel:Column) {
        log:printInfo(response.toString());
    }
}
