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
    json values = [["a3"], ["c3"], ["aa"]];
    int columnIndex = 1;
    
    excel:Column|error response = excelClient->createColumn(workbookIdOrPath, "sheetName", "tableName", values, 
    columnIndex);
    if (response is excel:Column) {
        log:printInfo(response.toString());
    }
}
