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
    excel:Worksheet sheet = {position: 1};

    excel:Worksheet|error response = excelClient->updateWorksheet(workbookIdOrPath, "worksheetName", sheet);
    if (response is excel:Worksheet) {
        log:printInfo(response.toString());
    }
}
