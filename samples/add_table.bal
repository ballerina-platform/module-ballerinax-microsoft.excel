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
    excel:TableConfiguration 'table = {address: "A1:C3"};

    excel:Table|error response = excelClient->addTable(workbookIdOrPath, "testSheet", 'table);
    if (response is excel:Table) {
        log:printInfo(response.toString());
    }
}
