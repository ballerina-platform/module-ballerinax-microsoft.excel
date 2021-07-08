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
    excel:Chart|error response = excelClient->addChart(workbookIdOrPath, "worksheetName", "ColumnStacked", "A1:B2",
    excel:AUTO);
    if (response is excel:Chart) {
        log:printInfo(response.toString());
    }
}
