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
    excel:Chart updateChart = {
        height: 99,
        left: 99
    };

    excel:Worksheet|error response = excelClient->updateChart(workbookIdOrPath, "worksheetName", "chartName", 
    updateChart);
    if (response is excel:Worksheet) {
        log:printInfo(response.toString());
    }
}
