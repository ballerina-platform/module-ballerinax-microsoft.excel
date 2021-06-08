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
    excel:ChartPosition chartPosition = {startCell: "D3"};

    error? response = excelClient->setChartPosition(workbookIdOrPath, "worksheetName", "chartName", chartPosition);
    if !(response is error) {
        log:printInfo("Chart position set");
    }
}
