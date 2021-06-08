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
    excel:Data data = {
        sourceData: "A1:B3",
        seriesBy: excel:AUTO
    };

    error? response = excelClient->resetChartData(workbookIdOrPath, "worksheetName", "chartName", data);
    if !(response is error) {
        log:printInfo("Chart reset");
    }
}
