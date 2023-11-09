// Copyright (c) 2021, WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
//
// WSO2 Inc. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/os;
import ballerina/io;
import ballerina/log;
import ballerina/test;
import microsoft.excel.excel;
import ballerina/http;

configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string workbookIdOrPath = os:getEnv("WORKBOOK_PATH");

excel:ConnectionConfig configuration = {
    auth: {
        clientId: clientId,
        clientSecret: clientSecret,
        refreshToken: refreshToken,
        refreshUrl: refreshUrl
    }
};

Client excelClient = check new (configuration);
string workBookId = workbookIdOrPath;
string worksheetName = "testSheet";
string tableName = EMPTY_STRING;
string chartName = EMPTY_STRING;
string sessionId = EMPTY_STRING;
string columnName = EMPTY_STRING;
string rowId = EMPTY_STRING;
int sheetPosition = 1;
excel:Worksheet sheet = {position: sheetPosition};
int rowIndex = 2;
boolean showHeaders = false;
int columnInputIndex = 2;
excel:Table updateTable = {
    showHeaders: showHeaders,
    showTotals: false
};

@test:BeforeSuite
function testCreateSession() {
    excel:Session|error response = excelClient->createSession(workBookId, {persistChanges: false});
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testAddWorksheet() {
    excel:Worksheet|error response = excelClient->addWorksheet(workBookId, {name: worksheetName}, sessionId);
    if response is excel:Worksheet {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Unmatch worksheet name");
    } else {
        test:assertFail(response.toString());
    } 
}

@test:Config {dependsOn: [testAddWorksheet]}
function testGetWorksheet() {
    excel:Worksheet|error response = excelClient->getWorksheet(workBookId, worksheetName, sessionId);
    if response is excel:Worksheet {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Worksheet not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testListWorksheets() {
    excel:Worksheet[]|error response = excelClient->listWorksheets(workBookId, sessionId = sessionId);
    if response is excel:Worksheet[] {
        string responseWorksheetName = response[0]?.name ?: EMPTY_STRING;
        test:assertNotEquals(responseWorksheetName, EMPTY_STRING, "Found 0 worksheets");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testDeleteTable]}
function testUpdateWorksheet() {
    excel:Worksheet|error response = excelClient->updateWorksheet(workBookId, worksheetName, sheet, sessionId);
    if response is excel:Worksheet {
        int responsePosition = response?.position ?: 0;
        test:assertEquals(responsePosition, sheetPosition, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testGetCell() {
    excel:Range|error response = excelClient->getWorksheetCell(workBookId, worksheetName, rowIndex, 7, sessionId);
    if response is excel:Range {
        int row = <int>response.rowIndex;
        test:assertEquals(row, rowIndex, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

@test:AfterSuite {}
function testDeleteWorksheet() {
    http:Response|error response = excelClient->deleteWorksheet(workBookId, worksheetName, sessionId);
    if response is http:Response {
        if response.statusCode != 404 {
            test:assertFail(response.statusCode.toBalString());
        }
    } else if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testAddTable() {
    excel:Table|error response = excelClient->addWorksheetTable(workBookId, worksheetName, {address: "A1:C3"}, sessionId = sessionId);
    if response is excel:Table {
        tableName = response?.name ?: EMPTY_STRING;
        test:assertNotEquals(tableName, EMPTY_STRING, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddTable]}
function testGetTable() {
    excel:Table|error response = excelClient->getWorksheetTable(workBookId, worksheetName, tableName, sessionId = sessionId);
    if response is excel:Table {
        string responseTableName = response?.name ?: EMPTY_STRING;
        test:assertEquals(tableName, responseTableName, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetTable]}
function testListTable() {
    log:printInfo("excelClient -> listTables()");
    excel:Table[]|error response = excelClient->listWorkbookTables(workBookId, sessionId = sessionId);
    if response is excel:Table[] {
        string responseTableName = response[0]?.name ?: EMPTY_STRING;
        test:assertNotEquals(responseTableName, EMPTY_STRING, "Found 0 tables");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetTable]}
function testUpdateTable() {
    excel:Table|error response = excelClient->updateWorksheetTable(workBookId, worksheetName, tableName, {style: "TableStyleMedium2"}, sessionId);
    if response is excel:Table {
        boolean responseTable = response?.showHeaders ?: true;
        test:assertEquals(responseTable, showHeaders, "Table is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

int rowInputIndex = 1;

@test:Config {dependsOn: [testUpdateTable]}
function testCreateRow() {
    excel:Row|error response = excelClient->createWorksheetTableRow(workBookId, worksheetName, tableName, {values: [[1, 2, 3]], index: rowInputIndex}, sessionId);
    if response is excel:Row {
        rowId = <string>response.id;
        int responseIndex = <int>response.index;
        test:assertEquals(responseIndex, rowInputIndex, "Row is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testListRows() {
    excel:Row[]|error response = excelClient->listWorksheetTableRows(workBookId, worksheetName, tableName, sessionId = sessionId);
    if response is excel:Row[] {
        int responseIndex = <int>response[1].index;
        test:assertEquals(responseIndex, rowInputIndex, "Found 0 rows");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testUpdateRow() {
    string value = "testValue";
    excel:Row|error response = excelClient->updateWorksheetTableRow(workBookId, worksheetName, tableName, rowInputIndex, {values: [[(), (), value]]},sessionId);
    if response is excel:Row {
        (string|int|decimal?)[][]? values = response.values;
        if values is () {
            test:assertFail("Row is not updated");
        }
        json updatedValue = values[0][2];
        test:assertEquals(updatedValue.toString(), value, "Row is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateRow, testListRows]}
function testDeleteRow() {
    error|http:Response response = excelClient->deleteWorksheetTableRow(workBookId, worksheetName, tableName, rowInputIndex, sessionId);
    if response is error {
        test:assertFail(response.toString());
    } else if (response.statusCode != 204) {
        test:assertFail(response.statusCode.toBalString());
    }
}

@test:Config {dependsOn: [testDeleteRow]}
function testCreateColumn() {
    excel:Column|error response = excelClient->createWorksheetTableColumn(workBookId, worksheetName, tableName, {index: columnInputIndex, values : [["a3"], ["c3"], ["aa"]]},  sessionId);
    if response is excel:Column {
        int responseIndex = <int>response.index;
        columnName = <string>response.name;
        test:assertEquals(responseIndex, columnInputIndex, "Column is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testListColumn() {
    excel:Column[]|error response = excelClient->listWorksheetTableColumns(workBookId, worksheetName, tableName, sessionId = sessionId);
    if response is excel:Column[] {
        int responseIndex = <int>response[2].index;
        test:assertEquals(responseIndex, columnInputIndex, "Found 0 columns");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testUpdateColumn() {
    string value = "testName";
    io:println(columnName);
    excel:Column|error response = excelClient->updateWorksheetTableColumn(workBookId, worksheetName, tableName, columnName, {values: [[()], [()], [value]]}, sessionId = sessionId);
    if response is excel:Column {
        test:assertEquals(response.values, value, "Column is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateColumn, testListColumn]}
function testDeleteColumn() {
    error|http:Response  response = excelClient->deleteWorksheetTableColumn(workBookId, worksheetName, tableName, columnName, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testDeleteColumn, testDeleteRow, testListTable, testUpdateTable]}
function testDeleteTable() {
    error|http:Response  response = excelClient->deleteWorksheetTable(workBookId, worksheetName, tableName, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testAddChart() {
    excel:Chart|error response = excelClient->addChart(workBookId, worksheetName, {'type: "ColumnStacked" , sourceData: "A1:B2", seriesBy: "Auto"}, sessionId);
    if response is excel:Chart {
        chartName = <string>response?.name;
        test:assertNotEquals(chartName, EMPTY_STRING, "Chart is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddChart]}
function testGetChart() {
    excel:Chart|error response = excelClient->getChart(workBookId, worksheetName, chartName, sessionId = sessionId);
    if response is excel:Chart {
        string chartId = response?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Chart not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChart]}
function testListChart() {
    excel:Chart[]|error response = excelClient->listCharts(workBookId, worksheetName, sessionId = sessionId);
    if response is excel:Chart[] {
        string chartId = response[0]?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Found 0 charts");
    } else {
        test:assertFail(response.toString());
    }
}

decimal height = 99;
excel:Chart updateChart = {
    height: height,
    left: 99
};

@test:Config {dependsOn: [testListChart]}
function testUpdateChart() {
    excel:Chart|error response = excelClient->updateChart(workBookId, worksheetName, chartName, updateChart, sessionId);
    if response is excel:Chart {
        decimal responseHeight = response?.height ?: 0;
        test:assertEquals(responseHeight, height, "Chart is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateChart]}
function testGetChartImage() {
    excel:Image|error response = excelClient->getChartImage(workBookId, worksheetName, chartName, sessionId = sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChartImage]}
function testResetData() {
    error|http:Response response = excelClient->resetChartData(workBookId, worksheetName, chartName, {sourceData:"A1:B3", seriesBy: "Auto"}, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testResetData]}
function testSetPosition() {
    error|http:Response response = excelClient->setChartPosition(workBookId, worksheetName, chartName, { startCell: "D3" }, sessionId = sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testSetPosition]}
function testDeleteChart() {
    error|http:Response response = excelClient->deleteChart(workBookId, worksheetName, chartName, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testGetWorkbookApplication() {
    excel:Application|error response = excelClient->getApplication(workBookId, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testCalculateWorkbookApplication() {
    error|http:Response response = excelClient->calculateApplication(workBookId, {calculationMode: "Full"}, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}
