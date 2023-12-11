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
import ballerina/log;
import ballerina/test;
import ballerina/http;

configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string workbookIdOrPath = os:getEnv("WORKBOOK_PATH");

ConnectionConfig configuration = {
    auth: {
        clientId: clientId,
        clientSecret: clientSecret,
        refreshToken: refreshToken,
        refreshUrl: refreshUrl
    }
};

string workbookId = workbookIdOrPath;

Client excelClient = check new (configuration);
string workBookId = workbookIdOrPath;
string itemId = "E1E6029AA48AC90!997";
string worksheetName = "testSheet";
string tableName = EMPTY_STRING;
string chartName = EMPTY_STRING;
string sessionId = EMPTY_STRING;
string columnName = EMPTY_STRING;
string rowId = EMPTY_STRING;
int:Signed32 sheetPosition = 1;
Worksheet sheet = {position: sheetPosition};
int rowIndex = 2;
boolean showHeaders = false;
int columnInputIndex = 2;

@test:BeforeSuite
function testCreateSession() {
    SessionInfo|error response = excelClient->createSession(workBookId, {persistChanges: false});
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testAddWorksheet() {
    Worksheet|error response = excelClient->createWorksheet(workBookId, {name: worksheetName});
    if response is Worksheet {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Unmatch worksheet name");
    } else {
        test:assertFail(response.toString());
    } 
}

@test:Config {dependsOn: [testAddWorksheet]}
function testGetWorksheet() {
    Worksheet|error response = excelClient->getWorksheet(workBookId, worksheetName);
    if response is Worksheet {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Worksheet not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testListWorksheets() {
    Worksheets|error response = excelClient->listWorksheets(workBookId);
    if response is Worksheets {
        test:assertNotEquals(response.value.toBalString(), EMPTY_STRING, "Found 0 worksheets");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testUpdateWorksheet() {
    Worksheet|error response = excelClient->updateWorksheet(workBookId, worksheetName, sheet);
    if response is Worksheet {
        int responsePosition = response?.position ?: 0;
        test:assertEquals(responsePosition, sheetPosition, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testGetCell() {
    Range|error response = excelClient->getWorksheetCell(workBookId, worksheetName, rowIndex, 7, sessionId);
    if response is Range {
        int row = response?.rowIndex ?: 0;
        test:assertEquals(row, rowIndex, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

@test:AfterSuite {}
function testDeleteWorksheet() {
    http:Response|error response = excelClient->deleteWorksheet(workBookId, worksheetName);
    if response is http:Response {
        if response.statusCode != 204 {
            test:assertFail(response.statusCode.toBalString());
        }
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testAddTable() {
    Table|error response = excelClient->addWorksheetTable(workBookId, worksheetName, {address: "A1:C3"}, sessionId);
    if response is Table {
        tableName = response?.name ?: EMPTY_STRING;
        test:assertNotEquals(tableName, EMPTY_STRING, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddTable]}
function testGetTable() {
    Table|error response = excelClient->getWorksheetTable(workBookId, worksheetName, tableName, sessionId);
    if response is Table {
        string responseTableName = response?.name ?: EMPTY_STRING;
        test:assertEquals(tableName, responseTableName, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetTable]}
function testListTable() {
    log:printInfo("excelClient -> listTables()");
    Tables|error response = excelClient->listWorkbookTables(workBookId, sessionId);
    if response is Tables {
        Table[] tables = response?.value ?: [];
        string responseTableName = tables[0]?.name ?: EMPTY_STRING;
        test:assertNotEquals(responseTableName, EMPTY_STRING, "Found 0 tables");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetTable]}
function testUpdateTable() {
    Table|error response = excelClient->updateWorksheetTable(workBookId, worksheetName, tableName, {style: "TableStyleMedium2"}, sessionId);
    if response is Table {
        test:assertEquals(response?.style, "TableStyleMedium2", "Table is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

int rowInputIndex = 1;

@test:Config {dependsOn: [testUpdateTable]}
function testCreateRow() {
    Row|error response = excelClient->addWorksheetTableRow(workBookId, worksheetName, tableName, {values: [[1, 2, 3]], index: <int:Signed32>rowInputIndex}, sessionId);
    if response is Row {
        int responseIndex = <int>response.index;
        test:assertEquals(responseIndex, rowInputIndex, "Row is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testListRows() {
    Rows|error response = excelClient->listWorksheetTableRows(workBookId, worksheetName, tableName, sessionId);
    if response is Rows {
        Row[] rows = response?.value ?: [];
        test:assertEquals(rows[1].index, rowInputIndex, "Found 0 rows");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testUpdateRow() {
    string value = "testValue";
    Row|error response = excelClient->updateWorksheetTableRow(workBookId, worksheetName, tableName, rowInputIndex.toString(), {index: rowInputIndex, values: [[(), 6, value]]},sessionId);
    if response is Row {
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
    error|http:Response response = excelClient->deleteWorksheetTableRow(workBookId, worksheetName, tableName, rowInputIndex.toString(), sessionId);
    if response is error {
        test:assertFail(response.toString());
    } else if (response.statusCode != 204) {
        test:assertFail(response.statusCode.toBalString());
    }
}

@test:Config {dependsOn: [testDeleteRow]}
function testCreateColumn() {
    Column|error response = excelClient->createWorksheetTableColumn(workBookId, worksheetName, tableName, {index: columnInputIndex, values : [["a3"], ["a4"], ["a5"], ["a1"]]},  sessionId);
    if response is Column {
        int responseIndex = <int>response.index;
        test:assertEquals(responseIndex, columnInputIndex, "Column is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testListColumn() {
    Columns|error response = excelClient->listWorksheetTableColumns(workBookId, worksheetName, tableName, sessionId);
    if response is Columns {
        Column[] columns = response?.value ?: [];
        int responseIndex = columns[2]?.index ?: 0;
        test:assertEquals(responseIndex, columnInputIndex, "Found 0 columns");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testUpdateColumn() {
    string value = "testName";
    Column|error response = excelClient->updateWorksheetTableColumn(workBookId, worksheetName, tableName, columnName, {values: [[()], [()], [value], [()]]}, sessionId);
    if response is Column {
        (string|int|decimal?)[][]? values = response.values;
        if values is () {
            test:assertFail("Column is not updated");
        }
        test:assertEquals(values[2][0], value, "Column is not updated");
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
    Chart|error response = excelClient->addWorksheetChart(workBookId, worksheetName, {'type: "ColumnStacked" , sourceData: "A1:B2", seriesBy: "Auto"}, sessionId);
    if response is Chart {
        chartName = <string>response?.name;
        test:assertNotEquals(chartName, EMPTY_STRING, "Chart is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddChart]}
function testGetChart() {
    Chart|error response = excelClient->getWorksheetChart(workBookId, worksheetName, chartName, sessionId);
    if response is Chart {
        string chartId = response?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Chart not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChart]}
function testListChart() {
    Charts|error response = excelClient->listWorksheetCharts(workBookId, worksheetName, sessionId);
    if response is Charts {
        Chart[] charts = response?.value ?: [];
        string chartId = charts[0]?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Found 0 charts");
    } else {
        test:assertFail(response.toString());
    }
}

decimal height = 99;
Chart updateChart = {
    height: height,
    left: 99
};

@test:Config {dependsOn: [testListChart]}
function testUpdateChart() {
    Chart|error response = excelClient->updateWorksheetChart(workBookId, worksheetName, chartName, updateChart, sessionId);
    if response is Chart {
        decimal|string|int responseHeight = response?.height ?: 0;
        test:assertEquals(responseHeight, height, "Chart is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateChart]}
function testGetChartImage() {
    Image|error response = excelClient->getWorksheetChartImage(workBookId, worksheetName, chartName, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChartImage]}
function testResetData() {
    error|http:Response response = excelClient->setWorksheetChartData(workBookId, worksheetName, chartName, {sourceData:"A1:B3", seriesBy: "Auto"}, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testResetData]}
function testSetPosition() {
    error|http:Response response = excelClient->setWorksheetChartPosition(workBookId, worksheetName, chartName, { startCell: "D3" }, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testSetPosition]}
function testDeleteChart() {
    error|http:Response response = excelClient->deleteWorksheetChart(workBookId, worksheetName, chartName, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testGetWorkbookApplication() {
    Application|error response = excelClient->getWorkbookApplication(workBookId, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testCalculateWorkbookApplication() {
    error|http:Response response = excelClient->calculateWorkbookApplication(workBookId, {calculationType: "Full"}, sessionId);
    if response is error {
        test:assertFail(response.toString());
    }
}
