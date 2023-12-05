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

Client excelClient = check new (configuration);
string workBookId = workbookIdOrPath;
string worksheetName = "testSheet";
string tableName = EMPTY_STRING;
string chartName = EMPTY_STRING;
string sessionId = EMPTY_STRING;

@test:BeforeSuite
function testCreateSession() {
    log:printInfo("excelClient -> createSession()");
    string|error response = excelClient->createSession(workBookId);
    if (response is string) {
        sessionId = response;
        test:assertNotEquals(response, EMPTY_STRING, "Session is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testAddWorksheet() {
    log:printInfo("excelClient -> addWorksheet()");
    Worksheet|error response = excelClient->addWorksheet(workBookId, worksheetName, sessionId);
    if (response is Worksheet) {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Unmatch worksheet name");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddWorksheet]}
function testGetWorksheet() {
    Worksheet|error response = excelClient->getWorksheet(workBookId, worksheetName, sessionId);
    if (response is Worksheet) {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Worksheet not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testListWorksheets() {
    log:printInfo("excelClient -> listWorksheets()");
    Worksheet[]|error response = excelClient->listWorksheets(workBookId, sessionId = sessionId);
    if (response is Worksheet[]) {
        string responseWorksheetName = response[0]?.name ?: EMPTY_STRING;
        test:assertNotEquals(responseWorksheetName, EMPTY_STRING, "Found 0 worksheets");
    } else {
        test:assertFail(response.toString());
    }
}

int sheetPosition = 1;
Worksheet sheet = {position: sheetPosition};

@test:Config {dependsOn: [testDeleteTable]}
function testUpdateWorksheet() {
    log:printInfo("excelClient -> updateWorksheet()");
    Worksheet|error response = excelClient->updateWorksheet(workBookId, worksheetName, sheet, sessionId);
    if (response is Worksheet) {
        int responsePosition = response?.position ?: 0;
        test:assertEquals(responsePosition, sheetPosition, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

int rowIndex = 2;

@test:Config {dependsOn: [testGetWorksheet]}
function testGetCell() {
    log:printInfo("excelClient -> getCell()");
    Cell|error response = excelClient->getCell(workBookId, worksheetName, rowIndex, 7, sessionId);
    if (response is Cell) {
        int row = response.rowIndex;
        test:assertEquals(row, rowIndex, "Unmatch worksheet position");
    } else {
        test:assertFail(response.toString());
    }
}

@test:AfterSuite {}
function testDeleteWorksheet() {
    log:printInfo("excelClient -> deleteWorksheet()");
    error? response = excelClient->deleteWorksheet(workBookId, worksheetName, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testAddTable() {
    log:printInfo("excelClient -> addTable()");
    Table|error response = excelClient->addTable(workBookId, worksheetName, "A1:C3", sessionId = sessionId);
    if (response is Table) {
        tableName = response?.name ?: EMPTY_STRING;
        test:assertNotEquals(tableName, EMPTY_STRING, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddTable]}
function testGetTable() {
    log:printInfo("excelClient -> getTable()");
    Table|error response = excelClient->getTable(workBookId, worksheetName, tableName, sessionId = sessionId);
    if (response is Table) {
        string responseTableName = response?.name ?: EMPTY_STRING;
        test:assertEquals(tableName, responseTableName, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetTable]}
function testListTable() {
    log:printInfo("excelClient -> listTables()");
    Table[]|error response = excelClient->listTables(workBookId, sessionId = sessionId);
    if (response is Table[]) {
        string responseTableName = response[0]?.name ?: EMPTY_STRING;
        test:assertNotEquals(responseTableName, EMPTY_STRING, "Found 0 tables");
    } else {
        test:assertFail(response.toString());
    }
}

boolean showHeaders = false;
Table updateTable = {
    showHeaders: showHeaders,
    showTotals: false
};

@test:Config {dependsOn: [testGetTable]}
function testUpdateTable() {
    log:printInfo("excelClient -> updateTable()");
    Table|error response = excelClient->updateTable(workBookId, worksheetName, tableName, updateTable, sessionId);
    if (response is Table) {
        boolean responseTable = response?.showHeaders ?: true;
        test:assertEquals(responseTable, showHeaders, "Table is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

int rowInputIndex = 1;

@test:Config {dependsOn: [testUpdateTable]}
function testCreateRow() {
    log:printInfo("excelClient -> createRow()");
    Row|error response = excelClient->createRow(workBookId, worksheetName, tableName, [[1, 2, 3]], rowInputIndex,
    sessionId);
    if (response is Row) {
        int responseIndex = response.index;
        test:assertEquals(responseIndex, rowInputIndex, "Row is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testListRows() {
    log:printInfo("excelClient -> listRows()");
    Row[]|error response = excelClient->listRows(workBookId, worksheetName, tableName, sessionId = sessionId);
    if (response is Row[]) {
        int responseIndex = response[1].index;
        test:assertEquals(responseIndex, rowInputIndex, "Found 0 rows");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testUpdateRow() {
    string value = "testValue";
    log:printInfo("excelClient -> updateRow()");
    Row|error response = excelClient->updateRow(workBookId, worksheetName, tableName, rowInputIndex, [[(), (), value]],
    sessionId);
    if (response is Row) {
        json updatedValue = response.values[0][2];
        test:assertEquals(updatedValue.toString(), value, "Row is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateRow, testListRows]}
function testDeleteRow() {
    log:printInfo("excelClient -> deleteRow()");
    error? response = excelClient->deleteRow(workBookId, worksheetName, tableName, rowInputIndex, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

int columnInputIndex = 2;

@test:Config {dependsOn: [testDeleteRow]}
function testCreateColumn() {
    log:printInfo("excelClient -> createColumn()");
    Column|error response = excelClient->createColumn(workBookId, worksheetName, tableName, [["a3"], ["c3"], ["aa"]], 
    columnInputIndex, sessionId);
    if (response is Column) {
        int responseIndex = response.index;
        test:assertEquals(responseIndex, columnInputIndex, "Column is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testListColumn() {
    log:printInfo("excelClient -> listColumns()");
    Column[]|error response = excelClient->listColumns(workBookId, worksheetName, tableName, sessionId = sessionId);
    if (response is Column[]) {
        int responseIndex = response[2].index;
        test:assertEquals(responseIndex, columnInputIndex, "Found 0 columns");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testUpdateColumn() {
    string value = "testName";
    log:printInfo("excelClient -> updateColumn()");
    Column|error response = excelClient->updateColumn(workBookId, worksheetName, tableName, columnInputIndex, 
    [[()], [()], [value]], sessionId = sessionId);
    if (response is Column) {
        json updatedValue = response.values[2][0];
        test:assertEquals(updatedValue.toString(), value, "Column is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateColumn, testListColumn]}
function testDeleteColumn() {
    log:printInfo("excelClient -> deleteColumn()");
    error? response = excelClient->deleteColumn(workBookId, worksheetName, tableName, columnInputIndex, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testDeleteColumn, testDeleteRow, testListTable, testUpdateTable]}
function testDeleteTable() {
    log:printInfo("excelClient -> deleteTable()");
    error? response = excelClient->deleteTable(workBookId, worksheetName, tableName, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testAddChart() {
    log:printInfo("excelClient -> addChart()");
    Chart|error response = excelClient->addChart(workBookId, worksheetName, "ColumnStacked", "A1:B2", AUTO, sessionId);
    if (response is Chart) {
        chartName = <string>response?.name;
        test:assertNotEquals(chartName, EMPTY_STRING, "Chart is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddChart]}
function testGetChart() {
    log:printInfo("excelClient -> getChart()");
    Chart|error response = excelClient->getChart(workBookId, worksheetName, chartName, sessionId = sessionId);
    if (response is Chart) {
        string chartId = response?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Chart not found");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChart]}
function testListChart() {
    log:printInfo("excelClient -> listCharts()");
    Chart[]|error response = excelClient->listCharts(workBookId, worksheetName, sessionId = sessionId);
    if (response is Chart[]) {
        string chartId = response[0]?.id ?: EMPTY_STRING;
        test:assertNotEquals(chartId, EMPTY_STRING, "Found 0 charts");
    } else {
        test:assertFail(response.toString());
    }
}

float height = 99;
Chart updateChart = {
    height: height,
    left: 99
};

@test:Config {dependsOn: [testListChart]}
function testUpdateChart() {
    log:printInfo("excelClient -> updateChart()");
    Chart|error response = excelClient->updateChart(workBookId, worksheetName, chartName, updateChart, sessionId);
    if (response is Chart) {
        float responseHeight = response?.height ?: 0;
        test:assertEquals(responseHeight, height, "Chart is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateChart]}
function testGetChartImage() {
    log:printInfo("excelClient -> getChartImage()");
    string|error response = excelClient->getChartImage(workBookId, worksheetName, chartName, sessionId = sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetChartImage]}
function testResetData() {
    log:printInfo("excelClient -> resetChartData()");
    error? response = excelClient->resetChartData(workBookId, worksheetName, chartName, "A1:B3", AUTO, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testResetData]}
function testSetPosition() {
    log:printInfo("excelClient -> setChartPosition()");
    error? response = excelClient->setChartPosition(workBookId, worksheetName, chartName, "D3", sessionId = sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testSetPosition]}
function testDeleteChart() {
    log:printInfo("excelClient -> deleteChart()");
    error? response = excelClient->deleteChart(workBookId, worksheetName, chartName, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testGetWorkbookApplication() {
    log:printInfo("excelClient -> getWorkbookApplication()");

    WorkbookApplication|error response = excelClient->getWorkbookApplication(workBookId, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testCalculateWorkbookApplication() {
    log:printInfo("excelClient -> calculateWorkbookApplication()");
    error? response = excelClient->calculateWorkbookApplication(workBookId, FULL, sessionId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}
