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
import ballerina/lang.runtime;
import ballerina/log;
import ballerina/test;

configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string workbookIdOrPath = os:getEnv("WORKBOOK_PATH");

ExcelConfiguration configuration = {
    authConfig: {
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

@test:Config {}
function testAddWorksheet() {
    log:printInfo("excelClient -> addWorksheet()");
    Worksheet|error response = excelClient->addWorksheet(workBookId, worksheetName);
    if (response is Worksheet) {
        string name = response?.name ?: EMPTY_STRING;
        test:assertEquals(name, worksheetName, "Unmatch worksheet name");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddWorksheet]}
function testGetWorksheet() {
    runtime:sleep(5);
    log:printInfo("excelClient -> getWorksheet()");
    Worksheet|error response = excelClient->getWorksheet(workBookId, worksheetName);
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
    Worksheet[]|error response = excelClient->listWorksheets(workBookId);
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
    Worksheet|error response = excelClient->updateWorksheet(workBookId, worksheetName, sheet);
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
    Cell|error response = excelClient->getCell(workBookId, worksheetName, rowIndex, 7);
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
    error? response = excelClient->deleteWorksheet(workBookId, worksheetName);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testGetWorksheet]}
function testAddTable() {
    log:printInfo("excelClient -> addTable()");
    TableConfiguration 'table = {address: "A1:C3"};
    Table|error response = excelClient->addTable(workBookId, worksheetName, 'table);
    if (response is Table) {
        tableName = response?.name ?: EMPTY_STRING;
        test:assertNotEquals(tableName, EMPTY_STRING, "Table is not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddTable]}
function testGetTable() {
    runtime:sleep(5);
    log:printInfo("excelClient -> getTable()");
    Table|error response = excelClient->getTable(workBookId, worksheetName, tableName);
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
    Table[]|error response = excelClient->listTables(workBookId);
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
    Table|error response = excelClient->updateTable(workBookId, worksheetName, tableName, updateTable);
    if (response is Table) {
        boolean responseTable = response?.showHeaders ?: true;
        test:assertEquals(responseTable, showHeaders, "Table is not updated");
    } else {
        test:assertFail(response.toString());
    }
}

int rowInputIndex = 0;
Row row = {
    index: rowInputIndex,
    values: [[1, 2, 3]]
};

@test:Config {dependsOn: [testUpdateTable]}
function testCreateRow() {
    log:printInfo("excelClient -> createRow()");
    Row|error response = excelClient->createRow(workBookId, worksheetName, tableName, row);
    if (response is Row) {
        int responseIndex = response?.index ?: 1;
        test:assertEquals(responseIndex, rowInputIndex, "Row is not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateRow]}
function testListRows() {
    log:printInfo("excelClient -> listRows()");
    Row[]|error response = excelClient->listRows(workBookId, worksheetName, tableName);
    if (response is Row[]) {
        int responseIndex = response[0]?.index ?: 1;
        test:assertEquals(responseIndex, rowInputIndex, "Found 0 rows");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testListRows]}
function testDeleteRow() {
    log:printInfo("excelClient -> deleteRow()");
    error? response = excelClient->deleteRow(workBookId, worksheetName, tableName, rowInputIndex);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

int columnInputIndex = 3;
Column column = {
    index: columnInputIndex,
    values: [["a3"], ["c3"], ["aa"]]
};

@test:Config {dependsOn: [testUpdateTable]}
function testCreateColumn() {
    log:printInfo("excelClient -> createColumn()");
    Row|error response = excelClient->createColumn(workBookId, worksheetName, tableName, column);
    if (response is Row) {
        int responseIndex = response?.index ?: 1;
        test:assertEquals(responseIndex, columnInputIndex, "Column not added");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testCreateColumn]}
function testListColumn() {
    log:printInfo("excelClient -> listColumns()");
    Column[]|error response = excelClient->listColumns(workBookId, worksheetName, tableName);
    if (response is Column[]) {
        int responseIndex = response[0]?.index ?: 1;
        test:assertEquals(responseIndex, rowInputIndex, "Found 0 columns");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testListColumn]}
function testDeleteColumn() {
    log:printInfo("excelClient -> deleteColumn()");
    error? response = excelClient->deleteColumn(workBookId, worksheetName, tableName, columnInputIndex);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testDeleteColumn, testDeleteRow]}
function testDeleteTable() {
    log:printInfo("excelClient -> deleteTable()");
    error? response = excelClient->deleteTable(workBookId, worksheetName, tableName);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

ChartConfiguration chart = {
    'type: "ColumnStacked",
    sourceData: "A1:B2",
    seriesBy: AUTO
};

@test:Config {dependsOn: [testCreateRow]}
function testAddChart() {
    runtime:sleep(5);
    log:printInfo("excelClient -> addChart()");
    Chart|error response = excelClient->addChart(workBookId, worksheetName, chart);
    if (response is Chart) {
        chartName = <string>response?.name;
        test:assertNotEquals(chartName, EMPTY_STRING, "Chart not created");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testAddChart]}
function testGetChart() {
    runtime:sleep(5);
    log:printInfo("excelClient -> getChart()");
    Chart|error response = excelClient->getChart(workBookId, worksheetName, chartName);
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
    Chart[]|error response = excelClient->listCharts(workBookId, worksheetName);
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
    Chart|error response = excelClient->updateChart(workBookId, worksheetName, chartName, updateChart);
    if (response is Chart) {
        float responseHeight = response?.height ?: 0;
        test:assertEquals(responseHeight, height, "Chart not updated");
    } else {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testUpdateChart]}
function testGetChartImage() {
    log:printInfo("excelClient -> getChartImage()");
    ChartImage|error response = excelClient->getChartImage(workBookId, worksheetName, chartName);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

Data data = {
    sourceData: "A1:B3",
    seriesBy: AUTO
};

@test:Config {dependsOn: [testGetChartImage]}
function testResetData() {
    log:printInfo("excelClient -> resetChartData()");
    error? response = excelClient->resetChartData(workBookId, worksheetName, chartName, data);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

ChartPosition chartPosition = {startCell: "D3"};

@test:Config {dependsOn: [testResetData]}
function testSetPosition() {
    log:printInfo("excelClient -> setChartPosition()");
    error? response = excelClient->setChartPosition(workBookId, worksheetName, chartName, chartPosition);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {dependsOn: [testSetPosition]}
function testDeleteChart() {
    log:printInfo("excelClient -> deleteChart()");
    error? response = excelClient->deleteChart(workBookId, worksheetName, chartName);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testGetWorkbookApplication() {
    log:printInfo("excelClient -> getWorkbookApplication()");

    Application|error response = excelClient->getWorkbookApplication(workBookId);
    if (response is error) {
        test:assertFail(response.toString());
    }
}

@test:Config {}
function testCalculateWorkbookApplication() {
    log:printInfo("excelClient -> calculateWorkbookApplication()");
    error? response = excelClient->calculateWorkbookApplication(workBookId, FULL);
    if (response is error) {
        test:assertFail(response.toString());
    }
}
