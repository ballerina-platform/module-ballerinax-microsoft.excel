// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
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
import ballerina/test;

final boolean isLiveServer = os:getEnv("IS_LIVE_SERVER") == "true";

final string serviceUrl = isLiveServer ? "https://graph.microsoft.com/v1.0" : "http://localhost:9090";

final string token = isLiveServer ? os:getEnv("MS_EXCEL_ACCESS_TOKEN") : "test_token";

final string driveId = isLiveServer ? os:getEnv("MS_EXCEL_DRIVE_ID") : "b!testDriveId";
final string driveItemId = isLiveServer ? os:getEnv("MS_EXCEL_ITEM_ID") : "01BYE5RZYRQEQ6EJ6DVJHZDF6TYMLRQKAP";
final string worksheetId = isLiveServer ? os:getEnv("MS_EXCEL_WORKSHEET_ID") : "{00000000-0001-0000-0000-000000000000}";
final string tableId = isLiveServer ? os:getEnv("MS_EXCEL_TABLE_ID") : "1";
final string rowId = isLiveServer ? os:getEnv("MS_EXCEL_ROW_ID") : "0";
final string chartId = isLiveServer ? os:getEnv("MS_EXCEL_CHART_ID") : "{6F6D2F1A-0000-0000-0000-000000000000}";

final Client excelClient = check initClient();

isolated function initClient() returns Client|error {
    return new ({auth: {token}}, serviceUrl);
}

// ---------------------------------------------------------------- workbook --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetWorkbook() returns error? {
    Workbook response = check excelClient->getWorkbook(driveId, driveItemId);
    test:assertTrue(response?.id !is (), "workbook should have an id");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetApplication() returns error? {
    Application response = check excelClient->getApplication(driveId, driveItemId);
    test:assertTrue(response?.calculationMode !is (), "application should report a calculation mode");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testCalculateApplication() returns error? {
    error? response = excelClient->calculateApplication(driveId, driveItemId, {calculationType: "Full"});
    test:assertTrue(response is (), "calculate should complete without a body");
}

// ---------------------------------------------------------------- sessions --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testCreateSession() returns error? {
    SessionInfoResponse response = check excelClient->createSession(driveId, driveItemId, {persistChanges: true});
    test:assertTrue(response !is (), "createSession should return session info");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testRefreshSession() returns error? {
    string sessionId = check createTestSession();
    error? response =
        excelClient->refreshSession(driveId, driveItemId, {workbookSessionId: sessionId});
    test:assertTrue(response is (), "refreshSession should complete without a body");
    check excelClient->closeSession(driveId, driveItemId, {workbookSessionId: sessionId});
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testCloseSession() returns error? {
    string sessionId = check createTestSession();
    error? response =
        excelClient->closeSession(driveId, driveItemId, {workbookSessionId: sessionId});
    test:assertTrue(response is (), "closeSession should complete without a body");
}

# Creates a fresh non-persistent workbook session and returns its ID. Each test that needs a
# session creates its own, so no session ID is shared between tests.
#
# + return - The ID of the newly created session
isolated function createTestSession() returns string|error {
    SessionInfoResponse session =
        check excelClient->createSession(driveId, driveItemId, {persistChanges: false});
    if session !is SessionInfo {
        return error("createSession did not return session information");
    }
    return session?.id ?: "";
}

// -------------------------------------------------------------- worksheets --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testListWorksheets() returns error? {
    WorksheetCollectionResponse response = check excelClient->listWorksheets(driveId, driveItemId);
    test:assertTrue(response.value !is (), "worksheet collection should have a value array");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetWorksheet() returns error? {
    Worksheet response = check excelClient->getWorksheet(driveId, driveItemId, worksheetId);
    test:assertTrue(response?.name !is (), "worksheet should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testAddWorksheet() returns error? {
    AddWorksheetResponse response = check excelClient->addWorksheet(driveId, driveItemId, {name: "Sheet1"});
    test:assertTrue(response !is (), "addWorksheet should return the created worksheet");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testUpdateWorksheet() returns error? {
    Worksheet payload = {atOdataType: "#microsoft.graph.workbookWorksheet", name: "Renamed"};
    Worksheet response = check excelClient->updateWorksheet(driveId, driveItemId, worksheetId, payload);
    test:assertTrue(response?.name !is (), "updated worksheet should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testDeleteWorksheet() returns error? {
    AddWorksheetResponse created =
        check excelClient->addWorksheet(driveId, driveItemId, {name: "ToDelete"});
    if created !is Worksheet {
        return error("addWorksheet did not return the created worksheet");
    }
    error? response = excelClient->deleteWorksheet(driveId, driveItemId, created.id ?: "");
    test:assertTrue(response is (), "deleteWorksheet should return no content");
}

// ------------------------------------------------------------------ tables --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testListWorksheetTables() returns error? {
    TableCollectionResponse response = check excelClient->listWorksheetTables(driveId, driveItemId, worksheetId);
    test:assertTrue(response.value !is (), "table collection should have a value array");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetWorksheetTable() returns error? {
    Table response = check excelClient->getWorksheetTable(driveId, driveItemId, worksheetId, tableId);
    test:assertTrue(response?.name !is (), "table should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testAddWorksheetTable() returns error? {
    AddTableResponse response = check excelClient->addWorksheetTable(driveId, driveItemId, worksheetId,
        {address: "Sheet1!A1:C3", hasHeaders: true});
    test:assertTrue(response !is (), "addWorksheetTable should return the created table");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testUpdateWorksheetTable() returns error? {
    Table payload = {atOdataType: "#microsoft.graph.workbookTable", name: "RenamedTable"};
    Table response = check excelClient->updateWorksheetTable(driveId, driveItemId, worksheetId, tableId, payload);
    test:assertTrue(response?.name !is (), "updated table should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testDeleteWorksheetTable() returns error? {
    AddTableResponse created = check excelClient->addWorksheetTable(
            driveId, driveItemId, worksheetId, {address: "A1:C3", hasHeaders: true});
    if created !is Table {
        return error("addWorksheetTable did not return the created table");
    }
    error? response =
        excelClient->deleteWorksheetTable(driveId, driveItemId, worksheetId, created.id ?: "");
    test:assertTrue(response is (), "deleteWorksheetTable should return no content");
}

// -------------------------------------------------------------------- rows --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testListRows() returns error? {
    TableRowCollectionResponse response = check excelClient->listRows(driveId, driveItemId, worksheetId, tableId);
    test:assertTrue(response.value !is (), "row collection should have a value array");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testAddRow() returns error? {
    TableRowOperationResponse response = check excelClient->addRow(driveId, driveItemId, worksheetId, tableId,
        {index: 0, values: [["EMEA", "Q1", 15000]]});
    test:assertTrue(response !is (), "addRow should return the created row");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testDeleteRow() returns error? {
    error? response = excelClient->deleteRow(driveId, driveItemId, worksheetId, tableId, rowId);
    test:assertTrue(response is (), "deleteRow should return no content");
}

// ----------------------------------------------------------------- columns --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testListColumns() returns error? {
    TableColumnCollectionResponse response = check excelClient->listColumns(driveId, driveItemId, worksheetId, tableId);
    test:assertTrue(response.value !is (), "column collection should have a value array");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testAddColumn() returns error? {
    AddColumnResponse response = check excelClient->addColumn(driveId, driveItemId, worksheetId, tableId,
        {name: "Region", index: 0, values: [["Region"], ["EMEA"]]});
    test:assertTrue(response !is (), "addColumn should return the created column");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testDeleteColumn() returns error? {
    AddColumnResponse created = check excelClient->addColumn(
            driveId, driveItemId, worksheetId, tableId, {name: "ToDelete", values: [["ToDelete"]]});
    if created !is TableColumn {
        return error("addColumn did not return the created column");
    }
    error? response = excelClient->deleteColumn(
            driveId, driveItemId, worksheetId, tableId, created.id ?: "");
    test:assertTrue(response is (), "deleteColumn should return no content");
}

// ------------------------------------------------------------------ charts --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testListCharts() returns error? {
    ChartCollectionResponse response = check excelClient->listCharts(driveId, driveItemId, worksheetId);
    test:assertTrue(response.value !is (), "chart collection should have a value array");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetChart() returns error? {
    Chart response = check excelClient->getChart(driveId, driveItemId, worksheetId, chartId);
    test:assertTrue(response?.name !is (), "chart should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testAddChart() returns error? {
    AddChartResponse response = check excelClient->addChart(driveId, driveItemId, worksheetId,
        {'type: "ColumnClustered", sourceData: "Sheet1!A1:C3", seriesBy: "Auto"});
    test:assertTrue(response !is (), "addChart should return the created chart");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testUpdateChart() returns error? {
    Chart payload = {atOdataType: "#microsoft.graph.workbookChart", name: "Revenue by region"};
    Chart response = check excelClient->updateChart(driveId, driveItemId, worksheetId, chartId, payload);
    test:assertTrue(response?.name !is (), "updated chart should have a name");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testDeleteChart() returns error? {
    AddChartResponse created = check excelClient->addChart(driveId, driveItemId, worksheetId,
            {'type: "ColumnClustered", sourceData: "A1:C3", seriesBy: "Auto"});
    if created !is Chart {
        return error("addChart did not return the created chart");
    }
    error? response =
        excelClient->deleteChart(driveId, driveItemId, worksheetId, created.id ?: "");
    test:assertTrue(response is (), "deleteChart should return no content");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetChartImage() returns error? {
    ChartImageResponse response = check excelClient->getChartImage(driveId, driveItemId, worksheetId, chartId);
    test:assertTrue(response?.value !is (), "chart image should return a base64 payload");
}

// ------------------------------------------------------------------ ranges --

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetRange() returns error? {
    RangeResponse response = check excelClient->getRange(driveId, driveItemId, worksheetId);
    test:assertTrue(response !is (), "getRange should return a range");
}

@test:Config {groups: ["live_tests", "mock_tests"]}
isolated function testGetUsedRange() returns error? {
    RangeResponse response = check excelClient->getUsedRange(driveId, driveItemId, worksheetId);
    test:assertTrue(response !is (), "getUsedRange should return the used range");
}
