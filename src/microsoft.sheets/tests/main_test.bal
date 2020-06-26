// Copyright (c) 2020 WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
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

import ballerina/config;
import ballerina/test;

// Create the Microsoft Graph Client configuration by reading the config file.
MicrosoftGraphConfiguration msGraphConfig = {
    baseUrl: config:getAsString("MS_BASE_URL"),
    msInitialAccessToken: config:getAsString("MS_ACCESS_TOKEN"),
    msClientID: config:getAsString("MS_CLIENT_ID"),
    msClientSecret: config:getAsString("MS_CLIENT_SECRET"),
    msRefreshToken: config:getAsString("MS_REFRESH_TOKEN"),
    msRefreshURL: config:getAsString("MS_REFRESH_URL"),
    trustStorePath: config:getAsString("TRUST_STORE_PATH"),
    trustStorePassword: config:getAsString("TRUST_STORE_PASSWORD"),
    bearerToken: config:getAsString("MS_ACCESS_TOKEN"),
    clientConfig: {
        accessToken: config:getAsString("MS_ACCESS_TOKEN"),
        refreshConfig: {
            clientId: config:getAsString("MS_CLIENT_ID"),
            clientSecret: config:getAsString("MS_CLIENT_SECRET"),
            refreshToken: config:getAsString("MS_REFRESH_TOKEN"),
            refreshUrl: config:getAsString("MS_REFRESH_URL")
        }
    }
};

MSSpreadsheetClient msSpreadsheetClient = new (msGraphConfig);
Workbook workBook = new (msGraphConfig, "", "");
Worksheet worksheet = new (msGraphConfig, "", "", "", "", 0);

@test:Config {}
function testOpenWorkBook() {
    Workbook|error wb = msSpreadsheetClient->openWorkbook("/", config:getAsString("WORK_BOOK_NAME"));

    if (wb is Workbook) {
        WorkbookProperties properties = wb.getProperties();
        test:assertEquals(properties.workbookName, config:getAsString("WORK_BOOK_NAME"),
            msg = "Failed to open the workbook.");
        workBook = wb;
    } else {
        test:assertFail(msg = <string>wb.detail()["message"]);
    }
}

@test:Config {
    dependsOn: ["testOpenWorkBook"]
}
function testCreateSpreadsheet() {
    Worksheet|error ws = workBook->createWorksheet(config:getAsString("WORK_SHEET_NAME"));

    if (ws is Worksheet) {
        test:assertEquals(ws.getProperties().worksheetName, config:getAsString("WORK_SHEET_NAME"),
            msg = "Failed to create the worksheet.");
        worksheet = ws;
    } else {
        test:assertFail(msg = <string>ws.detail()["message"]);
    }
}

@test:Config {
    dependsOn: ["testCreateSpreadsheet"]
}
function testCreateTable() {
    Table|error tbl = worksheet->createTable(config:getAsString("TABLE_NAME"), "A1:D1");

    if (tbl is Table) {
        test:assertEquals(tbl.getProperties().tableName, config:getAsString("TABLE_NAME"),
            msg = "Failed to create the table.");
    } else {
        test:assertFail(msg = <string>tbl.detail()["message"]);
    }
}

@test:Config {
    dependsOn: ["testCreateTable"]
}
function testRemoveSpreadsheet() {
    error? result = workBook->removeWorksheet(config:getAsString("WORK_SHEET_NAME"));

    if (result is ()) {
        test:assertEquals(result, (), msg = "Failed to delete worksheet");
    } else {
        test:assertFail(msg = <string>result.detail()["message"]);
    }
}