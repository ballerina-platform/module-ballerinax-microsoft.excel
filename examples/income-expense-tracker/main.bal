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

import ballerina/http;
import ballerina/os;
import ballerinax/microsoft.excel;

configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string workbookIdOrPath = os:getEnv("WORKBOOK_PATH");

public function main() returns error? {

    excel:Client excelClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            refreshToken: refreshToken,
            refreshUrl: refreshUrl
        }
    });
    
    excel:Session session = check excelClient->createSession(workbookIdOrPath, {persistChanges: true});
    string sessionId = <string>session.id;

    http:Response _ = check excelClient->deleteWorksheet(workbookIdOrPath, "sheetName", sessionId);

    excel:Worksheet _ = check excelClient->addWorksheet(workbookIdOrPath, {name: "sheetName"}, sessionId);

    excel:Table tableValue = check excelClient->addWorksheetTable(workbookIdOrPath, "sheetName", {address: "A1:C4"}, sessionId);
    string tableName = <string>tableValue.name;

    excel:Row _ = check excelClient->addWorksheetTableRow(workbookIdOrPath, "sheetName", tableName, {values: [["Year", "Income($)", "Expense($)"], [2020,5000,4000], [2021,7000,3000], [2022,3000,5000]]}, sessionId);

    excel:Chart _ = check excelClient->addChart(workbookIdOrPath, "sheetName", {'type: "ColumnStacked", sourceData: "A1:C3", seriesBy: "Auto"});

    http:Response _ = check excelClient->deleteWorksheetTable(workbookIdOrPath, "sheetName", tableName, sessionId);

    http:Response _ = check excelClient->deleteWorksheet(workbookIdOrPath, "sheetName", sessionId);
}
