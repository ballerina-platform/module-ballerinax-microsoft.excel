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

isolated function createRequestParams(string[] pathParameters, string workbookIdOrPath, string? sessionId = (), string? query = ()) 
returns [string, map<string|string[]>?]|error {
    string path = check createRequestPath(pathParameters, workbookIdOrPath, query);
    map<string|string[]>? headers = createRequestHeader(sessionId);
    return [path, headers];
}

isolated function createRequestPath(string[] pathParameters, string workbookIdOrPath, string? query = ()) 
returns string|error {
    string path = EMPTY_STRING;
    string[] baseParameters = workbookIdOrPath.endsWith(".xlsx") ? [ME, DRIVE, ROOT + COLON, workbookIdOrPath + COLON, 
    WORKBOOK] : [ME, DRIVE, ITEMS, workbookIdOrPath, WORKBOOK];

    path = check createPath(path, baseParameters);
    path = check createPath(path, pathParameters);
    path = query is string ? (path + query) : path;
    return path;
}

isolated function createRequestHeader(string? sessionId = ()) returns map<string|string[]>? {
    if sessionId is string {
        map<string|string[]> headers = {};
        headers[WORKBOOK_SESSION_ID] = sessionId;
        return headers;
    }
    return;
}

isolated function createPath(string currentpath, string[] pathParameters) returns string|error {
    string path = currentpath;
    if (pathParameters.length() > 0) {
        foreach string element in pathParameters {
            if (!element.startsWith(FORWARD_SLASH)) {
                path = path + FORWARD_SLASH;
            }
            path += element;
        }
    }
    return path;
}

isolated function setOptionalParamsToPath(string currentPath, int? width, int? height, string? fittingMode) 
returns string {
    string path = currentPath;
    if (width is int) {
        if (height is int) {
            if (fittingMode is string) {
                path = currentPath + OPEN_ROUND_BRACKET + WIDTH + EQUAL_SIGN + width.toString() + COMMA + HEIGHT + 
                EQUAL_SIGN + height.toString() + COMMA + FITTING_MODE + EQUAL_SIGN + fittingMode.toString() + 
                CLOSE_ROUND_BRACKET;
            } else {
                path = currentPath + OPEN_ROUND_BRACKET + WIDTH + EQUAL_SIGN + width.toString() + COMMA + HEIGHT + 
                EQUAL_SIGN + height.toString() + CLOSE_ROUND_BRACKET;
            }
        } else {
            path = currentPath + OPEN_ROUND_BRACKET + WIDTH + EQUAL_SIGN + width.toString() + CLOSE_ROUND_BRACKET;
        }
    }
    return path;
}

isolated function handleResponse(http:Response httpResponse) returns map<json>|error {
    if (httpResponse.statusCode == http:STATUS_OK || httpResponse.statusCode == http:STATUS_CREATED) {
        final json jsonResponse = check httpResponse.getJsonPayload();
        return <map<json>>jsonResponse;
    } else if (httpResponse.statusCode == http:STATUS_NO_CONTENT) {
        return {};
    }
    json jsonResponse = check httpResponse.getJsonPayload();
    return error(jsonResponse.toString());
}

isolated function getWorksheetArray(http:Response response) returns Worksheet[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse[VALUE].cloneWithType(WorkSheetArray);
}

isolated function getRowArray(http:Response response) returns Row[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse[VALUE].cloneWithType(RowArray);
}

isolated function getColumnArray(http:Response response) returns Column[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse[VALUE].cloneWithType(ColumnArray);
}

isolated function getTableArray(http:Response response) returns Table[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse[VALUE].cloneWithType(TableArray);
}

isolated function getChartArray(http:Response response) returns Chart[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse[VALUE].cloneWithType(ChartArray);
}

isolated function createSubPath(string ItemIdOrPath) returns string {
    return ItemIdOrPath.endsWith(".xlsx") ?  string `items/${ItemIdOrPath}` : string `root:/${ItemIdOrPath}:`;
}

type WorkSheetArray Worksheet[];

type RowArray Row[];

type ColumnArray Column[];

type TableArray Table[];

type ChartArray Chart[];
