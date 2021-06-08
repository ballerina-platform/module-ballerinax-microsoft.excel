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

isolated function createRequestPath(string[] pathParameters, string workbookIdOrPath, Query? query = ()) 
returns string|error {
    string path = EMPTY_STRING;
    string[] baseParameters = workbookIdOrPath.endsWith(".xlsx") ? [ME, DRIVE, ROOT + COLON, workbookIdOrPath + COLON, 
    WORKBOOK] : [ME, DRIVE, ITEMS, workbookIdOrPath, WORKBOOK];

    path = check createPath(path, baseParameters);
    path = check createPath(path, pathParameters);
    path = query is Query ? addQueryParameters(path, query) : path;
    return path;
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

isolated function addQueryParameters(string path, Query query) returns string {
    string queryPath = path + QUESTION_MARK;
    if (query?.count is boolean) {
        queryPath = queryPath + DOLLAR_SIGN + COUNT + EQUAL_SIGN + query?.count.toString();
    }
    if (query?.expand is string) {
        queryPath = queryPath + DOLLAR_SIGN + EXPAND + EQUAL_SIGN + query?.expand.toString();
    }
    if (query?.filter is string) {
        queryPath = queryPath + DOLLAR_SIGN + FILTER + EQUAL_SIGN + query?.filter.toString();
    }
    if (query?.format is string) {
        queryPath = queryPath + DOLLAR_SIGN + FORMAT + EQUAL_SIGN + query?.count.toString();
    }
    if (query?.orderBy is string) {
        queryPath = queryPath + DOLLAR_SIGN + ORDER_BY + EQUAL_SIGN + query?.orderBy.toString();
    }
    if (query?.search is string) {
        queryPath = queryPath + DOLLAR_SIGN + SEARCH + EQUAL_SIGN + query?.search.toString();
    }
    if (query?.'select is string) {
        queryPath = queryPath + DOLLAR_SIGN + SELECT + EQUAL_SIGN + query?.'select.toString();
    }
    if (query?.skip is int) {
        queryPath = queryPath + DOLLAR_SIGN + SKIP + EQUAL_SIGN + query?.skip.toString();
    }
    if (query?.top is int) {
        queryPath = queryPath + DOLLAR_SIGN + TOP + EQUAL_SIGN + query?.top.toString();
    }
    return queryPath;
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
        json jsonResponse = check httpResponse.getJsonPayload();
        return <map<json>>jsonResponse;
    } else if (httpResponse.statusCode == http:STATUS_NO_CONTENT) {
        return {};
    }
    json jsonResponse = check httpResponse.getJsonPayload();
    return error(jsonResponse.toString());
}

isolated function getWorksheetArray(http:Response response) returns Worksheet[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse["value"].cloneWithType(WorkSheetArray);
}

isolated function getRowArray(http:Response response) returns Row[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse["value"].cloneWithType(RowArray);
}

isolated function getColumnArray(http:Response response) returns Column[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse["value"].cloneWithType(ColumnArray);
}

isolated function getTableArray(http:Response response) returns Table[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse["value"].cloneWithType(TableArray);
}

isolated function getChartArray(http:Response response) returns Chart[]|error {
    map<json> handledResponse = check handleResponse(response);
    return check handledResponse["value"].cloneWithType(ChartArray);
}

type WorkSheetArray Worksheet[];

type RowArray Row[];

type ColumnArray Column[];

type TableArray Table[];

type ChartArray Chart[];
