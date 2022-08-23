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

# Ballerina Microsoft Excel connector provides the capability to access Microsoft Graph Excel API
# It provides capability to perform perform CRUD (Create, Read, Update, and Delete) operations on 
# [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in 
# Microsoft OneDrive. If you have more than one call to make within a certain period of time, Microsoft recommends to 
# create a session and pass the session ID with each request. By default, this connector uses sessionless.
@display {label: "Microsoft Excel", iconPath: "icon.png"}
public isolated client class Client {
    private final http:Client excelClient;

    # Initializes the connector. During initialization you can pass either http:BearerTokenConfig if you have a bearer 
    # token or http:OAuth2RefreshTokenGrantConfig if you have OAuth tokens.
    # Create a Microsoft account and obtain tokens following 
    # [this guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols)
    # 
    # + configuration - Configurations required to initialize the client
    # + return - An error on failure of initialization or else `()`
    public isolated function init(ConnectionConfig config) returns error? {
        http:ClientConfiguration httpClientConfig = {
            auth: config.auth,
            httpVersion: config.httpVersion,
            http1Settings: {...config.http1Settings},
            http2Settings: config.http2Settings,
            timeout: config.timeout,
            forwarded: config.forwarded,
            poolConfig: config.poolConfig,
            cache: config.cache,
            compression: config.compression,
            circuitBreaker: config.circuitBreaker,
            retryConfig: config.retryConfig,
            responseLimits: config.responseLimits,
            secureSocket: config.secureSocket,
            proxy: config.proxy,
            validation: config.validation
        };
        self.excelClient = check new (BASE_URL, httpClientConfig);
    }

    # Creates a session.
    # Excel APIs supports two types of sessions
    # Persistent session - All changes made to the workbook are persisted (saved). This is the most efficient and 
    # performant mode of operation.
    # Non-persistent session - Changes made by the API are not saved to the source location. Instead, the Excel backend
    # server keeps a temporary copy of the file that reflects the changes made during that particular API session.
    # When the Excel session expires, the changes are lost. This mode is useful for apps that need to do analysis or
    # obtain the results of a calculation or a chart image, but not affect the document state.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + persistChanges - All changes made to the workbook are persisted or not?
    # + return - Session ID or error
    @display {label: "Creates Session"}
    remote isolated function createSession(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Persistent Session"} boolean persistChanges = true) 
                                            returns @display {label: "Session ID"} string|error {
        string path = check createRequestPath([CREATE_SESSION], workbookIdOrPath);
        json payload = {persistChanges: persistChanges};
        record {string id;} response = check self.excelClient->post(path, payload);
        return response.id;
    }

    # Adds a new worksheet to the workbook.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetName - The name of the worksheet to be added. If specified, name should be unqiue. If not specified, 
    # Excel determines the name of the new worksheet
    # + sessionId - Session ID
    # + return - `Worksheet` record or else an `error` if failed
    @display {label: "Add Worksheet"}
    remote isolated function addWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name"} string? worksheetName = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Worksheet|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, ADD], workbookIdOrPath, 
        sessionId);
        json payload = {name: worksheetName};
        return check self.excelClient->post(path, payload, headers, targetType = Worksheet);
    }

    # Retrieves the properties of a worksheet.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + sessionId - Session ID
    # + return - `Worksheet` record or else an `error` if failed
    @display {label: "Get Worksheet"}
    remote isolated function getWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Worksheet|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId], 
        workbookIdOrPath, sessionId);
        return check self.excelClient->get(path, headers, targetType = Worksheet);
    }

    # Retrieves a list of worksheets.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Worksheet` record list or else an `error` if failed
    @display {label: "List Worksheets"}
    remote isolated function listWorksheets(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Query"} string? query = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns @display {label: "Worksheet List"} Worksheet[]|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS], workbookIdOrPath, 
        sessionId, query);
        http:Response response = check self.excelClient->get(path, headers);
        return getWorksheetArray(response);
    }

    # Update the properties of worksheet.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + worksheet - 'Worksheet' record contains  values for relevant fields that should be updated
    # + sessionId - Session ID
    # + return - `Worksheet` record or else an `error` if failed
    @display {label: "Update Worksheet"}
    remote isolated function updateWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                                @display {label: "Values need to be Updated"} Worksheet worksheet, 
                                                @display {label: "Session ID"} string? sessionId = ()) 
                                                returns Worksheet|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId], 
        workbookIdOrPath, sessionId);
        json payload = check worksheet.cloneWithType(json);
        return check self.excelClient->patch(path, payload, headers, targetType = Worksheet);
    }

    # Gets the range object containing the single cell based on row and column numbers.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + row - number of the cell to be retrieved. Zero-indexed
    # + column - Column number of the cell to be retrieved. Zero-indexed
    # + sessionId - Session ID
    # + return - `Cell` record or else an `error` if failed
    @display {label: "Get Cell"}
    remote isolated function getCell(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Row Number"} int row, 
                                        @display {label: "Column Number"} int column, 
                                        @display {label: "Session ID"} string? sessionId = ()) returns Cell|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, CELL 
        + OPEN_ROUND_BRACKET + ROW + EQUAL_SIGN + row.toString() + COMMA + COLUMN + EQUAL_SIGN + column.toString() 
        + CLOSE_ROUND_BRACKET], workbookIdOrPath, sessionId);
        return check self.excelClient->get(path, headers, targetType = Cell);
    }

    # Deletes the worksheet from the workbook.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Delete Worksheet"}
    remote isolated function deleteWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                                @display {label: "Session ID"} string? sessionId = ()) 
                                                returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId], 
        workbookIdOrPath, sessionId);
        http:Response response = check self.excelClient->delete(path, headers = headers);
        _ = check handleResponse(response);
    }

    # Creates a new table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + address - Range address or name of the range object representing the data source
    # + hasHeaders - Boolean value that indicates whether the data being imported has column labels. If the source does
    # not contain headers (i.e,. when this property set to false), Excel will automatically generate
    # header shifting the data down by one row
    # + sessionId - Session ID
    # + return - `Table` record or else an `error` if failed
    @display {label: "Add Table"}
    remote isolated function addTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Range Address"} string address, 
                                        @display {label: "Has Column Labels?"} boolean? hasHeaders = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Table|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, ADD], workbookIdOrPath, sessionId);
        json payload = {address: address, hasHeaders: hasHeaders};
        return check self.excelClient->post(path, payload, headers, targetType = Table);
    }

    # Retrieves the properties of table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Table` record or else an `error` if failed
    @display {label: "Get Table"}
    remote isolated function getTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Query"} string? query = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Table|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([TABLES, tableNameOrId], 
        workbookIdOrPath, sessionId, query);
        return check self.excelClient->get(path, headers, targetType = Table);
    }

    # Retrieves a list of tables.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Table` record list or else an `error` if failed
    @display {label: "List Tables"}
    remote isolated function listTables(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string? worksheetNameOrId = (), 
                                        @display {label: "Query"} string? query = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns @display {label: "Table List"} Table[]|error {
        [string, map<string|string[]>?] [path, headers] = worksheetNameOrId is string ? check createRequestParams(
            [WORKSHEETS, worksheetNameOrId, TABLES], workbookIdOrPath, query) : check createRequestParams([TABLES], 
            workbookIdOrPath, sessionId, query);
        http:Response response = check self.excelClient->get(path, headers);
        return getTableArray(response);
    }

    # Updates the properties of a table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + 'table - `Table` record contains the values for relevant fields that should be updated
    # + sessionId - Session ID    
    # + return - `Table` record or else an `error` if failed
    @display {label: "Update Table"}
    remote isolated function updateTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Values need to be Updated"} Table 'table, 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Table|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId], workbookIdOrPath, sessionId);
        json payload = check 'table.cloneWithType(json);
        return check self.excelClient->patch(path, payload, headers, targetType = Table);
    }

    # Adds rows to the end of the table. Note that the API can accept multiple rows data using this operation. Adding 
    # one row at a time could lead to performance degradation. The recommended approach would be to batch the rows
    # together in a single call rather than doing single row insertion. For best results, collect the rows to be
    # inserted on the application side and perform single rows add operation. 
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + values - A 2-dimensional array of unformatted values of the table rows (boolean or string or number).
    # + index - Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below
    # the inserted row are shifted downwards. Zero-indexed
    # + sessionId - Session ID
    # + return - `Row` record or else an `error` if failed
    @display {label: "Create Row"}
    remote isolated function createRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Values"} json values, 
                                        @display {label: "Index"} int? index = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Row|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, ROWS, ADD], workbookIdOrPath, sessionId);
        json payload = {values: values, index: index};
        return check self.excelClient->post(path, payload, headers, targetType = Row);
    }

    # Retrieves a list of table rows.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Row` record list or else an `error` if failed
    @display {label: "List Rows"}
    remote isolated function listRows(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Query"} string? query = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns @display {label: "Row List"} Row[]|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, ROWS], workbookIdOrPath, sessionId, query);
        http:Response response = check self.excelClient->get(path, headers);
        return getRowArray(response);
    }

    # Updates the properties of a row.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - The index number of the row within the rows collection of the table
    # + values - A 2-dimensional array of unformatted values of the table rows (boolean or string or number). Provide
    # values for relevant fields that should be updated. Existing properties that are not included in the request will
    # maintain their previous values or be recalculated based on changes to other property values. For best performance
    # you shouldn't include existing values that haven't changed
    # + sessionId - Session ID
    # + return - `Row` record or else an `error` if failed
    @display {label: "Update Row"}
    remote isolated function updateRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Index"} int index, 
                                        @display {label: "Values"} json[][] values, 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Row|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, ROWS, ITEM_AT + OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() + 
        CLOSE_ROUND_BRACKET], workbookIdOrPath, sessionId);
        json payload = {values: values};
        return check self.excelClient->patch(path, payload, headers, targetType = Row);
    }

    # Deletes the row from the table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - Row index
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Delete Row"}
    remote isolated function deleteRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Index"} int index, 
                                        @display {label: "Session ID"} string? sessionId = ()) returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, ROWS, ITEM_AT + OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() + 
        CLOSE_ROUND_BRACKET], workbookIdOrPath, sessionId);
        http:Response response = check self.excelClient->delete(path, headers = headers);
        _ = check handleResponse(response);
    }

    # Creates a new table column.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + values - A 2-dimensional array of unformatted values of the table columns (boolean or string or number).
    # + index - The index number of the column within the columns collection of the table
    # + sessionId - Session ID
    # + return - `Column` record or else an `error` if failed
    @display {label: "Create Column"}
    remote isolated function createColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Values"} json values, 
                                            @display {label: "Index"} int? index = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Column|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, COLUMNS, ADD], workbookIdOrPath, sessionId);
        json payload = {values: values, index: index};
        return check self.excelClient->post(path, payload, headers, targetType = Column);
    }

    # Retrieves a list of table columns.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Column` record list or else an `error` if failed
    @display {label: "List Columns"}
    remote isolated function listColumns(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Query"} string? query = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns @display {label: "Column List"} Column[]|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, COLUMNS], workbookIdOrPath, sessionId, query);
        http:Response response = check self.excelClient->get(path, headers);
        return getColumnArray(response);
    }

    # Updates the properties of a column.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a workbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - The index number of the column within the columns collection of the table
    # + values - A 2-dimensional array of unformatted values of the table rows (boolean or string or number). Provide
    # values for relevant fields that should be updated. Existing properties that are not included in the request will
    # maintain their previous values or be recalculated based on changes to other property values. For best performance
    # you shouldn't include existing values that haven't changed
    # + name - The name of the table column
    # + sessionId - Session ID
    # + return - `Column` record or else an `error` if failed
    @display {label: "Update Column"}
    remote isolated function updateColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Index"} int index, 
                                            @display {label: "Values"} json[][]? values = (), 
                                            @display {label: "Column Name"} string? name = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Column|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, COLUMNS, ITEM_AT + OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() 
        + CLOSE_ROUND_BRACKET], workbookIdOrPath, sessionId);
        map<json> payload = {index: index};
        if (name is string) {
            payload["name"] = name;
        }
        if (values is json[][]) {
            payload["values"] = values;
        }
        return check self.excelClient->patch(path, payload, headers, targetType = Column);
    }

    # Deletes a column from the table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - The index number of the column within the columns collection of the table
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Delete Column"}
    remote isolated function deleteColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Index"} int index, 
                                            @display {label: "Session ID"} string? sessionId = ()) returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId, COLUMNS, ITEM_AT + OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() 
        + CLOSE_ROUND_BRACKET], workbookIdOrPath, sessionId);
        http:Response response = check self.excelClient->delete(path, headers = headers);
        _ = check handleResponse(response);
    }

    # Deletes a table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Delete Table"}
    remote isolated function deleteTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Session ID"} string? sessionId = ()) returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        TABLES, tableNameOrId], workbookIdOrPath, sessionId);
        http:Response response = check self.excelClient->delete(path, headers = headers);
        _ = check handleResponse(response);
    }

    # Creates a new chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + type - Represents the type of a chart. The possible values are: ColumnClustered, ColumnStacked,
    # ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, 
    # LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.
    # + sourceData - The Range object corresponding to the source data
    # + seriesBy - Specifies the way columns or rows are used as data series on the chart
    # + sessionId - Session ID
    # + return - `Chart` record or else an `error` if failed
    @display {label: "Add Chart"}
    remote isolated function addChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Chart Type"} string 'type, 
                                        @display {label: "Data"} json sourceData, SeriesBy? seriesBy = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Chart|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, ADD], workbookIdOrPath, sessionId);
        json payload = {'type: 'type, sourceData: sourceData, seriesBy: seriesBy};
        return check self.excelClient->post(path, payload, headers, targetType = Chart);
    }

    # Retrieves the properties of a chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Chart` record or else an `error` if failed
    @display {label: "Get Chart"}
    remote isolated function getChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Chart Name"} string chartName, 
                                        @display {label: "Query"} string? query = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns Chart|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName], workbookIdOrPath, sessionId, query);
        return check self.excelClient->get(path, headers, targetType = Chart);
    }

    # Retrieve a list of charts.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + sessionId - Session ID
    # + return - `Chart` record list or else an `error` if failed
    @display {label: "List Charts"}
    remote isolated function listCharts(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Query"} string? query = (), 
                                        @display {label: "Session ID"} string? sessionId = ()) 
                                        returns @display {label: "Chart List"} Chart[]|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS], workbookIdOrPath, sessionId, query);
        http:Response response = check self.excelClient->get(path, headers);
        return getChartArray(response);
    }

    # Updates the properties of chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + chart - 'Chart' record contains values for relevant fields that should be updated
    # + sessionId - Session ID
    # + return - `Chart` record or else an `error` if failed
    @display {label: "Update Chart"}
    remote isolated function updateChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Values need to be Updated"} Chart chart, 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns Chart|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName], workbookIdOrPath, sessionId);
        json payload = check chart.cloneWithType(json);
        return check self.excelClient->patch(path, payload, headers, targetType = Chart);
    }

    # Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + width - The desired width of the resulting image
    # + height - The desired height of the resulting image.
    # + fittingMode - The method used to scale the chart to the specified dimensions (if both height and width are set)
    # + sessionId - Session ID
    # + return - Base-64 image string or else an `error` if failed
    @display {label: "Get Chart Image"}
    remote isolated function getChartImage(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Chart Width"} int? width = (), 
                                            @display {label: "Chart Height"} int? height = (), 
                                            FittingMode? fittingMode = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns @display {label: "Base-64 Chart Image"} string|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName, IMAGE], workbookIdOrPath, sessionId);
        path = setOptionalParamsToPath(path, width, height, fittingMode);
        http:Response response = check self.excelClient->get(path, headers);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse[VALUE].toString();
    }

    # Resets the source data for the chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + sourceData - The Range object corresponding to the source data
    # + seriesBy - Specifies the way columns or rows are used as data series on the chart
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Reset Chart Data"}
    remote isolated function resetChartData(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Data"} json sourceData, SeriesBy? seriesBy = (), 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName, SET_DATA], workbookIdOrPath, sessionId);
        json payload = {sourceData: sourceData, seriesBy: seriesBy};
        http:Response response = check self.excelClient->post(path, payload, headers);
        _ = check handleResponse(response);
    }

    # Positions the chart relative to cells on the worksheet.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + startCell - The start cell. This is where the chart will be moved to. The start cell is the top-left or 
    # top-right cell, depending on the user's right-to-left display settings.
    # + endCell - The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Set Chart Position"}
    remote isolated function setChartPosition(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                                @display {label: "Chart Name"} string chartName, 
                                                @display {label: "Start Cell"} string startCell, 
                                                @display {label: "End Cell"} string? endCell = (), 
                                                @display {label: "Session ID"} string? sessionId = ()) 
                                                returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName, SET_POSITION], workbookIdOrPath, sessionId);
        json payload = {startCell: startCell, endCell: endCell};
        http:Response response = check self.excelClient->post(path, payload, headers);
        _ = check handleResponse(response);
    }

    # Deletes a chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Delete Chart"}
    remote isolated function deleteChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Session ID"} string? sessionId = ()) 
                                            returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([WORKSHEETS, worksheetNameOrId, 
        CHARTS, chartName], workbookIdOrPath, sessionId);
        http:Response response = check self.excelClient->delete(path, headers = headers);
        _ = check handleResponse(response);
    }

    # Retrieves the properties of a workbookApplication.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + sessionId - Session ID
    # + return - `WorkbookApplication` or else an `error` if failed
    @display {label: "Get Workbook Application"}
    remote isolated function getWorkbookApplication(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                    @display {label: "Session ID"} string? sessionId = ()) 
                                                    returns WorkbookApplication|error {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([APPLICATION], workbookIdOrPath, 
        sessionId);
        WorkbookApplication response = check self.excelClient->get(path, headers, targetType = WorkbookApplication);
        return response;
    }

    # Recalculates all currently opened workbooks in Excel.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + 'type - Specifies the calculation type to use
    # + sessionId - Session ID
    # + return - `()` or else an `error` if failed
    @display {label: "Calculate Workbook Application"}
    remote isolated function calculateWorkbookApplication(@display {label: "Workbook ID or Path"} string 
                                                            workbookIdOrPath, CalculationType 'type, 
                                                            @display {label: "Session ID"} string? sessionId = ()) 
                                                            returns error? {
        [string, map<string|string[]>?] [path, headers] = check createRequestParams([APPLICATION, CALCULATE], 
        workbookIdOrPath, sessionId);
        json payload = {calculationType: 'type};
        http:Response response = check self.excelClient->post(path, payload, headers);
        _ = check handleResponse(response);
    }
}

