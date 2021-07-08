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

# Microsoft Excel connector client endpoint
@display {label: "Microsoft Excel", iconPath: "logo.svg"}
public client class Client {
    http:Client excelClient;

    # Initializes the Excel connector client.
    #
    # + configuration - Configurations required to initialize the `Client` endpoint
    public isolated function init(ExcelConfiguration configuration) returns error? {
        self.excelClient = check new (BASE_URL, {
            auth: configuration.authConfig,
            secureSocket: configuration?.secureSocketConfig
        });
    }

    # Adds a new worksheet to the workbook.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetName - The name of the worksheet to be added. If specified, name should be unqiue. If not specified, 
    # Excel determines the name of the new worksheet
    # + return - `Worksheet` record or error
    @display {label: "Add Worksheet"}
    remote isolated function addWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name"} string? worksheetName = ()) 
                                            returns Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, ADD], workbookIdOrPath);
        json payload = {name: worksheetName};
        return check self.excelClient->post(path, payload, targetType = Worksheet);
    }

    # Retrieves the properties of a worksheet.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + return - `Worksheet` record or error
    @display {label: "Get Worksheet"}
    remote isolated function getWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId) 
                                            returns Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId], workbookIdOrPath);
        return check self.excelClient->get(path, targetType = Worksheet);
    }

    # Retrieves a list of worksheets.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + return - `Worksheet` record list or error
    @display {label: "List Worksheets"}
    remote isolated function listWorksheets(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Query"} string? query = ()) 
                                            returns @display {label: "Worksheet List"} Worksheet[]|error {
        string path = check createRequestPath([WORKSHEETS], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getWorksheetArray(response);
    }

    # Update the properties of worksheet.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + worksheet - 'Worksheet' record contains  values for relevant fields that should be updated
    # + return - `Worksheet` record or error
    @display {label: "Update Worksheet"}
    remote isolated function updateWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                                @display {label: "Values need to be Updated"} Worksheet worksheet) 
                                                returns Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId], workbookIdOrPath);
        json payload = check worksheet.cloneWithType(json);
        return check self.excelClient->patch(path, payload, targetType = Worksheet);
    }

    # Gets the range object containing the single cell based on row and column numbers.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + row - number of the cell to be retrieved. Zero-indexed
    # + column - Column number of the cell to be retrieved. Zero-indexed
    # + return - `Cell` record or error
    @display {label: "Get Cell"}
    remote isolated function getCell(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Row Number"} int row, 
                                        @display {label: "Column Number"} int column) returns Cell|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CELL + OPEN_ROUND_BRACKET + ROW + 
        EQUAL_SIGN + row.toString() + COMMA + COLUMN + EQUAL_SIGN + column.toString() + CLOSE_ROUND_BRACKET], 
        workbookIdOrPath);
        return check self.excelClient->get(path, targetType = Cell);
    }

    # Deletes the worksheet from the workbook.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + return - nil or error
    @display {label: "Delete Worksheet"}
    remote isolated function deleteWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId) 
                                                returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    # Creates a new table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + address - Range address or name of the range object representing the data source
    # + hasHeaders - Boolean value that indicates whether the data being imported has column labels. If the source does
    # not contain headers (i.e,. when this property set to false), Excel will automatically generate
    # header shifting the data down by one row
    # + return - `Table` record or error
    @display {label: "Add Table"}
    remote isolated function addTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Range Address"} string address, 
                                        @display {label: "Has Column Labels?"} boolean? hasHeaders = ()) 
                                        returns Table|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, ADD], workbookIdOrPath);
        json payload = {address: address, hasHeaders: hasHeaders};
        return check self.excelClient->post(path, payload, targetType = Table);
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
    # + return - `Table` record or error
    @display {label: "Get Table"}
    remote isolated function getTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Query"} string? query = ()) 
                                        returns Table|error {
        string path = check createRequestPath([TABLES, tableNameOrId], workbookIdOrPath, query);
        return check self.excelClient->get(path, targetType = Table);
    }

    # Retrieves a list of tables.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + return - `Table` record list or error
    @display {label: "List Tables"}
    remote isolated function listTables(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string? worksheetNameOrId = (), 
                                        @display {label: "Query"} string? query = ()) 
                                        returns @display {label: "Table List"} Table[]|error {
        string path = worksheetNameOrId is string ? check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES], 
        workbookIdOrPath, query) : check createRequestPath([TABLES], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getTableArray(response);
    }

    # Updates the properties of a table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + 'table - `Table` record contains the values for relevant fields that should be updated
    # + return - `Table` record or error
    @display {label: "Update Table"}
    remote isolated function updateTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Values need to be Updated"} Table 'table) 
                                            returns Table|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId], workbookIdOrPath);
        json payload = check 'table.cloneWithType(json);
        return check self.excelClient->patch(path, payload, targetType = Table);
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
    # + return - `Row` record or error
    @display {label: "Create Row"}
    remote isolated function createRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Values"} json values, 
                                        @display {label: "Index"} int? index = ()) 
                                        returns Row|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, ROWS, ADD], 
        workbookIdOrPath);
        json payload = {values: values, index: index};
        return check self.excelClient->post(path, payload, targetType = Row);
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
    # + return - `Row` record list or error
    @display {label: "List Rows"}
    remote isolated function listRows(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Query"} string? query = ()) 
                                        returns @display {label: "Row List"} Row[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, ROWS], workbookIdOrPath, 
        query);
        http:Response response = check self.excelClient->get(path);
        return getRowArray(response);
    }

    # Deletes the row from the table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - Row index
    # + return - nil or error
    @display {label: "Delete Row"}
    remote isolated function deleteRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Table Name or ID"} string tableNameOrId, 
                                        @display {label: "Index"} int index) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, ROWS, ITEM_AT + 
        OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() + CLOSE_ROUND_BRACKET], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
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
    # + return - `Column` record or error
    @display {label: "Create Column"}
    remote isolated function createColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Values"} json values, 
                                            @display {label: "Index"} int? index = ()) 
                                            returns Column|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, COLUMNS, ADD], 
        workbookIdOrPath);
        json payload = {values: values, index: index};
        return check self.excelClient->post(path, payload, targetType = Column);
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
    # + return - `Column` record list or error
    @display {label: "List Columns"}
    remote isolated function listColumns(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Query"} string? query = ()) 
                                            returns @display {label: "Column List"} Column[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, COLUMNS], 
        workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getColumnArray(response);
    }

    # Deletes a column from the table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + index - The index number of the column within the columns collection of the table
    # + return - nil or error
    @display {label: "Delete Column"}
    remote isolated function deleteColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId, 
                                            @display {label: "Index"} int index) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId, COLUMNS, ITEM_AT + 
        OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() + CLOSE_ROUND_BRACKET], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    # Deletes a table.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + tableNameOrId - Table name or ID
    # + return - nil or error
    @display {label: "Delete Table"}
    remote isolated function deleteTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Table Name or ID"} string tableNameOrId) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, TABLES, tableNameOrId], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
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
    # + return - `Chart` record or error
    @display {label: "Add Chart"}
    remote isolated function addChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Chart Type"} string 'type, 
                                        @display {label: "Data"} json sourceData, SeriesBy? seriesBy = ()) 
                                        returns Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, ADD], workbookIdOrPath);
        json payload = {'type: 'type, sourceData: sourceData, seriesBy: seriesBy};
        return check self.excelClient->post(path, payload, targetType = Chart);
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
    # + return - `Chart` record or error
    @display {label: "Get Chart"}
    remote isolated function getChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Chart Name"} string chartName, 
                                        @display {label: "Query"} string? query = ()) 
                                        returns Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName], workbookIdOrPath, 
        query);
        return check self.excelClient->get(path, targetType = Chart);
    }

    # Retrieve a list of charts.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + query - Query string that can control the amount of data returned in a response. String should start with `?` 
    # and followed by query parameters. Example: `?$top=2&$count=true`. For more information about query 
    # parameters, refer https://docs.microsoft.com/en-us/graph/query-parameters
    # + return - `Chart` record list or error
    @display {label: "List Charts"}
    remote isolated function listCharts(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                        @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                        @display {label: "Query"} string? query = ()) 
                                        returns @display {label: "Chart List"} Chart[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getChartArray(response);
    }

    # Updates the properties of chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + chart - 'Chart' record contains values for relevant fields that should be updated
    # + return - `Chart` record or error
    @display {label: "Update Chart"}
    remote isolated function updateChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Values need to be Updated"} Chart chart) 
                                            returns Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName], workbookIdOrPath);
        json payload = check chart.cloneWithType(json);
        return check self.excelClient->patch(path, payload, targetType = Chart);
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
    # + return - Base-64 image string or error
    @display {label: "Get Chart Image"}
    remote isolated function getChartImage(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Chart Width"} int? width = (), 
                                            @display {label: "Chart Height"} int? height = (), 
                                            FittingMode? fittingMode = ()) 
                                            returns @display {label: "Base-64 Chart Image"} string|error {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName, IMAGE], 
        workbookIdOrPath);
        path = setOptionalParamsToPath(path, width, height, fittingMode);
        http:Response response = check self.excelClient->get(path);
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
    # + return - nil or error
    @display {label: "Reset Chart Data"}
    remote isolated function resetChartData(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName, 
                                            @display {label: "Data"} json sourceData, SeriesBy? seriesBy = ()) 
                                            returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName, SET_DATA], 
        workbookIdOrPath);
        json payload = {sourceData: sourceData, seriesBy: seriesBy};
        http:Response response = check self.excelClient->post(path, payload);
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
    # + return - nil or error
    @display {label: "Set Chart Position"}
    remote isolated function setChartPosition(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                                @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                                @display {label: "Chart Name"} string chartName, 
                                                @display {label: "Start Cell"} string startCell, 
                                                @display {label: "End Cell"} string? endCell = ()) 
                                                returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName, SET_POSITION], 
        workbookIdOrPath);
        json payload = {startCell: startCell, endCell: endCell};
        http:Response response = check self.excelClient->post(path, payload);
        _ = check handleResponse(response);
    }

    # Deletes a chart.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + worksheetNameOrId - Worksheet name or ID
    # + chartName - Chart name
    # + return - nil or error
    @display {label: "Delete Chart"}
    remote isolated function deleteChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet Name or ID"} string worksheetNameOrId, 
                                            @display {label: "Chart Name"} string chartName) 
                                            returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetNameOrId, CHARTS, chartName], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    # Retrieve the properties of a workbookApplication.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + return - `WorkbookApplication` or error
    @display {label: "Get Workbook Application"}
    remote isolated function getWorkbookApplication(@display {label: "Workbook ID or Path"} string workbookIdOrPath) 
                                                    returns WorkbookApplication|error {
        string path = check createRequestPath([APPLICATION], workbookIdOrPath);
        WorkbookApplication response = check self.excelClient->get(path, targetType = WorkbookApplication);
        return response;
    }

    # Recalculate all currently opened workbooks in Excel.
    #
    # + workbookIdOrPath - Workbook ID or file path. Path should be with the `.xlsx` extension from root. If a worksbook
    # is in root, path will be `<FILE_NAME>.xlsx`)
    # + 'type - Specifies the calculation type to use
    # + return - nil or error
    @display {label: "Calculate Workbook Application"}
    remote isolated function calculateWorkbookApplication(@display {label: "Workbook ID or Path"} string 
                                                            workbookIdOrPath, CalculationType 'type) returns error? {
        string path = check createRequestPath([APPLICATION, CALCULATE], workbookIdOrPath);
        json payload = {calculationType: 'type};
        http:Response response = check self.excelClient->post(path, payload);
        _ = check handleResponse(response);
    }
}

# Record used to create excel client
#
# + authConfig - Client configuration  
# + secureSocketConfig - Secure socket configuration
@display {label: "Excel Configuration"}
public type ExcelConfiguration record {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig authConfig;
    http:ClientSecureSocket secureSocketConfig?;
};
