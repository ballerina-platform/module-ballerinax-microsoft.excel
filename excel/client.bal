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

    # Initializes the Excel connector client endpoint.
    # 
    # + configuration - Configurations required to initialize the `Client` endpoint
    public isolated function init(ExcelConfiguration configuration) returns error? {
        self.excelClient = check new (BASE_URL, {
            auth: configuration.authConfig,
            secureSocket: configuration?.secureSocketConfig
        });
    }

    @display {label: "Add Worksheet"}
    remote isolated function addWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet Name"} string? worksheetName = ()) 
                                            returns @display {label: "Worksheet"} Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, ADD], workbookIdOrPath);
        json payload = {name: worksheetName};
        http:Response response = check self.excelClient->post(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Worksheet);
    }

    @display {label: "Get Worksheet"}
    remote isolated function getWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName)
                                            returns @display {label: "Worksheet"} Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName], workbookIdOrPath);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Worksheet);
    }

    @display {label: "List Worksheets"}
    remote isolated function listWorksheets(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Query"} Query? query = ())
                                            returns @display {label: "Worksheet List"} Worksheet[]|error {
        string path = check createRequestPath([WORKSHEETS], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getWorksheetArray(response);
    }

    @display {label: "Update Worksheet"}
    remote isolated function updateWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                                @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                                @display {label: "Worksheet"} Worksheet worksheet)
                                                returns @display {label: "Worksheet"} Worksheet|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName], workbookIdOrPath);
        json payload = check worksheet.cloneWithType(json);
        http:Response response = check self.excelClient->patch(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Worksheet);
    }

    @display {label: "Get Cell"}
    remote isolated function getCell(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Row Number"} int row,
                                        @display {label: "Column Number"} int column) returns Cell|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CELL + OPEN_ROUND_BRACKET + ROW + 
        EQUAL_SIGN + row.toString() + COMMA + COLUMN + EQUAL_SIGN + column.toString() + CLOSE_ROUND_BRACKET], 
        workbookIdOrPath);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Cell);
    }

    @display {label: "Delete Worksheet"}
    remote isolated function deleteWorksheet(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                                @display {label: "Worksheet ID or Name"} string worksheetIdOrName)
                                                returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    @display {label: "Add Table"}
    remote isolated function addTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Table Configuration"} TableConfiguration 'table) returns
                                        @display {label: "Table"} Table|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, ADD], workbookIdOrPath);
        json payload = check 'table.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Table);
    }

    @display {label: "Get Table"}
    remote isolated function getTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Table Name"} string tableName,
                                        @display {label: "Query"} Query? query = ())
                                        returns @display {label: "Table"} Table|error {
        string path = check createRequestPath([TABLES, tableName], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Table);
    }

    @display {label: "List Tables"}
    remote isolated function listTables(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string? worksheetIdOrName = (),
                                        @display {label: "Query"} Query? query = ()) 
                                        returns @display {label: "Table List"} Table[]|error {
        string path = worksheetIdOrName is string ? check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES], 
        workbookIdOrPath, query) : check createRequestPath([TABLES], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getTableArray(response);
    }

    @display {label: "Update Table"}
    remote isolated function updateTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Table Name"} string tableName, 
                                            @display {label: "Table Configuration"} Table 'table)
                                            returns @display {label: "Table"} Table|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName], workbookIdOrPath);
        json payload = check 'table.cloneWithType(json);
        http:Response response = check self.excelClient->patch(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Table);
    }

    @display {label: "Create Row"}
    remote isolated function createRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Table Name"} string tableName,
                                        @display {label: "Row Configuration"} Row row) 
                                        returns @display {label: "Row"} Row|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, ROWS, ADD], 
        workbookIdOrPath);
        json payload = check row.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Row);
    }

    @display {label: "List Rows"}
    remote isolated function listRows(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Table Name"} string tableName,
                                        @display {label: "Query"} Query? query = ())
                                        returns @display {label: "Row List"} Row[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, ROWS], workbookIdOrPath, 
        query);
        http:Response response = check self.excelClient->get(path);
        return getRowArray(response);
    }

    @display {label: "Delete Row"}
    remote isolated function deleteRow(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Table Name"} string tableName,
                                        @display {label: "Row Index"} int rowIndex) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, ROWS, ITEM_AT + 
        OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + rowIndex.toString() + CLOSE_ROUND_BRACKET], workbookIdOrPath);

        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    @display {label: "Create Column"}
    remote isolated function createColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Table Name"} string tableName,
                                            @display {label: "Column Configuration"} Column column)
                                            returns @display {label: "Column"} Column|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, COLUMNS, ADD], 
        workbookIdOrPath);
        json payload = check column.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Column);
    }

    @display {label: "List Columns"}
    remote isolated function listColumns(@display {label: "Workbook ID or Path"} string workbookIdOrPath, 
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName, 
                                            @display {label: "Table Name"} string tableName, 
                                            @display {label: "Query"} Query? query = ()) returns
                                            @display {label: "Column"} Column[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, COLUMNS], 
        workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getColumnArray(response);
    }

    @display {label: "Delete Column"}
    remote isolated function deleteColumn(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Worksheet ID or Name"} string tableName,
                                            @display {label: "Index"} int index) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName, COLUMNS, ITEM_AT + 
        OPEN_ROUND_BRACKET + INDEX + EQUAL_SIGN + index.toString() + CLOSE_ROUND_BRACKET], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    @display {label: "Delete Table"}
    remote isolated function deleteTable(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Table Name"} string tableName) returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, TABLES, tableName], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    @display {label: "Add Chart"}
    remote isolated function addChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Chart Configuration"} ChartConfiguration chart) 
                                        returns @display {label: "Chart"} Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, ADD], workbookIdOrPath);
        json payload = check chart.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Chart);
    }

    @display {label: "Get Chart"}
    remote isolated function getChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Chart Name"} string chartName,
                                        @display {label: "Query"} Query? query = ())
                                        returns @display {label: "Chart"} Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName], workbookIdOrPath, 
        query);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Chart);
    }

    @display {label: "List Charts"}
    remote isolated function listCharts(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                        @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                        @display {label: "Query"} Query? query = ()) 
                                        returns @display {label: "Chart List"} Chart[]|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS], workbookIdOrPath, query);
        http:Response response = check self.excelClient->get(path);
        return getChartArray(response);
    }

    @display {label: "Update Chart"}
    remote isolated function updateChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Chart Name"} string chartName,
                                            @display {label: "Chart Configurtion"} Chart chart)
                                            returns @display {label: "Chart"} Chart|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName], workbookIdOrPath);
        json payload = check 
        chart.cloneWithType(json);
        http:Response response = check self.excelClient->patch(path, payload);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Chart);
    }

    @display {label: "Get Chart Image"}
    remote isolated function getChartImage(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Chart Name"} string chartName,
                                            @display {label: "Chart Width"} int? width = (),
                                            @display {label: "Chart Height"} int? height = (),
                                            @display {label: "Chart Fitting Mode"} string? fittingMode = ()) 
                                            returns @display {label: "Chart Image"} ChartImage|error {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName, IMAGE], 
        workbookIdOrPath);
        path = setOptionalParamsToPath(path, width, height, fittingMode);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(ChartImage);
    }

    @display {label: "Reset Chart Data"}
    remote isolated function resetChartData(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Chart Name"} string chartName,
                                            @display {label: "Chart Data"} Data data) 
                                            returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName, SET_DATA], 
        workbookIdOrPath);
        json payload = check data.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        _ = check handleResponse(response);
    }

    @display {label: "Set Chart Position"}
    remote isolated function setChartPosition(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                                @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                                @display {label: "Chart Name"} string chartName, 
                                                @display {label: "Chart Position"} ChartPosition position) 
                                                returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName, SET_POSITION], 
        workbookIdOrPath);
        json payload = check position.cloneWithType(json);
        http:Response response = check self.excelClient->post(path, payload);
        _ = check handleResponse(response);
    }

    @display {label: "Delete Chart"}
    remote isolated function deleteChart(@display {label: "Workbook ID or Path"} string workbookIdOrPath,
                                            @display {label: "Worksheet ID or Name"} string worksheetIdOrName,
                                            @display {label: "Chart Name"} string chartName) 
                                            returns error? {
        string path = check createRequestPath([WORKSHEETS, worksheetIdOrName, CHARTS, chartName], workbookIdOrPath);
        http:Response response = check self.excelClient->delete(path);
        _ = check handleResponse(response);
    }

    @display {label: "Get Workbook Application"}
    remote isolated function getWorkbookApplication(@display {label: "Workbook ID or Path"} string workbookIdOrPath) 
                                                    returns @display {label: "Workbook Application"} Application|error {
        string path = check createRequestPath([APPLICATION], workbookIdOrPath);
        http:Response response = check self.excelClient->get(path);
        map<json> handledResponse = check handleResponse(response);
        return handledResponse.cloneWithType(Application);
    }

    @display {label: "Calculate Workbook Application"}
    remote isolated function calculateWorkbookApplication(@display {label: "Workbook ID or Path"} string
                                                            workbookIdOrPath, @display {label: "Calculation Type"}
                                                            CalculationType 'type) returns error? {
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
public type ExcelConfiguration record {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig authConfig;
    http:ClientSecureSocket secureSocketConfig?;
};
