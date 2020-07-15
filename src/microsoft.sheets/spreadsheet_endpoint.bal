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

import ballerina/http;
import ballerina/log;
import ballerina/oauth2;

# Microsoft Spreadsheet Client Object.
public type MsSpreadsheetClient client object {
    private http:Client msSpreadsheetClient;
    private MicrosoftGraphConfiguration microsoftGraphConfig;

    public function init(MicrosoftGraphConfiguration msGraphConfig) {
        self.microsoftGraphConfig = msGraphConfig;
        oauth2:OutboundOAuth2Provider oauth2Provider3 = new ({
            accessToken: msGraphConfig.msInitialAccessToken,
            refreshConfig: {
                clientId: msGraphConfig.msClientId,
                clientSecret: msGraphConfig.msClientSecret,
                refreshToken: msGraphConfig.msRefreshToken,
                refreshUrl: msGraphConfig.msRefreshUrl,
                clientConfig: {
                    secureSocket: {
                    }
                }
            }
        });
        http:BearerAuthHandler oauth2Handler3 = new (oauth2Provider3);

        self.msSpreadsheetClient = new (msGraphConfig.baseUrl, {
                auth: {
                    authHandler: oauth2Handler3
                },
                followRedirects: {
                    enabled: true,
                    maxCount: 100
                },
                secureSocket: {
                }
            });
    }

    # Open a Workbook by the given name.
    # + path - Path to the workbook file
    # + workbookName - Name of the Workbook
    # + return - A Workbook client object on success, else returns an error
    public remote function openWorkbook(string path, string workbookName) returns Workbook|Error {
        Workbook workBook = new (self.msSpreadsheetClient, path, workbookName);

        return workBook;
    }
};

# Workbook Client Object.
public type Workbook client object {
    private http:Client workbookClient;
    private WorkbookProperties properties = {"path": "", "workbookName": ""};
    
    public function init(http:Client wbClient, string path, string workbookName) {
        self.workbookClient = wbClient;
        self.properties = {"path": path, "workbookName": workbookName};
    }

    # Get the properties of the workbook.
    # + return - Properties of the Workbook
    public function getProperties() returns WorkbookProperties {
        return self.properties;
    }

    # Open a worksheet on this workbook.
    # + worksheetName - name of the worksheet to be opened
    # + return - A Worksheet client object on success, else returns an error
    public remote function openWorksheet(string worksheetName) returns @tainted (Worksheet|Error) {
        http:Request request = new;
        http:Response|error response = self.workbookClient->get("/v1.0/me/drive/root:" + self.properties.path +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (response is error) {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", response);
        }

        http:Response httpResponse = <http:Response>response;

        if (httpResponse.statusCode != http:STATUS_CREATED) {
            return HttpResponseHandlingError("Error occurred while opening the worksheet.");
        }

        //If the worksheet is available we will get a JSON response with the worksheet's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            return TypeConversionError("Invalid identifier; expected a `map<json>` found " + typeOfRespone.toString());
        }

        map<json> payload = <map<json>>responseJson;

        json|error identifier = payload.id;

        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            return TypeConversionError("Invalid identifier; expected a `string` found " + typeOfIdentifier.toString());
        }

        string sheetId = <string>identifier;

        json|error sheetPosition = payload.position;

        if !(sheetPosition is int) {
            typedesc<any|error> typeOfPosition = typeof sheetPosition;
            return TypeConversionError("Invalid sheet position; expected a `int` found " + typeOfPosition.toString());
        }

        int position = <int>sheetPosition;

        //Populate a new worksheet object using the information from the properties as well as from the received JSON object
        Worksheet workSheet = new (self.workbookClient, self.properties.path,
            self.properties.workbookName, sheetId, worksheetName, position);
        return workSheet;
    }

    # Create a worksheet on this workbook.
    # + worksheetName - name of the worksheet to be created
    # + return - A Worksheet client object on success, else returns an error
    public remote function createWorksheet(string worksheetName) returns @tainted (Worksheet|Error) {
        //Make a POST request and create the worksheet
        http:Request request = new;
        json payload = {"name": worksheetName};
        request.setJsonPayload(payload);
        http:Response|error response = self.workbookClient->post("/v1.0/me/drive/root:" +
            self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets", request);

        if (response is error) {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", response);
        }

        http:Response httpResponse = <http:Response>response;

        if (httpResponse.statusCode != http:STATUS_CREATED) {
            log:printDebug("Error occurred while creating the worksheet. HTTP Status Code: " + 
            httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
            return HttpResponseHandlingError("Error occurred while creating the worksheet.");
        }

        //If the worksheet was created we will get a JSON response with the newly created worksheet's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            return TypeConversionError("Invalid response; expected a `map<json>` found " + typeOfRespone.toString());
        }

        map<json> responsePayload = <map<json>>responseJson;

        json|error identifier = responsePayload.id;

        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            return TypeConversionError("Invalid identifier; expected a `string` found " + typeOfIdentifier.toString());
        }

        string sheetId = <string>identifier;

        json|error sheetPosition = responsePayload.position;

        if !(sheetPosition is int) {
            typedesc<any|error> typeOfPosition = typeof sheetPosition;
            return TypeConversionError("Invalid sheet position; expected an `int` found " + typeOfPosition.toString());
        }

        int position = <int>sheetPosition;

        //Populate a new worksheet object using the information from the properties as well as from the received JSON object
        Worksheet workSheet = new (self.workbookClient, self.properties.path,
            self.properties.workbookName, sheetId, worksheetName, position);
        return workSheet;
    }

    # Remove a worksheet from this workbook.
    # + worksheetName - name of the worksheet to be removed
    # + return - boolean true on success, else returns an error
    public remote function removeWorksheet(string worksheetName) returns @tainted Error? {
        http:Request request = new;
        http:Response|error httpResponse = self.workbookClient->delete("/v1.0/me/drive/root:/" +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_NO_CONTENT) {
                return ();
            } else {
                log:printDebug("Error occurred while deleting the worksheet. HTTP Status Code: " + 
                httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
                return HttpResponseHandlingError("Error occurred while deleting the worksheet.");
            }
        } else {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", httpResponse);
        }
    }
};

# Worksheet Client Object.
public type Worksheet client object {
    private http:Client worksheetClient;
    private WorksheetProperties properties;

    public function init(http:Client wsClient, string path, string workbookName, string sheetId,
        string worksheetName, int position) {
        self.worksheetClient = wsClient;
        self.properties = {
            path: path,
            workbookName: workbookName,
            sheetId: sheetId,
            worksheetName: worksheetName,
            position: position
        };
    }

    # Get the properties of the Worksheet.
    # + return - Properties of the Worksheet
    public function getProperties() returns WorksheetProperties {
        return self.properties;
    }

    # Create a new Table on this Worksheet.
    # + tableName - name of the table to be created
    # + address - The location where the table should be created
    # + return - A Table client object on success, else returns an error
    public remote function createTable(string tableName, string address) returns @tainted (Table|Error) {
        //Make a POST request and create the table
        http:Request request = new;
        request.setJsonPayload({"name": tableName, "address": address, "hasHeaders": false});
        http:Response|error response = self.worksheetClient->post(<@untainted>("/v1.0/me/drive/root:" +
            self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets/" +
            self.properties.worksheetName + "/tables/add"), request);

        if (response is error) {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", response);
        }

        http:Response httpResponse = <http:Response>response;

        if (httpResponse.statusCode != http:STATUS_CREATED) {
            log:printDebug("Error occurred while creating the table. HTTP Status Code: " + 
            httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
            return HttpResponseHandlingError("Error occurred while creating the table.");
        }

        //If the table was created we will get a JSON response with the newly created table's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            return TypeConversionError("Invalid response; expected a `map<json>` found " + typeOfRespone.toString());
        }

        map<json> payload = <map<json>>responseJson;

        json|error nameItem = payload.name;

        if !(nameItem is string) {
            typedesc<any|error> typeOfNameItem = typeof nameItem;
            return TypeConversionError("Invalid name; expected a `string` found " + typeOfNameItem.toString());
        }

        string createdTableName = <string>nameItem;

        json|error newTableId = payload.id;

        if !(newTableId is string) {
            typedesc<any|error> typeOfTableId = typeof newTableId;
            return TypeConversionError("Invalid table ID; expected a `string` found " + typeOfTableId.toString());
        }

        //Populate a new Table object using the information from the properties as well as from the received JSON object
        Table resultsTable = <@untainted>new (self.worksheetClient, self.properties.path,
            self.properties.workbookName, self.properties.sheetId, self.properties.worksheetName, newTableId.toString(),
            address, createdTableName);

        if (createdTableName == tableName) {
            return resultsTable;
        }

        log:printDebug("Table created (" + createdTableName + ") carries different name than what " +
            "was passed as the table name (" + tableName + "). Now patching the table with the correct " +
            "table name.");

        Error? renameResult = resultsTable->renameTable(tableName);

        if (renameResult is ()) {
            return resultsTable;
        } else {
            return TableError("Error ocurred while renaming the created table.", renameResult);
        }
    }

    # Open a Table.
    # + tableName - name of the table to be opened
    # + return - A Table client object on success, else returns an error
    public remote function openTable(string tableName) returns @tainted (Table|Error) {
        //Make a GET request and retrieve the table information
        http:Request request = new;
        http:Response|error response = self.worksheetClient->get(<@untainted>("/v1.0/me/drive/root:" + self.properties.path +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" +
            self.properties.worksheetName + "/tables/" + tableName), request);

        if (response is error) {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", response);
        }

        http:Response httpResponse = <http:Response>response;

        if (httpResponse.statusCode != http:STATUS_OK) {
            log:printDebug("Error occurred while inserting data into table. HTTP Status Code: " + 
            httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
            return HttpResponseHandlingError("Error occurred while inserting data into table.");
        }

        //If the table exists we will get a JSON response with the table's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            return TypeConversionError("Invalid response; expected a `map<json>` found " + typeOfRespone.toString());
        }

        map<json> payload = <map<json>>responseJson;

        json|error identifier = payload.id;
        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            return TypeConversionError("Invalid identifier; expected a `string` found " + typeOfIdentifier.toString());
        }

        string sheetIdentifier = <string>identifier;

        //Address is not returned from the above API call. Hence the address is initialized to an empty string
        string address = "";

        Table resultsTable = <@untainted>new (self.worksheetClient, self.properties.path,
            self.properties.workbookName, self.properties.sheetId, self.properties.worksheetName,
            sheetIdentifier, address, tableName);

        return resultsTable;
    }
};

# Table Client Object.
public type Table client object {
    private http:Client tableClient;
    private TableProperties properties;

    public function init(http:Client tblClient, string path, string workbookName, string sheetId,
        string worksheetName, string tableId, string address, string tableName) {
        self.tableClient = tblClient;
        self.properties = {
            "path": path,
            "workbookName": workbookName,
            "sheetId": sheetId,
            "worksheetName": worksheetName,
            "tableId": tableId,
            "address": address,
            "tableName": tableName
        };
    }

    # Get the properties of the table.
    # + return - Properties of the Table
    public function getProperties() returns TableProperties {
        return self.properties;
    }

    # Insert data into the table.
    # + data - data to be inserted into the table
    # + return - boolean true on success, else returns an error
    public remote function insertDataIntoTable(json data) returns Error? {
        http:Request request = new;
        json payload = data;
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->post(<@untainted>("/v1.0/me/drive/root:" +
            self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets/" +
            self.properties.worksheetName + "/tables/" + self.properties.tableName + "/rows/add"), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_CREATED) {
                return ();
            } else {
                log:printDebug("Error occurred while inserting data into table. HTTP Status Code: " + 
                httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
                return HttpResponseHandlingError("Error occurred while inserting data into table.");
            }
        } else {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", httpResponse);
        }
    }

    # Rename the table.
    # + newTableName - new name to be used with the table
    # + return - boolean true on success, else returns an error
    public remote function renameTable(string newTableName) returns @tainted Error? {
        http:Request request = new;
        json payload = {"name": newTableName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch("/v1.0/me/drive/root:" + self.properties.path +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName +
            "/tables/" + self.properties.tableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_OK) {
                self.properties.tableName = newTableName;
                return ();
            } else {
                log:printDebug("Error occurred while renaming the table. HTTP Status Code: " + 
                httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
                return HttpResponseHandlingError("Error occurred while renaming the table.");
            }
        } else {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", httpResponse);
        }
    }

    # Set a table's header.
    # + columnID - ID of the tableColumn to change
    # + headerName - new name of the table header
    # + return - boolean true on success, else returns an error
    public remote function setTableHeader(int columnID, string headerName) returns Error? {
        http:Request request = new;
        json payload = {"name": headerName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch(<@untainted>("/v1.0/me/drive/root:" +
            self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets/" +
            self.properties.worksheetName + "/tables/" + self.properties.tableName + "/columns/" +
            columnID.toString()), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_OK) {
                return ();
            } else {
                log:printDebug("Error occurred while setting the table header. HTTP Status Code: " + 
                httpResponse.statusCode.toJsonString() + ", Reason : " + httpResponse.reasonPhrase);
                return HttpResponseHandlingError("Error occurred while setting the table header.");
            }
        } else {
            return HttpError("Error occurred while accessing the Microsoft Graph API.", httpResponse);
        }
    }
};

# Microsoft Graph client configuration.
# + baseUrl - The Microsoft Graph endpoint URL
# + msInitialAccessToken - Initial access token
# + msClientId - Microsoft client identifier
# + msClientSecret - client secret
# + msRefreshToken - refresh token
# + msRefreshUrl - refresh URL
# + trustStorePath - trust store path
# + trustStorePassword - trust store password
# + bearerToken - bearer token
# + clientConfig - OAuth2 direct token configuration
public type MicrosoftGraphConfiguration record {
    string baseUrl;
    string msInitialAccessToken;
    string msClientId;
    string msClientSecret;
    string msRefreshToken;
    string msRefreshUrl;
    string bearerToken;
};
