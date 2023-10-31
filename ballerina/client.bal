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
import microsoft.excel.excel;

# Ballerina Microsoft Excel client provides the capability to access Microsoft Graph Excel API to perform 
# CRUD (Create, Read, Update, and Delete) operations on Excel workbooks stored in Microsoft OneDrive for Business,
# SharePoint site or Group drive.
public isolated client class Client {

    final excel:Client excelClient;

    # Gets invoked to initialize the `connector`.
    #
    # + config - The configurations to be used when initializing the `connector` 
    # + serviceUrl - URL of the target service 
    # + return - An error if connector initialization failed 
    public isolated function init(excel:ConnectionConfig config, string serviceUrl = "https://graph.microsoft.com/v1.0/") returns error? {
        self.excelClient = check new (config, serviceUrl);
    }

    # Creates a new session for a workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + session - The payload to be create session
    # + return - A `excel:Session` or else an error on failure 
    remote isolated function createSession(string itemIdOrPath, excel:Session session) returns excel:Session|error {
        return self.excelClient->createSession(createSubPath(itemIdOrPath), session);
    }

    # Refresh the existing workbook session.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `http:Response` or error on failure  
    remote isolated function refreshSession(string itemIdOrPath, string sessionId) returns http:Response|error {
        return self.excelClient->refreshSession(createSubPath(itemIdOrPath), sessionId);
    }

    # Close the existing workbook session.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else error on failure 
    remote isolated function closeSession(string itemIdOrPath, string sessionId) returns http:Response|error {
        return self.excelClient->closeSession(createSubPath(itemIdOrPath), sessionId);
    }

    # Recalculate all currently opened workbooks in Excel.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + calculationMode - The payload to be calculate application
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else error on failure 
    remote isolated function calculateApplication(string itemIdOrPath, excel:CalculationMode calculationMode, string? sessionId = ()) returns http:Response|error {
        return self.excelClient->calculateApplication(createSubPath(itemIdOrPath), calculationMode, sessionId);
    }

    # Get the properties and relationships of the application.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `excel:Application` or else an error on failure 
    remote isolated function getApplication(string itemIdOrPath, string? sessionId = ()) returns excel:Application|error {
        return self.excelClient->getApplication(createSubPath(itemIdOrPath), sessionId);
    }

    # Retrieve a list of comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `excel:Comments` or else an error on failure 
    remote isolated function listComments(string itemIdOrPath, string? sessionId = ()) returns excel:Comments|error {
        return self.excelClient->listComments(createSubPath(itemIdOrPath), sessionId);
    }

    # Retrieve the properties and relationships of the comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + commentId - The ID of the comment to get
    # + sessionId - The ID of the session
    # + return - An `excel:Comment` or else an error on failure 
    remote isolated function getComment(string itemIdOrPath, string commentId, string? sessionId = ()) returns excel:Comment|error {
        return self.excelClient->getComment(createSubPath(itemIdOrPath), commentId, sessionId);
    }

    # Create a new reply of the comment.
    #
    # + itemId - The ID of the drive containing the workbook
    # + commentId - The ID of the comment to get
    # + reply - The payload to be create reply
    # + sessionId - The ID of the session
    # + return - Created. 
    remote isolated function createCommentReply(string itemId, string commentId, excel:Reply reply, string? sessionId = ()) returns excel:Reply|error {
        return self.excelClient->createCommentReply(itemId, commentId, reply, sessionId);
    }
}

