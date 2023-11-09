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
    # + session - The properties of the session to be created 
    # + return - A `excel:Session` or else an error on failure 
    remote isolated function createSession(string itemIdOrPath, excel:Session session) returns excel:Session|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createSessionWithItemPath(itemIdOrPath, session);
        }
        return self.excelClient->createSession(itemIdOrPath, session);
    }

    # Refreshes the existing workbook session.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `http:Response` or error on failure  
    remote isolated function refreshSession(string itemIdOrPath, string sessionId) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->refreshSessionWithItemPath(itemIdOrPath, sessionId);
        }
        return self.excelClient->refreshSession(itemIdOrPath, sessionId);
    }

    # Closes the existing workbook session.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else error on failure 
    remote isolated function closeSession(string itemIdOrPath, string sessionId) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->closeSessionWithItemPath(itemIdOrPath, sessionId);
        }
        return self.excelClient->closeSession(itemIdOrPath, sessionId);
    }

    # Retrieves a list of the worksheets.
    #
    # + itemId - The ID of the drive containing the workbooks
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Worksheets` or else an error on failure 
    remote isolated function listWorksheets(string itemId, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Worksheet[]|error {
        excel:Worksheets worksheets = check self.excelClient->listWorksheets(itemId, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        excel:Worksheet[]? value = worksheets.value;
        return value is excel:Worksheet[] ? value : [];
    }

    # Recalculates all currently opened workbooks in Excel.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + calculationMode - Details of the mode used to calculate the application
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else error on failure 
    remote isolated function calculateApplication(string itemIdOrPath, excel:CalculationMode calculationMode, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->calculateApplicationWithItemPath(itemIdOrPath, calculationMode, sessionId);
        }
        return self.excelClient->calculateApplication(itemIdOrPath, calculationMode, sessionId);
    }

    # Gets the properties and relationships of the application.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `excel:Application` or else an error on failure 
    remote isolated function getApplication(string itemIdOrPath, string? sessionId = ()) returns excel:Application|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getApplicationWithItemPath(itemIdOrPath, sessionId);
        }
        return self.excelClient->getApplication(itemIdOrPath, sessionId);
    }

    # Retrieves a list of comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + return - An `excel:Comments` or else an error on failure 
    remote isolated function listComments(string itemIdOrPath, string? sessionId = ()) returns excel:Comment[]|error {
        excel:Comments comments;
        if isItemPath(itemIdOrPath) {
            comments = check self.excelClient->listCommentsWithItemPath(itemIdOrPath, sessionId);
        } else {
            comments = check self.excelClient->listComments(itemIdOrPath, sessionId);
        }
        excel:Comment[]? value = comments.value;
        return value is excel:Comment[] ? value : [];
    }

    # Retrieves the properties and relationships of the comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + commentId - The ID of the comment to get
    # + sessionId - The ID of the session
    # + return - An `excel:Comment` or else an error on failure 
    remote isolated function getComment(string itemIdOrPath, string commentId, string? sessionId = ()) returns excel:Comment|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getCommentWithItemPath(itemIdOrPath, commentId, sessionId);
        }
        return self.excelClient->getComment(itemIdOrPath, commentId, sessionId);
    }

    # Creates a new reply of the comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + commentId - The ID of the comment to get
    # + reply - The properties of the reply to be created
    # + sessionId - The ID of the session
    # + return - An `excel:Reply` or else an error on failure
    remote isolated function createCommentReply(string itemIdOrPath, string commentId, excel:Reply reply, string? sessionId = ()) returns excel:Reply|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createCommentReplyWithItemPath(itemIdOrPath, commentId, reply, sessionId);
        }
        return self.excelClient->createCommentReply(itemIdOrPath, commentId, reply, sessionId);
    }

    # Retrieves the properties and relationships of the reply.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + commentId - The ID of the comment to get
    # + replyId - The ID of the reply
    # + sessionId - The ID of the session
    # + return - An `excel:Reply` or else an error on failure 
    remote isolated function getCommentReply(string itemIdOrPath, string commentId, string replyId, string? sessionId = ()) returns excel:Reply|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getCommentReplyWithItemPath(itemIdOrPath, commentId, replyId, sessionId);
        }
        return self.excelClient->getCommentReply(itemIdOrPath, commentId, replyId, sessionId);
    }

    # Lists the replies of the comment.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + commentId - The ID of the comment to get
    # + sessionId - The ID of the session
    # + return - An `excel:Replies` or else an error on failure 
    remote isolated function listCommentReplies(string itemIdOrPath, string commentId, string? sessionId = ()) returns excel:Reply[]|error {
        excel:Replies replies;
        if isItemPath(itemIdOrPath) {
            replies = check self.excelClient->listCommentRepliesWithItemPath(itemIdOrPath, commentId, sessionId);
        } else {
            replies = check self.excelClient->listCommentReplies(itemIdOrPath, commentId, sessionId);
        }
        excel:Reply[]? value = replies.value;
        return value is excel:Reply[] ? value : [];
    }

    # Retrieves a list of table row in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Rows` or else an error on failure  
    remote isolated function listWorkbookTableRows(string itemIdOrPath, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Row[]|error {
        excel:Rows rows; 
        if isItemPath(itemIdOrPath) {
            rows = check self.excelClient->listWorkbookTableRowsWithItemPath(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
           rows = check self.excelClient->listWorkbookTableRows(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Row[]? value = rows.value;
        return value is excel:Row[] ? value : [];
    }

    # Adds rows to the end of a table in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + row - The properties of the row to be added
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function createWorkbookTableRow(string itemIdOrPath, string tableIdOrName, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createWorkbookTableRowWithItemPath(itemIdOrPath, tableIdOrName, row, sessionId);
        }
        return self.excelClient->createWorkbookTableRow(itemIdOrPath, tableIdOrName, row, sessionId);
    }

    # Retrieves the properties and relationships of the table row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function getWorkbookTableRow(string itemIdOrPath, string tableIdOrName, int index, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableRowWithItemPath(itemIdOrPath, tableIdOrName, index, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorkbookTableRow(itemIdOrPath, tableIdOrName, index, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Deletes the row from the workbook table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorkbookTableRow(string itemIdOrPath, string tableIdOrName, int index, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorkbookTableRowWithItemPath(itemIdOrPath, tableIdOrName, index, sessionId);
        }
        return self.excelClient->deleteWorkbookTableRow(itemIdOrPath, tableIdOrName, index, sessionId);
    }

    # Gets the range associated with the entire row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorkbookTableRowRange(string itemIdOrPath, string tableIdOrName, int index, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableRowRangeWithItemPath(itemIdOrPath, tableIdOrName, index, sessionId);
        }
        return self.excelClient->getWorkbookTableRowRange(itemIdOrPath, tableIdOrName, index, sessionId);
    }

    # Updates the properties of table row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved
    # + row - Details of the table row to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function updateWorkbookTableRow(string itemIdOrPath, string tableIdOrName, int index, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorkbookTableRowWithItemPath(itemIdOrPath, tableIdOrName, index, row, sessionId);
        }
        return self.excelClient->updateWorkbookTableRow(itemIdOrPath, tableIdOrName, index, row, sessionId);
    }

    # Gets a row based on its position in the collection.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function getWorkbookTableRowWithIndex(string itemIdOrPath, string tableIdOrName, int index, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableRowWithIndexItemPath(itemIdOrPath, tableIdOrName, index, sessionId);
        }
        return self.excelClient->getWorkbookTableRowWithIndex(itemIdOrPath, tableIdOrName, index, sessionId);
    }

    # Adds a new worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheet - The properties of the worksheet to be created
    # + sessionId - The ID of the session
    # + return - An `excel:Worksheet` or else an error on failure 
    remote isolated function addWorksheet(string itemIdOrPath, excel:NewWorksheet worksheet, string? sessionId = ()) returns excel:Worksheet|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorksheetWithItemPath(itemIdOrPath, worksheet, sessionId);
        }
        return self.excelClient->addWorksheet(itemIdOrPath, worksheet, sessionId);
    }

    # Retrieves the properties and relationships of the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Worksheet` or else an error on failure 
    remote isolated function getWorksheet(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Worksheet|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheet(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Gets the used range of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + valuesOnly - A value indicating whether to return only the values in the used range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getNameUsedRange(string itemIdOrPath, string name, boolean valuesOnly, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameUsedRangeWithItemPath(itemIdOrPath, name, valuesOnly, sessionId);
        }
        return self.excelClient->getNameUsedRange(itemIdOrPath, name, valuesOnly, sessionId);
    }

    # Updates the properties of the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + worksheet - The properties of the worksheet to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Worksheet` or else an error on failure 
    remote isolated function updateWorksheet(string itemIdOrPath, string worksheetIdOrName, excel:Worksheet worksheet, string? sessionId = ()) returns excel:Worksheet|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetWithItemPath(itemIdOrPath, worksheetIdOrName, worksheet, sessionId);
        }
        return self.excelClient->updateWorksheet(itemIdOrPath, worksheetIdOrName, worksheet, sessionId);
    }

    # Deletes worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else error on failure  
    remote isolated function deleteWorksheet(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorksheetWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->deleteWorksheet(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Gets the range containing the single cell based on row and column numbers.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + row - Row number of the cell to be retrieved
    # + column - Column number of the cell to be retrieved
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetCell(string itemIdOrPath, string worksheetIdOrName, int row, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetCellWithItemPath(itemIdOrPath, worksheetIdOrName, row, column, sessionId);
        }
        return self.excelClient->getWorksheetCell(itemIdOrPath, worksheetIdOrName, row, column, sessionId);
    }

    # Gets the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetRange(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheetRange(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Retrieves a list of table in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Tables` or else an error on failure  
    remote isolated function listWorksheetTables(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Table[]|error {
        excel:Tables tables;
        if isItemPath(itemIdOrPath) {
            tables = check self.excelClient->listWorksheetTablesWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            tables = check self.excelClient->listWorksheetTables(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Table[]? value = tables.value;
        return value is excel:Table[] ? value : [];
    }

    # Adds a new table in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + 'table - Properties to create table
    # + sessionId - The ID of the session
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function addWorksheetTable(string itemIdOrPath, string worksheetIdOrName, excel:NewTable 'table, string? sessionId = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorksheetTableWithItemPath(itemIdOrPath, worksheetIdOrName, 'table, sessionId);
        }
        return self.excelClient->addWorksheetTable(itemIdOrPath, worksheetIdOrName, 'table, sessionId);
    }

    # Retrieves a list of charts.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Charts` or else an error on failure  
    remote isolated function listCharts(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Chart[]|error {
        excel:Charts charts;
        if isItemPath(itemIdOrPath) {
            charts = check self.excelClient->listChartsWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            charts = check self.excelClient->listCharts(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Chart[]? value = charts.value;
        return value is excel:Chart[] ? value : []; 
    }

    # Creates a new chart.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chart - Properties to create chart
    # + sessionId - The ID of the session
    # + return - An `excel:Chart` or else an error on failure  
    remote isolated function addChart(string itemIdOrPath, string worksheetIdOrName, excel:NewChart chart, string? sessionId = ()) returns excel:Chart|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addChartWithItemPath(itemIdOrPath, worksheetIdOrName, chart, sessionId);
        }
        return self.excelClient->addChart(itemIdOrPath, worksheetIdOrName, chart, sessionId);
    }

    # Retrieves a list of named items associated with the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function listWorksheetNames(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:NamedItem[]|error {
        excel:NamedItems namedItems;
        if isItemPath(itemIdOrPath) {
            namedItems = check self.excelClient->listWorksheetNamedItemWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            namedItems = check self.excelClient->listWorksheetNamedItem(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:NamedItem[]? value = namedItems.value;
        return value is excel:NamedItem[] ? value : [];
    }
    
    # Retrieves a list of the workbook pivot table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:PivotTables` or else an error on failure 
    remote isolated function listWorksheetPivotTables(string itemIdOrPath, string worksheetIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:PivotTable[]|error {
        excel:PivotTables pivotTables;
        if isItemPath(itemIdOrPath) {
            pivotTables = check self.excelClient->listPivotTablesWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            pivotTables = check self.excelClient->listPivotTables(itemIdOrPath, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:PivotTable[]? value = pivotTables.value;
        return value is excel:PivotTable[] ? value : [];
    }

    # Retrieves the properties and relationships of the pivot table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + pivotTableId - The ID of the pivot table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:PivotTables` or else an error on failure 
    remote isolated function getPivotTable(string itemIdOrPath, string worksheetIdOrName, string pivotTableId, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:PivotTable|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getPivotTableWithItemPath(itemIdOrPath, worksheetIdOrName, pivotTableId, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getPivotTable(itemIdOrPath, worksheetIdOrName, pivotTableId, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Refreshes the pivot table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + pivotTableId - The ID of the pivot table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function refreshPivotTable(string itemIdOrPath, string worksheetIdOrName, string pivotTableId, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->refreshPivotTableWithItemPath(itemIdOrPath, worksheetIdOrName, pivotTableId, sessionId);
        }
        return self.excelClient->refreshPivotTable(itemIdOrPath, worksheetIdOrName, pivotTableId, sessionId);
    }

    # Refreshes all pivot tables within given worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure
    remote isolated function refreshAllPivotTables(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->refreshAllPivotTablesWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->refreshAllPivotTables(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Retrieves the properties and relationships of range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetRangeWithAddress(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeWithAddressItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheetRangeWithAddress(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Updates the properties of range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + range - Details of the range to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateWorksheetRangeWithAddress(string itemIdOrPath, string worksheetIdOrName, string address, excel:Range range, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetRangeWithAddressItemPath(itemIdOrPath, worksheetIdOrName, address, range, sessionId);
        }
        return self.excelClient->updateWorksheetRangeWithAddress(itemIdOrPath, worksheetIdOrName, address, range, sessionId);
    }

    # Retrieves the properties and relationships of range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Updates the properties of range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + range - Details of the range to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:Range range, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, range, sessionId);
        }
        return self.excelClient->updateColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, range, sessionId);
    }

    # Updates the properties of range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + namedItemName - The name of the named item
    # + range - The properties of the range to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateNameRange(string itemIdOrPath, string namedItemName, excel:Range range, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateNameRangeWithItemPath(itemIdOrPath, namedItemName, range, sessionId);
        }
        return self.excelClient->updateNameRange(itemIdOrPath, namedItemName, range, sessionId);
    }

    # Gets the range containing the single cell based on row and column numbers.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + namedItemName - The name of the named item
    # + row - Row number of the cell to be retrieved
    # + column - Column number of the cell to be retrieved
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getNameRangeCell(string itemIdOrPath, string namedItemName, int row, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeCell(itemIdOrPath, namedItemName, row, column, sessionId);
        }
        return self.excelClient->getNameRangeCell(itemIdOrPath, namedItemName, row, column, sessionId);
    }

    # Gets the range containing the single cell based on row and column numbers.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + row - Row number of the cell to be retrieved
    # + column - Column number of the cell to be retrieved
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetRangeCell(string itemIdOrPath, string worksheetIdOrName, int row, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeCellWithItemPath(itemIdOrPath, worksheetIdOrName, row, column, sessionId);
        }
        return self.excelClient->getWorksheetRangeCell(itemIdOrPath, worksheetIdOrName, row, column, sessionId);
    }

    # Gets the range containing the single cell based on row and column numbers.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + row - Row number of the cell to be retrieved
    # + column - Column number of the cell to be retrieved
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetRangeCellWithAddress(string itemIdOrPath, string worksheetIdOrName, string address, int row, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeCellWithAddressItemPath(itemIdOrPath, worksheetIdOrName, address, row, column, sessionId);
        }
        return self.excelClient->getWorksheetRangeCellWithAddress(itemIdOrPath, worksheetIdOrName, address, row, column, sessionId);
    }

    # Gets the range object containing the single cell based on row and column numbers.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + row - Row number of the cell to be retrieved. Zero-indexed.
    # + column - Column number of the cell to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getColumnRangeCell(string itemIdOrPath, string tableIdOrName, string columnIdOrName, int row, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeCellWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, row, column, sessionId);
        }
        return self.excelClient->getColumnRangeCell(itemIdOrPath, tableIdOrName, columnIdOrName, row, column, sessionId);
    }

    # Gets a column contained in the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + column - Column number of the cell to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeColumn(string itemIdOrPath, string name, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeColumnWithItemPath(itemIdOrPath, name, column, sessionId);
        }
        return self.excelClient->getNameRangeColumn(itemIdOrPath, name, column, sessionId);
    }

    # Gets a column contained in the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + column - Column number of the cell to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeColumn(string itemIdOrPath, string worksheetIdOrName, string address, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeColumnWithItemPath(itemIdOrPath, worksheetIdOrName, address, column,sessionId);
        }
        return self.excelClient->getWorksheetRangeColumn(itemIdOrPath, worksheetIdOrName, address, column, sessionId);
    }

    # Gets a column contained in the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + column - Column number of the cell to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, int column, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, column,sessionId);
        }
        return self.excelClient->getColumnRangeColumn(itemIdOrPath, tableIdOrName, columnIdOrName, column, sessionId);
    }

    # Gets a certain number of columns to the right of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetColumnsAfterRange(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetColumnsAfterRangeWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetColumnsAfterRange(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Gets a certain number of columns to the right of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + columnCount - The number of columns to include in the resulting range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetColumnsAfterRangeWithCount(string itemIdOrPath, string worksheetIdOrName, int columnCount, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetColumnsAfterRangeWithCountItemPath(itemIdOrPath, worksheetIdOrName, columnCount, sessionId);
        }
        return self.excelClient->getWorksheetColumnsAfterRangeWithCount(itemIdOrPath, worksheetIdOrName, columnCount, sessionId);
    }

    # Gets a certain number of columns to the left of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetColumnsBeforeRange(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetColumnsBeforeRangeWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetColumnsBeforeRange(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Gets a certain number of columns to the left of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + columnCount - The number of columns to include in the resulting range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetColumnsBeforeRangeWithCount(string itemIdOrPath, string worksheetIdOrName, int columnCount, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetColumnsBeforeRangeWithCountItemPath(itemIdOrPath, worksheetIdOrName, columnCount, sessionId);
        }
        return self.excelClient->getWorksheetColumnsBeforeRangeWithCount(itemIdOrPath, worksheetIdOrName, columnCount, sessionId);
    }

    # Gets the range that represents the entire column of the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + namedItemName - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeEntireColumn(string itemIdOrPath, string namedItemName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeEntireColumnWithItemPath(itemIdOrPath, namedItemName, sessionId);
        }
        return self.excelClient->getNameRangeEntireColumn(itemIdOrPath, namedItemName, sessionId);
    }

    # Gets the range that represents the entire column of the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeEntireColumn(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeEntireColumnWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getWorksheetRangeEntireColumn(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Gets the range that represents the entire column of the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeEntireColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeEntireColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getColumnRangeEntireColumn(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range that represents the entire row of the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + namedItemName - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeEntireRow(string itemIdOrPath, string namedItemName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeEntireRowWithItemPath(itemIdOrPath, namedItemName, sessionId);
        }
        return self.excelClient->getNameRangeEntireRow(itemIdOrPath, namedItemName, sessionId);
    }

    # Gets the range that represents the entire row of the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getRangeEntireRow(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeEntireRowWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getWorksheetRangeEntireRow(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Gets the range that represents the entire row of the range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeEntireRow(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeEntireRowWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getColumnRangeEntireRow(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the last cell within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeLastCell(string itemIdOrPath, string name, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeLastCellWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->getNameRangeLastCell(itemIdOrPath, name, sessionId);
    }

    # Gets the last cell within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getRangeLastCell(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeLastCellWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getWorksheetRangeLastCell(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Gets the last cell within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeLastCell(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeLastCellWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getColumnRangeLastCell(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the last column within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeLastColumn(string itemIdOrPath, string name, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeLastColumnWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->getNameRangeLastColumn(itemIdOrPath, name, sessionId);
    }

    # Gets the last column within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeLastColumn(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeLastColumnWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getWorksheetRangeLastColumn(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Gets the last column within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeLastColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeLastColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getColumnRangeLastColumn(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the last row within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getNameRangeLastRow(string itemIdOrPath, string name, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeLastRowWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->getNameRangeLastRow(itemIdOrPath, name, sessionId);
    }

    # Gets the last row within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeLastRow(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeLastRowWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getWorksheetRangeLastRow(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Gets the last row within the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getColumnRangeLastRow(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeLastRowWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getColumnRangeLastRow(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets a certain number of rows above a given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRowsAboveRange(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRowsAboveRangeWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetRowsAboveRange(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Gets a certain number of rows above a given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + rowCount - The number of rows to include in the resulting range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRowsAboveRangeWithCount(string itemIdOrPath, string worksheetIdOrName, int rowCount, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRowsAboveRangeWithCountItemPath(itemIdOrPath, worksheetIdOrName, rowCount, sessionId);
        }
        return self.excelClient->getWorksheetRowsAboveRangeWithCount(itemIdOrPath, worksheetIdOrName, rowCount, sessionId);
    }

    # Gets a certain number of columns to the left of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeRowsBelow(string itemIdOrPath, string worksheetIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRowsBelowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetRowsBelowRange(itemIdOrPath, worksheetIdOrName, sessionId);
    }

    # Gets a certain number of columns to the left of the given range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + rowCount - The number of rows to include in the resulting range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetRangeRowsBelowWithCount(string itemIdOrPath, string worksheetIdOrName, int rowCount, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRowsBelowRangeWithCountItemPath(itemIdOrPath, worksheetIdOrName, rowCount, sessionId);
        }
        return self.excelClient->getWorksheetRowsBelowRangeWithCount(itemIdOrPath, worksheetIdOrName, rowCount, sessionId);
    }

    # Gets the used range of the worksheet with in the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + valuesOnly - A value indicating whether to return only the values in the used range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetUsedRange(string itemIdOrPath, string worksheetIdOrName, string address, boolean valuesOnly, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetUsedRangeWithItemPath(itemIdOrPath, worksheetIdOrName, address, valuesOnly, sessionId);
        }
        return self.excelClient->getWorksheetUsedRange(itemIdOrPath, worksheetIdOrName, address, valuesOnly, sessionId);
    }

    # Get the used range of the given range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + valuesOnly - A value indicating whether to return only the values in the used range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getColumnUsedRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, boolean valuesOnly, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnUsedRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, valuesOnly, sessionId);
        }
        return self.excelClient->getColumnUsedRange(itemIdOrPath, tableIdOrName, columnIdOrName, valuesOnly, sessionId);
    }

    # Clear range values such as format, fill, and border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + applyTo - Determines the type of clear action
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function clearNameRange(string itemIdOrPath, string name, excel:ApplyTo applyTo, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearNameRangeWithItemPath(itemIdOrPath, name, applyTo, sessionId);
        }
        return self.excelClient->clearNameRange(itemIdOrPath, name, applyTo, sessionId);
    }

    # Clear range values such as format, fill, and border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + applyTo - Determines the type of clear action
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function clearWorksheetRange(string itemIdOrPath, string worksheetIdOrName, string address, excel:ApplyTo applyTo, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearWorksheetRangeWithItemPath(itemIdOrPath, worksheetIdOrName, address, applyTo, sessionId);
        }
        return self.excelClient->clearWorksheetRange(itemIdOrPath, worksheetIdOrName, address, applyTo, sessionId);
    }

    # Clear range values such as format, fill, and border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + applyTo - Determines the type of clear action
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function clearColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:ApplyTo applyTo, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, applyTo, sessionId);
        }
        return self.excelClient->clearColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, applyTo, sessionId);
    }

    # Deletes the cells associated with the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteNameRangeCell(string itemIdOrPath, string name, excel:Shift shift, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteNameRangeCellWithItemPath(itemIdOrPath, name, shift, sessionId);
        }
        return self.excelClient->deleteNameRangeCell(itemIdOrPath, name, shift, sessionId);
    }

    # Deletes the cells associated with the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorksheetRangeCell(string itemIdOrPath, string worksheetIdOrName, string address, excel:Shift shift, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorksheetRangeCellWithItemPath(itemIdOrPath, worksheetIdOrName, address, shift, sessionId);
        }
        return self.excelClient->deleteWorksheetRangeCell(itemIdOrPath, worksheetIdOrName, address, shift, sessionId);
    }

    # Deletes the cells associated with the range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteColumnRangeCell(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:Shift shift, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteColumnRangeCellWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, shift, sessionId);
        }
        return self.excelClient->deleteColumnRangeCell(itemIdOrPath, tableIdOrName, columnIdOrName, shift, sessionId);
    }

    # Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function insertNameRange(string itemIdOrPath, string name, excel:Shift shift, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->insertNameRangeWithItemPath(itemIdOrPath, name, shift, sessionId);
        }
        return self.excelClient->insertNameRange(itemIdOrPath, name, shift, sessionId);
    }

    # Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function insertWorksheetRange(string itemIdOrPath, string worksheetIdOrName, string address, excel:Shift shift, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->insertWorksheetRangeWithItemPath(itemIdOrPath, worksheetIdOrName, address, shift, sessionId);
        }
        return self.excelClient->insertWorksheetRange(itemIdOrPath, worksheetIdOrName, address, shift, sessionId);
    }

    # Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + shift - Represents the ways to shift the cells
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function insertColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:Shift shift, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->insertColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, shift, sessionId);
        }
        return self.excelClient->insertColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, shift, sessionId);
    }

    # Merge the range cells into one region in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + across - The properties to the merge range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function mergeNameRange(string itemIdOrPath, string name, excel:Across across, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->mergeNameRangeWithItemPath(itemIdOrPath, name, across, sessionId);
        }
        return self.excelClient->mergeNameRange(itemIdOrPath, name, across, sessionId);
    }

    # Merge the range cells into one region in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + across - The properties to the merge range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function mergeWorksheetRange(string itemIdOrPath, string worksheetIdOrName, string address, excel:Across across, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->mergeWorksheetRangeWithItemPath(itemIdOrPath, worksheetIdOrName, address, across, sessionId);
        }
        return self.excelClient->mergeWorksheetRange(itemIdOrPath, worksheetIdOrName, address, across, sessionId);
    }

    # Merge the range cells into one region in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + across - The properties to the merge range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function mergeColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:Across across, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->mergeColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, across, sessionId);
        }
        return self.excelClient->mergeColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, across, sessionId);
    }

    # Unmerge the range cells into separate cells.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function unmergeNameRange(string itemIdOrPath, string name, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->unmergeNameRangeWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->unmergeNameRange(itemIdOrPath, name, sessionId);
    }

    # Unmerge the range cells into separate cells.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function unmergeWorksheetRange(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->unmergeWorksheetRangeWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->unmergeWorksheetRange(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Unmerge the range cells into separate cells.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function unmergeColumnRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->unmergeColumnRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->unmergeColumnRange(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Retrieve the properties and relationships of the range format
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getNameRangeFormat(string itemIdOrPath, string name, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNameRangeFormatWithItemPath(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getNameRangeFormat(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Update the properties of range format.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + rangeFormat - Properties of the range format to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateNameRangeFormat(string itemIdOrPath, string name, excel:RangeFormat rangeFormat, string? sessionId = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateNameRangeFormat(itemIdOrPath, name, rangeFormat, sessionId);
        }
        return self.excelClient->updateNameRangeFormat(itemIdOrPath, name, rangeFormat, sessionId);
    }

    # Retrieve the properties and relationships of the range format
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getWorksheetRangeFormat(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetRangeFormat(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheetRangeFormat(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Update the properties of range format.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + rangeFormat - Properties of the range format to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateWorksheetRangeFormat(string itemIdOrPath, string worksheetIdOrName, string address, excel:RangeFormat rangeFormat, string? sessionId = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetRangeFormatWithItemPath(itemIdOrPath, worksheetIdOrName, address, rangeFormat, sessionId);
        }
        return self.excelClient->updateWorksheetRangeFormat(itemIdOrPath, worksheetIdOrName, address, rangeFormat, sessionId);
    }

    # Retrieve the properties and relationships of the range format
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function getColumnRangeFormat(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getColumnRangeFormatWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getColumnRangeFormat(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Update the properties of range format.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + rangeFormat - Properties of the range format to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure 
    remote isolated function updateColumnRangeFormat(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:RangeFormat rangeFormat, string? sessionId = ()) returns excel:RangeFormat|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateColumnRangeFormatWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, rangeFormat, sessionId);
        }
        return self.excelClient->updateColumnRangeFormat(itemIdOrPath, tableIdOrName, columnIdOrName, rangeFormat, sessionId);
    }

    # Create a new range border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + rangeBorder - Details of the range border to be created 
    # + sessionId - The ID of the session
    # + return - An `excel:RangeBorder` or else an error on failure 
    remote isolated function createNameRangeBorder(string itemIdOrPath, string name, excel:RangeBorder rangeBorder, string? sessionId = ()) returns excel:RangeBorder|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createNameRangeBorderWithItemPath(itemIdOrPath, name, rangeBorder, sessionId);
        }
        return self.excelClient->createNameRangeBorder(itemIdOrPath, name, rangeBorder, sessionId);
    }

    # Create a new range border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + rangeBorder - Details of the range border to be created 
    # + sessionId - The ID of the session
    # + return - An `excel:RangeBorder` or else an error on failure 
    remote isolated function createWorksheetRangeBorder(string itemIdOrPath, string worksheetIdOrName, string address, excel:RangeBorder rangeBorder, string? sessionId = ()) returns excel:RangeBorder|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createWorksheetRangeBorderWithItemPath(itemIdOrPath, worksheetIdOrName, address, rangeBorder, sessionId);
        }
        return self.excelClient->createWorksheetRangeBorder(itemIdOrPath, worksheetIdOrName, address, rangeBorder, sessionId);
    }

    # Create a new range border.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + rangeBorder - Properties to create range border 
    # + sessionId - The ID of the session
    # + return - An `excel:RangeBorder` or else an error on failure 
    remote isolated function createColumnRangeBorder(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:RangeBorder rangeBorder, string? sessionId = ()) returns excel:RangeBorder|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createColumnRangeBorderWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, rangeBorder, sessionId);
        }
        return self.excelClient->createColumnRangeBorder(itemIdOrPath, tableIdOrName, columnIdOrName, rangeBorder, sessionId);
    }

    # Retrieves a list of range borders.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:RangeBorders` or else an error on failure 
    remote isolated function listNameRangeBorders(string itemIdOrPath, string name, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeBorder[]|error {
        excel:RangeBorders rangeBorders;
        if isItemPath(itemIdOrPath) {
            rangeBorders = check self.excelClient->listNameRangeBordersWithItemPath(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            rangeBorders = check self.excelClient->listNameRangeBorders(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:RangeBorder[]? value = rangeBorders.value;
        return value is excel:RangeBorder[] ? value : [];
    }

    # Retrieves a list of range borders.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:RangeBorders` or else an error on failure 
    remote isolated function listColumnRangeBorders(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeBorder[]|error {
        excel:RangeBorders rangeBorders;
        if isItemPath(itemIdOrPath) {
            rangeBorders = check self.excelClient->listColumnRangeBordersWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        rangeBorders = check self.excelClient->listColumnRangeBorders(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        excel:RangeBorder[]? value = rangeBorders.value;
        return value is excel:RangeBorder[] ? value : [];
    }

    # Retrieves a list of range borders.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:RangeBorders` or else an error on failure 
    remote isolated function listRangeBorders(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:RangeBorder[]|error {
        excel:RangeBorders rangeBorders;
        if isItemPath(itemIdOrPath) {
            rangeBorders = check self.excelClient->listColumnRangeBordersWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            rangeBorders = check self.excelClient->listColumnRangeBorders(itemIdOrPath, worksheetIdOrName, address, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:RangeBorder[]? value = rangeBorders.value;
        return value is excel:RangeBorder[] ? value : [];
    }

    # Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitNameRangeColumns(string itemIdOrPath, string name, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitNameRangeColumnsWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->autofitNameRangeColumns(itemIdOrPath, name, sessionId);
    }
    
    # Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitWorksheetRangeColumns(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitWorksheetRangeColumnsWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->autofitWorksheetRangeColumns(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }
    
    # Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitColumnRangeColumns(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitColumnRangeColumnsWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->autofitColumnRangeColumns(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitNameRangeRows(string itemIdOrPath, string name, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitNameRangeRowsWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->autofitNameRangeRows(itemIdOrPath, name, sessionId);
    }
    
    # Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitWorksheetRangeRows(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitWorksheetRangeRowsWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->autofitWorksheetRangeRows(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function autofitColumnRangeRows(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->autofitColumnRangeRowsWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->autofitColumnRangeRows(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Perform a sort operation.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item
    # + rangeSort - The properties to the sort operation
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function performNameRangeSort(string itemIdOrPath, string name, excel:RangeSort rangeSort, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->performNameRangeSortWithItemPath(itemIdOrPath, name, rangeSort, sessionId);
        }
        return self.excelClient->performNameRangeSort(itemIdOrPath, name, rangeSort, sessionId);
    }
    
    # Perform a sort operation.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + rangeSort - The properties to the sort operation
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function performRangeSort(string itemIdOrPath, string worksheetIdOrName, string address, excel:RangeSort rangeSort, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->performWorksheetRangeSortWithItemPath(itemIdOrPath, worksheetIdOrName, address, rangeSort, sessionId);
        }
        return self.excelClient->performWorksheetRangeSort(itemIdOrPath, worksheetIdOrName, address, rangeSort, sessionId);
    }

    # Perform a sort operation.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + rangeSort - The properties to the perform sort operation
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure 
    remote isolated function performColumnRangeSort(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:RangeSort rangeSort, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->performColumnRangeSortWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, rangeSort, sessionId);
        }
        return self.excelClient->performColumnRangeSort(itemIdOrPath, tableIdOrName, columnIdOrName, rangeSort, sessionId);
    }

    # Get the resized range of a range.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + deltaRows - The number of rows to expand or contract the bottom-right corner of the range by. If deltaRows is positive, the range will be expanded. If deltaRows is negative, the range will be contracted.
    # + deltaColumns - The number of columns to expand or contract the bottom-right corner of the range by. If deltaColumns is positive, the range will be expanded. If deltaColumns is negative, the range will be contracted.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getResizedRange(string itemIdOrPath, string worksheetIdOrName, int deltaRows, int deltaColumns, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getResizedRangeWithItemPath(itemIdOrPath, worksheetIdOrName, deltaRows, deltaColumns, sessionId);
        }
        return self.excelClient->getResizedRange(itemIdOrPath, worksheetIdOrName, deltaRows, deltaColumns, sessionId);
    }

    # Get the range visible from a filtered range
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + address - The address of the range
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getVisibleView(string itemIdOrPath, string worksheetIdOrName, string address, string? sessionId = ()) returns excel:RangeView|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getVisibleViewWithItemPath(itemIdOrPath, worksheetIdOrName, address, sessionId);
        }
        return self.excelClient->getVisibleView(itemIdOrPath, worksheetIdOrName, address, sessionId);
    }

    # Retrieve the properties and relationships of table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function getWorkbookTable(string itemIdOrPath, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableWithItemPath(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorkbookTable(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Retrieve the properties and relationships of table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function getWorksheetTable(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheetTable(itemIdOrPath, tableIdOrName, worksheetIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Create a new table in the workbook
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + table - The properties to create table
    # + sessionId - The ID of the session
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function addWorkbookTable(string itemIdOrPath, excel:NewTable 'table, string? sessionId = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorkbookTableWithItemPath(itemIdOrPath, 'table, sessionId);
        }
        return self.excelClient->addWorkbookTable(itemIdOrPath, 'table, sessionId);
    }

    # Retrieve a list of table in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Tables` or else an error on failure  
    remote isolated function listWorkbookTables(string itemIdOrPath, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Table[]|error {
        excel:Tables tables;
        if isItemPath(itemIdOrPath) {
            tables = check self.excelClient->listWorkbookTablesWithItemPath(itemIdOrPath, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            tables = check self.excelClient->listWorkbookTables(itemIdOrPath, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Table[]? value = tables.value;
        return value is excel:Table[] ? value : [];
    }

    # Deletes the table from the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorkbookTable(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorkbookTableWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->deleteWorkbookTable(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Update the properties of table in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + table - The properties of the table to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function updateWorkbookTable(string itemIdOrPath, string tableIdOrName, excel:Table 'table, string? sessionId = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorkbookTableWithItemPath(itemIdOrPath, tableIdOrName, 'table, sessionId);
        }
        return self.excelClient->updateWorkbookTable(itemIdOrPath, tableIdOrName, 'table, sessionId);
    }

    # Deletes the table from the worksheet
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorksheetTable(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorksheetTableWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->deleteWorksheetTable(itemIdOrPath, tableIdOrName, worksheetIdOrName, sessionId);
    }

    # Update the properties of table in the worksheet
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + table - The properties of the table to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Table` or else an error on failure  
    remote isolated function updateWorksheetTable(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, excel:Table 'table, string? sessionId = ()) returns excel:Table|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetTableWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, 'table, sessionId);
        }
        return self.excelClient->updateWorksheetTable(itemIdOrPath, worksheetIdOrName, tableIdOrName, 'table, sessionId);
    }

    # Gets the range associated with the data body of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorkbookTableBodyRange(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableBodyRangeWithItemPath(itemIdOrPath, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorkbookTableBodyRange(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Gets the range associated with the data body of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetTableBodyRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableBodyRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorksheetTableBodyRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }

    # Gets the range associated with header row of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorkbookTableHeaderRowRange(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableHeaderRowRangeWithItemPath(itemIdOrPath, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorkbookTableHeaderRowRange(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Gets the range associated with header row of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetTableHeaderRowRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableHeaderRowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorksheetTableHeaderRowRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }

    # Gets the range associated with totals row of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorkbookTableTotalRowRange(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableTotalRowRangeWithItemPath(itemIdOrPath, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorkbookTableTotalRowRange(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Gets the range associated with totals row of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetTableTotalRowRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableTotalRowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName,sessionId);
        }
        return self.excelClient->getWorksheetTableTotalRowRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }

    # Clears all the filters currently applied on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function clearWorkbookTableFilters(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearWorkbookTableFiltersWithItemPath(itemIdOrPath, tableIdOrName,sessionId);
        }
        return self.excelClient->clearWorkbookTableFilters(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Clears all the filters currently applied on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Responsee` or else an error on failure  
    remote isolated function clearWorksheetTableFilters(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearWorksheetTableFiltersWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName,sessionId);
        }
        return self.excelClient->clearWorksheetTableFilters(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }

    # Converts the table into a normal range of cells. All data is preserved.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function convertWorkbookTableToRange(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->convertWorkbookTableToRangeWithItemPath(itemIdOrPath, tableIdOrName,sessionId);
        }
        return self.excelClient->convertWorkbookTableToRange(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Converts the table into a normal range of cells. All data is preserved.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function convertWorksheetTableToRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->convertWorksheetTableToRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName,sessionId);
        }
        return self.excelClient->convertWorksheetTableToRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }

    # Reapplies all the filters currently on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function reapplyWorkbookTableFilters(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->reapplyWorkbookTableFiltersWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->reapplyWorkbookTableFilters(itemIdOrPath, tableIdOrName, sessionId);
    } 

    # Reapplies all the filters currently on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function reapplyWorksheetTableFilters(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->reapplyWorksheetTableFiltersWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->reapplyWorksheetTableFilters(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    } 

    # Retrieve the properties and relationships of table sort.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderby - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:TableSort` or else an error on failure  
    remote isolated function getWorkbookTableSort(string itemIdOrPath, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderby = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:TableSort|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableSortWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->getWorkbookTableSort(itemIdOrPath, tableIdOrName, sessionId);
    }  

    # Retrieve the properties and relationships of table sort.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderby - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:TableSort` or else an error on failure  
    remote isolated function getWorksheetTableSort(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderby = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:TableSort|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableSortWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetTableSort(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }  

    # Perform a sort operation to the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function performWorkbookTableSort(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->performWorkbookTableSortWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->performWorkbookTableSort(itemIdOrPath, tableIdOrName, sessionId);
    }  

    # Perform a sort operation to the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + tableSort - The properties to the sort operation
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function performWorksheetTableSort(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, excel:TableSort tableSort,  string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->performWorksheetTableSortWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, tableSort, sessionId);
        }
        return self.excelClient->performWorksheetTableSort(itemIdOrPath, worksheetIdOrName, tableIdOrName, tableSort, sessionId);
    }  

    # Clears the sorting that is currently on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function clearWorkbookTableSort(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearWorkbookTableSortWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->clearWorkbookTableSort(itemIdOrPath, tableIdOrName, sessionId);
    }  

    # Clears the sorting that is currently on the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function clearWorksheetTableSort(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->clearWorksheetTableSortWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->clearWorksheetTableSort(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }  

    # Reapplies the current sorting parameters to the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function reapplyWorkbookTableSort(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->reapplyWorkbookTableSortWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->reapplyWorkbookTableSort(itemIdOrPath, tableIdOrName, sessionId);
    }  

    # Reapplies the current sorting parameters to the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function reapplyWorksheetTableSort(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->reapplyWorksheetTableSortWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->reapplyWorksheetTableSort(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }  

    # Get the range associated with the entire table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorkbookTableRange(string itemIdOrPath, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableRangeWithItemPath(itemIdOrPath, tableIdOrName, sessionId);
        }
        return self.excelClient->getWorkbookTableRange(itemIdOrPath, tableIdOrName, sessionId);
    }

    # Get the range associated with the entire table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetTableRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetTableRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId);
    }  

    # Retrieve a list of table row in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Rows` or else an error on failure  
    remote isolated function listWorksheetTableRows(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Row[]|error {
        excel:Rows rows; 
        if isItemPath(itemIdOrPath) {
            rows = check self.excelClient->listWorksheetTableRowsWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            rows = check self.excelClient->listWorksheetTableRows(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Row[]? value = rows.value;
        return value is excel:Row[] ? value : [];
    }

    # Adds rows to the end of a table in the worksheet.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + row - The properties of the table row to be created
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function createWorksheetTableRow(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createWorksheetTableRowWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, row, sessionId);
        }
        return self.excelClient->createWorksheetTableRow(itemIdOrPath, worksheetIdOrName, tableIdOrName, row, sessionId);
    }

    # Update the properties of table row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + rowIndex - The index of the table row
    # + row - The properties of the table row to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function updateWorksheetTableRow(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, int rowIndex, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetTableRowWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, rowIndex, row, sessionId);
        }
        return self.excelClient->updateWorksheetTableRow(itemIdOrPath, worksheetIdOrName, tableIdOrName, rowIndex, row, sessionId);
    }

    # Adds rows to the end of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + row - The properties of the table row to be added
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function addWorkbookTableRow(string itemIdOrPath, string tableIdOrName, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorkbookTableRowWithItemPath(itemIdOrPath, tableIdOrName, row, sessionId);
        }
        return self.excelClient->addWorkbookTableRow(itemIdOrPath, tableIdOrName, row, sessionId);
    }

    # Adds rows to the end of the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + row - The properties of the table row to be added
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function addWorksheetTableRow(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, excel:Row row, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorksheetTableRowWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, row, sessionId);
        }
        return self.excelClient->addWorksheetTableRow(itemIdOrPath, worksheetIdOrName, tableIdOrName, row, sessionId);
    }

    # Gets a row based on its position in the collection.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function getWorksheetTableRowWithIndex(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, int index, string? sessionId = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableRowWithIndexItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
        }
        return self.excelClient->getWorksheetTableRowWithIndex(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
    }

    # Retrieve the properties and relationships of table row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Row` or else an error on failure  
    remote isolated function getWorksheetTableRow(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, int index, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Row|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableRowWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getWorksheetTableRow(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Deletes the row from the workbook table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorksheetTableRow(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, int index, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorksheetTableRowWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
        }
        return self.excelClient->deleteWorksheetTableRow(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
    }

    # Get the range associated with the entire row.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getWorksheetTableRowRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, int index, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableRowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
        }
        return self.excelClient->getWorksheetTableRowRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, index, sessionId);
    }

    # Retrieve a list of table column in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Columns` or else an error on failure  
    remote isolated function listWorkbookTableColumns(string itemIdOrPath, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Column[]|error {
        excel:Columns columns;
        if isItemPath(itemIdOrPath) {
            columns = check self.excelClient->listWorkbookTableColumnsWithItemPath(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            columns = check  self.excelClient->listWorkbookTableColumns(itemIdOrPath, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Column[]? value = columns.value;
        return value is excel:Column[] ? value : [];
    }

    # Retrieve a list of table column in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Columns` or else an error on failure  
    remote isolated function listWorksheetTableColumns(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Column[]|error {
        excel:Columns columns;
        if isItemPath(itemIdOrPath) {
            columns = check self.excelClient->listWorksheetTableColumnsWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            columns = check  self.excelClient->listWorksheetTableColumns(itemIdOrPath, worksheetIdOrName, tableIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:Column[]? value = columns.value;
        return value is excel:Column[] ? value : [];
    }

    # Create a new table column in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + column - The properties of the table column to be created
    # + sessionId - The ID of the session
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function createWorkbookTableColumn(string itemIdOrPath, string tableIdOrName, excel:Column column, string? sessionId = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createWorkbookTableColumnWithItemPath(itemIdOrPath, tableIdOrName, column, sessionId);
        }
        return self.excelClient->createWorkbookTableColumn(itemIdOrPath, tableIdOrName, column, sessionId);
    }

    # Create a new table column in the workbook.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + column - The properties of the table column to be created
    # + sessionId - The ID of the session
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function createWorksheetTableColumn(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, excel:Column column, string? sessionId = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->createWorksheetTableColumnWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, column, sessionId);
        }
        return self.excelClient->createWorksheetTableColumn(itemIdOrPath, worksheetIdOrName, tableIdOrName, column, sessionId);
    }

    # Deletes the column from the table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorkbookTableColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorkbookTableColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->deleteWorkbookTableColumn(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Delete a column from a table.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteWorksheetTableColumn(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteWorksheetTableColumnWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->deleteWorksheetTableColumn(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
    }

    # Retrieve the properties and relationships of table column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderby - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function getWorkbookTableColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderby = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorkbookTableColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getWorkbookTableColumn(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Retrieve the properties and relationships of table column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves related resources
    # + filter - Filters results
    # + format - Returns the results in the specified media format
    # + orderby - Orders results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function getWorksheetTableColumn(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderby = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getWorksheetTableColumnWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getWorksheetTableColumn(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
    }

    # Update the properties of table column
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + column - The properties of the table column to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function updateWorksheetTableColumn(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, excel:Column column, string? sessionId = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorksheetTableColumnWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, column, sessionId);
        }
        return self.excelClient->updateWorksheetTableColumn(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, column, sessionId);
    }


    # Update the properties of table column
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + column - 
    # + sessionId - The ID of the session
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function updateWorkbookTableColumn(string itemIdOrPath, string tableIdOrName, string columnIdOrName, excel:Column column, string? sessionId = ()) returns excel:Column|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateWorkbookTableColumnWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, column, sessionId);
        }
        return self.excelClient->updateWorkbookTableColumn(itemIdOrPath, tableIdOrName, columnIdOrName, column, sessionId);
    }

    # Gets the range associated with the data body of the column
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Column` or else an error on failure  
    remote isolated function getworkbookTableColumnsDataBodyRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworkbookTableColumnsDataBodyRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworkbookTableColumnsDataBodyRange(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range associated with the data body of the column
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getworksheetTableColumnsDataBodyRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworksheetTableColumnsDataBodyRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworksheetTableColumnsDataBodyRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range associated with the header row of the column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getworkbookTableColumnsHeaderRowRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworkbookTableColumnsHeaderRowRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworkbookTableColumnsHeaderRowRange(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range associated with the header row of the column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getworksheetTableColumnsHeaderRowRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworksheetTableColumnsHeaderRowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworksheetTableColumnsHeaderRowRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range associated with the totals row of the column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getworkbookTableColumnsTotalRowRange(string itemIdOrPath, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworkbookTableColumnsTotalRowRangeWithItemPath(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworkbookTableColumnsTotalRowRange(itemIdOrPath, tableIdOrName, columnIdOrName, sessionId);
    }

    # Gets the range associated with the totals row of the column.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + tableIdOrName - The ID or name of the table
    # + columnIdOrName - The ID or name of the column
    # + sessionId - The ID of the session
    # + return - An `excel:Range` or else an error on failure  
    remote isolated function getworksheetTableColumnsTotalRowRange(string itemIdOrPath, string worksheetIdOrName, string tableIdOrName, string columnIdOrName, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getworksheetTableColumnsTotalRowRangeWithItemPath(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
        }
        return self.excelClient->getworksheetTableColumnsTotalRowRange(itemIdOrPath, worksheetIdOrName, tableIdOrName, columnIdOrName, sessionId);
    }

    # Retrieve the properties and relationships of chart.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + sessionId - The ID of the session
    # + return - An `excel:Chart` or else an error on failure  
    remote isolated function getChart(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, string? sessionId = ()) returns excel:Chart|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
        }
        return self.excelClient->getChart(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
    }

    # Deletes the chart.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function deleteChart(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->deleteChartWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
        }
        return self.excelClient->deleteChart(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
    }

    # Update the properties of chart.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + chart - The properties of the chart to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:Chart` or else an error on failure  
    remote isolated function updateChart(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, excel:Chart chart, string? sessionId = ()) returns excel:Chart|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateChartWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, chart, sessionId);
        }
        return self.excelClient->updateChart(itemIdOrPath, worksheetIdOrName, chartIdOrName, chart, sessionId);
    }

    # Resets the source data for the chart.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + resetData - The properties of the reset data
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function resetChartData(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, excel:ResetData resetData, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->resetChartDataWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, resetData, sessionId);
        }
        return self.excelClient->resetChartData(itemIdOrPath, worksheetIdOrName, chartIdOrName, resetData, sessionId);
    }

    # Positions the chart relative to cells on the worksheet
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + position - the properties of the position
    # + sessionId - The ID of the session
    # + return - An `http:Response` or else an error on failure  
    remote isolated function setChartPosition(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, excel:Position position, string? sessionId = ()) returns http:Response|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->setChartPositionWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, position, sessionId);
        }
        return self.excelClient->setChartPosition(itemIdOrPath, worksheetIdOrName, chartIdOrName, position, sessionId);
    }

    # Retrieve a list of chart series .
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:CollectionOfChartSeries` or else an error on failure  
    remote isolated function listChartSeries(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:ChartSeries[]|error {
        excel:CollectionOfChartSeries chartSeries;
        if isItemPath(itemIdOrPath) {
            chartSeries = check self.excelClient->listChartSeriesWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            chartSeries = check self.excelClient->listChartSeries(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:ChartSeries[]? value = chartSeries.value;
        return value is excel:ChartSeries[] ? value : [];
    }

    # Gets a chart based on its position in the collection.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + index - Index value of the object to be retrieved. Zero-indexed.
    # + sessionId - The ID of the session
    # + return - An `excel:Chart` or else an error on failure  
    remote isolated function getChartBasedOnPosition(string itemIdOrPath, string worksheetIdOrName, int index, string? sessionId = ()) returns excel:Chart|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartBasedOnPositionWithItemPath(itemIdOrPath, worksheetIdOrName, index, sessionId);
        }
        return self.excelClient->getChartBasedOnPosition(itemIdOrPath, worksheetIdOrName, index, sessionId);
    }

    # Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + sessionId - The ID of the session
    # + return - An `excel:Image` or else an error on failure  
    remote isolated function getChartImage(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, string? sessionId = ()) returns excel:Image|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartImageWithItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
        }
        return self.excelClient->getChartImage(itemIdOrPath, worksheetIdOrName, chartIdOrName, sessionId);
    }

    # Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + width - The desired width of the resulting image.
    # + sessionId - The ID of the session
    # + return - An `excel:Image` or else an error on failure  
    remote isolated function getChartImageWithWidth(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, int width, string? sessionId = ()) returns excel:Image|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartImageWithWidthItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, sessionId);
        }
        return self.excelClient->getChartImageWithWidth(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, sessionId);
    }

    # Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + width - The desired width of the resulting image.
    # + height - The desired height of the resulting image.
    # + sessionId - The ID of the session
    # + return - An `excel:Image` or else an error on failure  
    remote isolated function getChartImageWithWidthHeight(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, int width, int height, string? sessionId = ()) returns excel:Image|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartImageWithWidthHeightItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, height, sessionId);
        }
        return self.excelClient->getChartImageWithWidthHeight(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, height, sessionId);
    }

    # Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + chartIdOrName - The ID or name of the chart
    # + width - The desired width of the resulting image.
    # + height - The desired height of the resulting image.
    # + fittingMode - The method used to scale the chart to the specified dimensions (if both height and width are set)."
    # + sessionId - The ID of the session
    # + return - An `excel:Image` or else an error on failure  
    remote isolated function getChartImageWithWidthHeightFittingMode(string itemIdOrPath, string worksheetIdOrName, string chartIdOrName, int width, int height, "Fit"|"FitAndCenter"|"Fill" fittingMode, string? sessionId = ()) returns excel:Image|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getChartImageWithWidthHeightFittingModeItemPath(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, height, fittingMode, sessionId);
        }
        return self.excelClient->getChartImageWithWidthHeightFittingMode(itemIdOrPath, worksheetIdOrName, chartIdOrName, width, height, fittingMode, sessionId);
    }

    # Retrieves a list of named items.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:NamedItems` or else an error on failure  
    remote isolated function listNamedItem(string itemIdOrPath, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:NamedItem[]|error {
        excel:NamedItems namedItems;
        if isItemPath(itemIdOrPath) {
            namedItems = check self.excelClient->listNamedItemWithItemPath(itemIdOrPath, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        } else {
            namedItems = check self.excelClient->listNamedItem(itemIdOrPath, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        excel:NamedItem[]? value = namedItems.value;
        return value is excel:NamedItem[] ? value : [];
    }

    # Adds a new name to the collection of the given scope using the user's locale for the formula.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + namedItem - The properties of the named item to be added
    # + sessionId - The ID of the session
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function addWorkbookNamedItem(string itemIdOrPath, excel:NewNamedItem namedItem, string? sessionId = ()) returns excel:NamedItem|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorkbookNamedItemWithItemPath(itemIdOrPath, namedItem, sessionId);
        }
        return self.excelClient->addWorkbookNamedItem(itemIdOrPath, namedItem, sessionId);
    }

    # Adds a new name to the collection of the given scope using the user's locale for the formula.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + worksheetIdOrName - The ID or name of the worksheet
    # + namedItem - The properties of the named item to be added
    # + sessionId - The ID of the session
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function addWorksheetNamedItem(string itemIdOrPath, string worksheetIdOrName, excel:NewNamedItem namedItem, string? sessionId = ()) returns excel:NamedItem|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->addWorksheetNamedItemWithItemPath(itemIdOrPath, worksheetIdOrName, namedItem, sessionId);
        }
        return self.excelClient->addWorksheetNamedItem(itemIdOrPath, worksheetIdOrName, namedItem, sessionId);
    }

    # Retrieve the properties and relationships of the named item.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item to get.
    # + sessionId - The ID of the session
    # + count - Retrieves the total count of matching resources
    # + expand - Retrieves the related resources
    # + filter - Filters the results
    # + format - Returns the results in the specified media format
    # + orderBy - Orders the results
    # + search - Returns results based on search criteria
    # + 'select - Filters properties(columns)
    # + skip - Indexes into a result set
    # + top - Sets the page size of results
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function getNamedItem(string itemIdOrPath, string name, string? sessionId = (), string? count = (), string? expand = (), string? filter = (), string? format = (), string? orderBy = (), string? search = (), string? 'select = (), int? skip = (), int? top = ()) returns excel:NamedItem|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNamedItemWithItemPath(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
        }
        return self.excelClient->getNamedItem(itemIdOrPath, name, sessionId, count, expand, filter, format, orderBy, search, 'select, skip, top);
    }

    # Update the properties of the named item.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item to get
    # + namedItem - The properties of the named item to be updated
    # + sessionId - The ID of the session
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function updateNamedItem(string itemIdOrPath, string name, excel:NamedItem namedItem, string? sessionId = ()) returns excel:NamedItem|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->updateNamedItemWithItemPath(itemIdOrPath, name, namedItem, sessionId);
        }
        return self.excelClient->updateNamedItem(itemIdOrPath, name, namedItem, sessionId);
    }

    # Retrieve the range object that is associated with the name.
    #
    # + itemIdOrPath - The ID of the drive containing the workbook or the path to the workbook
    # + name - The name of the named item to get.
    # + sessionId - The ID of the session
    # + return - An `excel:NamedItem` or else an error on failure  
    remote isolated function getNamedItemRange(string itemIdOrPath, string name, string? sessionId = ()) returns excel:Range|error {
        if isItemPath(itemIdOrPath) {
            return self.excelClient->getNamedItemRangeWithItemPath(itemIdOrPath, name, sessionId);
        }
        return self.excelClient->getNamedItemRange(itemIdOrPath, name, sessionId);
    }
}
