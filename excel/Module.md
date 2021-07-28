## Overview
Ballerina connector for Microsoft Excel is connecting the  Microsoft Graph Excel API via Ballerina language. It provides capability to perform CRUD (Create, Read, Update, and Delete) operations on [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in Microsoft OneDrive. 

This module supports [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) v1.0 version and only allows to perform functions behalf of the currently logged in user.

## Prerequisites
Before using this connector in your Ballerina application, complete the following:
1. Create a [Microsoft Office365 account](https://signup.live.com/)
2. Obtain token - Follow the steps [here](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps)

## Quickstart
To use the Microsoft Excel connector in your Ballerina application, update the .bal file as follows:

### Step 1: Import connector
Import the `ballerinax/microsoft.excel` module into the Ballerina project.
```ballerina
import ballerinax/microsoft.excel;
```
### Step 2: Create a new connector instance
Create a `excel:Configuration` with the OAuth2 tokens obtained, and initialize the connector with it. 
```ballerina
excel:Configuration configuration = {
    clientConfig: {
        refreshUrl: <REFRESH_URL>,
        refreshToken : <REFRESH_TOKEN>,
        clientId : <CLIENT_ID>,
        clientSecret : <CLIENT_SECRET>
    }
};

excel:Client excelClient = check new (configuration);
```
### Step 3: Invoke connector operation
1. Now you can use the operations available within the connector. Note that they are in the form of remote operations.  
Following is an example on how to add a worksheet using the connector.
    ```ballerina
    public function main() returns error? {
        excel:Worksheet response = check excelClient->addWorksheet("workbookIdOrPath", "sheetName");
    }
    ```

2. Use `bal run` command to compile and run the Ballerina program.

**[You can find a list of samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/tree/master/samples)**
