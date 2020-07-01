This module allows users to connect to a [Microsoft Office 365 Workbook](https://www.microsoft.com/en-ww/microsoft-365) located on [Microsoft OneDrive](https://docs.microsoft.com/en-us/graph/onedrive-concept-overview).

# Module Overview
This module contains operations to perform CRUD (Create, Read, Update, and Delete) operations on [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in Microsoft OneDrive.

## Supported Operations
- Open a workbook
- Create a worksheet
- Open a worksheet
- Remove a worksheet
- Create a table
- Rename a table
- Set a table header
- Insert data into a table

## Compatibility
|                     |    Version     |
|:-------------------:|:--------------:|
| Ballerina Language  | Swan Lake Preview1   |
| Microsoftgraph REST API | v1.0          |

## Sample
Instantiate the connector by giving authentication details in an HTTP client config. The HTTP client config has built-in support for BasicAuth and OAuth 2.0. Microsoft Graph API uses OAuth 2.0 to authenticate and authorize requests. 

**Obtaining configuration information**
The Microsoft Sheets connector can be minimally instantiated in the HTTP client config using the access token (`<MS_ACCESS_TOKEN>`), the client ID (`<MS_CLIENT_ID>`), the client secret (`<MS_CLIENT_SECRET>`), and the refresh token (`<MS_REFRESH_TOKEN>`). Specific details on obtaining these values are mentioned in the [README](https://github.com/miyurud/module-ballerinax-microsoft.sheets/blob/master/README.md).

**Add project configurations file**

Add the project configuration file by creating a `ballerina.conf` file under the root path of the project structure. This file should have the following configurations below. Add the tokens obtained in the previous step to the `ballerina.conf` file. Make sure to configure the path to the Ballerina trust store and to set the trust store password.

```
MS_BASE_URL="https://graph.microsoft.com"
MS_CLIENT_ID="<MS_CLIENT_ID>"
MS_CLIENT_SECRET="<MS_CLIENT_SECRET>"
MS_REFRESH_TOKEN="<MS_REFRESH_TOKEN>"
MS_REFRESH_URL="https://login.microsoftonline.com/common/oauth2/v2.0/token"
MS_ACCESS_TOKEN="<MS_ACCESS_TOKEN>"
TRUST_STORE_PATH=""
TRUST_STORE_PASSWORD=""
```

**Example Code**
Creating a `microsoft.sheets:MsSpreadsheetClient` by giving the HTTP client config details. The `microsoft.sheets` module 
is referred by the `sheets` module prefix.

```
    import ballerinax/microsoft.sheets;

    sheets:MicrosoftGraphConfiguration msGraphConfig = {
        baseUrl: config:getAsString("MS_BASE_URL"),
        msInitialAccessToken: config:getAsString("MS_ACCESS_TOKEN"),
        msClientId: config:getAsString("MS_CLIENT_ID"),
        msClientSecret: config:getAsString("MS_CLIENT_SECRET"),
        msRefreshToken: config:getAsString("MS_REFRESH_TOKEN"),
        msRefreshURL: config:getAsString("MS_REFRESH_URL"),
        trustStorePath: config:getAsString("TRUST_STORE_PATH"),
        trustStorePassword: config:getAsString("TRUST_STORE_PASSWORD"),
        bearerToken: config:getAsString("MS_ACCESS_TOKEN"),
        clientConfig: {
            accessToken: config:getAsString("MS_ACCESS_TOKEN"),
            refreshConfig: {
                clientId: config:getAsString("MS_CLIENT_ID"),
                clientSecret: config:getAsString("MS_CLIENT_SECRET"),
                refreshToken: config:getAsString("MS_REFRESH_TOKEN"),
                refreshUrl: config:getAsString("MS_REFRESH_URL")
            }
        }
    };

    sheets:MsSpreadsheetClient msSpreadsheetClient = new(msGraphConfig);
```

Open an existing workbook named `Book.xlsx` (no need of specifying the `.xlsx` workbook extension here).

```sheets:Workbook|error workbook = msSpreadsheetClient->openWorkbook("/", "Book");```

Creating a new worksheet named `WS`

```sheets:Worksheet|error sheet = workbook->createWorksheet("WS");```

Opening an existing worksheet named `WS`

```sheets:Worksheet|error sheet = workbook->openWorksheet("WS");```

Creating a new `table1` table on a worksheet. The table has five columns

```sheets:Table|error resultTable = sheet->createTable("table1", <@untainted> ("A1:E1"));```

Setting a table header. Here, the header of the first column of the table is set to `ID`.

```error? resultHeader = resultTable->setTableHeader(1, "ID");```

Inserting data to a table. Here, data is a JSON variable holding rows of the table inside a JSON array attribute.

```error? result = resultTable->insertDataIntoTable(<@untainted> data); ```

Remove a worksheet named `WS`.

```error? resultRemove = workbook->removeWorksheet("WS");```