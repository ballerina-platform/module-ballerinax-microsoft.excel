# Module Microsoft Sheets

The `ballerinax/module-ballerinax-microsoft.sheets` module contains operations to perform CRUD (Create, Read, Update, and Delete) operations on [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in Microsoft OneDrive.

## Compatibility
|                     |    Version     |
|:-------------------:|:--------------:|
| Ballerina Language  | Swan Lake Preview1   |
| Microsoftgraph REST API | v1.0          |

## Getting started

1.  Download and install Ballerina. For instructions, go to [Installing Ballerina](https://ballerina.io/learn/installing-ballerina/).

2.  Provide the following configuration information in the `ballerina.conf` file to use the Microsoft Graph API.

       - MS_CLIENT_ID
       - MS_CLIENT_SECRET
       - MS_ACCESS_TOKEN
       - MS_REFRESH_TOKEN
       - TRUST_STORE_PATH
       - TRUST_STORE_PASSWORD
       - WORK_BOOK_NAME
       - WORK_SHEET_NAME
       - TABLE_NAME

    Follow the steps below to obtain the configuration information mentioned above.

    Before you run the following steps, create an account in [OneDrive](https://onedrive.live.com). Next, sign into [Azure Portal - App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade). You can use your personal or work or school account to register.

    In the App registrations page, click **New registration** and enter a meaningful name in the name field.

    !["Figure 1: App registrations page"](images/step1.jpg)

    *Figure 1: App registrations page*

    In the Supported account types section, select **Accounts** in any organizational directory under personal Microsoft accounts (e.g., Skype, Xbox, Outlook.com). Click **Register** to create the application.

    !["Figure 2: Accounts type selection"](images/step2.jpg)
    
    *Figure 2: Accounts type selection*

    Copy the Application (client) ID (\<MS_CLIENT_ID>). This is the unique identifier for your app.
    In the application's list of pages (under the **Manage** tab in left hand side menu), select **Authentication**.
    Under **Platform configurations**, click **Add a platform**.

    !["Figure 3: Add a platform"](images/step3.jpg)
    
    *Figure 3: Add a platform*

    Under **Configure platforms**, click the **Web** button located under **Web applications**.

    Under the **Redirect URIs text box**, put [OAuth2 Native Client](https://login.microsoftonline.com/common/oauth2/nativeclient).
    Under **Implicit grant**, select **Access tokens**.
    Click on **Configure**.

    !["Figure 4: Update security configurations"](images/step4.jpg)
    
    *Figure 4: Update security configurations*

    Under **Certificates & Secrets**, create a new client secret (\<MS_CLIENT_SECRET>). This requires providing a description and a period of expiry. Next, click **Add**.

    Next, you need to obtain an access token and a refresh token to invoke the Microsoft Graph API.
    First, in a new browser, enter the below URL by replacing the \<MS_CLIENT_ID> with the application ID.

    ```
    https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=<MS_CLIENT_ID>&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite offline_access
    ```

    This will prompt you to enter the username and password for signing into the Azure Portal App.

    Once the username and password pair is successfully entered, this will give a URL as follows on the browser address bar.

    `https://login.microsoftonline.com/common/oauth2/nativeclient?code=M95780001-0fb3-d138-6aa2-0be59d402f32`

    Copy the code parameter (M95780001-0fb3-d138-6aa2-0be59d402f32 in the above example) and in a new terminal, enter the following CURL command by replacing the \<MS_CODE> with the code received from the above step. The \<MS_CLIENT_ID> and \<MS_CLIENT_SECRET> parameters are the same as above.

    ```
    curl -X POST --header "Content-Type: application/x-www-form-urlencoded" --header "Host:login.microsoftonline.com" -d "client_id=<MS_CLIENT_ID>&client_secret=<MS_CLIENT_SECRET>&grant_type=authorization_code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&code=<MS_CODE>&scope=Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite offline_access" https://login.microsoftonline.com/common/oauth2/v2.0/token
    ```

    The above CURL command should result in a response as follows.
    ```
    {
    "token_type": "Bearer",
    "scope": "Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite",
    "expires_in": 3600,
    "ext_expires_in": 3600,
    "access_token": "<MS_ACCESS_TOKEN>",
    "refresh_token": "<MS_REFRESH_TOKEN>",
    "id_token": "<ID_TOKEN>"
    }
    ```

    Set the path to your Ballerina distribution's trust store as the \<TURST_STORE_PATH>. This is by default located in the following path.

    `$BALLERINA_HOME/distributions/ballerina-<BALLERINA_VERSION>/bre/security/ballerinaTruststore.p12`

    The default `TRUST_STORE_PASSWORD` is set to "ballerina".

    The `WORK_BOOK_NAME`, `WORK_SHEET_NAME`, and `TABLE_NAME` correspond to the workbook file name (without the .xlsx extension), worksheet name, and table name respectively. Make sure you create a workbook with the same `WORK_BOOK_NAME` as on Microsoft OneDrive before using the connector.

3. Create a new Ballerina project by executing the following command.

	```shell
	<PROJECT_ROOT_DIRECTORY>$ ballerina init
	```

4. Import the Microsoft Graph connector to your Ballerina program as follows.

    The following sample program creates a new worksheet on an existing workbook on Microsoft OneDrive. Prior to running this application, create a workbook on your Microsoft OneDrive account with the name "MyShop.xlsx". There needs to be at least one worksheet (i.e., a tab) on the workbook for this sample program to work. 

    The sample application first tries to delete an existing worksheet named "Sales" from the workbook. If its not available, it may throw an error and continue executing the rest of the program. This error will get thrown during the very first round of running the sample application. Make sure that you keep the `ballerina.conf` file with the above-mentioned configuration information before running the sample application.

	```
    import ballerina/config;
    import ballerina/log;
    import ballerina/time;
    import ballerinax/microsoft.sheets1 as sheets;
    import ballerinax/microsoft.onedrive1 as onedrive;

    // Create the Microsoft Graph Client configuration by reading the config file.
    sheets:MicrosoftGraphConfiguration msGraphConfig = {
        baseUrl: config:getAsString("MS_BASE_URL"),
        msInitialAccessToken: config:getAsString("MS_ACCESS_TOKEN"),
        msClientID: config:getAsString("MS_CLIENT_ID"),
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

    sheets:MSSpreadsheetClient msSpreadsheetClient = new(msGraphConfig);
    
    string WORK_BOOK_NAME = "MyShop";
    string WORK_SHEET_NAME = "Sales";
    string TABLE_NAME = "tbl";

    public function main() {
        sheets:Workbook|error workbookResponse = msSpreadsheetClient->openWorkbook("/", WORK_BOOK_NAME);

        if workbookResponse is error {
            log:printInfo("Error opening workbook.");
            return;
        }

        sheets:Workbook workbook = <sheets:Workbook> workbookResponse;

        error? resultRemove = workbook->removeWorksheet(WORK_SHEET_NAME);

        if !(resultRemove is ()) {
            log:printError("Could not delete the Worksheet, but will continue execution", err = resultRemove);
        }

        sheets:Worksheet|error sheetResponse = workbook->createWorksheet(WORK_SHEET_NAME);

        if !(sheetResponse is sheets:Worksheet) {
            log:printError("Could not create the Worksheet", err = sheetResponse);
            return;
        }

        sheets:Worksheet sheet = <sheets:Worksheet> sheetResponse;

        sheets:Table|error resultTableResponse = sheet->createTable(TABLE_NAME, <@untainted> ("A1:E1"));

        if !(resultTableResponse is sheets:Table) {
            log:printError("Could not create the Table", err = resultTableResponse);
            return;
        }

        sheets:Table resultTable = <sheets:Table> resultTableResponse;

        error? resultHeader = resultTable->setTableHeader(1, "ID");

        if !(resultHeader is ()) {
            log:printError("Could not set the Table header of colunm 1", err = resultHeader);
            return;
        }

        resultHeader = resultTable->setTableHeader(2, "DateSold");

        if !(resultHeader is ()) {
            log:printError("Could not set the Table header of colunm 2", err = resultHeader);
            return;
        }

        resultHeader = resultTable->setTableHeader(3, "ItemID");

        if !(resultHeader is ()) {
            log:printError("Could not set the Table header of colunm 3", err = resultHeader);
            return;
        }

        resultHeader = resultTable->setTableHeader(4, "ItemName");

        if !(resultHeader is ()) {
            log:printError("Could not set the Table header of colunm 4", err = resultHeader);
            return;
        }

        resultHeader = resultTable->setTableHeader(5, "Price");

        if !(resultHeader is ()) {
            log:printError("Could not set the Table header of colunm 5", err = resultHeader);
            return;
        }        

        json[][] valuesString=[];
        time:Time time = time:currentTime();
        string|error cString1 = time:format(time, "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
        string customTimeString = "";
        if (cString1 is string) {
            customTimeString = cString1;
        }

        foreach int counter in 1...5 {
            int itemID = counter + 100;
            json[] arr = [ counter.toString(), customTimeString, 
            itemID.toString(), "Item-" + itemID.toString(), "10" ];
            valuesString.push(arr);
        }
        json data = {"values": valuesString};
        error? result = resultTable->insertDataIntoTable(<@untainted> data);

        if !(result is ()) {
            log:printError("Error inserting data into the table");
            return;
        }
    }
	```