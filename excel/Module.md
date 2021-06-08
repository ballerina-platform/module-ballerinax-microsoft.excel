## Overview
Ballerina connector for Microsoft Excel is connecting the Excel API in Microsoft Graph v1.0 via Ballerina language easily. It provides capability to perform perform CRUD (Create, Read, Update, and Delete) operations on [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in Microsoft OneDrive. 

The connector is developed on top of Microsoft Graph is a REST web API that empowers you to access Microsoft Cloud service resources. This version of the connector only supports the access to the resources and information of a specific account (currently logged in user).

## Obtaining tokens

Follow the following steps below to obtain the configurations.

1. Before you run the following steps, create an account in [OneDrive](https://onedrive.live.com). Next, sign into [Azure Portal - App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade). You can use your personal or work or school account to register.

2. In the App registrations page, click **New registration** and enter a meaningful name in the name field.

3. In the **Supported account types** section, select **Accounts** in any organizational directory under personal Microsoft accounts (e.g., Skype, Xbox, Outlook.com). Click **Register** to create the application.
    
4. Copy the Application (client) ID (`<CLIENT_ID>`). This is the unique identifier for your app.
    
5. In the application's list of pages (under the **Manage** tab in left hand side menu), select **Authentication**.
    Under **Platform configurations**, click **Add a platform**.

6. Under **Configure platforms**, click **Web** located under **Web applications**.

7. Under the **Redirect URIs text box**, register the https://login.microsoftonline.com/common/oauth2/nativeclient url.
   Under **Implicit grant**, select **Access tokens**.
   Click **Configure**.

8. Under **Certificates & Secrets**, create a new client secret (`<CLIENT_SECRET>`). This requires providing a description and a period of expiry. Next, click **Add**.

9. Next, you need to obtain an access token and a refresh token to invoke the Microsoft Graph API.
First, in a new browser, enter the below URL by replacing the `<CLIENT_ID>` with the application ID. Here you can use `Files.ReadWrite` or `Files.ReadWrite.All` according to your preference. `Files.ReadWrite` will allow you to access to only your files and `Files.ReadWrite.All` will allow you to access all files you can access.

    ```
    https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=<CLIENT_ID>&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=Files.ReadWrite offline_access
    ```

10. This will prompt you to enter the username and password for signing into the Azure Portal App.

11. Once the username and password pair is successfully entered, this will give a URL as follows on the browser address bar.

    `https://login.microsoftonline.com/common/oauth2/nativeclient?code=xxxxxxxxxxxxxxxxxxxxxxxxxxx`

12. Copy the code parameter (`xxxxxxxxxxxxxxxxxxxxxxxxxxx` in the above example) and in a new terminal, enter the following cURL command by replacing the `<CODE>` with the code received from the above step. The `<CLIENT_ID>` and `<CLIENT_SECRET>` parameters are the same as above.

    ```
    curl -X POST --header "Content-Type: application/x-www-form-urlencoded" --header "Host:login.microsoftonline.com" -d "client_id=<CLIENT_ID>&client_secret=<CLIENT_SECRET>&grant_type=authorization_code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&code=<CODE>&scope=Files.ReadWrite offline_access" https://login.microsoftonline.com/common/oauth2/v2.0/token
    ```

    The above cURL command should result in a response as follows.
    ```
    {
      "token_type": "Bearer",
      "scope": "Files.ReadWrite",
      "expires_in": 3600,
      "ext_expires_in": 3600,
      "access_token": "<ACCESS_TOKEN>",
      "refresh_token": "<REFRESH_TOKEN>",
    }
    ```

13. Provide the following configuration information in the `Config.toml` file to use the Microsoft Excel connector.

    ```ballerina
    clientId = <CLIENT_ID>
    clientSecret = <CLIENT_SECRET>
    refreshUrl = <REFRESH_URL>
    refreshToken = <REFRESH_TOKEN>
    workbookIdOrPath = <WORKBOOK_ID_OR_PATH>
    ```

    The `workbookIdOrPath` is workbook id or file path (with the `.xlsx` extension from root. If you have a file in root directory with name of `Work.xlsx`, you need to pass it as `Work.xlsx`). Make sure you create a workbook in Microsoft OneDrive and pass the correct `workbookIdOrPath` before using the connector.

## Compatibility & Limitations
### Compatibility
|                                                                                    | Version               |
|------------------------------------------------------------------------------------|-----------------------|
| Ballerina Language Version                                                         | **Swan Lake Beta 1**  |
| [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) Version     | **v1.0**              |
| Java Development Kit (JDK)                                                         | 11                    |

### Limitations
- Connector only allows to perform functions behalf of the currently logged in user.

## Quickstart

### Create a worksheet in Workbook
#### Step 1: Import Sheet module
First, import the ballerinax/microsoft.excel module into the Ballerina project.
```ballerina
import ballerinax/microsoft.excel;
```
#### Step 2: Configure the connection credentials.
You can now make the connection configuration using the OAuth2 refresh token grant config.
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
#### Step 3: Create a worksheet in a workbook
You can provide either workbook id or path as `workbookIdOrPath` parameter and file name as 'sheetName` parameter.

```ballerina
public function main() {
    excel:Worksheet|error response = excelClient->addWorksheet(workbookIdOrPath, sheetName);
    if (response is excel:Worksheet) {
        log:printInfo(response.toString());
    }
}
```

### [You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/tree/master/samples)
