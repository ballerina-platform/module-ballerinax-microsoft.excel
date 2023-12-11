Ballerina Microsoft Excel Connector
===================

[![Build Status](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/workflows/CI/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/actions?query=workflow%3ACI)
[![codecov](https://codecov.io/gh/ballerina-platform/module-ballerinax-microsoft.excel/branch/master/graph/badge.svg)](https://codecov.io/gh/ballerina-platform/module-ballerinax-microsoft.excel)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerinax-microsoft.excel.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/commits/master)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/actions/workflows/build-with-bal-test-native.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel/actions/workflows/build-with-bal-test-native.yml)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

[Excel](https://www.microsoft.com/en-ww/microsoft-365/excel) is a widely used spreadsheet application developed by Microsoft, enabling users to organize, analyze, and visualize data in a tabular format.

The `ballerinax/microsoft.excel` package offers APIs to connect and interact with [Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0 ) to allow web and mobile applications to read and modify Excel workbooks stored in OneDrive for Business, SharePoint site or Group drive.

## Set up Excel API

To use the Excel connector, You have to use the [Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/) to authenticate Excel APIs. If you do not have a valid Microsoft Azure AD account, you can sign up for one through the [Azure Portal](https://portal.azure.com/).

### Step 1: Register your app in Azure Active Directory (Azure AD)

1. Go to the [Azure portal](https://portal.azure.com/) and sign in with a global administrator account.

2. Select **Microsoft Entra ID** from the left-hand navigation menu.

3. Select **App registrations** from the **Manage** section of the left-hand navigation menu.

4. Select **New registration**.

5. In the Register an application page, enter a name for your app.

6. Select the **supported account types** is **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)** .

7. Select the **Web application** type and enter the URL under the **Redirect URI**.

8. Click on **Register**.

   ![Register App](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/master/docs/setup/resources/register_app.png)

### Step 2: Configure app permissions

1. In the left pane, select **App registrations**. 

2. Go to the **Owned applications** and select the app you registered in Step 1. 

3. In the Manage section, select **API permissions**.

4. Click on **Add a permission**. 

5. Select **Microsoft Graph**. 

6. Select the **Delegated permissions scope**. 

7. Click on **Add permissions**.

   ![Add App Permissions](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/master/docs/setup/resources/add_app_permission.png)

### Step 3: Create a client secret

1. In the left pane, select **App registrations**.
2. Go to the **Owned applications** and select the app you registered in Step 1.
3. In the Manage section, select **Certificates & secrets**. 
4. Click on **New client secret**. 
5. Enter a **Description** for the client secret and select the **Expires** time period according to your purpose. 
6. Click on **Add**. 
7. Copy the **value** of the client secret. You will need this value later.

   ![Create Client Secret](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/master/docs/setup/resources/create_client_secrets.png)

### Step 4: Get the renew refresh token

Use the following table to change the value of the fields in the sections below:

   | Field                    | Values of your app                                         |
   |------------------------------------------------------------| -------------------- |
   | Tenant ID        | Directory (tenant) ID                                      |
   | Name                     | Application (client)                                       |
   | Authorized Redirect URIs | Redirect URI eg:http%3a%2f%2flocalhost%3a8080              |
   | Scope                    | Delegated permissions scope eg: Files.Read Files.ReadWrite |

   ![Details of App](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.excel/master/docs/setup/resources/deatils_of_app.png)

#### Get an authorization code

Open a new tab in your browser and enter the following URL after replacing the values of the query parameters with the values of your application.
   ```
   https://login.microsoftonline.com/{Tenant ID}/oauth2/v2.0/authorize?client_id={AppReg ID}
   &response_type=code
   &redirect_uri={Redirect URI}
   &response_mode=query
   &scope={SCOPE}&sso_reload=true
   ```
#### Get a refresh token

After replacing the values of the query parameters with the values of your application, open a terminal and enter the following curl command.

```
  curl -X POST -H "Content-Type: application/x-www-form-urlencoded" -d 'client_id={AppReg ID}
  &scope={SCOPE} openid profile offline_access
  &code={authorization code}
  &redirect_uri={Redirect URI}
  &grant_type=authorization_code
  &client_secret={AppReg Secret}' 'https://login.microsoftonline.com/{Tenant ID}/oauth2/v2.0/token'
```

#### Get a refresh token to renew an expired access token

Using the following cul command, you can get a refresh token to renew an expired access token.

```
  curl --location --request POST 'https://login.microsoftonline.com/{Tenant ID}/oauth2/v2.0/token' \
   --header 'Content-Type: application/x-www-form-urlencoded' \
   --data-urlencode 'client_id={AppReg ID}' \
   --data-urlencode 'scope={SCOPE}' \
   --data-urlencode 'refresh_token={Refresh token} \
   --data-urlencode 'grant_type=refresh_token' \
   --data-urlencode 'client_secret={AppReg Secret}'
```

## Quickstart

To use the `excel` connector in your Ballerina application, modify the `.bal` file as follows:

### Step 1: Import the connector

Import the `ballerinax/microsoft.excel` package into your Ballerina project.

```ballerina
import ballerinax/microsoft.excel;
```

### Step 2: Instantiate a new connector

Create a `excel:ConnectionConfig` with the obtained OAuth2.0 tokens and initialize the connector with it.

```ballerina
excel:Client excelClient = check new excel:Client(
    config = {
        auth: {
            clientId: "<client_id>",
            clientSecret: "<client_secret>",
            refreshToken: "<refresh_token>",
            refreshUrl: "<refresh_url>"
        }
    }
);
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations.

#### Add worksheet

```ballerina
excel:Worksheet|error response = excelClient->createWorksheet(workBookId, {name: "test"});
```

#### Get Worksheet

```ballerina
excel:Worksheet|error response = excelClient->getWorksheet(itemId, workBookId, worksheetName);
```

## Building from the source

### Setting up the prerequisites

1. Download and install Java SE Development Kit (JDK) version 11. You can install either [OpenJDK](https://adoptopenjdk.net/) or [Oracle JDK](https://www.oracle.com/java/technologies/javase-jdk11-downloads.html).

    > **Note:** Set the JAVA_HOME environment variable to the path name of the directory into which you installed JDK.

2. Download and install [Ballerina Swan Lake](https://ballerina.io/). 

### Building the source
Execute the commands below to build from the source after installing Ballerina.

1. To build the package:
    ```    
    bal build ./ballerina
    ```
2. To run tests after build:
    ```
    bal test ./ballerina
    ```
## Contributing to Ballerina
As an open source project, Ballerina welcomes contributions from the community. 

For more information, go to the [contribution guidelines](https://github.com/ballerina-platform/ballerina-lang/blob/main/CONTRIBUTING.md).

## Code of conduct
All contributors are encouraged to read the [Ballerina Code of Conduct](https://ballerina.io/code-of-conduct).

## Useful links
* Discuss about code changes of the Ballerina project in [ballerina-dev@googlegroups.com](mailto:ballerina-dev@googlegroups.com).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
