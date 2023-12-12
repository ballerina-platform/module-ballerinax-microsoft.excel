# Examples

The `excel` connector provides practical example to demonstrates how to use the connector to connect to Excel files and work with data within them. The connector allows you to import and export data, manipulate workbooks and worksheets, and even create new Excel files from scratch.

This example covers several key features of the connector, including:

- Creating new worksheets and table
- Writing data to Excel sheets
- Create a chart

## Prerequisites

1. Follow the [instructions](https://github.com/ballerina-platform/module-ballerinax-microsoft.excel#set-up-excel-api) to set up the Excel API.

2. Create a `config.toml` file with your credential. Here's an example of how your `config.toml` file should look:

    ```toml
    refreshToken="<Refresh Token>"
    clientId="<Client Id>"
    clientSecret="<Client Secret>"
    workbookId="<Id of the Workbook>"
    refreshUrl="<Refresh URL>"
    ```

## Running an Example

Execute the following commands to build an example from the source:

* To build an example:

    ```bash
    bal build
    ```

* To run an example:

    ```bash
    bal run
    ```

## Building the Examples with the Local Module

**Warning**: Due to the absence of support for reading local repositories for single Ballerina files, the Bala of the module is manually written to the central repository as a workaround. Consequently, the bash script may modify your local Ballerina repositories.

Execute the following commands to build all the examples against the changes you have made to the module locally:

* To build all the examples:

    ```bash
    ./build.sh build
    ```

* To run all the examples:

    ```bash
    ./build.sh run
    ```
