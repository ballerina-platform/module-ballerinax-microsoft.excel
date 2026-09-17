# Expense table ingestion

Sets up a fresh expense sheet from scratch: adds a worksheet, defines a table over its header range, and appends expense rows to that table. This is the shape of a recurring ingestion job that writes periodic records into a workbook rather than reading from one.

## Prerequisites

- Ballerina Swan Lake 2201.12.0 or later
- Make the connector available to the example. Until this version is published to
  Ballerina Central, run the helper script from the `examples` directory, which packs the
  connector and stages it where the examples can resolve it:
  ```bash
  cd .. && ./build.sh build
  ```
- Follow the [Setup guide](../../ballerina/README.md#setup-guide) to register an application and obtain a refresh token.
- Create a `Config.toml` in this directory:
  ```toml
  clientId = "<CLIENT_ID>"
  clientSecret = "<CLIENT_SECRET>"
  refreshToken = "<REFRESH_TOKEN>"
  driveId = "<DRIVE_ID>"
  driveItemId = "<DRIVE_ITEM_ID>"
  ```

## Run the example

```bash
bal run
```

To build or run every example at once, use the helper script from the `examples`
directory instead — `./build.sh build` or `./build.sh run`.
