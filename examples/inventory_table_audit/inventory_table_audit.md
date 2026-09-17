# Inventory table audit

Audits the tables on an inventory worksheet by listing each table, its columns, and the addresses of its header and total rows. A data-quality check can run this before a downstream job to confirm the sheet still has the structure that job expects.

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
  worksheetId = "<WORKSHEET_ID>"
  ```

## Run the example

```bash
bal run
```

To build or run every example at once, use the helper script from the `examples`
directory instead — `./build.sh build` or `./build.sh run`.
