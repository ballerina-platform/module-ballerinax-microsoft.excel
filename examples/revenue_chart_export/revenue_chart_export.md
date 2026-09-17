# Revenue chart export

Builds a clustered column chart from a range of revenue figures, anchors it to a cell range on the worksheet, and exports it as a base64-encoded image. The exported image can be embedded directly in a report or an email without the recipient needing access to the workbook.

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
