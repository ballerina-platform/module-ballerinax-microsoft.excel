# Workbook session summary

Opens a workbook session, lists every worksheet in the workbook, and reports the used range of each one. Running the calls inside a single session is what Microsoft recommends whenever an integration makes more than one request in a short period, and reading the used range first lets a reporting job size each sheet before deciding what to read in full.

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
