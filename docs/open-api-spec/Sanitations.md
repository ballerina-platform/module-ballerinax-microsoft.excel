# Sanitizations for Open API Specification
This connector is generated using OpenApi description for an OData service metadata [API version graph1.0](https://github.com/microsoft/OpenAPI.NET.OData/blob/master/docs/oas3_0_0/graph1.0.json), and the following sanitizations were applied to the specification before client generation.

1. Modified parameters and records names:

   | Name                                           | New Name                   | 
   |----------------------------|----------------------------------------------------|
   | driveItem-id                                       | workbook-id                |
   | workbookWorksheet-id                               | workbook-worksheet-id      |
   | workbookChart-id                                   | workbook-chart-id          |
   | workbookNamedItem-id                               | workbook-named-item-id     |
   | workbookTable-id                                   | workbook-table-id          |
   | workbookTableColumn-id                             | workbook-table-column-id   |
   | microsoft.graph.workbookWorksheet                  | Worksheet                  |
   | microsoft.graph.workbookTable                      | Table                      |
   | microsoft.graph.workbookRange                      | Range                      |
   | microsoft.graph.workbookChart                      | Chart                      |
   | microsoft.graph.workbookTableColumn                | TableColumn                |
   | microsoft.graph.workbookTableRow                   | TableRow                   |
   | microsoft.graph.entity                             | Entity                     |
   | microsoft.graph.workbookPivotTable                 | PivotTable                 |
   | microsoft.graph.workbookRangeFormat                | RangeFormat                |
   | microsoft.graph.workbookChartGridlines             | ChartGridlines             |
   | microsoft.graph.workbookSortField                  | SortField                  |
   | microsoft.graph.workbookChartGridlines             | ChartGridlines             |
   | microsoft.graph.workbookApplication                | Application                |
   | microsoft.graph.workbookChartAxisTitleFormat       | ChartAxisTitleFormat       |
   | microsoft.graph.workbookChartFont                  | ChartFont                  |
   | microsoft.graph.workbookChartTitleFormat           | ChartTitleFormat           |
   | microsoft.graph.workbookRangeView                  | RangeView                  |
   | microsoft.graph.workbookChartLineFormat            | ChartLineFormat            |
   | microsoft.graph.workbookChartAxisTitle             | ChartAxisTitle             |
   | microsoft.graph.workbookChartSeriesFormat          | ChartSeriesFormat          |
   | microsoft.graph.workbookChartSeries                | ChartSeries                |
   | microsoft.graph.workbookChartAreaFormat            | ChartAreaFormat            |
   | microsoft.graph.workbookRangeFont                  | RangeFont                  |
   | microsoft.graph.workbookChartDataLabelFormat       | ChartDataLabelFormat       |
   | microsoft.graph.workbookSortField                  | SortField                  |
   | microsoft.graph.workbookIcon                       | Icon                       |
   | microsoft.graph.workbookRangeBorder                | RangeBorder                |
   | microsoft.graph.workbookChartFill                  | ChartFill                  |
   | microsoft.graph.workbookPivotTable                 | PivotTable                 |
   | microsoft.graph.workbookRangeSort                  | RangeSort                  |
   | microsoft.graph.workbookRangeFill                  | RangeFill                  |
   | microsoft.graph.workbookWorksheetProtection        | WorksheetProtection        |
   | microsoft.graph.workbookChartPointFormat           | ChartPointFormat           |
   | microsoft.graph.workbookChartGridlinesFormat       | ChartGridlinesFormat       |
   | microsoft.graph.workbookChartLegend                | ChartLegend                |
   | microsoft.graph.workbookChartLegendFormat          | ChartLegendFormat          |
   | microsoft.graph.workbookFormatProtection           | FormatProtection           |
   | microsoft.graph.workbookChartPoint                 | ChartPoint                 |
   | microsoft.graph.workbookChartAxisFormat            | hartAxisFormat             |
   | microsoft.graph.workbookChartDataLabels            | ChartDataLabels            |
   | microsoft.graph.workbookChartAxis                  | ChartAxis                  |
   | microsoft.graph.workbookWorksheetProtectionOptions | WorksheetProtectionOptions |
   | microsoft.graph.workbookChartAxes                  | ChartAxes                  |
   | microsoft.graph.workbookChartTitle                 | ChartTitle                 |
   | microsoft.graph.workbookFilter                     | Filter                     |
   | microsoft.graph.workbookFilterCriteria             | FilterCriteria             |

2. Change following URL prefix:
    - `/workbooks/{workbook-id}/workbook`  changed to `/me/drive/items/{workbook-id}/workbook`
    - `/me/insights/shared/{sharedInsight-id}/resource/Range/` changed to `/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/`

3. This open API definition is related to OData Service. Hence,  we have to add missing properties manually. `securitySchemes` under the `components` and `seesionId` parameter under the `parameters`. 
   ```
    "securitySchemes": {
        "OAuth2": {
            "type": "oauth2",
            "flows": {
            "authorizationCode": {
                "tokenUrl": "https://login.microsoftonline.com/organizations/oauth2/v2.0/token",
                "authorizationUrl": "https://login.microsoftonline.com/organizations/oauth2/v2.0/token",
                "scopes": {
                "write": "Files.ReadWrite"
                }
            }
            }
        }
        }
   ```
   ```
    "sessionId": {
        "name": "workbook-session-id",
        "description": "The ID of the session",
        "in": "header",
        "required": false,
        "schema": {
            type: "string"
        }
    }
   ```
 4. Update the following paths oprationId with the given name in the below table, add the `sessionId` as the parameter and add the another tag value as `workbook` in the `tags`.

   | Path                    | Operation Id                                       |
   |------------------------------------------------------------| -------------------- |
   |/me/drive/items/{workbook-id}/workbook/application |  getWorkbookApplication           |
   |/me/drive/items/{workbook-id}/workbook/application/calculate| calculateWorkbookApplication|
   |/me/drive/items/{workbook-id}/workbook/names | listWorkbookNames|
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}| getWorkbookNamedItem , updateWorkbookNamedItem, deleteWorkbookNamedItem|
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/worksheet/charts| listWorkbookNamedItemCharts|
   |/me/drive/items/{workbook-id}/{workbook-id}/workbook/names/add | addWorkbookName|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/names/addFormulaLocal|addWorkbookNameFormulaLocal|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables |listWorkbookTables|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}| getWorkbookTable, updateWorkbookTable, deleteWorkbookTable |
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/columns| listWorkbookTableColumns, createWorkbookTableColumns|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/columns/{workbook-table-column-id}|getWorkbookTableColumn, updateWorkbbokColumn, deleteWorkbookColumn|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/columns/{workbook-table-column-id}/range| getWorkbookTableRange|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/rows|listWorkbookTableRows|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/rows/{workbook-table-row-id}|getWorkbookTableRow, updateWorkbookTableRow, deleteWorkbookTableRow|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/rows/add|addWorkbookTableRow|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts|listWorkbookTableCharts|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts/{workbookChart-id}|getWorkbookTableWorksheetChart, updateWorkbookTableWorksheetChart,  deleteWorkbookTableWorksheetWChart|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts/{workbookChart-id}/setData|setWorkbooktableWorksheetChartData|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts/{workbookChart-id}/setPosition|setWorkbookTableWorksheeetChartPosition|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts/{workbookChart-id}/series|listWorkbookTableWorksheetChartSeries|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/charts/add|addWorkbookTableWorksheetChart|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/cell(row={row},column={column})|getWorkbookTableCell|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/names/{workbook-named-item-id}|getWorkbookTableWorksheetNamedItem, updateWorkbookTableWorksheetNamedItem, deleteWorkbookTableWorksheetNamedItem|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/names/add| addWorkbookTableWorksheetName|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/names/addFormulaLocal|addWorkbookTableWorksheetFormula|
   |/me/drive/items/{workbook-id}/workbook/tables/{workbook-table-id}/worksheet/tables/add|addWorkbookTableWorksheetTable|
   |/me/drive/items/{workbook-id}/workbook/tables/add| addWorkbookTable|
   |/me/drive/items/{workbook-id}/workbook/worksheets| listWorksheets, createWorksheets|
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}| getWorksheet updateWorksheet deleteWorksheet|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts| listWorksheetCharts|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}| getWorksheetChart updateWorksheetChart deleteWorksheetChart|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/image|getWorksheetChartImage|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/image(width={width},height={height},fittingMode={fittingMode})|getWorksheetChartImageWithWidthHeightFittingMode|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/image(width={width},height={height})| getWorksheetChartImageWithWidthHeight|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/image(width={width})|getWorksheetChartImageWithWidth|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/setData|setWorksheetChartData|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/setPosition|setWorksheetChartPosition|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/{workbookChart-id}/series|listWorksheetChartSeries createWorksheetChartSeries|
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/add|addWorksheetChart|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/charts/itemAt(index={index})| getWorksheetChartItemAt|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/cell(row={row},column={column})| getWorksheetCell|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range(address={address})|getWorksheetRangeWithAddress|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/names/{workbook-named-item-id}|getWorksheetNamedItem updateWorksheetNamedItem deleteWorksheetNamedItem|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/names/add|addWorksheetName|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/names/addFormulaLocal|addWorksheetFormula|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/pivotTables|listWorksheetPivotTables|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables|listWorksheetTables|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}|getWorksheetTable updateWorksheetTable  deleteWorksheetTable|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/columns|listWorksheetTableColumns|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/columns/{workbook-table-column-id}/range|getWorksheetTableColumnRange|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/range| getWorksheetTableRange|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/rows|listWorksheetTableRows|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/rows/{workbook-table-row-id}|getWorksheetTableRow, updateWorksheetTableRow, deleteWorksheetTableRow
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/{workbook-table-id}/rows/add|addWorksheetTableRow|
   /me/drive/items/{workbook-id}/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/tables/add|addWorksheetTable|
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/rowsAbove  |   getWorksheetRangeRowAbove          |
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/rowsBelow  | getWorksheetRangeRowsBelow |
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/rowsBelow(count={count}) |  getWorksheetRangeRowBelowWithCount           |
   |/me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/rowsAbove(count={count})| getWorksheetRangeRowAboveWithCount |
   | /me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/columnsAfter| getWorksheetColumnsAfterRange            |
   | /me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/columnsAfter(count={count})  |    getWorksheetColumnsAfterRangeWithCount         |
   | /me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/columnsBefore | getWorksheetColumnsBeforeRange            |
   | /me/drive/items/{workbook-id}/workbook/worksheets/{workbook-worksheet-id}/range/columnsBefore(count={count}) |  getWorksheetColumnsBeforeRangeWithCount           |

   Additionally, you must create two copies of each of the APIs below and change the path and parameters to include the missing API.

   Eg: If the `path` is `/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/clear` and `operationId` is `clearWorkbookNamedItemRange`, `path` and `id` of copied APIs should be changed according to the below format:

   | Path                    | Operation Id                                       |
   |------------------------------------------------------------| -------------------- |
   |/me/drive/items/{workbook-id}/**workbook/worksheets/{workbook-worksheet-id}/range(address='<address>')**/clear|clear**Worksheet**Range|
   |/me/drive/items/{workbook-id}/**workbook/tables/{workbook-table-id}/columns/{workbook-table-column-id}/range**/clear|clear**WorkbookTableColumn**Range|


   | Path                    | Operation Id                                       |
   |------------------------------------------------------------| -------------------- |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/boundingRect(anotherRange={anotherRange})       | getWorkbookNamedItemRangeBoundingRect                                    |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/cell(row={row},column={column})                 | getWorkbookNamedItemRangeCell                                       |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/clear | clearWorkbookNamedItemRange              |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/column(column={column})                   |getWorkbookNamedItemRangeColumn|
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/column(column={column}) |  getWorkbookNamedItemRangeColumn           |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/delete  |    deleteWorkbookNamedItemRange         |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/entireColumn |   getWorkbookNamedItemRangeEntireColumn         |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/entireRow |  getWorkbookNameRangeEntireRow           |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/insert | insertWorkbookNamedItemRange            |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/intersection(anotherRange={anotherRange}) |   getWorkbookNamedItemRangeIntersection          |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/lastCell |  getWorkbookNamedItemRangeLastCell           |
   | /me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/lastColumn |   getWorkbookNamedItemRangeLastColumn          |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/lastRow  |    getWorkbookNamedItemRangeLastRow         |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/merge |  mergeWorkbookNamedItemRange           |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/offsetRange(rowOffset={rowOffset},columnOffset={columnOffset})  |  getWorkbookNamedItemOffsetRange           |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/resizedRange(deltaRows={deltaRows},deltaColumns={deltaColumns})  |   getWorkbookNamedItemResizedRange          |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/row(row={row})  |  getWorkbookNamedItemRangeRow           |
   |/me/drive/items/{workbook-id}/workbook/tables/{workbook-table-id}/columns/{workbook-table-column-id}/range/unmerge |  unmergeWorkbookNamedItemRange           |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/usedRange |   getWorkbookNamedItemUsedRnge          |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/usedRange(valuesOnly={valuesOnly}) | getWorkbookNamedItemUsedRngeWithValuesOnly            |
   |/me/drive/items/{workbook-id}/workbook/names/{workbook-named-item-id}/range/visibleView |   getWorkbookNamedItemRangeVisibleView          |

5. Since `Object` type converts to `Record` when generating client, change `values` field defintion of `Row` and `Column` from `Json` to the following defintion
   ```
        "values": {
            "type": "array",
            "description": "The values in the table row",
            "items": {
                "type": "array",
                "items": {
                "oneOf": [
                    {
                    "type": "string",
                    "nullable": true
                    },
                    {
                    "type": "integer",
                    "nullable": true
                    },
                    {
                    "type": "number",
                    "nullable": true
                    }
                ]
                }
            }
        }
   ```
6. Remove `"format": "int32"`.
6. Run the following OpenAPI CLI command to generate the client:
    ```
        bal openapi -i docs/open-api-spec/openapi.json --mode client --client-methods remote --tags workbook -o ballerina/
    ```
7. Clean the generated `types.bal` file. Easily, you can do it by removing the records which start with `microsoft`.
