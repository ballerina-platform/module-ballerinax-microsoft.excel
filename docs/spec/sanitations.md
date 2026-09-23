_Author_:  DimuthuMadushan \
_Created_: 2026/09/16 \
_Updated_: 2026/09/17 \
_Edition_: Swan Lake

# Sanitation for OpenAPI specification

This document records the sanitation done on top of the official OpenAPI specification from Microsoft Excel. 
The OpenAPI specification is obtained from the [Microsoft Graph v1.0 OpenAPI description](https://github.com/wso2/api-specs/blob/main/openapi/microsoft/graph/v1.0/openapi.yaml).
These changes are done in order to improve the overall usability, and as workarounds for some known language limitations.

1. Extract the Excel workbook surface from the upstream Microsoft Graph specification
- **Original**: The upstream [Microsoft Graph v1.0 OpenAPI description](https://github.com/wso2/api-specs/blob/main/openapi/microsoft/graph/v1.0/openapi.yaml) is a single ~44 MB document describing 11,546 paths across the whole Graph surface.
- **Updated**: `docs/resources/script.py` selects every path whose key contains `/workbook` and writes them verbatim, in upstream order, together with the transitive `$ref` closure of their components. The result is 1,442 paths and 1,765 operations, with `info.title`/`info.description` replaced by the Excel text. It takes no arguments; paths resolve relative to the script, so it runs from anywhere, and it writes `openapi.yaml` beside itself:

  ```bash
  python3 docs/resources/script.py
  ```

  The source is the copy published in [`wso2/api-specs`](https://github.com/wso2/api-specs/tree/main/openapi/microsoft/graph/v1.0), which is versioned alongside the other specs this org generates connectors from. It is downloaded on first run and cached beside the script as `msgraph-v1.0-openapi.yaml` (~44 MB, untracked); delete that file to pick up a newer upstream. The 1,442-path output is the input to the hand-trimming in item 2.
- **Reason**: Makes the starting point reproducible. A regeneration workflow can fetch the description from `wso2/api-specs` and rebuild this spec instead of depending on a checked-in copy of unknown provenance.

2. Remove operations outside the Excel workbook surface
- **Original**: The Excel-scoped specification produced by item 1 exposes 1,765 operations across 1,442 paths (9.6 MB), which generates a 1.8 MB `client.bal` with 944 records.
- **Updated**: 1,574 operations were removed, leaving 191 operations across 135 paths (660 KB). Every retained operation, `operationId` and all 73 component schemas were passed through unchanged; only whole operations were removed.

  | Removed | Ops | Reason |
  |---|---:|---|
  | Workbook-level table duplicates (`/workbook/tables/...`) | 394 | Byte-identical to the worksheet-scoped tree. The workbook-scoped collection itself is retained, so tables can still be listed workbook-wide. |
  | Worksheet functions (`/workbook/functions/...`) | 369 | 366 spreadsheet formula endpoints. Formula evaluation is a separate concern; shipping it would triple the connector. |
  | Range navigation helpers | 320 | `boundingRect`, `intersection`, `offsetRange`, `entireRow/Column`, `lastCell`, and similar. Each returns another range and is reachable via `range(address=...)`. |
  | Deep chart formatting sub-trees | 184 | Cosmetic styling nested 6+ levels deep under `axes`, `gridlines`, `dataLabels`, `format`. |
  | Range sub-operations on non-primary anchors | 143 | `format`/`sort`/`clear`/`insert`/`delete`/`merge`/`unmerge` kept only on worksheet-level anchors. |
  | `item(name=)` / `itemAt(index=)` variants | 86 | Alternate addressing for resources already reachable by id. |
  | Chained range anchors | 43 | e.g. `range()/usedRange()/cell(...)` — an anchor applied to an anchor. |
  | `DELETE` on singleton navigation properties | 19 | OData codegen noise: Graph cannot delete `format`, `sort`, `filter`, `protection`, `axes`, or the workbook itself. |
  | Range `worksheet` back-links | 16 | `range()/worksheet` — reachable directly. |
  | **Total** | **1,574** | Paths 1,442 → 135. File 9.6 MB → 660 KB. |
- **Reason**: Generating from the upstream description directly yields a connector that is unusable in practice. The retained 191 operations are a superset of the 30 the connector exposed previously.

3. Add an OAuth2 security scheme and a global security requirement
- **Original**: The specification declared no `securitySchemes` and no `security`.
- **Updated**: Added an `oAuth2` scheme to `components.securitySchemes` using the Microsoft identity platform authorization code flow (`authorizationUrl`, `tokenUrl` and `refreshUrl` under `login.microsoftonline.com/common/oauth2/v2.0`, with the `Files.Read`, `Files.ReadWrite` and `offline_access` scopes), plus a document-level `security` requirement of `oAuth2: [Files.ReadWrite]`.
- **Reason**: Without a security scheme the generated `ConnectionConfig` has no `auth` field at all and the client cannot authenticate. Adding it makes the tool emit the connector's existing auth model, `http:BearerTokenConfig|OAuth2RefreshTokenGrantConfig auth`, preserving both the bearer token and OAuth2 refresh token flows the connector supports today.

  Only the `authorizationCode` flow is declared. A `clientCredentials` (app-only) flow is deliberately **not** added: the Microsoft Graph Excel API is delegated-only. Its permissions table lists `Files.ReadWrite` for delegated work/school and personal accounts, and **`Not supported.`** for Application permissions, and the [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) overview lists only the delegated `Files.Read` and `Files.ReadWrite` scopes. Declaring an app-only flow would advertise an authentication mode in which every workbook call fails. This is where this connector differs from the Microsoft Teams connector, whose API does support application permissions and which therefore declares both flows.

4. Add the `workbook-session-id` request header
- **Original**: No operation declared the header, although the specification models `createSession`, `closeSession` and `refreshSession`.
- **Updated**: An optional `workbook-session-id` string header parameter was appended to 190 of the 191 operations — every operation except `POST /{driveId}/items/{driveItemId}/workbook/createSession`.
- **Reason**: Excel sessions are a core feature of this connector; Microsoft recommends creating a session and passing its ID with each request when making more than one call in a period. Without the header the session endpoints return an ID the generated client has no typed way to use, and every call silently runs sessionless. `createSession` is excluded because it mints the session ID and cannot take a pre-existing one.

5. Remove the unresolvable `discriminator` from the `Entity` schema
- **Original**: `Entity` carried a `discriminator` with a `mapping` of 1,162 entries enumerating the full Microsoft Graph type hierarchy — 124 KB, about 18% of the aligned specification.
- **Updated**: The `discriminator` was removed from `Entity`. Its `@odata.type` and `id` properties are unchanged, and the 45 references to `Entity` still resolve.
- **Reason**: Every one of the 1,162 mapping targets was removed by the operation trim above, so none of them resolved; an implicit mapping by schema name would not resolve either. The dead mapping accounted for all 1,162 dangling `$ref` strings in the specification, now zero.

6. Make `@odata.type` optional on every schema
- **Original**: The flatten + align step marked `@odata.type` as `required` on every schema where it appears — 53 `required` arrays across 51 schemas, 8 at the top level and 45 nested inside `allOf[1]` — so `bal openapi` generated `atOdataType` as a required field on 51 records in `ballerina/types.bal`.
- **Updated**: `@odata.type` was removed from every `required` array in `docs/spec/aligned_ballerina_openapi.json`; each array contained nothing else, so all 53 were dropped entirely. The `@odata.type` property definitions themselves are unchanged, so the wire format is unchanged. `atOdataType` is now optional (`string atOdataType?;`) on all 51 records.
- **Reason**: Microsoft Graph returns `@odata.type` only for polymorphic or derived instances and omits it from most responses, so response binding failed whenever it was absent. Unlike the Teams connector — which keeps it required on the two abstract request types `ConversationMember` and `TeamworkNotificationRecipient` — the Excel workbook surface has no polymorphic request payloads, so no schema needs it. Verified by regenerating from the pre-fix specification against a mock that omits `@odata.type`: 21 compilation errors, 8 of them `missing non-defaultable required record field 'atOdataType'`. After the fix the same suite passes. Because the change lives in the specification, regenerating the client reproduces it.

7. Replace the OData "navigation property" doc-comment boilerplate with resource-specific wording
- **Original**: Microsoft Graph's OData metadata describes relationship operations in generic OData terms, and `bal openapi` copied that phrasing verbatim into the generated method docs — summaries such as `Update the navigation property workbook in drives` and `Create new navigation property to comments for drives`, request-body descriptions `New navigation property` and `New navigation property values`, and response descriptions `Retrieved navigation property` and `Created navigation property`. The wording does not say which resource is affected, and every summary ends in "in drives" even when the target is a chart, table, or range. 149 doc lines in `ballerina/client.bal` were affected.
- **Updated**: Rewrote all 149 occurrences in `docs/spec/aligned_ballerina_openapi.json` — 56 operation `summary` fields, 44 `requestBody` descriptions, and 49 response descriptions — into resource-specific text derived from each operation's own path: summaries such as `Update range format`, `Create chart series`, and `Delete worksheet named item`; request bodies such as `The chart to create` and `The workbook comment properties to update`; responses such as `The created chart` and `The retrieved worksheet`. The 56 summaries are unique across the 56 operations. Collections that exist at both workbook and worksheet scope keep their qualifier (`named item` vs `worksheet named item`), and the two range anchors are distinguished as `range` and `addressed range`, matching the `updateRangeFormat` / `updateRangeByAddressFormat` operation IDs.
- **Reason**: The generic OData boilerplate carried no useful information and was frequently misleading about the target resource. Unlike the Microsoft Teams connector, which applied the equivalent change as a post-generation edit to `client.bal`, this change lives in the specification, so regenerating the client reproduces it and no re-application is required.

8. Describe the return value on operations labeled with the generic `Success` response description
- **Original**: Microsoft Graph's OData-to-OpenAPI metadata labels most success responses with the bare description `Success`, and `bal openapi` copies each response `description` verbatim into the generated `# + return -` doc. 122 responses carried it, 72 of them on operations that return a typed body — so the doc line said nothing and did not match the declared return type.
- **Updated**: Rewrote the success-response `description` for those 72 operations in `docs/spec/aligned_ballerina_openapi.json`, deriving the wording from each operation's own ID: `The updated chart`, `The created table row`, `The number of table columns`, `The retrieved range at the given address`, `The chart image as a base64-encoded string, scaled to the given width and height`, and so on. Nouns match the summaries introduced in item 6 (`table column`, not `column`). The remaining 50 responses are the no-content (HTTP 204) operations, every one of which generates `returns error?`; they **keep** `Success`, which correctly describes a no-content success. The 20 `list*`/`count*` operations reference shared response components (`components/responses/...`) whose descriptions — `Retrieved collection` and `The count of the resource` — were left unchanged.
- **Reason**: The generic `Success` gave no indication of what a call returns and contradicted the typed return. Because the change lives in the specification's `description` fields, regenerating the client reproduces the descriptive return docs and no post-generation re-application is required.

9. Document the `refreshUrl` tenant caveat rather than removing the generated default
- **Original**: `OAuth2RefreshTokenGrantConfig.refreshUrl` is generated with the default `https://login.microsoftonline.com/common/oauth2/v2.0/token`. The `/common/` endpoint works only for multi-tenant app registrations; a single-tenant app must authenticate against `https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token`, and the default produces an authentication failure that does not name the endpoint as the cause.
- **Updated**: The default is left in place and the caveat is documented in `ballerina/README.md` — the setup guide calls it out explicitly, and the quickstart takes `refreshUrl` as a `configurable` value instead of hardcoding `/common/`.
- **Reason**: The Microsoft Teams connector drops the default by editing `ballerina/types.bal` after generation, which it records as an edit that must be re-applied on every regeneration. There is no specification-level equivalent here: `bal openapi` derives the `refreshUrl` default from the flow's `tokenUrl` and always emits one — verified by removing `refreshUrl` from the security scheme (no change) and then by altering `tokenUrl` (the default followed it). A post-generation edit would therefore be silently lost on the next regeneration, so the caveat is documented where callers actually read it instead.

10. Keep the canonical Microsoft Graph base URL
- **Original**: `bal openapi align` folds the shared `/drives` prefix into the base URL, producing `servers[0].url = https://graph.microsoft.com/v1.0/drives` with every path starting `/{driveId}/items/...`.
- **Updated**: The base URL is left at `https://graph.microsoft.com/v1.0` and the `/drives` segment moved back into the 135 path keys, so they read `/drives/{driveId}/items/{driveItemId}/workbook/...`.
- **Reason**: `https://graph.microsoft.com/v1.0` is the well-known Microsoft Graph service root, and `/drives` is a resource collection rather than a version or namespace segment. Keeping the canonical root matches the sibling `ballerinax/microsoft.teams` connector, which declines the equivalent fold for `/teams`, so every Graph connector presents the same `serviceUrl`. The change is cosmetic with respect to the requests sent — the resolved URLs are identical either way — and must be re-applied after each `align`, which re-folds the prefix.

11. Change `MainError target` to nullable
- **Original**: The `target` field in `MainError` was `not nullable`.
- **Updated**: The `target` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

12. Change `AddTableRequest address` to nullable
- **Original**: The `address` field in `AddTableRequest` was `not nullable`.
- **Updated**: The `address` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

13. Change `AddWorksheetRequest name` to nullable
- **Original**: The `name` field in `AddWorksheetRequest` was `not nullable`.
- **Updated**: The `name` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

14. Change `AddRowRequest index` to nullable
- **Original**: The `index` field in `AddRowRequest` was `not nullable`.
- **Updated**: The `index` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

15. Change `InnerError date` to nullable
- **Original**: The `date` field in `InnerError` was `not nullable`.
- **Updated**: The `date` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

16. Change `InnerError client-request-id` to nullable
- **Original**: The `client-request-id` field in `InnerError` was `not nullable`.
- **Updated**: The `client-request-id` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

17. Change `InnerError request-id` to nullable
- **Original**: The `request-id` field in `InnerError` was `not nullable`.
- **Updated**: The `request-id` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

18. Change `AddNamedItemRequest name` to nullable
- **Original**: The `name` field in `AddNamedItemRequest` was `not nullable`.
- **Updated**: The `name` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

19. Change `AddNamedItemRequest comment` to nullable
- **Original**: The `comment` field in `AddNamedItemRequest` was `not nullable`.
- **Updated**: The `comment` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

20. Change `ChartImageResponse value` to nullable
- **Original**: The `value` field in `ChartImageResponse` was `not nullable`.
- **Updated**: The `value` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

21. Change `AddColumnRequest name` to nullable
- **Original**: The `name` field in `AddColumnRequest` was `not nullable`.
- **Updated**: The `name` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

22. Change `AddColumnRequest index` to nullable
- **Original**: The `index` field in `AddColumnRequest` was `not nullable`.
- **Updated**: The `index` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

23. Change `ApplyCellColorFilterRequest color` to nullable
- **Original**: The `color` field in `ApplyCellColorFilterRequest` was `not nullable`.
- **Updated**: The `color` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

24. Change `OperationError code` to nullable
- **Original**: The `code` field in `OperationError` was `not nullable`.
- **Updated**: The `code` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

25. Change `OperationError message` to nullable
- **Original**: The `message` field in `OperationError` was `not nullable`.
- **Updated**: The `message` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

26. Change `SortField color` to nullable
- **Original**: The `color` field in `SortField` was `not nullable`.
- **Updated**: The `color` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

27. Change `FilterCriteria color` to nullable
- **Original**: The `color` field in `FilterCriteria` was `not nullable`.
- **Updated**: The `color` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

28. Change `FilterCriteria criterion2` to nullable
- **Original**: The `criterion2` field in `FilterCriteria` was `not nullable`.
- **Updated**: The `criterion2` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

29. Change `FilterCriteria criterion1` to nullable
- **Original**: The `criterion1` field in `FilterCriteria` was `not nullable`.
- **Updated**: The `criterion1` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

30. Change `AddNamedItemFormulaLocalRequest name` to nullable
- **Original**: The `name` field in `AddNamedItemFormulaLocalRequest` was `not nullable`.
- **Updated**: The `name` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

31. Change `AddNamedItemFormulaLocalRequest formula` to nullable
- **Original**: The `formula` field in `AddNamedItemFormulaLocalRequest` was `not nullable`.
- **Updated**: The `formula` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

32. Change `AddNamedItemFormulaLocalRequest comment` to nullable
- **Original**: The `comment` field in `AddNamedItemFormulaLocalRequest` was `not nullable`.
- **Updated**: The `comment` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

33. Change `SessionInfo persistChanges` to nullable
- **Original**: The `persistChanges` field in `SessionInfo` was `not nullable`.
- **Updated**: The `persistChanges` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

34. Change `SessionInfo id` to nullable
- **Original**: The `id` field in `SessionInfo` was `not nullable`.
- **Updated**: The `id` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

35. Change `ApplyCustomFilterRequest criteria1` to nullable
- **Original**: The `criteria1` field in `ApplyCustomFilterRequest` was `not nullable`.
- **Updated**: The `criteria1` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

36. Change `ApplyCustomFilterRequest criteria2` to nullable
- **Original**: The `criteria2` field in `ApplyCustomFilterRequest` was `not nullable`.
- **Updated**: The `criteria2` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

37. Change `ErrorDetails target` to nullable
- **Original**: The `target` field in `ErrorDetails` was `not nullable`.
- **Updated**: The `target` field has been updated to be `nullable`.
- **Reason**: The API can return a null value for this field.
<!-- auto-generated -->

## OpenAPI cli command

The following command was used to generate the Ballerina client from the OpenAPI specification. The command should be executed from the repository root directory.

```bash
# 1. Rebuild the Excel-scoped spec from the Graph description in wso2/api-specs (item 1)
python3 docs/resources/script.py

# 2. Apply the specification changes in items 2-8 to docs/spec/openapi.yaml

# 3. Flatten and align
bal openapi flatten -i docs/spec/openapi.yaml -o docs/spec
bal openapi align   -i docs/spec/flattened_openapi.yaml -o docs/spec

# 4. Generate the client
bal openapi -i docs/spec/aligned_ballerina_openapi.json -o ballerina --mode client --license docs/license.txt
```

Step 2 is not yet scripted. `docs/resources/script.py` reproduces the input to item 2 — the full 1,765-operation workbook surface, written to `docs/resources/openapi.yaml` — but the reduction to the 191 operations this connector ships, and the changes recorded in items 3 to 8, are still applied by hand.

Note: The license header in `docs/license.txt` is dated 2026; update it if necessary.
