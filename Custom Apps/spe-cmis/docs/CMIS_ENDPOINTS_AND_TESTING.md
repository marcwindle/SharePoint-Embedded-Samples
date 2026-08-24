# SPE CMIS Adapter — Endpoints & Testing Report

_Last updated: 2026-08-03 (feature-gap fill-in pass)_

> **Note:** This is an internal development log kept for historical context on
> bugs found and fixed during implementation. It is not end-user setup
> documentation — see the main [README](../README.md) and
> [docs/SETUP_AZURE.md](./SETUP_AZURE.md) for that. Identifiers below are
> placeholders; the original testing was done against a private development
> tenant.

This report summarizes the CMIS Browser Binding endpoints implemented by this Azure
Functions app, and the live testing performed against a real SharePoint Embedded (SPE)
tenant to validate them.

## Test environment

| Setting | Value |
|---|---|
| Tenant | `<your-tenant>.onmicrosoft.com` |
| Container Type ID | `<container-type-id>` |
| Test repository | "CMIS Test" (`<repository-id>`) |
| Auth | Basic Auth (ROPC → delegated Graph token), `admin@<your-tenant>.onmicrosoft.com` |
| Server | `func start` (local, `http://localhost:7071`), auto-rebuilt via `npm run watch` |
| Client | `tests/cmis_client.py` (Python, `cmislib` + Browser Binding) plus ad-hoc scripts |

Route base for all endpoints: `/api/storage/fileStorage/containerTypes/{containerTypeId}/cmis/browser`

## Endpoint inventory

| # | Function | Method | Route | Selector / Action | Purpose | Status |
|---|---|---|---|---|---|---|
| 1 | `getRepositories.ts` | GET | `.../cmis/browser` | *(repository listing)* | List SPE containers as CMIS repositories (paging via `skipCount`/`maxItems`, name-prefix filter) | ✅ Implemented & tested |
| 2 | `getRepositoryInfo.ts` | GET | `.../cmis/browser/{repositoryId}` | *(repository info)* | Repository capabilities, root folder id, versions | ✅ Implemented & tested |
| 3 | `getRepositoryInfo.ts` | POST | `.../cmis/browser/{repositoryId}` | any (`query`, `createType`, etc.) | Repository-level actions | ✅ **New**: returns proper CMIS `notSupported` (405) instead of a generic 404 — verified live |
| 4 | `rootFolder.ts` | GET | `.../cmis/browser/{repositoryId}/root` | `object` | Get properties of an object by `objectId` (defaults to `root`) | ✅ Fixed & tested |
| 5 | `rootFolder.ts` | GET | `.../cmis/browser/{repositoryId}/root` | `children` | List children of an object by `objectId` (defaults to `root`) | ✅ Fixed & tested |
| 6 | `rootFolder.ts` | GET | `.../cmis/browser/{repositoryId}/root` | `content` | Stream a document's content by `objectId` | ✅ Fixed & tested |
| 7 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `createDocument` | Upload a new file under a parent (`objectId`, default `root`) | ✅ Implemented & tested |
| 8 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `createFolder` | Create a new folder under a parent (`objectId`, default `root`) | ✅ Implemented & tested |
| 9 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `delete` | Delete a single object by `objectId` | ✅ **New**: implemented & tested |
| 10 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `deleteTree` | Recursively delete a folder by `objectId` | ✅ **New**: implemented & tested |
| 11 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `update` | Rename / update properties by `objectId` | ✅ **New**: implemented & tested |
| 12 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `move` | Move an object (`objectId`, `targetFolderId`) | ✅ **New**: implemented & tested |
| 13 | `rootFolder.ts` | POST | `.../cmis/browser/{repositoryId}/root` | `setContent` | Replace a document's content by `objectId` | ✅ **New**: implemented |
| 14 | `objectByPath.ts` | GET | `.../cmis/browser/{repositoryId}/root/{*path}` | `object` / `children` / `content` | Same as above, addressed by path — OR by `objectId` if provided (id takes precedence) | ✅ Fixed & tested |
| 15 | `objectByPath.ts` | POST | `.../cmis/browser/{repositoryId}/root/{*path}` | `createDocument` / `createFolder` | Same as rootFolder, with path-based parent lookup as fallback when no `objectId` | ✅ Implemented & tested |
| 16 | `objectByPath.ts` | POST | `.../cmis/browser/{repositoryId}/root/{*path}` | `setContent` | Path-based or `objectId`-based content replace | ✅ Implemented |
| 17 | `objectByPath.ts` | POST | `.../cmis/browser/{repositoryId}/root/{*path}` | `delete` / `deleteTree` / `update` / `move` | Mirrors `rootFolder.ts` (needed due to an Azure Functions routing quirk — see below) | ✅ **New**: implemented & tested |
| 18 | `getRepositoryInfo.ts` | GET | `.../cmis/browser/{repositoryId}` | `typeChildren` / `typeDescendants` / `typeDefinition` (+`typeId`) | Static type definitions for `cmis:document`/`cmis:folder` (no custom subtypes) | ✅ **New**: implemented & tested |
| 19 | `getRepositoryInfo.ts` | POST | `.../cmis/browser/{repositoryId}` | `cmisaction=query` (`q`, `maxItems`, `skipCount`) | Minimal CMIS SQL subset: `SELECT * FROM cmis:document\|folder\|item [WHERE cmis:name (=\|LIKE) '...']`, BFS folder-tree walk from root | ✅ **New**: implemented & tested |
| 20 | `rootFolder.ts` / `objectByPath.ts` | GET | `.../root` / `.../root/{*path}` | `includeACL=true` on the `object` selector | Returns the object's `acl` (via Graph `permissions`) alongside its properties | ✅ **New**: implemented & tested |
| 21 | `rootFolder.ts` / `objectByPath.ts` | POST | `.../root` / `.../root/{*path}` | `checkOut` | Checks out a document (Graph `checkout`), marks it checked-out in the response | ✅ **New**: implemented & tested |
| 22 | `rootFolder.ts` / `objectByPath.ts` | POST | `.../root` / `.../root/{*path}` | `cancelCheckOut` | Discards a checkout (Graph `discardCheckout`) | ✅ **New**: implemented & tested |
| 23 | `rootFolder.ts` / `objectByPath.ts` | POST | `.../root` / `.../root/{*path}` | `checkin` | Checks in a PWC, optionally with new content/name/`checkinComment` (Graph `checkin`) | ✅ **New**: implemented & tested |
| 24 | `rootFolder.ts` / `objectByPath.ts` | POST | `.../root` / `.../root/{*path}` | `applyACL` | Add/remove ACEs (Graph `invite`/permission delete), returns resulting ACL | ✅ **New**: implemented & tested |

**Still not implemented (structural gaps — SPE has no data model for these):** relationships, policies, multi-filing/unfiling. Capabilities advertised by `getRepositoryInfo` were updated accordingly: `capabilityQuery: 'metadataonly'`, `capabilityACL: 'manage'` (previously `capabilityAcl: 'discover'` — the field name itself was also buggy, see below).

## Key bug found and fixed this session

**Symptom:** Recursive folder listings in the test client appeared to show an exponentially growing, duplicated nested "Fabrikam" folder structure.

**Root cause:** `rootFolder.ts` and `objectByPath.ts` GET handlers were **ignoring the `objectId` parameter entirely** — every "list children of folder X" request actually returned the **root** folder's children again, regardless of which folder was requested. Recursing into what looked like sub-folders just kept re-listing root, producing the illusion of infinite nested duplication. **No data was actually duplicated in SPE.**

**Fix:** Both handlers now extract `objectId` (defaulting to `'root'` when absent) and pass it through to `getDriveItem` / `listChildren` for the `object`, `children`, and `content` selectors, and use it (or fall back to path lookup) as the parent for `createDocument` / `createFolder`.

**Routing quirk discovered along the way:** A bare request to `.../root?objectId=X&cmisselector=children` (no extra path segment) is routed by Azure Functions to `objectByPath` (route `root/{*path}` matching an empty path) rather than to `rootFolder` (route `root` exact match). Because either function can receive these requests, both files implement identical `objectId`-handling logic via shared helpers in `src/lib/cmis/objectActions.ts`.

## Feature-gap fill-in pass (versioning, ACL, query, type definitions)

Implemented the 5 previously-known gaps that were structurally fillable (relationships, policies,
and multi-filing/unfiling remain out of scope — SPE's single-parent-tree data model has no
equivalent). Validated end-to-end with a new ad-hoc script, `tests/test_gaps.py`, against
the same live "CMIS Test" repository: type definitions, query, ACL discover+manage, and
checkOut/checkIn/cancelCheckOut all passed after fixing four real bugs surfaced only by live
testing (not caught by TypeScript compilation or code review):

1. **Missing `repositoryUrl`** — `RepositoryInfo` only ever returned `rootFolderUrl`; cmislib's
   `getRepositoryUrl()`/`query()` require `repositoryUrl` and threw `KeyError: 'repositoryUrl'`.
   Fixed by adding the field to the type and to `mapContainerToRepository`.
2. **`capabilityAcl` → `capabilityACL` casing bug** — cmislib strips the `'capability'` prefix from
   wire keys and looks up the capitalized `['ACL']` key exactly; the old `capabilityAcl` casing
   produced `'Acl'`, causing `KeyError: 'ACL'` in `getACL()`.
3. **Repository-level POST only parsed `application/x-www-form-urlencoded`** — cmislib's
   `repo.query()` actually POSTs `multipart/form-data`, so `cmisaction=query` always failed with
   `"cmisaction is required"`. Fixed by extracting the multipart-form parser (previously duplicated
   in `rootFolder.ts`/`objectByPath.ts`) into a shared `src/lib/http/multipartForm.ts` module and
   using it in `getRepositoryInfo.ts` too.
4. **Case-sensitive `includeACL`/`includeAllowableActions` check** — the server compared the query
   param via `=== 'true'`, but Python's `requests` library serializes a Python `True` as the query
   string literal `True` (capital T), not `true`. This silently disabled ACL/allowable-actions
   inclusion for any Python client. Fixed to a case-insensitive comparison in both `rootFolder.ts`
   and `objectByPath.ts`.

## Testing performed

### 1. Full client test suite (`cmis_client.py`)
Ran against the live "CMIS Test" repository:
- Listed 6 repositories via `getRepositories`.
- Located the "CMIS Test" repository and fetched repository info.
- Listed root contents — **flat, correct structure** (9 items: "Fabrikam" folder + 8 files), confirming the objectId fix.
- **Deleted all 9 items** (files + the "Fabrikam" folder, exercising `delete`) — all succeeded.
- **Re-uploaded** the full `docs/` sample set (files + nested "Fabrikam" folder with 7 files inside), exercising `createDocument`/`createFolder`.
- Listed root contents again — same correct flat 9-item structure, no duplication.
- Result: **"All tests completed successfully!"**

### 2. Ad-hoc `update`/`move` smoke test
A small standalone script (created, run, then removed) validated the newly implemented actions directly:
1. Create a document (`rename-test.txt`) at root.
2. `update` action: rename it to `renamed-test.txt` — verified the new name round-tripped correctly.
3. Create a folder (`move-target`).
4. `move` action: move the renamed document into the new folder.
5. List the folder's children — confirmed the document is now inside it.
6. Clean up (delete document + folder).

Result: **PASSED** — both `update` and `move` work correctly end-to-end against the live tenant.

### 3. Repository-level POST (`query`, etc.) — superseded, see section 4
Initially sent a direct `POST .../cmis/browser/{repositoryId}` with `cmisaction=query` via
`Invoke-WebRequest` and confirmed it returned an HTTP 405 `notSupported` error body (the route
wasn't registered for POST at all, falling through to a generic 404, before that fix). `query` has
since been fully implemented — see section 4.

### 4. Feature-gap validation (`test_gaps.py`)
Ran the new `tests/test_gaps.py` script against the live "CMIS Test" repository to
validate all 5 newly-implemented feature gaps, iterating until every section passed:
- **Type definitions:** `getTypeDefinition('cmis:document')` (22 properties), `getTypeDefinition('cmis:folder')`, `getTypeChildren()`, `getTypeDescendants()` — all returned correctly.
- **Query:** `SELECT * FROM cmis:document WHERE cmis:name LIKE 'Fabrikam%'` correctly matched all 7 "Fabrikam Expenses" documents.
- **ACL:** created a test folder, `getACL()` returned an empty ACE list, `applyACL()` (adding a read ACE for the test user) succeeded and the resulting ACL reflected 1 ace.
- **Versioning:** created a document, `checkout()` returned a valid PWC id, `checkin(checkinComment=...)` succeeded; a second document was checked out and `cancelCheckout()` succeeded.
- Cleanup (`deleteTree` + individual `delete`) removed all test artifacts.

Result: **All 4 sections + cleanup PASSED** end-to-end against the live tenant.

## Notes on test methodology
- Terminal output capture (`get_terminal_output` / killing background terminals) proved unreliable for reading full script output — snapshots were sometimes stale or contained old scrollback. **Redirecting script output directly to a file** (`... > out.txt 2>&1`) and reading the file was the reliable approach used for final verification.
- The local `func start` process does **not** hot-reload compiled `dist/` output. The `npm run watch` background task only recompiles TypeScript; the running `func start` host must be manually killed and restarted after every `npm run build`/watch-triggered rebuild, or it silently keeps serving stale code.

## Real third-party client testing: Apache Chemistry OpenCMIS Workbench

To prove protocol compliance against something closer to legacy CMIS clients (SAP, Doc Center,
etc.) than our own Python test scripts, we connected the official **Apache Chemistry OpenCMIS
Workbench** (a Java Swing GUI reference client for the CMIS Browser Binding) directly to the local
server (`http://localhost:7071/api/storage/fileStorage/containerTypes/{id}/cmis/browser`, Basic
Auth). This surfaced a real interoperability bug that none of the `cmislib`-based tests had caught.

### Bug found: `cmisaction` not read from the query string on POST

**Symptom:** Workbench's "Create folder" and "Query" actions both failed with
`"cmisaction is required"`, even though the server already had logic to parse `cmisaction` out of
both `multipart/form-data` and `application/x-www-form-urlencoded` POST bodies.

**Root cause:** `rootFolder.ts`/`objectByPath.ts`'s `handlePost` only fell back to
`request.query.get('cmisaction')` in the `else` branch for an *unrecognized* content-type — for the
`multipart/form-data` and `application/x-www-form-urlencoded` branches, `cmisaction` was read
**exclusively** from the parsed body. OpenCMIS Workbench sends `cmisaction` as a query-string
parameter on the POST URL (e.g. `POST .../root?cmisaction=createFolder`) while still using an
url-encoded/multipart body for the rest of the fields (`propertyId[0]`, `propertyValue[0]`, etc.) —
which is valid per the CMIS Browser Binding spec, but the server ignored the query string whenever
a recognized form content-type was present. (`getRepositoryInfo.ts`'s `postRepositoryInfo`, fixed
earlier this session for the multipart/`query` gap, already merged query + body correctly — the gap
was isolated to `rootFolder.ts`/`objectByPath.ts`.)

**Fix:** Both `handlePost` functions now initialize `cmisaction` from `request.query.get('cmisaction')`
**first**, then let the parsed multipart/url-encoded body override it if present — treating the query
string and form body as a single merged parameter source rather than either/or.

**Verification:** A Python `requests` smoke test replicating Workbench's exact request shape
(`cmisaction` in the query string, `propertyId[0]`/`propertyValue[0]` in an url-encoded body) against
`POST .../root?cmisaction=createFolder` returned `201 Created` with the new folder's properties. The
full `test_gaps.py` suite was re-run afterward with no regressions (type definitions, query, ACL,
versioning, cleanup all passed).

## Addendum: chunked Transfer-Encoding request bodies were silently empty

**Symptom:** Even after the fix above, Workbench's "Create folder" and "Query" actions still
intermittently failed with `"cmisaction is required"`.

**Root cause:** Azure Functions Node.js isolated-worker HTTP triggers return an **empty** request
body when a client sends `Transfer-Encoding: chunked` with no `Content-Length` header — a known
platform limitation (see `Azure/azure-functions-host#7930` and
`Azure/azure-functions-nodejs-library#171`). OpenCMIS Workbench's Java HTTP client sends chunked
requests by default, so its POST bodies were vanishing before reaching the server's form-parsing
code. Proven with two independent HTTP clients (Python `requests` forced into chunked mode via a
generator body, and `curl --data-binary "@-" -H "Transfer-Encoding: chunked"`) reproducing the
identical failure.

**Fix:** Enabled Node.js "HTTP Streams" in `src/index.ts` (`app.setup({ enableHttpStream: true })`).
This requires Azure Functions Core Tools `>= 4.0.5530` and runtime `>= 4.28` locally (the
`@azure/functions` npm package only needs `>= 4.3.0`, already satisfied). Verified via both smoke
tests flipping from `400` to `200`, plus a full `test_gaps.py` regression run with no breakage, and
manual end-to-end testing directly against the real OpenCMIS Workbench GUI.

