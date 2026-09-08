# Changelog

## [1.2.0] - 2026-09-08 — final 1.x release

**This is the last release on the 1.x line, and it is the 2.x code.** The entries below document
what is in it; the 1.x entries that used to head this file are on the `v1.1.4` tag.

The 1.x line is now end-of-life. Rather than backport fixes to it one at a time, this release
carries the whole 2.x tree, so a 1.x consumer gets every fix in one upgrade.

### Changed

- Internals moved from `@modelcontextprotocol/sdk` v1 to the v2 SDK, and the package now serves both
  the 2025 and 2026-07-28 protocol revisions. See the 2.x entries below for what changed.

### Migrating to 2.x

`npm install @mcp-z/mcp-outlook@latest`. If you import types from `@mcp-z/server`, two names moved:
`McpError` → `ProtocolError` and `RequestHandlerExtra` → `ServerContext`.

### Support

None. There will be no further 1.x releases, including for security. Fixes land on 2.x.

## [2.2.0] - 2026-09-07

### Changed

- Clients speaking the 2026-07-28 protocol revision now receive cache hints on list results: `tools/list`, `prompts/list`, `resources/templates/list` and `server/discover` carry a five-minute TTL and `cacheScope: 'public'`, while `resources/list` and `resources/read` stay `private` with no TTL because they vary by account. Previously every cacheable result used the SDK's conservative `ttlMs: 0` / `private` default, which caches nothing. 2025-era clients are unaffected — the fields do not exist on that revision.
- Tools, resources and prompts are now registered in name order, so `tools/list` returns the same order from every connection and a client can keep a cached catalog valid across a reconnect. The listed order differs from previous releases; no tool is added, removed or renamed.

## [2.1.0] - 2026-09-06

### Added

- Serves the 2026-07-28 MCP protocol revision alongside the existing 2025 revision, over both HTTP and stdio. A client speaking either one reaches the same tools; the 2026 revision is stateless, so such a client sends no `initialize` handshake. Legacy 2025 clients are unaffected.

### Fixed

- The MCP server is now built per request (HTTP) and per connection (stdio) rather than shared. The SDK caches the negotiated protocol revision on the server instance, so a shared one pinned itself to whichever revision arrived first and answered the other with `-32601 Method not found`.

## [2.0.0] - 2026-09-06

### Changed

- Migrated to the v2 MCP SDK by way of `@mcp-z/server` 2.x. `McpError`/`ErrorCode` become `ProtocolError`/`ProtocolErrorCode` (the wire codes are unchanged), and `RequestHandlerExtra` becomes `ServerContext`. The MCP surface (`McpServer`, `ResourceTemplate`, etc.) is imported from `@mcp-z/server` rather than `@modelcontextprotocol/sdk` directly.
- The 1.x line is now maintained on `support/1.x`, cut at v1.1.3; releases from that branch publish under the `support-1` dist-tag rather than `latest`.

## [1.1.4] - 2026-09-06

Release infrastructure only, on the `support/1.x` line; nothing a consumer notices. Publishing from that line now requires its `support-1` dist-tag, so a release there can no longer move `latest` onto it.

## [1.1.3] - 2026-09-05

### Fixed

- **Security:** the HTTP transport's Origin check never actually ran. An app-level `cors()` middleware was mounted ahead of `/mcp` and answered its CORS preflight first, so a page loaded from any origin could reach the tools that hold the user's Outlook credentials; the server also bound every network interface instead of loopback-only. `/mcp` now validates Origin and binds to loopback, deriving the allowed origin/host from `BASE_URL` for public deployments. Requires `@mcp-z/server` ^1.2.0 — against older versions the new options are silently ignored and the protection is lost with no error.

## [1.1.2] - 2026-08-31

Documentation formatting only; no functional changes.

## [1.1.1] - 2026-08-30

### Added

- README "Storage backends" section documenting `TOKEN_STORE_URI`/`DCR_STORE_URI` adapter resolution (e.g. installing `@keyv/redis` alongside the server for `redis://`).
- `-v`/`-h` short flags for `--version`/`--help`.

### Fixed

- `keyv-file` is now a direct dependency, so the default `file://` token/DCR store resolves reliably for a globally installed server.

### Removed

- The `version` subcommand added in 1.0.17. Use `--version`/`-v` instead.

## [1.1.0] - 2026-08-29

Internal test and tooling changes only; no consumer-visible effect.

## [1.0.17] - 2026-08-29

### Changed

- `--version`, `--help`, and the `version` subcommand now resolve before the OAuth/MCP SDK dependency graph loads, so they return instantly instead of paying full server-startup cost.

## [1.0.16] - 2026-08-28

Dependency updates only.

## [1.0.15] - 2026-08-23

### Fixed

- Removed two debug `console.log` calls (added in 1.0.13, in `message-search` and its underlying search execution) that wrote to stdout, corrupting the JSON-RPC stream for stdio-transport clients. HTTP transport was unaffected.
- A query requiring multiple search terms (a `$all` clause) no longer builds a `(A AND B)` Microsoft Graph `$search` expression, which Graph silently ignores. Only the first term is now sent to `$search`; the rest are enforced by the client-side filter added in 1.0.12.

## [1.0.14] - 2026-08-23

Internal only: added an MCP registry publish workflow step and refactored internal query-filter typing. No consumer-visible changes.

## [1.0.13] - 2026-01-06

### Changed

- Full-text search terms across `subject`, `text`, and `body` are now combined with `OR` instead of `AND` when built into a Microsoft Graph `$search` query, so a query naming multiple fields matches messages hitting any of them rather than requiring all.

## [1.0.12] - 2026-01-06

### Added

- `message-search` and `messages-export-csv` can now combine full-text search (`subject`, `text`, `body`, `exactPhrase`, `kqlQuery`) with structured filters (`from`, `to`, `cc`, `bcc`, `categories`, `label`, `hasAttachment`, `importance`, `date`) in the same query. Structured filters are applied client-side against the Graph `$search` results, since Graph cannot combine `$search` with those `$filter` fields itself.

## [1.0.11] - 2026-01-06

### Added

- The `query` parameter on `message-search` and `messages-export-csv` now also accepts a JSON string, in addition to a structured object, and a `kqlQuery` field for raw Microsoft Graph KQL search syntax.

## [1.0.10] - 2026-01-04

### Removed

- The `ttl`/`ttlSeconds` store-URI query parameters added in 1.0.9.

## [1.0.9] - 2026-01-04

### Added

- Store URIs (`TOKEN_STORE_URI`, `DCR_STORE_URI`) accept `ttl`/`ttlSeconds` query parameters to set a default expiry for stored entries. (Removed again in 1.0.10.)

## [1.0.8] - 2026-01-04

### Changed

- CSV export and `/files` file serving are now disabled entirely in DCR (self-hosted) mode, so exported mail is never written to disk in that shared-deployment mode; `resourceStoreUri` is ignored when set.

## [1.0.7] - 2026-01-04

### Changed

- **Breaking:** the `--storage-dir`/`STORAGE_DIR` option for CSV export storage is renamed to `--resource-store-uri`/`RESOURCE_STORE_URI` and now takes a `file://` URI (a bare filesystem path is still accepted and normalized to one). `TOKEN_STORE_URI` is now also documented as a supported environment variable.

## [1.0.6] - 2026-01-03

Dependency updates only.

## [1.0.5] - 2026-01-03

### Changed

- Clarified that `BASE_URL` is used for OAuth/DCR callback endpoints as well as file links, not just HTTP file serving (documentation and `server.json` wording only).

## [1.0.4] - 2026-01-02

### Added

- Tool input schemas and `Input` types for `categories-list` and `labels-list` are now exported from the package root, along with the `AuthMiddleware`, `OAuthAdapters`, and `OAuthRuntimeDeps` types.

## [1.0.3] - 2026-01-02

### Added

- Storage location (`~/.mcp-z` by default) is now resolved by searching upward from the current working directory for a `.mcp.json` file, instead of always using the process's working directory.
- The HTTP transport now mounts a loopback OAuth callback router when a redirect URI is configured.

## [1.0.2] - 2025-12-29

Dependency updates only.

## [1.0.1] - 2025-12-29

Version bump only; no code changes.

## [1.0.0] - 2025-12-29

First stable release, published to npm via the MCP registry publish workflow.

## [0.0.0] - 2025-12-28

Initial scaffold: OAuth2 (loopback, device code, and self-hosted DCR modes), stdio and HTTP transport, and the message search/get/send, label/category, and CSV export toolset for Outlook.
