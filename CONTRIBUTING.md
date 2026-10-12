# Contributing to @mcp-z/mcp-outlook

MCP server for Outlook integration with OAuth authentication, message search, and batch operations

## Before Starting

A few conventions here differ from what you might expect:

- **Breaking changes over compatibility.** This project has no compatibility burden yet. Do not add back-compat layers, migration utilities, or wrappers for deprecated APIs - change the API cleanly and bump the major.
- **Keep it approachable.** This is a small community project, not an enterprise codebase. Prefer the simplest solution that fits in the existing files over new abstractions, frameworks, or shared infrastructure.
- **Tests run against real services, not mocks.** Suites call live provider APIs with real credentials, so you need your own test account configured through the shell, CI, or an optional `.env.test` file copied from `.env.test.example`. File values override matching inherited values. A test that fails on credentials is reported, not skipped or loosened.
- **Never write to stdout in server code.** MCP speaks JSON-RPC over stdio; a stray `console.log` corrupts the protocol stream. Use the injected logger, which writes to stderr.
- **Test scratch goes in the package's gitignored `.tmp/`**, never `os.tmpdir()`.

## Branches

`master` is the only maintained release line. The 1.x line is retired and receives no fixes. All changes target `master`.

## Pre-Commit Commands

Install ts-dev-stack globally if not already installed:

```bash
npm install -g ts-dev-stack
```

Run before committing - this builds, type-checks, lints, and tests:

```bash
tsds validate
```

`tsds validate` also runs automatically on `npm publish` via the `prepublishOnly` hook; a failure blocks the publish.

## Testing

```bash
npm run test:setup    # Generate OAuth tokens (interactive, run once)
npm test              # Run the test suite
npm run test:engines  # Run the suite across every supported Node version
```

Specs live in `test/unit/`, mirroring `src/`. Cross-service specs live in `test/integration/`. Both run under `npm test`.

## Package Development

See `README.md` for package overview and usage.

## GitHub Actions

Trusted master CI runs the complete unattended suite on Ubuntu, then Windows. Each fresh platform job uses Node 26 for `npm ci`, `prepublishOnly`, and `npm test -- --bail`, followed by `npm run test:engines -- --bail` across the supported Node versions. The current Node >=20 range selects major even engines 20, 22, 24 and 26. Windows starts only after the complete Ubuntu job succeeds. macOS coverage runs locally. `npm test`, `test:engines`, `test:ci` and `test:ci:engines` all retain full discovery.

Pull requests receive no provider credentials. Their Ubuntu/Windows jobs validate the package and run these eleven existing local files through canonical TSDS and the supported-engine sweep:

```bash
npm test -- "test/exports/*.test.*" "test/unit/**/*.test.ts" "test/integration/http-server.test.ts" "test/integration/server.test.ts" "test/integration/outlook-server-spawn.test.ts"
npm run test:engines -- "test/exports/*.test.*" "test/unit/**/*.test.ts" "test/integration/http-server.test.ts" "test/integration/server.test.ts" "test/integration/outlook-server-spawn.test.ts"
```

The selection includes all export and unit tests, local HTTP/file-serving checks and both stdio startup/discovery suites. The PR job supplies the explicit nonsecret startup configuration `MS_CLIENT_ID=test-client-id`, `MS_TENANT_ID=common` and `HEADLESS=true`. These tests do not invoke provider tools. The eleven suites in `test/integration/mcp/tools/` call real Graph APIs and run in trusted master CI, not PRs. A green PR check does not certify provider behavior.

The normal Outlook suite has no authored browser-consent skip cases. `npm run test:setup` is the existing interactive token-provisioning command and stays local. Trusted tests use the already-consented refresh token. `TEST_INCLUDE_MANUAL=false` and headless mode remain explicit in CI; neither setting turns missing or invalid credentials into success.

### Trusted tests

Master CI calls the existing **Tests** workflow for serial Ubuntu and Windows jobs. The same workflow remains available for manual dispatch from master with a Linux or Windows runner choice. Both paths run the complete unattended suite and supported engines with `--bail`, using this repository's existing credential setup, persistence and cleanup.

The canonical engine runner executes one runtime at a time. `--bail` stops tests after a failure within that runtime; nvu still continues later engine versions before returning the aggregate failure. This can cause further provider requests after an authentication or quota failure within the sweep. The failure stays red; do not start another provider run to retry it. Windows does not follow a failed Ubuntu job.

This repository owns its `live-test` GitHub environment, restricted to master. Automatic and manual token-consuming jobs hold the same `live-provider-tests` concurrency group through persistence and cleanup, with active cancellation disabled and up to 100 pending jobs queued. Environment secrets are read when each job starts, after earlier token jobs finish. There is no cross-repository coordinator: Microsoft OAuth or other repositories can still share provider accounts, refresh grants and quotas. Avoid overlapping those consumers or local interactive setup against the same grant.

Configure these individual environment secrets from the existing test configuration and token stores:

- `CI_SECRETS_TOKEN`
- `TEST_ACCOUNT_ID`
- `TEST_REFRESH_TOKEN`
- `TEST_SCOPE`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_TENANT_ID`

`CI_SECRETS_TOKEN` is a fine-grained GitHub PAT with this repository selected and **Environments: Read and write**. GitHub's default workflow token cannot update environment secrets; see the [environment-secret API permissions](https://docs.github.com/en/rest/actions/secrets#create-or-update-an-environment-secret). Keep the PAT's expiry visible to the maintainer; do not use the local broad GitHub login as a CI secret. Provider refresh tokens and GitHub authorization have separate lifetimes.

The seeder validates configuration, renews provider credentials, saves private runner files, and writes replacement refresh tokens back to environment secrets before tests run. Separate always-run steps persist subsequent replacements and remove owned runner credential files even when tests fail. Forced cancellation, a job timeout or runner loss can prevent finalization, so verify both outcomes in the run. Credentials are not cached or uploaded as artifacts. Missing configuration, failed renewal and failed persistence fail the job; CI never opens a consent screen. After provider revocation or expiry, reauthorize locally with the existing setup command and reseed this repository's refresh-token secrets.
