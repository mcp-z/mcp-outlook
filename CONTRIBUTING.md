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

CI follows the Linux/Windows template used by each-package: Node 26, `npm ci`, `prepublishOnly`, a current-runtime test run, and the supported-engine sweep. macOS coverage runs locally. Pull requests receive no provider credentials.

`npm run test:ci` and `npm run test:ci:engines` run the credential-free selection. They exclude `test/integration/mcp/**`, `test/integration/server.test.ts` and `test/integration/outlook-server-spawn.test.ts`; `test/integration/http-server.test.ts` starts local HTTP servers without credentials and runs in this selection. `test/integration/mcp/**` calls the provider with configured credentials. `test/integration/server.test.ts` and `test/integration/outlook-server-spawn.test.ts` make no provider requests; they start the stdio server and check discovery, but the server requires `MS_CLIENT_ID` to start, so they stay outside this selection. Local `npm test` and `npm run test:engines` run all of these; in CI they run through the manually dispatched **Live provider tests** workflow (`.github/workflows/live.yml`, `workflow_dispatch`, described below). Interactive consent suites are separate and run only locally with `TEST_INCLUDE_MANUAL=true`.

`npm test` and `npm run test:engines` retain full discovery. CI sets `TEST_INCLUDE_MANUAL=false`; consent tests require a person and run locally with `TEST_INCLUDE_MANUAL=true`. A green credential-free check does not certify live-provider behavior. Release evidence must include the configured live suites and relevant manual OAuth flows.

### Live provider tests

Run the **Live provider tests** workflow from master, selecting Linux or Windows. It runs the full non-interactive suite once with `--bail`, without an engine matrix or automatic retries. Live runs are manual while request usage and service quotas are being measured. A green PR check covers only the credential-free selection above.

This repository owns its `live-test` GitHub environment, restricted to master, and its own live-job concurrency group. There is no cross-repository coordinator. Concurrent runs in different repositories can still share provider quotas; avoid starting several against the same account at once. Browser consent tests remain local and opt-in.

Configure these individual environment secrets from the existing test configuration and token stores:

- `CI_SECRETS_TOKEN`
- `TEST_ACCOUNT_ID`
- `TEST_REFRESH_TOKEN`
- `TEST_SCOPE`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_TENANT_ID`

`CI_SECRETS_TOKEN` is a fine-grained GitHub PAT with this repository selected and **Environments: Read and write**. GitHub's default workflow token cannot update environment secrets; see the [environment-secret API permissions](https://docs.github.com/en/rest/actions/secrets#create-or-update-an-environment-secret). Keep the PAT's expiry visible to the maintainer; do not use the local broad GitHub login as a CI secret. Provider refresh tokens and GitHub authorization have separate lifetimes.

The seeder validates configuration, renews provider credentials, saves private runner files, and writes replacement refresh tokens back to environment secrets before tests run. An always-run finalizer saves subsequent replacements even when tests fail, then removes runner credential files. Credentials are not cached or uploaded as artifacts. Missing configuration, failed renewal and failed persistence fail the job; CI never opens a consent screen. After provider revocation or expiry, reauthorize locally with the existing setup command and reseed that repository's refresh-token secrets.
