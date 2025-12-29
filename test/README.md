This folder contains the Outlook package's tests.

Run tests for this package from the package directory:

```bash
npm test
```

Notes:
- Tests run in the package context so Node will resolve `package.json` and dependencies correctly.
- Live integration tests load credentials via Node's `--env-file`. Place them in `.env/test/outlook`.
- Helpers used only for tests should live under `test/lib/` so they are executed within the package context.

Docs: https://mcp-z.github.io/mcp-outlook
