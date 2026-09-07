This folder contains the Outlook package's tests.

Run tests for this package from the package directory:

```bash
npm test
```

Notes:
- Tests run in the package context so Node will resolve `package.json` and dependencies correctly.
- Tests load credentials from this package's `.env.test`. Loading is done by `test/lib/env-loader.ts`, which must be the first import in every test file.
- Helpers used only for tests should live under `test/lib/` so they are executed within the package context.
