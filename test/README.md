This folder contains the Outlook package's tests.

Run tests for this package from the package directory:

```bash
npm test
```

Notes:
- Tests run in the package context so Node will resolve `package.json` and dependencies correctly.
- `test/lib/env-loader.ts` optionally loads `.env.test` through `portable-env` and must be the first import in every test file. File values override matching inherited values. The shell or CI can provide credentials without a file.
- Helpers used only for tests should live under `test/lib/` so they are executed within the package context.
