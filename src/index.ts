import { createConfig, handleVersionHelp } from './setup/config.ts';
import { createHTTPServer } from './setup/http.ts';
import { createStdioServer } from './setup/stdio.ts';
import type { ServerConfig } from './types.ts';

export { MS_SCOPE } from './constants.ts';
export * as mcp from './mcp/index.ts';
// Export input schemas for documentation visibility
export type { Input as CategoriesListInput } from './mcp/tools/categories-list.ts';
export { inputSchema as categoriesListInputSchema } from './mcp/tools/categories-list.ts';
export type { Input as LabelsListInput } from './mcp/tools/labels-list.ts';
export { inputSchema as labelsListInputSchema } from './mcp/tools/labels-list.ts';
export * as schemas from './schemas/index.ts';
export * as setup from './setup/index.ts';
export type { AuthMiddleware, OAuthAdapters, OAuthRuntimeDeps } from './setup/oauth-microsoft.ts';
export * from './types.ts';

export async function startServer(config: ServerConfig): Promise<void> {
  const { logger, close } = config.transport.type === 'stdio' ? await createStdioServer(config) : await createHTTPServer(config);

  process.on('SIGINT', async () => {
    await close();
    process.exit(0);
  });

  logger.info(`Server started with ${config.transport.type} transport`);
  await new Promise(() => {});
}

export default async function main(): Promise<void> {
  // Check for help/version flags FIRST, before config parsing
  const versionHelpResult = handleVersionHelp(process.argv);
  if (versionHelpResult.handled) {
    console.log(versionHelpResult.output);
    process.exit(0);
  }

  // Only parse config if no help/version flags
  const config = createConfig();
  await startServer(config);
}

if (process.argv[1] === new URL(import.meta.url).pathname) {
  main();
}
