import { type AccountAuthProvider, AccountServer } from '@mcp-z/oauth';
import { createDcrRouter, createLoopbackCallbackRouter, DcrOAuthProvider, DeviceCodeProvider, LoopbackOAuthProvider } from '@mcp-z/oauth-microsoft';
import type { Logger, PromptModule, ToolModule } from '@mcp-z/server';
import type { Router } from 'express';
import type { Keyv } from 'keyv';
import { MS_SCOPE } from '../constants.ts';
import type { ServerConfig } from '../types.ts';

/**
 * Outlook OAuth runtime dependencies
 *
 * Contains the required dependencies for OAuth functionality including logger,
 * token storage, and optional DCR store.
 */
export interface OAuthRuntimeDeps {
  logger: Logger;
  tokenStore: Keyv<unknown>;
  dcrStore?: Keyv<unknown>;
}

/**
 * Auth middleware wrapper with withToolAuth/withResourceAuth/withPromptAuth methods
 *
 * Provides authentication middleware methods for tools, resources, and prompts.
 */
export interface AuthMiddleware {
  withToolAuth<T extends ToolModule>(module: T): T;
  withResourceAuth<T>(module: T): T;
  withPromptAuth<T>(module: T): T;
}

/**
 * Result of OAuth adapter creation
 *
 * Contains the primary OAuth adapter, middleware, auth adapter,
 * account tools, prompts, and optional routers.
 */
export interface OAuthAdapters {
  primary: LoopbackOAuthProvider | DeviceCodeProvider | DcrOAuthProvider;
  middleware: AuthMiddleware;
  authAdapter: AccountAuthProvider;
  accountTools: ToolModule[];
  accountPrompts: PromptModule[];
  dcrRouter?: Router; // DCR OAuth endpoints (only for dcr mode)
  loopbackRouter?: Router;
}

/**
 * Create Outlook OAuth adapters based on transport configuration
 *
 * Returns primary adapter, pre-configured middleware, auth email provider,
 * and pre-selected account tools based on auth mode.
 * @param config - Outlook server configuration
 * @param deps - Runtime dependencies (logger, tokenStore, etc.)
 * @returns OAuth adapters with pre-configured middleware and account tools
 */
export async function createOAuthAdapters(config: ServerConfig, deps: OAuthRuntimeDeps, baseUrl?: string): Promise<OAuthAdapters> {
  const { logger, tokenStore, dcrStore } = deps;
  const oauthStaticConfig = {
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    scope: MS_SCOPE,
    tenantId: config.tenantId,
    auth: config.auth,
    headless: config.headless,
    redirectUri: config.transport.type === 'stdio' ? undefined : config.redirectUri,
    ...(baseUrl && { baseUrl }),
  };

  let primary: LoopbackOAuthProvider | DeviceCodeProvider | DcrOAuthProvider;

  // DCR mode - Dynamic Client Registration with HTTP-only support
  if (oauthStaticConfig.auth === 'dcr') {
    logger.debug('Creating DCR provider', { service: oauthStaticConfig.service });

    // DCR requires dcrStore and baseUrl
    if (!dcrStore) {
      throw new Error('DCR mode requires dcrStore to be configured');
    }
    if (!oauthStaticConfig.baseUrl) {
      throw new Error('DCR mode requires baseUrl to be configured');
    }

    // Create DcrOAuthProvider (stateless provider that receives tokens from verification context)
    primary = new DcrOAuthProvider({
      clientId: oauthStaticConfig.clientId,
      ...(oauthStaticConfig.clientSecret && { clientSecret: oauthStaticConfig.clientSecret }),
      tenantId: oauthStaticConfig.tenantId,
      scope: oauthStaticConfig.scope,
      verifyEndpoint: `${oauthStaticConfig.baseUrl}/oauth/verify`,
      logger,
    });

    // Create DCR OAuth router with authorization server endpoints
    const dcrRouter = createDcrRouter({
      store: dcrStore,
      issuerUrl: oauthStaticConfig.baseUrl,
      baseUrl: oauthStaticConfig.baseUrl,
      scopesSupported: oauthStaticConfig.scope.split(' '), // Convert space-separated scope to array
      clientConfig: {
        clientId: oauthStaticConfig.clientId,
        ...(oauthStaticConfig.clientSecret && { clientSecret: oauthStaticConfig.clientSecret }),
        tenantId: oauthStaticConfig.tenantId,
      },
    });

    // DCR uses bearer token authentication with middleware validation
    const middleware = primary.authMiddleware();

    // Create auth email provider (stateless, no token storage)
    const authAdapter: AccountAuthProvider = {
      getAccessToken: () => {
        throw new Error('DCR mode does not support getAccessToken - tokens are provided via bearer auth');
      },
      getUserEmail: () => {
        throw new Error('DCR mode does not support getUserEmail - tokens are provided via bearer auth');
      },
    };

    // No account management tools for DCR (multi-user management at client registration level)
    const accountTools: ToolModule[] = [];
    const accountPrompts: PromptModule[] = [];

    logger.info('DCR provider created for Outlook', {
      service: oauthStaticConfig.service,
      baseUrl: oauthStaticConfig.baseUrl,
    });

    return { primary, middleware: middleware as AuthMiddleware, authAdapter, accountTools, accountPrompts, dcrRouter };
  }

  // Device code mode - similar to service accounts (single static identity)
  if (oauthStaticConfig.auth === 'device-code') {
    logger.debug('Creating device code provider', { service: oauthStaticConfig.service });
    const deviceCodeProvider = new DeviceCodeProvider({
      service: oauthStaticConfig.service,
      clientId: oauthStaticConfig.clientId,
      tenantId: oauthStaticConfig.tenantId,
      scope: oauthStaticConfig.scope,
      logger,
      tokenStore,
      headless: oauthStaticConfig.headless,
    });
    primary = deviceCodeProvider;

    // Device code uses single-user middleware (no account management needed)
    const middleware = deviceCodeProvider.authMiddleware();

    // Create auth email provider
    const authAdapter: AccountAuthProvider = {
      getAccessToken: (accountId) => deviceCodeProvider.getAccessToken(accountId),
      getUserEmail: (accountId) => deviceCodeProvider.getUserEmail(accountId),
    };

    // No account management tools for device code (like service accounts)
    const accountTools: ToolModule[] = [];
    const accountPrompts: PromptModule[] = [];

    logger.info('Device code provider created for Outlook', {
      service: oauthStaticConfig.service,
    });

    return { primary, middleware: middleware as AuthMiddleware, authAdapter, accountTools, accountPrompts };
  }

  // Always create primary adapter for file-based OAuth (loopback mode)
  primary = new LoopbackOAuthProvider({
    service: oauthStaticConfig.service,
    clientId: oauthStaticConfig.clientId,
    clientSecret: oauthStaticConfig.clientSecret,
    scope: oauthStaticConfig.scope,
    tenantId: oauthStaticConfig.tenantId,
    headless: oauthStaticConfig.headless,
    logger,
    tokenStore,
    ...(oauthStaticConfig.redirectUri !== undefined && { redirectUri: oauthStaticConfig.redirectUri }),
  });

  // Create auth email provider (used by account management tools)
  const authAdapter: AccountAuthProvider = primary;

  // Select middleware AND account tools based on auth mode
  const middleware: ReturnType<LoopbackOAuthProvider['authMiddleware']> = primary.authMiddleware();

  // Loopback OAuth - multi-account mode
  const result = AccountServer.createLoopback({
    service: oauthStaticConfig.service,
    store: tokenStore,
    logger,
    auth: authAdapter,
  });
  const accountTools: ToolModule[] = result.tools as ToolModule[];
  const accountPrompts: PromptModule[] = result.prompts as PromptModule[];
  logger.debug('Loopback OAuth (multi-account mode)', { service: oauthStaticConfig.service });

  const loopbackRouter = oauthStaticConfig.redirectUri ? createLoopbackCallbackRouter(primary) : undefined;

  return { primary, middleware: middleware as AuthMiddleware, authAdapter, accountTools, accountPrompts, loopbackRouter };
}
