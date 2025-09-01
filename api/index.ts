/* -------------------------------------------------------------------- *
 *  src/index.ts   –   MS Graph MCP Server                              *
 * -------------------------------------------------------------------- */

import { MSGraphMCP } from './MSGraphMCP';
import { exchangeCodeForToken, refreshAccessToken } from './lib/msgraph-auth';
import { cors } from 'hono/cors';
import { Hono } from 'hono';
import { serve } from '@hono/node-server';
import dotenv from 'dotenv';
import logger from './lib/logger.js';

// Initialize environment variables
dotenv.config();
export { MSGraphMCP };

// Helper function to get Microsoft Graph OAuth endpoints
function getMSGraphAuthEndpoint(type: string): string {
  const tenantId = process.env.TENANT_ID || 'common';
  const endpoints = {
    authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
  };
  return endpoints[type as keyof typeof endpoints] || endpoints.authorize;
}

// Middleware for bearer token authentication
const msGraphBearerTokenAuthMiddleware = async (c: any, next: () => Promise<void>) => {
  const authHeader = c.req.header('Authorization');
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return c.json({ error: 'Missing or invalid Authorization header' }, 401);
  }
  await next();
};

// Client registration store
interface RegisteredClient {
  client_id: string;
  client_name: string;
  redirect_uris: string[];
  grant_types: string[];
  response_types: string[];
  scope?: string;
  token_endpoint_auth_method: string;
  created_at: number;
  azure_client_id?: string;
}
const registeredClients = new Map<string, RegisteredClient>();

const app = new Hono();

// CORS setup for LibreChat compatibility
app.use(
  '*',
  cors({
    origin: ['http://localhost:3000', 'https://librechat.example.com', 'http://localhost:3080'], // Added LibreChat default port
    allowHeaders: ['Content-Type', 'Authorization', 'x-refresh-token'],
    allowMethods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    credentials: true,
  })
);

// OAuth 2.0 metadata endpoint
app.get('/.well-known/oauth-authorization-server', (c) => {
  logger.info(`/.well-known/oauth-authorization-server endpoint hit, ${JSON.stringify(c.req.query())}`);
  
  if (!process.env.PUBLIC_BASE_URL) {
    logger.warn('PUBLIC_BASE_URL not set. Falling back to request URL, which may be incorrect in Docker/proxied environments. Set PUBLIC_BASE_URL to your external public URL (e.g., https://your-domain.com/mcp-server) for correct OAuth endpoints.');
  }
  
  const baseUrl = process.env.PUBLIC_BASE_URL || c.req.url.split('/.well-known')[0];
  return c.json({
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/token`,
    jwks_uri: `${baseUrl}/.well-known/jwks.json`,
    response_types_supported: ['code'],
    grant_types_supported: ['authorization_code', 'refresh_token'],
    token_endpoint_auth_methods_supported: ['client_secret_basic', 'client_secret_post', 'none'],
    scopes_supported: ['openid', 'profile', 'email', 'https://graph.microsoft.com/.default'],
    claims_supported: ['sub', 'iss', 'aud', 'exp', 'iat', 'auth_time', 'nonce', 'acr', 'amr', 'azp'],
    id_token_signing_alg_values_supported: ['RS256'],
    userinfo_endpoint: `${baseUrl}/userinfo`,
    end_session_endpoint: `${baseUrl}/logout`,
    registration_endpoint: `${baseUrl}/register`,
  });
});

// Client registration endpoint
app.post('/register', async (c) => {
  logger.info(`/register endpoint hit, ${JSON.stringify(await c.req.json())}`);
  try {
    const registration = await c.req.json();
    if (!registration.client_name || !registration.redirect_uris) {
      return c.json({ error: 'Missing required fields: client_name, redirect_uris' }, 400);
    }
    const client_id = crypto.randomUUID();
    const client: RegisteredClient = {
      client_id,
      client_name: registration.client_name || 'MCP Client',
      redirect_uris: Array.isArray(registration.redirect_uris)
        ? registration.redirect_uris
        : [registration.redirect_uris],
      grant_types: registration.grant_types || ['authorization_code'],
      response_types: registration.response_types || ['code'],
      scope: registration.scope || 'https://graph.microsoft.com/.default',
      token_endpoint_auth_method: registration.token_endpoint_auth_method || 'none',
      created_at: Date.now(),
      azure_client_id: process.env.CLIENT_ID,
    };
    registeredClients.set(client_id, client);
    return c.json({
      client_id: client.client_id,
      client_id_issued_at: client.created_at,
      client_secret_expires_at: 0,
      redirect_uris: client.redirect_uris,
      grant_types: client.grant_types,
      response_types: client.response_types,
      scope: client.scope,
      token_endpoint_auth_method: client.token_endpoint_auth_method,
    });
  } catch (error) {
    logger.error('Client registration error:', error);
    return c.json({ error: 'Invalid registration request' }, 400);
  }
});

// Authorization endpoint
app.get('/authorize', async (c) => {
  logger.info(`/authorize endpoint hit, ${JSON.stringify(c.req.query())}`);
  const { client_id, redirect_uri } = c.req.query();
  if (!client_id || !redirect_uri) {
    return c.json({ error: 'Missing required parameters: client_id, redirect_uri' }, 400);
  }
  const client = registeredClients.get(client_id as string);
  if (!client) {
    return c.json({ error: 'Invalid client_id' }, 400);
  }
  if (!client.redirect_uris.includes(redirect_uri as string)) {
    return c.json({ error: 'Invalid redirect_uri' }, 400);
  }
  const msGraphAuthUrl = new URL(getMSGraphAuthEndpoint('authorize'));
  const azureClientIdToUse = client.azure_client_id || process.env.CLIENT_ID;
  if (!azureClientIdToUse) {
    return c.json({ error: 'Missing Microsoft Client ID' }, 400);
  }
  const incoming = c.req.query();
  Object.keys(incoming).forEach((k) => {
    if (k !== 'client_id') {
      msGraphAuthUrl.searchParams.set(k, String(incoming[k]));
    }
  });
  msGraphAuthUrl.searchParams.set('client_id', azureClientIdToUse);
  return c.redirect(msGraphAuthUrl.toString());
});

// Token endpoint
app.post('/token', async (c) => {
  logger.info(`/token endpoint hit, ${JSON.stringify(c.req.query())}`);
  try {
    const body = await c.req.parseBody();
    const { grant_type, code, redirect_uri, client_id, refresh_token } = body;
    logger.info(`/token called with grant_type=${grant_type}, client_id=${client_id}, redirect_uri=${redirect_uri}`);
    if (grant_type === 'authorization_code') {
      if (!code || !redirect_uri || !client_id) {
        return c.json({ error: 'Missing required parameters' }, 400);
      }
      const regClient = registeredClients.get(client_id as string);
      const azureClientIdToUse = regClient?.azure_client_id || process.env.CLIENT_ID || '';
      const tokenAuthMethod = regClient?.token_endpoint_auth_method || 'client_secret_post';
      const secretToSend = tokenAuthMethod === 'none' ? undefined : process.env.CLIENT_SECRET || undefined;
      const tokenResponse = await exchangeCodeForToken(
        code as string,
        redirect_uri as string,
        azureClientIdToUse,
        secretToSend || '',
        body.code_verifier as string | undefined
      );
      return c.json(tokenResponse);
    } else if (grant_type === 'refresh_token') {
      if (!refresh_token) {
        return c.json({ error: 'Missing refresh_token' }, 400);
      }
      const regClient = registeredClients.get(client_id as string);
      const azureClientIdToUse = regClient?.azure_client_id || process.env.CLIENT_ID || '';
      const tokenAuthMethod = regClient?.token_endpoint_auth_method || 'client_secret_post';
      const secretToSend = tokenAuthMethod === 'none' ? undefined : process.env.CLIENT_SECRET || undefined;
      const tokenResponse = await refreshAccessToken(
        refresh_token as string,
        azureClientIdToUse,
        secretToSend || ''
      );
      return c.json(tokenResponse);
    } else {
      return c.json({ error: 'Unsupported grant_type' }, 400);
    }
  } catch (error) {
    logger.error('Token exchange error:', error);
    return c.json({ error: 'Token exchange failed' }, 500);
  }
});

// Userinfo endpoint
app.get('/userinfo', msGraphBearerTokenAuthMiddleware, async (c) => {
  logger.info(`/userinfo endpoint hit...`);
  return c.json({
    sub: 'user-id',
    name: 'User Name',
    email: 'user@example.com',
  });
});

// Logout endpoint
app.post('/logout', (c) => {
  return c.json({ message: 'Logged out successfully' });
});

// MCP endpoint for LibreChat
app.use('/mcp', msGraphBearerTokenAuthMiddleware);
app.route('/mcp', new Hono().mount('/', MSGraphMCP.serve().fetch));

// Health check endpoint
app.get('/health', (c) => {
  logger.info(`/health endpoint hit...`);
  return c.json({ status: 'ok', service: 'msgraph-mcp' });
});

export default app;

if (import.meta.url === `file://${process.argv[1]}`) {
  const port = process.env.PORT || 3001;
  logger.info(`🚀 Starting Microsoft Graph MCP Server on port ${port}`);
  serve({ fetch: app.fetch, port: Number(port) });
}