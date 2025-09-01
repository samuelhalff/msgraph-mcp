import { createMiddleware } from "hono/factory";
import { HTTPException } from "hono/http-exception";
import { Env } from "../../types";
import { logToken } from "./logger.js";

/**
 * msGraphBearerTokenAuthMiddleware validates that the request has a valid Microsoft Graph access token
 * The token is passed in the Authorization header as a Bearer token
 */
export const msGraphBearerTokenAuthMiddleware = createMiddleware<{
  Bindings: Env,
}>(async (c, next) => {
  const authHeader = c.req.header('Authorization');

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    throw new HTTPException(401, { message: 'Missing or invalid access token' });
  }
  const accessToken = authHeader.substring(7);

  // Extract refresh token from a custom header (if provided)
  const refreshToken = c.req.header('X-Refresh-Token') || '';

  // Store tokens in context for use in MSGraphService
  (c as any).set('authContext', {
    accessToken,
    refreshToken,
  });

  await next();
});

/**
 * Get Microsoft Graph OAuth endpoints
 */
export function getMSGraphAuthEndpoint(endpoint: string): string {
  const tenantId = process.env.TENANT_ID || 'common';
  const endpoints = {
    authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
  };
  return endpoints[endpoint as keyof typeof endpoints] || endpoints.authorize;
}

/**
 * Exchange authorization code for access token
 */
export async function exchangeCodeForToken(
  code: string,
  redirectUri: string,
  clientId: string,
  clientSecret?: string,
  codeVerifier?: string,
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const params = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    client_id: clientId,
  });

  if (codeVerifier) {
    params.append('code_verifier', codeVerifier);
  }

  if (clientSecret) {
    params.append('client_secret', clientSecret);
  }

  const response = await fetch(getMSGraphAuthEndpoint('token'), {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params,
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to exchange code for token: ${error}`);
  }

  const tokens = await response.json() as {
    access_token: string;
    token_type: string;
    scope: string;
    expires_in: number;
    refresh_token?: string;
  };

  logToken(tokens.access_token, "OAuth callback");

  return tokens;
}

/**
 * Refresh an access token
 */
export async function refreshAccessToken(
  refreshToken: string,
  clientId: string,
  clientSecret?: string,
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
  });

  if (clientSecret) {
    params.append('client_secret', clientSecret);
  }

  const response = await fetch(getMSGraphAuthEndpoint('token'), {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params,
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to refresh token: ${error}`);
  }

  const tokens = await response.json() as {
    access_token: string;
    token_type: string;
    scope: string;
    expires_in: number;
    refresh_token?: string;
  };

  logToken(tokens.access_token, "Token refresh");

  return tokens;
}