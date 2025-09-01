import { Context, Next } from 'hono';
import { HTTPException } from 'hono/http-exception';
import logger from './logger.js';
import { z } from 'zod';
import { StatusCode } from 'hono/utils/http-status';
import { Env } from '../../types.js';

// Define token response schema
const TokenResponseSchema = z.object({
  access_token: z.string(),
  token_type: z.string(),
  expires_in: z.number().optional(),
  refresh_token: z.string().optional(),
  scope: z.string().optional(),
});

type TokenResponse = z.infer<typeof TokenResponseSchema>;

export const msGraphBearerTokenAuthMiddleware = async (c: Context<{ Variables: Env }>, next: Next) => {
  const authHeader = c.req.header('Authorization');
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    logger.error('Missing or invalid Authorization header', { authHeader: authHeader || 'null' });
    throw new HTTPException(401, { message: 'Missing or invalid Authorization header' });
  }

  const token = authHeader.replace('Bearer ', '');
  try {
    // Placeholder for token validation (e.g., JWT verification)
    (c as any).set('msGraphAuth', { accessToken: token });
    await next();
  } catch (error) {
    logger.error('Token validation failed', { error: (error as Error).message });
    throw new HTTPException(401, { message: 'Invalid token' });
  }
};

export function getMSGraphAuthEndpoint(tenantId: string): string {
  // Build the Microsoft identity authorize endpoint for the given tenant
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
}

export async function refreshAccessToken(
    tenantId: string,
    refreshToken: string,
    clientId: string,
    clientSecret: string
): Promise<TokenResponse> {
    const response = await fetch(getMSGraphAuthEndpoint(tenantId), {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            client_id: clientId,
            client_secret: clientSecret
        }).toString()
    });

    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to refresh token: ${error}`);
    }

    const data = await response.json();
    return TokenResponseSchema.parse(data);
}

export async function exchangeCodeForToken(
    code: string,
    redirectUri: string,
    clientId: string,
    clientSecret: string,
    codeVerifier?: string
): Promise<TokenResponse> {
    const tokenEndpoint = `https://login.microsoftonline.com/common/oauth2/v2.0/token`;

    const body = new URLSearchParams({
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        client_id: clientId,
        client_secret: clientSecret,
    });

    if (codeVerifier) {
        body.append('code_verifier', codeVerifier);
    }

    const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: body.toString()
    });

    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to exchange code for token: ${error}`);
    }

    const data = await response.json();
    return TokenResponseSchema.parse(data);
}

export function getMicrosoftAuthEndpoint(type: 'authorize' | 'token'): string {
    if (type === 'authorize') {
        return 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
    } else {
        return 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    }
}