import {createMiddleware} from "hono/factory";
import {HTTPException} from "hono/http-exception";
import { Env } from "../../types";
import { logToken } from "./logger.js";

/**
 * spotifyBearerTokenAuthMiddleware validates that the request has a valid Spotify access token
 * The token is passed in the Authorization header as a Bearer token
 */
export const spotifyBearerTokenAuthMiddleware = createMiddleware<{
    Bindings: Env,
}>(async (c, next) => {
    const authHeader = c.req.header('Authorization')

    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        throw new HTTPException(401, {message: 'Missing or invalid access token'})
    }
    const accessToken = authHeader.substring(7);

    // For Spotify, we don't validate the token here - we'll let the API calls fail if it's invalid
    // and handle token refresh in the SpotifyService
    
    // Extract refresh token from a custom header (if provided)
    const refreshToken = c.req.header('X-Spotify-Refresh-Token') || ''

    // @ts-expect-error Props go brr
    c.executionCtx.props = {
        accessToken,
        refreshToken,
    }

    await next()
})

/**
 * Get Microsoft Graph OAuth endpoints
 */
export function getMSGraphAuthEndpoint(endpoint: string): string {
    const tenantId = process.env.TENANT_ID || 'common';
    const endpoints = {
        authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
        token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
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
    clientSecret: string,
    codeVerifier?: string
): Promise<{
    access_token: string
    token_type: string
    scope: string
    expires_in: number
    refresh_token: string
}> {
    const params = new URLSearchParams({
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        client_id: clientId,
    })

    // Add code_verifier for PKCE flow if provided
    if (codeVerifier) {
        params.append('code_verifier', codeVerifier)
    }

    // Only include client_secret when provided (confidential clients). Public clients must not send client_secret.
    if (clientSecret) {
        params.append('client_secret', clientSecret)
    }

    const response = await fetch(getMSGraphAuthEndpoint('token'), {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params
    })

    if (!response.ok) {
        const error = await response.text()
        throw new Error(`Failed to exchange code for token: ${error}`)
    }

    const tokens = await response.json() as {
        access_token: string;
        token_type: string;
        scope: string;
        expires_in: number;
        refresh_token: string;
    };

    // Log the received token
    logToken(tokens.access_token, "OAuth callback");

    return tokens;
}

/**
 * Refresh an access token
 */
export async function refreshAccessToken(
    refreshToken: string,
    clientId: string,
    clientSecret: string
): Promise<{
    access_token: string
    token_type: string
    scope: string
    expires_in: number
    refresh_token?: string
}> {
    const params = new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        client_id: clientId,
    })
    if (clientSecret) {
        params.append('client_secret', clientSecret)
    }

    const response = await fetch(getMSGraphAuthEndpoint('token'), {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params
    })

    if (!response.ok) {
        const error = await response.text()
        throw new Error(`Failed to refresh token: ${error}`)
    }

    const tokens = await response.json() as {
        access_token: string;
        token_type: string;
        scope: string;
        expires_in: number;
        refresh_token?: string;
    };

    // Log the refreshed token
    logToken(tokens.access_token, "Token refresh");

    return tokens;
} 