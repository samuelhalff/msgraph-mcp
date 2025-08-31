import { MSGraphMCP } from "./dist/api/MSGraphMCP.js";
import { exchangeCodeForToken, refreshAccessToken } from "./dist/api/lib/msgraph-auth.js";
import { cors } from "hono/cors";
import { Hono } from "hono";
import { serve } from "@hono/node-server";
import dotenv from "dotenv";

// Load environment variables
dotenv.config();

// Export the MSGraphMCP class
export { MSGraphMCP };

// Helper functions
function getMSGraphAuthEndpoint(type) {
    const endpoints = {
        authorize: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
        token: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    };
    return endpoints[type] || endpoints.authorize;
}

// Simple bearer token authentication middleware
const msGraphBearerTokenAuthMiddleware = async (c, next) => {
    const authHeader = c.req.header('Authorization');
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return c.json({ error: 'Missing or invalid Authorization header' }, 401);
    }
    await next();
};

// Store registered clients in memory
const registeredClients = new Map();

const app = new Hono();

// Enable CORS
app.use('*', cors({
    origin: ['http://localhost:3000', 'https://pbm-ai.ddns.net/mcp/msgraph/'],
    allowHeaders: ['Content-Type', 'Authorization'],
    allowMethods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    credentials: true
}));

// OAuth Discovery endpoint
app.get('/.well-known/oauth-authorization-server', (c) => {
    const baseUrl = c.req.url.split('/.well-known')[0];
    return c.json({
        issuer: baseUrl,
        authorization_endpoint: `${baseUrl}/authorize`,
        token_endpoint: `${baseUrl}/token`,
        jwks_uri: `${baseUrl}/.well-known/jwks.json`,
        response_types_supported: ['code'],
        grant_types_supported: ['authorization_code', 'refresh_token'],
        token_endpoint_auth_methods_supported: ['client_secret_basic', 'client_secret_post'],
        scopes_supported: ['openid', 'profile', 'email', 'https://graph.microsoft.com/.default'],
        claims_supported: ['sub', 'iss', 'aud', 'exp', 'iat', 'auth_time', 'nonce', 'acr', 'amr', 'azp'],
        id_token_signing_alg_values_supported: ['RS256'],
        userinfo_endpoint: `${baseUrl}/userinfo`,
        end_session_endpoint: `${baseUrl}/logout`,
        registration_endpoint: `${baseUrl}/register`,
        // Add PKCE support for LibreChat compatibility
        code_challenge_methods_supported: ['S256', 'plain']
    });
});

// Client registration endpoint
app.post('/register', async (c) => {
    try {
        const registration = await c.req.json();
        if (!registration.client_name || !registration.redirect_uris) {
            return c.json({ error: 'Missing required fields: client_name, redirect_uris' }, 400);
        }

        const client_id = `msgraph-mcp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
        const client = {
            client_id,
            client_name: registration.client_name,
            redirect_uris: Array.isArray(registration.redirect_uris) ? registration.redirect_uris : [registration.redirect_uris],
            grant_types: registration.grant_types || ['authorization_code'],
            response_types: registration.response_types || ['code'],
            scope: registration.scope || 'https://graph.microsoft.com/.default',
            token_endpoint_auth_method: registration.token_endpoint_auth_method || 'client_secret_post',
            created_at: Date.now()
        };

        registeredClients.set(client_id, client);
        return c.json({
            client_id: client.client_id,
            client_secret: 'msgraph-mcp-secret',
            client_id_issued_at: client.created_at,
            client_secret_expires_at: 0,
            redirect_uris: client.redirect_uris,
            grant_types: client.grant_types,
            response_types: client.response_types,
            scope: client.scope,
            token_endpoint_auth_method: client.token_endpoint_auth_method
        });
    } catch (error) {
        console.error('Client registration error:', error);
        return c.json({ error: 'Invalid registration request' }, 400);
    }
});

// Authorization endpoint
app.get('/authorize', async (c) => {
    const { client_id, redirect_uri, scope, state, response_type, code_challenge, code_challenge_method } = c.req.query();

    if (!client_id || !redirect_uri) {
        return c.json({ error: 'Missing required parameters: client_id, redirect_uri' }, 400);
    }

    const client = registeredClients.get(client_id);
    if (!client) {
        return c.json({ error: 'Invalid client_id' }, 400);
    }

    if (!client.redirect_uris.includes(redirect_uri)) {
        return c.json({ error: 'Invalid redirect_uri' }, 400);
    }

    // Store PKCE parameters for this client if provided
    if (code_challenge && code_challenge_method) {
        client.code_challenge = code_challenge;
        client.code_challenge_method = code_challenge_method;
        registeredClients.set(client_id, client);
    }

    const msGraphAuthUrl = new URL(getMSGraphAuthEndpoint('authorize'));
    msGraphAuthUrl.searchParams.set('client_id', process.env.CLIENT_ID || '');
    msGraphAuthUrl.searchParams.set('redirect_uri', redirect_uri);
    msGraphAuthUrl.searchParams.set('scope', scope || client.scope || 'https://graph.microsoft.com/.default');
    msGraphAuthUrl.searchParams.set('response_type', response_type || 'code');
    msGraphAuthUrl.searchParams.set('state', state || '');

    return c.redirect(msGraphAuthUrl.toString());
});

// Token endpoint
app.post('/token', async (c) => {
    try {
        const body = await c.req.parseBody();
        const grant_type = body.grant_type;
        const code = body.code;
        const redirect_uri = body.redirect_uri;
        const client_id = body.client_id;
        const client_secret = body.client_secret;
        const refresh_token = body.refresh_token;
        const code_verifier = body.code_verifier;

        if (grant_type === 'authorization_code') {
            if (!code || !redirect_uri || !client_id) {
                return c.json({ error: 'Missing required parameters' }, 400);
            }

            // Validate PKCE if code_verifier is provided
            if (code_verifier) {
                const client = registeredClients.get(client_id);
                if (client && client.code_challenge && client.code_challenge_method) {
                    // Validate PKCE challenge
                    const crypto = await import('crypto');
                    let expectedChallenge;

                    if (client.code_challenge_method === 'S256') {
                        // Create SHA256 hash of code_verifier
                        const hash = crypto.createHash('sha256');
                        hash.update(code_verifier);
                        expectedChallenge = hash.digest('base64url');
                    } else if (client.code_challenge_method === 'plain') {
                        expectedChallenge = code_verifier;
                    } else {
                        return c.json({ error: 'Unsupported code challenge method' }, 400);
                    }

                    if (expectedChallenge !== client.code_challenge) {
                        return c.json({ error: 'Invalid code verifier' }, 400);
                    }

                    // Clear PKCE data after successful validation
                    delete client.code_challenge;
                    delete client.code_challenge_method;
                    registeredClients.set(client_id, client);
                }
            }

            const tokenResponse = await exchangeCodeForToken(code, redirect_uri, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            return c.json(tokenResponse);
        } else if (grant_type === 'refresh_token') {
            if (!refresh_token) {
                return c.json({ error: 'Missing refresh_token' }, 400);
            }

            const tokenResponse = await refreshAccessToken(refresh_token, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            return c.json(tokenResponse);
        } else {
            return c.json({ error: 'Unsupported grant_type' }, 400);
        }
    } catch (error) {
        console.error('Token exchange error:', error);
        return c.json({ error: 'Token exchange failed' }, 500);
    }
});

// User info endpoint
app.get('/userinfo', msGraphBearerTokenAuthMiddleware, async (c) => {
    return c.json({
        sub: 'user-id',
        name: 'User Name',
        email: 'user@example.com'
    });
});

// Logout endpoint
app.post('/logout', (c) => {
    return c.json({ message: 'Logged out successfully' });
});

// MCP route
app.post('/mcp', async (c) => {
    const authHeader = c.req.header('Authorization');
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return c.json({ error: 'Missing or invalid Authorization header' }, 401);
    }

    const accessToken = authHeader.substring(7);

    try {
        const authContext = {
            accessToken: accessToken,
            refreshToken: c.req.header('X-Refresh-Token') || undefined
        };

        const env = {
            TENANT_ID: process.env.TENANT_ID,
            CLIENT_ID: process.env.CLIENT_ID,
            CLIENT_SECRET: process.env.CLIENT_SECRET,
            ACCESS_TOKEN: process.env.ACCESS_TOKEN,
            REDIRECT_URI: process.env.REDIRECT_URI,
            CERTIFICATE_PATH: process.env.CERTIFICATE_PATH,
            CERTIFICATE_PASSWORD: process.env.CERTIFICATE_PASSWORD,
            MS_GRAPH_CLIENT_ID: process.env.MS_GRAPH_CLIENT_ID,
            OAUTH_SCOPES: process.env.OAUTH_SCOPES,
            USE_GRAPH_BETA: process.env.USE_GRAPH_BETA,
            USE_INTERACTIVE: process.env.USE_INTERACTIVE,
            USE_CLIENT_TOKEN: process.env.USE_CLIENT_TOKEN,
            USE_CERTIFICATE: process.env.USE_CERTIFICATE
        };

        const mcp = new MSGraphMCP(env, authContext);
        await mcp.initialize();

        const server = mcp.server;
        const request = await c.req.json();

        return c.json({
            jsonrpc: "2.0",
            id: request.id || 1,
            result: {
                tools: [
                    { name: "microsoft-graph-api", description: "A versatile tool to interact with Microsoft Graph APIs" },
                    { name: "get-auth-status", description: "Check the current authentication status" }
                ]
            }
        });
    } catch (error) {
        console.error('MCP request error:', error);
        return c.json({ error: 'Internal server error' }, 500);
    }
});

// Health check endpoint
app.get('/health', (c) => {
    return c.json({ status: 'ok', service: 'msgraph-mcp' });
});

// Start the server
const port = process.env.PORT || 3001;
console.log(`🚀 Starting Microsoft Graph MCP Server on port ${port}`);

serve({
    fetch: app.fetch,
    port: Number(port),
});
