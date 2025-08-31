import { MSGraphMCP } from "./MSGraphMCP";
import { exchangeCodeForToken, refreshAccessToken } from "./lib/msgraph-auth";
import { cors } from "hono/cors";
import { Hono } from "hono";
import { serve } from "@hono/node-server";
import { MSGraphAuthContext, Env } from "../types";
import dotenv from "dotenv";
import logger from "./lib/logger.js";

// Load environment variables
dotenv.config();

// Export the MSGraphMCP class so the Worker runtime can find it
export {MSGraphMCP};

// Helper functions
function getMSGraphAuthEndpoint(type: string): string {
    // Microsoft Graph authorization endpoints
    const tenantId = process.env.TENANT_ID || 'common';
    const endpoints = {
        authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
        token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
    };
    return endpoints[type as keyof typeof endpoints] || endpoints.authorize;
}

// Simple bearer token authentication middleware for Microsoft Graph
const msGraphBearerTokenAuthMiddleware = async (c: any, next: any) => {
    const authHeader = c.req.header('Authorization');
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return c.json({ error: 'Missing or invalid Authorization header' }, 401);
    }
    // In a real implementation, you would validate the token here
    await next();
};

// Store registered clients in memory (in production, use a database)
interface RegisteredClient {
    client_id: string;
    client_name: string;
    redirect_uris: string[];
    grant_types: string[];
    response_types: string[];
    scope?: string;
    token_endpoint_auth_method: string;
    created_at: number;
}
const registeredClients = new Map<string, RegisteredClient>();

const app = new Hono();

// Enable CORS for all routes
app.use('*', cors({
    origin: ['http://localhost:3000', 'https://librechat.example.com'], // Add your LibreChat domain
    allowHeaders: ['Content-Type', 'Authorization'],
    allowMethods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    credentials: true
}));
app.get('/.well-known/oauth-authorization-server', (c) => {
    logger.info(`/.well-known/oauth-authorization-server endpoint hit, ${JSON.stringify(c.req.query())}`);
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
        registration_endpoint: `${baseUrl}/register`
    });
});

// Client registration endpoint
app.post('/register', async (c) => {
    logger.info(`/register endpoint hit, ${JSON.stringify(c.req.json())}`);
    try {
        const registration = await c.req.json();

        // Validate required fields
        if (!registration.client_name || !registration.redirect_uris) {
            return c.json({ error: 'Missing required fields: client_name, redirect_uris' }, 400);
        }

        // Generate client_id
        const client_id = `msgraph-mcp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

        const client: RegisteredClient = {
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
            client_secret: 'msgraph-mcp-secret', // In production, generate a real secret
            client_id_issued_at: client.created_at,
            client_secret_expires_at: 0, // Never expires
            redirect_uris: client.redirect_uris,
            grant_types: client.grant_types,
            response_types: client.response_types,
            scope: client.scope,
            token_endpoint_auth_method: client.token_endpoint_auth_method
        });
    } catch (error) {
        logger.error('Client registration error:', error);
        return c.json({ error: 'Invalid registration request' }, 400);
    }
});

// Authorization endpoint
app.get('/authorize', async (c) => {
    logger.info(`/authorize endpoint hit, ${JSON.stringify(c.req.query())}`);
    const { client_id, redirect_uri, scope, state, response_type } = c.req.query();

    if (!client_id || !redirect_uri) {
        return c.json({ error: 'Missing required parameters: client_id, redirect_uri' }, 400);
    }

    // Validate client
    const client = registeredClients.get(client_id as string);
    if (!client) {
        return c.json({ error: 'Invalid client_id' }, 400);
    }

    // Validate redirect_uri
    if (!client.redirect_uris.includes(redirect_uri as string)) {
        return c.json({ error: 'Invalid redirect_uri' }, 400);
    }

    // Redirect to Microsoft Graph authorization endpoint
    const msGraphAuthUrl = new URL(getMSGraphAuthEndpoint('authorize'));
    msGraphAuthUrl.searchParams.set('client_id', process.env.CLIENT_ID || '');
    msGraphAuthUrl.searchParams.set('redirect_uri', redirect_uri as string);
    msGraphAuthUrl.searchParams.set('scope', scope || client.scope || 'https://graph.microsoft.com/.default');
    msGraphAuthUrl.searchParams.set('response_type', response_type || 'code');
    msGraphAuthUrl.searchParams.set('state', state || '');

    return c.redirect(msGraphAuthUrl.toString());
});

// Token endpoint
app.post('/token', async (c) => {
    logger.info(`/token endpoint hit, ${JSON.stringify(c.req.query())}`);
    try {
        const body = await c.req.parseBody();
        const grant_type = body.grant_type as string;
        const code = body.code as string;
        const redirect_uri = body.redirect_uri as string;
        const client_id = body.client_id as string;
        const client_secret = body.client_secret as string;
        const refresh_token = body.refresh_token as string;

        if (grant_type === 'authorization_code') {
            if (!code || !redirect_uri || !client_id) {
                return c.json({ error: 'Missing required parameters' }, 400);
            }

            // Exchange code for token
            const tokenResponse = await exchangeCodeForToken(code, redirect_uri, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            return c.json(tokenResponse);
        } else if (grant_type === 'refresh_token') {
            if (!refresh_token) {
                return c.json({ error: 'Missing refresh_token' }, 400);
            }

            // Refresh token
            const tokenResponse = await refreshAccessToken(refresh_token, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            return c.json(tokenResponse);
        } else {
            return c.json({ error: 'Unsupported grant_type' }, 400);
        }
    } catch (error) {
        logger.error('Token exchange error:', error);
        return c.json({ error: 'Token exchange failed' }, 500);
    }
});

// User info endpoint
app.get('/userinfo', msGraphBearerTokenAuthMiddleware, async (c) => {
    logger.info(`/userinfo endpoint hit, ${JSON.stringify(c.req.query())}`);
    // This would typically fetch user info from Microsoft Graph
    return c.json({
        sub: 'user-id',
        name: 'User Name',
        email: 'user@example.com'
    });
});

// Logout endpoint
app.post('/logout', (c) => {
    // Handle logout
    return c.json({ message: 'Logged out successfully' });
});

// MCP route - receives bearer token from LibreChat
app.post('/mcp', async (c) => {
    logger.info(`/mcp endpoint hit, ${JSON.stringify(c.req.query())}`);
    const authHeader = c.req.header('Authorization');
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return c.json({ error: 'Missing or invalid Authorization header' }, 401);
    }

    const accessToken = authHeader.substring(7); // Remove 'Bearer ' prefix

    try {
        // Create auth context for MSGraphMCP
        const authContext: MSGraphAuthContext = {
            accessToken: accessToken,
            refreshToken: c.req.header('X-Refresh-Token') || undefined
        };

        // Create MSGraphMCP instance with the provided token
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

        const mcp = new MSGraphMCP(env as any, authContext);
        await mcp.initialize();

        // Get the MCP server instance
        const server = mcp.server;

        // Handle MCP request using the server's tool calling
        const request = await c.req.json();

        // For now, return a simple response - this would need proper MCP protocol handling
        return c.json({
            jsonrpc: "2.0",
            id: request.id || 1,
            result: {
                tools: [
                    {
                        name: "microsoft-graph-api",
                        description: "A versatile tool to interact with Microsoft Graph APIs"
                    },
                    {
                        name: "get-auth-status",
                        description: "Check the current authentication status"
                    }
                ]
            }
        });
    } catch (error) {
        logger.error('MCP request error:', error);
        return c.json({ error: 'Internal server error' }, 500);
    }
});

// Health check endpoint
app.get('/health', (c) => {
    logger.info(`/health endpoint hit, ${JSON.stringify(c.req.query())}`);
    return c.json({ status: 'ok', service: 'msgraph-mcp' });
});

export default app;

// Start the server if this file is run directly
if (import.meta.url === `file://${process.argv[1]}`) {
    const port = process.env.PORT || 3001;
    logger.info(`🚀 Starting Microsoft Graph MCP Server on port ${port}`);

    serve({
        fetch: app.fetch,
        port: Number(port),
    });
}

