import { MSGraphMCP } from "./dist/api/MSGraphMCP.js";
import { exchangeCodeForToken, refreshAccessToken } from "./dist/api/lib/msgraph-auth.js";
import { cors } from "hono/cors";
import { Hono } from "hono";
import { serve } from "@hono/node-server";
import dotenv from "dotenv";
import winston from "winston";

// Load environment variables
dotenv.config();

// Create logger instance
const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  defaultMeta: { service: 'msgraph-mcp' },
  transports: [],
  exceptionHandlers: [],
  rejectionHandlers: [],
});

// Add file transports with error handling
try {
  logger.add(new winston.transports.File({
    filename: 'logs/error.log',
    level: 'error',
    handleExceptions: true,
    handleRejections: true
  }));

  logger.add(new winston.transports.File({
    filename: 'logs/combined.log',
    handleExceptions: true,
    handleRejections: true
  }));
} catch (error) {
  console.warn('Failed to initialize file logging, falling back to console only:', error instanceof Error ? error.message : String(error));
}

// Always add console transport
logger.add(new winston.transports.Console({
  format: winston.format.combine(
    winston.format.colorize(),
    winston.format.simple()
  ),
  handleExceptions: true,
  handleRejections: true
}));

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
    logger.info('OAuth discovery endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        method: c.req.method,
        path: c.req.path
    });
    
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
        logger.info('Client registration attempt', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            userAgent: c.req.header('user-agent'),
            clientName: registration.client_name,
            redirectUris: registration.redirect_uris
        });
        
        if (!registration.client_name || !registration.redirect_uris) {
            logger.warn('Client registration failed: missing required fields', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                hasClientName: !!registration.client_name,
                hasRedirectUris: !!registration.redirect_uris
            });
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
        
        logger.info('Client registration successful', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            clientId: client_id,
            clientName: client.client_name
        });
        
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
        logger.error('Client registration error:', error);
        return c.json({ error: 'Invalid registration request' }, 400);
    }
});

// Authorization endpoint
app.get('/authorize', async (c) => {
    const { client_id, redirect_uri, scope, state, response_type, code_challenge, code_challenge_method } = c.req.query();
    
    logger.info('Authorization endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        clientId: client_id,
        redirectUri: redirect_uri,
        scope: scope,
        responseType: response_type,
        hasCodeChallenge: !!code_challenge,
        codeChallengeMethod: code_challenge_method
    });

    if (!client_id || !redirect_uri) {
        logger.warn('Authorization failed: missing required parameters', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            hasClientId: !!client_id,
            hasRedirectUri: !!redirect_uri
        });
        return c.json({ error: 'Missing required parameters: client_id, redirect_uri' }, 400);
    }

    const client = registeredClients.get(client_id);
    if (!client) {
        logger.warn('Authorization failed: invalid client_id', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            clientId: client_id
        });
        return c.json({ error: 'Invalid client_id' }, 400);
    }

    if (!client.redirect_uris.includes(redirect_uri)) {
        logger.warn('Authorization failed: invalid redirect_uri', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            clientId: client_id,
            redirectUri: redirect_uri,
            allowedUris: client.redirect_uris
        });
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

    // Forward PKCE parameters to Microsoft OAuth
    if (code_challenge) {
        msGraphAuthUrl.searchParams.set('code_challenge', code_challenge);
    }
    if (code_challenge_method) {
        msGraphAuthUrl.searchParams.set('code_challenge_method', code_challenge_method);
    }

    logger.info('Redirecting to Microsoft OAuth', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        clientId: client_id,
        msGraphUrl: msGraphAuthUrl.toString()
    });

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

        logger.info('Token endpoint accessed', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            userAgent: c.req.header('user-agent'),
            grantType: grant_type,
            clientId: client_id,
            hasCode: !!code,
            hasRefreshToken: !!refresh_token,
            hasCodeVerifier: !!code_verifier
        });

        if (grant_type === 'authorization_code') {
            if (!code || !redirect_uri || !client_id) {
                logger.warn('Token exchange failed: missing required parameters for authorization_code', {
                    ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                    hasCode: !!code,
                    hasRedirectUri: !!redirect_uri,
                    hasClientId: !!client_id
                });
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
                        logger.warn('Token exchange failed: unsupported code challenge method', {
                            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                            clientId: client_id,
                            method: client.code_challenge_method
                        });
                        return c.json({ error: 'Unsupported code challenge method' }, 400);
                    }

                    if (expectedChallenge !== client.code_challenge) {
                        logger.warn('Token exchange failed: invalid code verifier', {
                            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                            clientId: client_id
                        });
                        return c.json({ error: 'Invalid code verifier' }, 400);
                    }

                    // Clear PKCE data after successful validation
                    delete client.code_challenge;
                    delete client.code_challenge_method;
                    registeredClients.set(client_id, client);
                }
            }

            const tokenResponse = await exchangeCodeForToken(code, redirect_uri, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            
            logger.info('Token exchange successful', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                clientId: client_id,
                grantType: 'authorization_code',
                hasAccessToken: !!tokenResponse.access_token,
                hasRefreshToken: !!tokenResponse.refresh_token
            });
            
            return c.json(tokenResponse);
        } else if (grant_type === 'refresh_token') {
            if (!refresh_token) {
                logger.warn('Token refresh failed: missing refresh_token', {
                    ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                    clientId: client_id
                });
                return c.json({ error: 'Missing refresh_token' }, 400);
            }

            const tokenResponse = await refreshAccessToken(refresh_token, process.env.CLIENT_ID || '', process.env.CLIENT_SECRET || '');
            
            logger.info('Token refresh successful', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                clientId: client_id,
                grantType: 'refresh_token',
                hasAccessToken: !!tokenResponse.access_token,
                hasRefreshToken: !!tokenResponse.refresh_token
            });
            
            return c.json(tokenResponse);
        } else {
            logger.warn('Token exchange failed: unsupported grant_type', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                grantType: grant_type
            });
            return c.json({ error: 'Unsupported grant_type' }, 400);
        }
    } catch (error) {
        logger.error('Token exchange error:', error);
        return c.json({ error: 'Token exchange failed' }, 500);
    }
});

// User info endpoint
app.get('/userinfo', msGraphBearerTokenAuthMiddleware, async (c) => {
    logger.info('User info endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        hasAuth: !!c.req.header('Authorization')
    });
    
    return c.json({
        sub: 'user-id',
        name: 'User Name',
        email: 'user@example.com'
    });
});

// Logout endpoint
app.post('/logout', (c) => {
    logger.info('Logout endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        method: c.req.method
    });
    
    return c.json({ message: 'Logged out successfully' });
});

// MCP route
app.post('/mcp', async (c) => {
    const authHeader = c.req.header('Authorization');
    
    logger.info('MCP endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        hasAuth: !!authHeader,
        contentType: c.req.header('content-type')
    });
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        logger.warn('MCP request rejected: missing or invalid authorization', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            hasAuth: !!authHeader,
            authType: authHeader ? authHeader.split(' ')[0] : 'none'
        });
        
        // Return OAuth-specific error for MCP clients
        const baseUrl = c.req.url.split('/mcp')[0];
        return c.json({
            jsonrpc: "2.0",
            error: {
                code: -32002,
                message: "Authentication required",
                data: {
                    oauth_url: `${baseUrl}/.well-known/oauth-authorization-server`
                }
            }
        }, 401);
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

        logger.info('MCP request received', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            method: request.method,
            id: request.id,
            hasParams: !!request.params
        });

        // Handle MCP protocol messages properly
        if (request.method === 'initialize') {
            logger.info('MCP initialize request', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                protocolVersion: request.params?.protocolVersion
            });
            
            return c.json({
                jsonrpc: "2.0",
                id: request.id,
                result: {
                    protocolVersion: "2025-06-18",
                    capabilities: {
                        tools: { listChanged: true }
                    },
                    serverInfo: {
                        name: "Microsoft Graph Service",
                        version: "1.0.0"
                    }
                }
            });
        } else if (request.method === 'tools/list') {
            logger.info('MCP tools/list request', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown'
            });
            
            return c.json({
                jsonrpc: "2.0",
                id: request.id,
                result: {
                    tools: [
                        {
                            name: "microsoft-graph-api",
                            description: "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management",
                            inputSchema: {
                                type: "object",
                                properties: {
                                    apiType: { type: "string", enum: ["graph", "azure"] },
                                    path: { type: "string" },
                                    method: { type: "string", enum: ["get", "post", "put", "patch", "delete"] },
                                    apiVersion: { type: "string" },
                                    subscriptionId: { type: "string" },
                                    queryParams: { type: "object" },
                                    body: { type: "object" },
                                    graphApiVersion: { type: "string", enum: ["v1.0", "beta"] },
                                    fetchAll: { type: "boolean" },
                                    consistencyLevel: { type: "string" }
                                },
                                required: ["apiType", "path", "method"]
                            }
                        },
                        {
                            name: "get-auth-status",
                            description: "Check the current authentication status",
                            inputSchema: {
                                type: "object",
                                properties: {}
                            }
                        }
                    ]
                }
            });
        } else if (request.method === 'tools/call') {
            const { name, arguments: args } = request.params;
            
            logger.info('MCP tools/call request', {
                ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                toolName: name,
                hasArgs: !!args
            });

            try {
                let result;
                if (name === 'microsoft-graph-api') {
                    result = await mcp.msGraphServiceInstance.genericGraphRequest(
                        args.path,
                        args.method,
                        args.body,
                        args.queryParams,
                        args.graphApiVersion || 'v1.0',
                        args.fetchAll || false,
                        args.consistencyLevel
                    );
                    
                    logger.info('MCP tool execution successful', {
                        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                        toolName: name,
                        path: args.path,
                        method: args.method
                    });
                    
                } else if (name === 'get-auth-status') {
                    result = { status: 'authenticated', user: 'current-user' };
                    
                    logger.info('MCP auth status check', {
                        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown'
                    });
                    
                } else {
                    logger.warn('MCP unknown tool requested', {
                        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                        toolName: name
                    });
                    
                    throw new Error(`Unknown tool: ${name}`);
                }

                return c.json({
                    jsonrpc: "2.0",
                    id: request.id,
                    result: {
                        content: [{
                            type: "text",
                            text: JSON.stringify(result, null, 2)
                        }]
                    }
                });
            } catch (error) {
                logger.error('MCP tool execution failed', {
                    ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
                    toolName: name,
                    error: error.message
                });
                
                return c.json({
                    jsonrpc: "2.0",
                    id: request.id,
                    error: {
                        code: -32000,
                        message: error.message || 'Tool execution failed'
                    }
                });
            }
        }

        // Default response for unhandled methods
        logger.warn('MCP unknown method requested', {
            ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
            method: request.method
        });
        
        return c.json({
            jsonrpc: "2.0",
            id: request.id,
            error: {
                code: -32601,
                message: `Method not found: ${request.method}`
            }
        });

    } catch (error) {
        logger.error('MCP request error:', error);
        return c.json({
            jsonrpc: "2.0",
            error: {
                code: -32000,
                message: 'Internal server error'
            }
        }, 500);
    }
});

// Health check endpoint
app.get('/health', (c) => {
    logger.info('Health check endpoint accessed', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent')
    });
    
    return c.json({ status: 'ok', service: 'msgraph-mcp' });
});

// 404 handler for unmatched routes
app.notFound((c) => {
    logger.warn('404 - Route not found', {
        ip: c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'unknown',
        userAgent: c.req.header('user-agent'),
        method: c.req.method,
        path: c.req.path,
        query: c.req.query()
    });
    
    return c.json({ error: 'Not Found', message: `Route ${c.req.method} ${c.req.path} not found` }, 404);
});

// Start the server
const port = process.env.PORT || 3001;
logger.info(`🚀 Starting Microsoft Graph MCP Server on port ${port}`);

serve({
    fetch: app.fetch,
    port: Number(port),
});

