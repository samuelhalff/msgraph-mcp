import {McpServer} from '@modelcontextprotocol/sdk/server/mcp.js'
import {z} from 'zod'
import {MSGraphService, AuthManager, AuthMode, AuthConfig} from "./MSGraphService.js";
import {MSGraphAuthContext, Env} from "../types";
import logger from './lib/logger.js';

/**
 * Parameters for msgraph_request method
 */
interface MSGraphRequestParams {
    path: string;
    method: string;
    body?: unknown;
    queryParams?: Record<string, string>;
    apiVersion?: string;
    subscriptionId?: string;
    fetchAll?: boolean;
    apiType?: string;
    graphApiVersion?: string;
    consistencyLevel?: string;
}

/**
 * The `MSGraphMCP` class exposes Microsoft Graph API via the Model Context Protocol
 * for consumption by API Agents
 */
export class MSGraphMCP {
    private authManager: AuthManager | null = null;
    private msGraphService: MSGraphService | null = null;
    private env: Env;
    private authContext: MSGraphAuthContext;

    constructor(env: Env, authContext: MSGraphAuthContext) {
        this.env = env;
        this.authContext = authContext;
    }

    async initialize() {
        // Initialize authentication and Microsoft Graph service
        const authConfig: AuthConfig = {
            mode: this.determineAuthMode(),
            tenantId: this.env.TENANT_ID,
            clientId: this.env.CLIENT_ID,
            clientSecret: this.env.CLIENT_SECRET,
            accessToken: this.env.ACCESS_TOKEN,
            redirectUri: this.env.REDIRECT_URI,
            certificatePath: this.env.CERTIFICATE_PATH,
            certificatePassword: this.env.CERTIFICATE_PASSWORD
        };

        this.authManager = new AuthManager(authConfig);

        if (authConfig.mode === AuthMode.ClientProvidedToken && this.authContext.accessToken) {
            this.authManager.updateAccessToken(this.authContext.accessToken, undefined, this.authContext.refreshToken);
        }

        this.msGraphService = new MSGraphService(this.env, this.authContext, authConfig);
        await this.msGraphService.initialize();
    }

    private determineAuthMode(): AuthMode {
        if (this.env.USE_CLIENT_TOKEN === 'true') {
            return AuthMode.ClientProvidedToken;
        } else if (this.env.USE_INTERACTIVE === 'true') {
            return AuthMode.Interactive;
        } else if (this.env.USE_CERTIFICATE === 'true') {
            return AuthMode.Certificate;
        } else if (this.env.CLIENT_SECRET) {
            return AuthMode.ClientCredentials;
        } else {
            return AuthMode.Interactive; // Default
        }
    }

    get msGraphServiceInstance() {
        if (!this.msGraphService) {
            throw new Error("MSGraphService not initialized");
        }
        return this.msGraphService;
    }

    formatResponse = (description: string, data: unknown): {
        content: Array<{ type: 'text', text: string }>
    } => {
        return {
            content: [{
                type: "text",
                text: `Success! ${description}\n\nResult:\n${JSON.stringify(data, null, 2)}`
            }]
        };
    }

    get server() {
        const server = new McpServer({
            name: 'Microsoft Graph Service',
            version: '1.0.0',
        })

        // Main Microsoft Graph API tool
        server.tool(
            "microsoft-graph-api",
            "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. IMPORTANT: For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: \"eventual\"'.",
            {
                apiType: z.enum(["graph", "azure"]).optional().default("graph").describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management. Defaults to 'graph'."),
                path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
                method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
                apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
                subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
                queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
                body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
                graphApiVersion: z.enum(["v1.0", "beta"]).optional().default("v1.0").describe("Microsoft Graph API version to use (default: v1.0)"),
                fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
                consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
            },
            async ({
                apiType,
                path,
                method,
                queryParams,
                body,
                graphApiVersion,
                fetchAll,
                consistencyLevel
            }) => {
                try {
                    const responseData = await this.msGraphServiceInstance.genericGraphRequest(
                        path,
                        method,
                        body,
                        queryParams,
                        graphApiVersion,
                        fetchAll,
                        consistencyLevel
                    );

                    let resultText = `Result for ${apiType} API (${graphApiVersion}) - ${method} ${path}:\n\n`;
                    resultText += JSON.stringify(responseData, null, 2);

                    if (!fetchAll && method === 'get') {
                        const nextLinkKey = apiType === 'graph' ? '@odata.nextLink' : 'nextLink';
                        if (responseData && (responseData as Record<string, unknown>)[nextLinkKey]) {
                            resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
                        }
                    }

                    return {
                        content: [{ type: "text" as const, text: resultText }],
                    };
                } catch (err: unknown) {
                    logger.error('Error in MCP request handler', { message: err instanceof Error ? err.message : String(err), error: err });
                    return {
                        content: [{ type: "text" as const, text: `Error: ${err instanceof Error ? err.message : String(err)}` }],
                        isError: true
                    };
                }
            }
        );

        // Token management tools
        server.tool(
            "set-access-token",
            "Set or update the access token for Microsoft Graph authentication. Use this when the MCP Client has obtained a fresh token through interactive authentication.",
            {
                accessToken: z.string().describe("The access token obtained from Microsoft Graph authentication"),
                refreshToken: z.string().optional().describe("The refresh token for obtaining new access tokens"),
                expiresOn: z.string().optional().describe("Token expiration time in ISO format (optional, defaults to 1 hour from now)")
            },
            async ({ accessToken, refreshToken, expiresOn }) => {
                try {
                    if (this.authManager?.getAuthMode() === AuthMode.ClientProvidedToken) {
                        const expirationDate = expiresOn ? new Date(expiresOn) : undefined;
                        this.authManager.updateAccessToken(accessToken, expirationDate, refreshToken);

                        // Reinitialize the Graph client with the new token
                        this.msGraphService = new MSGraphService(this.env, {
                            ...this.authContext,
                            accessToken,
                            refreshToken
                        }, {
                            mode: AuthMode.ClientProvidedToken,
                            accessToken: accessToken
                        });
                        await this.msGraphService.initialize();

                        return {
                            content: [{
                                type: "text" as const,
                                text: "Access token updated successfully. You can now make Microsoft Graph requests on behalf of the authenticated user."
                            }],
                        };
                    } else {
                        return {
                            content: [{
                                type: "text" as const,
                                text: "Error: MCP Server is not configured for client-provided token authentication. Set USE_CLIENT_TOKEN=true in environment variables."
                            }],
                            isError: true
                        };
                    }
                } catch (error: unknown) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: `Error setting access token: ${error instanceof Error ? error.message : String(error)}`
                        }],
                        isError: true
                    };
                }
            }
        );

        server.tool(
            "get-auth-status",
            "Check the current authentication status and mode of the MCP Server and also returns the current graph permission scopes of the access token for the current session.",
            {},
            async () => {
                try {
                    const authMode = this.authManager?.getAuthMode() || "Not initialized";
                    const isReady = this.authManager !== null;
                    const tokenStatus = this.authManager ? await this.authManager.getTokenStatus() : { isExpired: false };

                    return {
                        content: [{
                            type: "text" as const,
                            text: JSON.stringify({
                                authMode,
                                isReady,
                                supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
                                tokenStatus: tokenStatus,
                                timestamp: new Date().toISOString()
                            }, null, 2)
                        }],
                    };
                } catch (error: unknown) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: `Error checking auth status: ${error instanceof Error ? error.message : String(error)}`
                        }],
                        isError: true
                    };
                }
            }
        );

        // Specific Microsoft Graph tools for common operations
        server.tool('getCurrentUserProfile', 'Get the current user\'s Microsoft Graph profile', {}, async () => {
            const profile = await this.msGraphServiceInstance.getCurrentUserProfile()
            return this.formatResponse('User profile retrieved', profile)
        })

        server.tool('getUsers', 'Get users from Microsoft Graph', {
            queryParams: z.record(z.string()).optional().describe('Query parameters for filtering users'),
            fetchAll: z.boolean().optional().default(false).describe('Fetch all users')
        }, async ({queryParams, fetchAll}) => {
            const users = await this.msGraphServiceInstance.getUsers(queryParams, fetchAll)
            return this.formatResponse('Users retrieved', users)
        })

        server.tool('getGroups', 'Get groups from Microsoft Graph', {
            queryParams: z.record(z.string()).optional().describe('Query parameters for filtering groups'),
            fetchAll: z.boolean().optional().default(false).describe('Fetch all groups')
        }, async ({queryParams, fetchAll}) => {
            const groups = await this.msGraphServiceInstance.getGroups(queryParams, fetchAll)
            return this.formatResponse('Groups retrieved', groups)
        })

        server.tool('getApplications', 'Get applications from Microsoft Graph', {
            queryParams: z.record(z.string()).optional().describe('Query parameters for filtering applications'),
            fetchAll: z.boolean().optional().default(false).describe('Fetch all applications')
        }, async ({queryParams, fetchAll}) => {
            const apps = await this.msGraphServiceInstance.getApplications(queryParams, fetchAll)
            return this.formatResponse('Applications retrieved', apps)
        })

        // Create an Outlook draft (POST /me/messages)
        server.tool('createDraftEmail', 'Create an email draft in Outlook (saved to Drafts)', {
            subject: z.string().optional().describe('Subject of the email'),
            body: z.string().optional().describe('Plain text body of the email'),
            contentType: z.enum(['Text', 'HTML']).optional().default('Text').describe('The content type of the body'),
            toRecipients: z.array(z.string()).optional().describe('Array of recipient email addresses'),
            ccRecipients: z.array(z.string()).optional().describe('Array of CC recipient email addresses'),
            bccRecipients: z.array(z.string()).optional().describe('Array of BCC recipient email addresses')
        }, async ({ subject, body, contentType, toRecipients, ccRecipients, bccRecipients }) => {
            try {
                const message = {
                    subject: subject || '',
                    body: { contentType: contentType || 'Text', content: body || '' },
                    toRecipients: (toRecipients || []).map((addr: string) => ({ emailAddress: { address: addr } })),
                    ccRecipients: (ccRecipients || []).map((addr: string) => ({ emailAddress: { address: addr } })),
                    bccRecipients: (bccRecipients || []).map((addr: string) => ({ emailAddress: { address: addr } }))
                };

                const response = await this.msGraphServiceInstance.genericGraphRequest(
                    '/me/messages',
                    'post',
                    message,
                    undefined,
                    'v1.0',
                    false,
                    undefined
                );

                return this.formatResponse('Draft created', response);
            } catch (err: unknown) {
                return {
                    content: [{ type: 'text', text: `Error creating draft: ${err instanceof Error ? err.message : String(err)}` }],
                    isError: true
                };
            }
        })

        return server
    }

    // Static methods for MCP server setup
    static serve() {
        return {
            fetch: async (request: Request) => {
                // Extract auth context from request headers
                const authHeader = request.headers.get('Authorization');
                let authContext: MSGraphAuthContext = { accessToken: '' };

                if (authHeader && authHeader.startsWith('Bearer ')) {
                    authContext = {
                        accessToken: authHeader.substring(7),
                        refreshToken: request.headers.get('X-Refresh-Token') || undefined
                    };
                }

                // Create env object from process.env
                const env: Env = {
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
                    USE_CERTIFICATE: process.env.USE_CERTIFICATE,
                };

                const mcp = new MSGraphMCP(env, authContext);
                await mcp.initialize();

                // Handle MCP protocol messages using the SDK
                if (request.method === 'POST') {
                    try {
                        const bodyText = await request.text();
                        if (!bodyText.trim()) {
                            return new Response(JSON.stringify({
                                jsonrpc: "2.0",
                                error: {
                                    code: -32700,
                                    message: "Parse error: empty request body"
                                },
                                id: null
                            }), {
                                status: 400,
                                headers: { 'Content-Type': 'application/json' }
                            });
                        }
                        
                        let body: { method?: string; id?: unknown; params?: unknown };
                        try {
                            body = JSON.parse(bodyText);
                        } catch {
                            return new Response(JSON.stringify({
                                jsonrpc: "2.0",
                                error: {
                                    code: -32700,
                                    message: "Parse error: invalid JSON"
                                },
                                id: null
                            }), {
                                status: 400,
                                headers: { 'Content-Type': 'application/json' }
                            });
                        }
                        
                        logger.info('MCP request received', { method: body.method, id: body.id });
                        logger.debug('Full MCP request body', { body });
                        logger.debug('Request headers', { headers: Object.fromEntries(request.headers.entries()) });
                        
                        // For discovery requests (initialize, tools/list, ping), don't require auth
                        if (body.method === 'initialize' || body.method === 'tools/list' || body.method === 'ping') {
                            logger.info('Processing discovery request', { method: body.method });
                            
                            if (body.method === 'initialize') {
                                // Return server capabilities
                                return new Response(JSON.stringify({
                                    jsonrpc: "2.0",
                                    id: body.id,
                                    result: {
                                        protocolVersion: "2024-11-05",
                                        capabilities: {
                                            tools: {}
                                        },
                                        serverInfo: {
                                            name: "Microsoft Graph Service",
                                            version: "1.0.0"
                                        }
                                    }
                                }), {
                                    headers: { 'Content-Type': 'application/json' }
                                });
                            } else if (body.method === 'tools/list') {
                                // Return list of available tools
                                const tools = [
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
                                        },
                                        examples: [
                                            {
                                                description: 'Get current user profile',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/me',
                                                    method: 'get',
                                                    graphApiVersion: 'v1.0'
                                                }
                                            },
                                            {
                                                description: 'List inbox messages (Graph)',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/me/messages',
                                                    method: 'get',
                                                    graphApiVersion: 'v1.0',
                                                    fetchAll: true
                                                }
                                            },
                                            {
                                                description: 'Send an email (Graph)',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/me/sendMail',
                                                    method: 'post',
                                                    body: {
                                                        message: {
                                                            subject: 'Hello',
                                                            body: { contentType: 'Text', content: 'Hi there' },
                                                            toRecipients: [{ emailAddress: { address: 'recipient@example.com' } }]
                                                        },
                                                        saveToSentItems: true
                                                    }
                                                }
                                            },
                                            {
                                                description: 'List users (directory)',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/users',
                                                    method: 'get',
                                                    graphApiVersion: 'v1.0',
                                                    fetchAll: true
                                                }
                                            },
                                            {
                                                description: 'List groups',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/groups',
                                                    method: 'get',
                                                    graphApiVersion: 'v1.0',
                                                    fetchAll: true
                                                }
                                            },
                                            {
                                                description: 'Create calendar event',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/me/events',
                                                    method: 'post',
                                                    graphApiVersion: 'v1.0',
                                                    body: {
                                                        subject: 'Team sync',
                                                        body: { contentType: 'Text', content: 'Weekly sync' },
                                                        start: { dateTime: '2025-09-01T10:00:00', timeZone: 'UTC' },
                                                        end: { dateTime: '2025-09-01T11:00:00', timeZone: 'UTC' },
                                                        attendees: [{ emailAddress: { address: 'alice@example.com' }, type: 'required' }]
                                                    }
                                                }
                                            },
                                            {
                                                description: 'Upload file to OneDrive (replace path)',
                                                value: {
                                                    apiType: 'graph',
                                                    path: "/me/drive/root:/Documents/example.txt:/content",
                                                    method: 'put',
                                                    graphApiVersion: 'v1.0',
                                                    body: { /* raw file bytes expected by the client when using HTTP */ }
                                                }
                                            },
                                            {
                                                description: 'Get group members',
                                                value: {
                                                    apiType: 'graph',
                                                    path: '/groups/{group-id}/members',
                                                    method: 'get',
                                                    graphApiVersion: 'v1.0',
                                                    fetchAll: true
                                                }
                                            },
                                            {
                                                description: 'Azure: list subscriptions (ARM)',
                                                value: {
                                                    apiType: 'azure',
                                                    path: '/subscriptions',
                                                    method: 'get',
                                                    apiVersion: '2020-01-01',
                                                    fetchAll: true
                                                }
                                            }
                                        ]
                                    },
                                    {
                                        name: "set-access-token",
                                        description: "Set or update the access token for Microsoft Graph authentication",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                accessToken: { type: "string" },
                                                refreshToken: { type: "string" },
                                                expiresOn: { type: "string" }
                                            },
                                            required: ["accessToken"]
                                        }
                                    },
                                    {
                                        name: "get-auth-status",
                                        description: "Check the current authentication status and mode of the MCP Server",
                                        inputSchema: {
                                            type: "object",
                                            properties: {}
                                        }
                                    },
                                    {
                                        name: "getCurrentUserProfile",
                                        description: "Get the current user's Microsoft Graph profile",
                                        inputSchema: {
                                            type: "object",
                                            properties: {}
                                        }
                                    },
                                    {
                                        name: "getUsers",
                                        description: "Get users from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    },
                                    {
                                        name: "getGroups",
                                        description: "Get groups from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    },
                                    {
                                        name: "getApplications",
                                        description: "Get applications from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    }
                                ];
                                
                                return new Response(JSON.stringify({
                                    jsonrpc: "2.0",
                                    id: body.id,
                                    result: { tools }
                                }), {
                                    headers: { 'Content-Type': 'application/json' }
                                });
                            } else if (body.method === 'ping') {
                                // Return empty result for ping (no content key)
                                return new Response(JSON.stringify({
                                    jsonrpc: '2.0',
                                    id: body.id,
                                    result: {}
                                }), { headers: { 'Content-Type': 'application/json' } });
                            }
                        } else if (body.method === 'notifications/initialized') {
                            logger.info('Received notifications/initialized from client', { id: body.id, params: body.params });
                            return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: {} }), { headers: { 'Content-Type': 'application/json' } });
                        }

                        // Map streamable-http "tools/call" to our tool execution flow.
                        if (body.method === 'tools/call') {
                            const callParams = body.params as Record<string, unknown> || {};
                            logger.info('Received tools/call', { callParams });

                            // Flexible extraction of tool name and input payload from different clients
                            const toolName = (callParams.tool as string) || (callParams.toolName as string) || (callParams.name as string) || (callParams.id as string) || (Array.isArray(callParams) && callParams[0] && (callParams[0] as Record<string, unknown>).tool as string) || undefined;
                            const input = (callParams.input as Record<string, unknown>) || (callParams.args as Record<string, unknown>) || (callParams.params as Record<string, unknown>) || (callParams.arguments as Record<string, unknown>) || (Array.isArray(callParams) ? callParams[0] as Record<string, unknown> : {}) || {};

                            if (!toolName) {
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32602, message: 'Invalid params: tool name missing in tools/call' } }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                            }

                            // Only allow authenticated tool calls (unless tool is discovery-type)
                            if (!authHeader || !authHeader.startsWith('Bearer ')) {
                                const baseUrl = (process.env.PUBLIC_BASE_URL || `http://localhost:3001`).replace(/\/$/, '');
                                const oauthDiscoveryUrl = `${baseUrl}/.well-known/oauth-authorization-server`;
                                return new Response(JSON.stringify({
                                    jsonrpc: '2.0',
                                    id: body.id,
                                    error: {
                                        code: -32002,
                                        message: `OAuth authentication required. Please check the server logs for the authentication URL. OAuth discovery endpoint: ${oauthDiscoveryUrl}`
                                    }
                                }), { status: 401, headers: { 'Content-Type': 'application/json' } });
                            }

                            try {
                                // Route common tool names to existing handlers
                                if (toolName === 'microsoft-graph-api') {
                                    const params = input || {};
                                    // default apiType
                                    if (!params.apiType) params.apiType = 'graph';

                                    // Validate minimal params (reuse same checks we have later)
                                    const paramsTyped = params as unknown as MSGraphRequestParams;
                                    const { path, method } = paramsTyped;
                                    if (!path || typeof path !== 'string') {
                                        return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32602, message: 'Invalid params: path is required and must be a string' } }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                    }

                                    if (!method || typeof method !== 'string') {
                                        return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32602, message: 'Invalid params: method is required' } }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                    }

                                    // Delegate to MSGraphService
                                    if (paramsTyped.apiType === 'azure') {
                                        const responseData = await mcp.msGraphServiceInstance.azureRequest(
                                            paramsTyped.path,
                                            paramsTyped.method as 'get' | 'post' | 'put' | 'patch' | 'delete',
                                            paramsTyped.body,
                                            paramsTyped.queryParams,
                                            paramsTyped.apiVersion,
                                            paramsTyped.subscriptionId,
                                            paramsTyped.fetchAll
                                        );

                                        return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                                    }

                                const responseData = await mcp.msGraphServiceInstance.genericGraphRequest(
                                    paramsTyped.path,
                                    paramsTyped.method as 'get' | 'post' | 'put' | 'patch' | 'delete',
                                    paramsTyped.body,
                                    paramsTyped.queryParams,
                                    (paramsTyped.graphApiVersion as 'v1.0' | 'beta') || 'v1.0',
                                    paramsTyped.fetchAll,
                                    paramsTyped.consistencyLevel
                                );                                    return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                                }

                                // Support direct mapped MCP tools
                                if (toolName === 'getCurrentUserProfile') {
                                    const profile = await mcp.msGraphServiceInstance.getCurrentUserProfile();
                                    return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(profile, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                                }

                                if (toolName === 'getUsers') {
                                    const users = await mcp.msGraphServiceInstance.getUsers(input.queryParams as Record<string, string>, input.fetchAll as boolean);
                                    return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(users, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                                }

                                // Unknown tool
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32601, message: `Method not found (tool): ${toolName}` } }), { status: 404, headers: { 'Content-Type': 'application/json' } });

                            } catch (err: unknown) {
                                logger.error('Error handling tools/call', { message: err instanceof Error ? err.message : String(err), error: err });
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32000, message: err instanceof Error ? err.message : String(err) } }), { status: 500, headers: { 'Content-Type': 'application/json' } });
                            }
                        }

                        // For tool calls, require authentication
                        if (!authHeader || !authHeader.startsWith("Bearer ")) {
                            const baseUrl = (process.env.PUBLIC_BASE_URL || `http://localhost:3001`).replace(/\/$/, '');
                            const oauthDiscoveryUrl = `${baseUrl}/.well-known/oauth-authorization-server`;
                            return new Response(JSON.stringify({
                                jsonrpc: "2.0",
                                error: {
                                    code: -32002,
                                    message: `OAuth authentication required. Please check the server logs for the authentication URL. OAuth discovery endpoint: ${oauthDiscoveryUrl}`
                                },
                                id: body.id
                            }), {
                                status: 401,
                                headers: { 'Content-Type': 'application/json' }
                            });
                        }
                        // Handle tool calls using the MCP server (direct dispatch for known tools)


                        try {
                            // Directly handle common tool calls by invoking the MSGraphService methods
                            const params = (body.params as Record<string, unknown>) || {};
                            logger.debug('Tool call parameters', { params });

                            // Expect the client/model to provide structured params per the tool schema (like Lokka).
                            // Default apiType to 'graph' if omitted for backward compatibility.
                            if (!params.apiType) (params as Record<string, unknown>).apiType = 'graph';
                            logger.debug('Final apiType after defaulting', { apiType: params.apiType });

                            if (body.method === 'microsoft-graph-api') {
                                logger.info('Executing microsoft-graph-api tool call');
                                const {
                                    apiType,
                                    path,
                                    method,
                                    apiVersion,
                                    subscriptionId,
                                    queryParams,
                                    body: requestBody,
                                    graphApiVersion,
                                    fetchAll,
                                    consistencyLevel
                                } = params as Record<string, unknown>;
                                
                                logger.debug('Extracted parameters', {
                                    apiType,
                                    path,
                                    method,
                                    apiVersion,
                                    subscriptionId,
                                    queryParams,
                                    graphApiVersion,
                                    fetchAll,
                                    consistencyLevel,
                                    hasBody: !!requestBody
                                });

                                // Validate required parameters
                                if (!path || typeof path !== 'string') {
                                    return new Response(JSON.stringify({
                                        jsonrpc: '2.0',
                                        id: body.id,
                                        error: { code: -32602, message: 'Invalid params: path is required and must be a string' }
                                    }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                }

                                if (!method || typeof method !== 'string' || !['get', 'post', 'put', 'patch', 'delete'].includes(method.toLowerCase())) {
                                    return new Response(JSON.stringify({
                                        jsonrpc: '2.0',
                                        id: body.id,
                                        error: { code: -32602, message: 'Invalid params: method is required and must be one of: get, post, put, patch, delete' }
                                    }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                }

                                if (apiType === 'azure') {
                                    logger.info('Making Azure API call', {
                                        pathPreview: path ? path.substring(0, 100) + (path.length > 100 ? '...' : '') : 'undefined',
                                        method,
                                        apiVersion,
                                        subscriptionId,
                                        hasQueryParams: !!queryParams,
                                        fetchAll
                                    });
                                    
                                    if (!apiVersion || typeof apiVersion !== 'string') {
                                        return new Response(JSON.stringify({
                                            jsonrpc: '2.0',
                                            id: body.id,
                                            error: { code: -32602, message: 'Invalid params: apiVersion is required for Azure API calls' }
                                        }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                    }

                                    const responseData = await mcp.msGraphServiceInstance.azureRequest(
                                        path as string,
                                        method as 'get' | 'post' | 'put' | 'patch' | 'delete',
                                        requestBody as unknown,
                                        queryParams as Record<string, string>,
                                        apiVersion as string,
                                        subscriptionId as string,
                                        fetchAll as boolean
                                    );

                                    return new Response(JSON.stringify({
                                        jsonrpc: '2.0',
                                        id: body.id,
                                        result: { content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }] }
                                    }), { headers: { 'Content-Type': 'application/json' } });
                                }

                                logger.info('Making Graph API call', {
                                    pathPreview: path ? path.substring(0, 100) + (path.length > 100 ? '...' : '') : 'undefined',
                                    method,
                                    graphApiVersion: graphApiVersion || 'v1.0',
                                    hasQueryParams: !!queryParams,
                                    fetchAll,
                                    consistencyLevel
                                });

                                const responseData = await mcp.msGraphServiceInstance.genericGraphRequest(
                                    path as string,
                                    method as 'get' | 'post' | 'put' | 'patch' | 'delete',
                                    requestBody as unknown,
                                    queryParams as Record<string, string>,
                                    (graphApiVersion as 'v1.0' | 'beta') || 'v1.0',
                                    fetchAll as boolean,
                                    consistencyLevel as string
                                );

                                return new Response(JSON.stringify({
                                    jsonrpc: '2.0',
                                    id: body.id,
                                    result: { content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }] }
                                }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'getCurrentUserProfile') {
                                logger.info('Executing getCurrentUserProfile tool call');
                                const profile = await mcp.msGraphServiceInstance.getCurrentUserProfile();
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(profile, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'getUsers') {
                                logger.info('Executing getUsers tool call', { queryParams: params.queryParams, fetchAll: params.fetchAll });
                                const { queryParams, fetchAll } = params as Record<string, unknown>;
                                const users = await mcp.msGraphServiceInstance.getUsers(queryParams as Record<string, string>, fetchAll as boolean);
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(users, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'getGroups') {
                                const { queryParams, fetchAll } = params as Record<string, unknown>;
                                const groups = await mcp.msGraphServiceInstance.getGroups(queryParams as Record<string, string>, fetchAll as boolean);
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(groups, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'getApplications') {
                                const { queryParams, fetchAll } = params as Record<string, unknown>;
                                const apps = await mcp.msGraphServiceInstance.getApplications(queryParams as Record<string, string>, fetchAll as boolean);
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(apps, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'get-auth-status') {
                                // Access private authManager via any to avoid type visibility issues
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                const authMode = (mcp as any).authManager?.getAuthMode ? (mcp as any).authManager.getAuthMode() : 'Not initialized';
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                const tokenStatus = (mcp as any).authManager ? await (mcp as any).authManager.getTokenStatus() : { isExpired: false };

                                // If no valid token, provide OAuth URL for authentication
                                let oauthUrl = null;
                                if (!tokenStatus || tokenStatus.isExpired || authMode === 'Not initialized') {
                                    const baseUrl = (process.env.PUBLIC_BASE_URL || `http://localhost:3001`).replace(/\/$/, '');
                                    oauthUrl = `${baseUrl}/.well-known/oauth-authorization-server`;
                                }

                                const status = {
                                    authMode,
                                    tokenStatus,
                                    oauthUrl,
                                    timestamp: new Date().toISOString(),
                                    message: oauthUrl ? 'OAuth authentication required. Use the oauthUrl to discover authorization endpoints.' : 'Authentication available'
                                };
                                return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: JSON.stringify(status, null, 2) }] } }), { headers: { 'Content-Type': 'application/json' } });
                            }

                            if (body.method === 'set-access-token') {
                                const { accessToken, refreshToken, expiresOn } = params as Record<string, unknown>;
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                if ((mcp as any).authManager?.getAuthMode && (mcp as any).authManager.getAuthMode() === AuthMode.ClientProvidedToken) {
                                    const expirationDate = expiresOn ? new Date(expiresOn as string | number | Date) : undefined;
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    (mcp as any).authManager.updateAccessToken(accessToken, expirationDate, refreshToken);
                                    // Reinitialize service with new token
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    (mcp as any).authContext = { ...(mcp as any).authContext, accessToken, refreshToken };
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    await (mcp as any).initialize();
                                    return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, result: { content: [{ type: 'text', text: 'Access token updated successfully.' }] } }), { headers: { 'Content-Type': 'application/json' } });
                                } else {
                                    return new Response(JSON.stringify({ jsonrpc: '2.0', id: body.id, error: { code: -32000, message: 'Server not configured for client-provided token auth' } }), { status: 400, headers: { 'Content-Type': 'application/json' } });
                                }
                            }

                            // Unknown tool
                            return new Response(JSON.stringify({
                                jsonrpc: '2.0',
                                id: body.id,
                                error: { code: -32601, message: `Method not found: ${body.method}` }
                            }), { status: 404, headers: { 'Content-Type': 'application/json' } });
                        } catch (err: unknown) {
                            logger.error('MCP tool execution error', {
                                method: body.method,
                                params: body.params,
                                error: err instanceof Error ? err.message : String(err),
                                stack: err instanceof Error ? err.stack : 'No stack trace',
                                errorType: typeof err,
                                errorName: err instanceof Error ? err.name : 'Unknown'
                            });
                            return new Response(JSON.stringify({
                                jsonrpc: '2.0',
                                id: body.id || null,
                                error: { code: -32000, message: err instanceof Error ? err.message : String(err) }
                            }), { status: 500, headers: { 'Content-Type': 'application/json' } });
                        }
                        
                    } catch (error: unknown) {
                        logger.error('MCP request processing error', {
                            error: error instanceof Error ? error.message : String(error),
                            stack: error instanceof Error ? error.stack : 'No stack trace',
                            errorType: typeof error,
                            errorName: error instanceof Error ? error.name : 'Unknown'
                        });
                        return new Response(JSON.stringify({
                            jsonrpc: "2.0",
                            error: {
                                code: -32603,
                                message: `Internal error: ${error instanceof Error ? error.message : 'Unknown error'}`
                            },
                            id: null
                        }), {
                            status: 500,
                            headers: { 'Content-Type': 'application/json' }
                        });
                    }
                }

                return new Response('MCP Server Running', { status: 200 });
            }
        };
    }

    static serveSSE() {
        return {
            fetch: async (request: Request) => {
                // Extract auth context from request headers
                const authHeader = request.headers.get('Authorization');
                let authContext: MSGraphAuthContext = { accessToken: '' };

                if (authHeader && authHeader.startsWith('Bearer ')) {
                    authContext = {
                        accessToken: authHeader.substring(7),
                        refreshToken: request.headers.get('X-Refresh-Token') || undefined
                    };
                }

                // Create env object from process.env
                const env: Env = {
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
                    USE_CERTIFICATE: process.env.USE_CERTIFICATE,
                };

                const mcp = new MSGraphMCP(env, authContext);
                await mcp.initialize();

                // Handle SSE connection
                if (request.method === 'GET') {
                    // Set up SSE headers
                    const headers = new Headers({
                        'Content-Type': 'text/event-stream',
                        'Cache-Control': 'no-cache',
                        'Connection': 'keep-alive',
                        'Access-Control-Allow-Origin': '*',
                        'Access-Control-Allow-Headers': 'Cache-Control',
                    });

                    const stream = new ReadableStream({
                        start(controller) {
                            // Send initial connection message
                            const initialMessage = {
                                jsonrpc: "2.0",
                                method: "connection/ready",
                                params: {
                                    protocolVersion: "2024-11-05",
                                    capabilities: {
                                        tools: {}
                                    },
                                    serverInfo: {
                                        name: "Microsoft Graph Service",
                                        version: "1.0.0"
                                    }
                                }
                            };

                            controller.enqueue(`data: ${JSON.stringify(initialMessage)}\n\n`);

                            // Handle incoming messages (this would need proper MCP protocol handling)
                            // For now, keep the connection alive
                            const keepAlive = setInterval(() => {
                                controller.enqueue(': keepalive\n\n');
                            }, 30000);

                            // Clean up on close
                            request.signal.addEventListener('abort', () => {
                                clearInterval(keepAlive);
                                controller.close();
                            });
                        }
                    });

                    return new Response(stream, { headers });
                }

                // Handle POST requests for tool calls over SSE
                if (request.method === 'POST') {
                    try {
                        const body = await request.json() as { method?: string; id?: unknown; params?: Record<string, unknown> };
                        logger.info('SSE MCP request received', { method: body.method, id: body.id });
                        logger.debug('SSE Full MCP request body', { body });

                        // Allow discovery requests without authentication
                        if (body.method === 'initialize' || body.method === 'tools/list' || body.method === 'ping') {
                            logger.info('Processing SSE discovery request', { method: body.method });
                            
                            if (body.method === 'initialize') {
                                return new Response(JSON.stringify({
                                    jsonrpc: "2.0",
                                    id: body.id,
                                    result: {
                                        protocolVersion: "2024-11-05",
                                        capabilities: {
                                            tools: {}
                                        },
                                        serverInfo: {
                                            name: "Microsoft Graph Service",
                                            version: "1.0.0"
                                        }
                                    }
                                }), {
                                    headers: { 'Content-Type': 'application/json' }
                                });
                            } else if (body.method === 'tools/list') {
                                // Return list of available tools
                                const tools = [
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
                                        name: "set-access-token",
                                        description: "Set or update the access token for Microsoft Graph authentication",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                accessToken: { type: "string" },
                                                refreshToken: { type: "string" },
                                                expiresOn: { type: "string" }
                                            },
                                            required: ["accessToken"]
                                        }
                                    },
                                    {
                                        name: "get-auth-status",
                                        description: "Check the current authentication status and mode of the MCP Server",
                                        inputSchema: {
                                            type: "object",
                                            properties: {}
                                        }
                                    },
                                    {
                                        name: "getCurrentUserProfile",
                                        description: "Get the current user's Microsoft Graph profile",
                                        inputSchema: {
                                            type: "object",
                                            properties: {}
                                        }
                                    },
                                    {
                                        name: "getUsers",
                                        description: "Get users from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    },
                                    {
                                        name: "getGroups",
                                        description: "Get groups from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    },
                                    {
                                        name: "getApplications",
                                        description: "Get applications from Microsoft Graph",
                                        inputSchema: {
                                            type: "object",
                                            properties: {
                                                queryParams: { type: "object" },
                                                fetchAll: { type: "boolean" }
                                            }
                                        }
                                    }
                                ];
                                
                                return new Response(JSON.stringify({
                                    jsonrpc: "2.0",
                                    id: body.id,
                                    result: { tools }
                                }), {
                                    headers: { 'Content-Type': 'application/json' }
                                });
                            } else if (body.method === 'ping') {
                                return new Response(JSON.stringify({
                                    jsonrpc: '2.0',
                                    id: body.id,
                                    result: {},
                                }), { headers: { 'Content-Type': 'application/json' } });
                            }
                        }

                        // For tool calls, require authentication
                        if (!authHeader || !authHeader.startsWith('Bearer ')) {
                            const baseUrl = (process.env.PUBLIC_BASE_URL || `http://localhost:3001`).replace(/\/$/, '');
                            const oauthDiscoveryUrl = `${baseUrl}/.well-known/oauth-authorization-server`;
                            return new Response(JSON.stringify({
                                jsonrpc: "2.0",
                                error: {
                                    code: -32002,
                                    message: `OAuth authentication required. Please check the server logs for the authentication URL. OAuth discovery endpoint: ${oauthDiscoveryUrl}`
                                },
                                id: body.id
                            }), {
                                status: 401,
                                headers: { 'Content-Type': 'application/json' }
                            });
                        }





                        // Handle tool calls using the MCP server
                        // This would need proper MCP protocol handling for tool calls
                        // For now, return a placeholder
                        return new Response(JSON.stringify({
                            jsonrpc: "2.0",
                            id: body.id,
                            result: { content: [{ type: "text", text: "Tool execution over SSE not yet implemented" }] }
                        }), {
                            headers: { 'Content-Type': 'application/json' }
                        });

                    } catch (error: unknown) {
                        logger.error('SSE MCP request error', {
                            error: error instanceof Error ? error.message : String(error),
                            stack: error instanceof Error ? error.stack : 'No stack trace',
                            errorType: typeof error,
                            errorName: error instanceof Error ? error.name : 'Unknown'
                        });
                        return new Response(JSON.stringify({
                            jsonrpc: "2.0",
                            error: {
                                code: -32603,
                                message: "Internal error"
                            },
                            id: null
                        }), {
                            status: 500,
                            headers: { 'Content-Type': 'application/json' }
                        });
                    }
                }

                return new Response('SSE Endpoint', { status: 200 });
            }
        };
    }
} 