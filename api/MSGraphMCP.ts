import {McpServer} from '@modelcontextprotocol/sdk/server/mcp.js'
import {StdioServerTransport} from '@modelcontextprotocol/sdk/server/stdio.js'
import {z} from 'zod'
import {MSGraphService, AuthManager, AuthMode, AuthConfig} from "./MSGraphService.js";
import {MSGraphAuthContext, Env} from "../types";

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
            tenantId: (this.env as any).TENANT_ID,
            clientId: (this.env as any).CLIENT_ID,
            clientSecret: (this.env as any).CLIENT_SECRET,
            accessToken: (this.env as any).ACCESS_TOKEN,
            redirectUri: (this.env as any).REDIRECT_URI,
            certificatePath: (this.env as any).CERTIFICATE_PATH,
            certificatePassword: (this.env as any).CERTIFICATE_PASSWORD
        };

        this.authManager = new AuthManager(authConfig);

        if (authConfig.mode === AuthMode.ClientProvidedToken && this.authContext.accessToken) {
            this.authManager.updateAccessToken(this.authContext.accessToken, undefined, this.authContext.refreshToken);
        }

        this.msGraphService = new MSGraphService(this.env, this.authContext, authConfig);
        await this.msGraphService.initialize();
    }

    private determineAuthMode(): AuthMode {
        if ((this.env as any).USE_CLIENT_TOKEN === 'true') {
            return AuthMode.ClientProvidedToken;
        } else if ((this.env as any).USE_INTERACTIVE === 'true') {
            return AuthMode.Interactive;
        } else if ((this.env as any).USE_CERTIFICATE === 'true') {
            return AuthMode.Certificate;
        } else if ((this.env as any).CLIENT_SECRET) {
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
                apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
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
                apiVersion,
                subscriptionId,
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
                        if (responseData && (responseData as any)[nextLinkKey]) {
                            resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
                        }
                    }

                    return {
                        content: [{ type: "text" as const, text: resultText }],
                    };
                } catch (error: any) {
                    return {
                        content: [{
                            type: "text",
                            text: JSON.stringify({
                                error: error instanceof Error ? error.message : String(error),
                                statusCode: error.statusCode || 'N/A',
                                errorBody: error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A',
                                attemptedPath: path
                            }),
                        }],
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
                        const authProvider = this.authManager.getGraphAuthProvider();
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
                } catch (error: any) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: `Error setting access token: ${error.message}`
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
                } catch (error: any) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: `Error checking auth status: ${error.message}`
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

        return server
    }

    // Static methods for MCP server setup
    static serve(path: string, options?: any) {
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

                // Create a simple HTTP handler for the MCP server
                if (request.method === 'POST') {
                    // Handle MCP protocol messages
                    const body: any = await request.json();
                    // This is a simplified implementation - in practice you'd need proper MCP protocol handling
                    return new Response(JSON.stringify({ 
                        jsonrpc: "2.0", 
                        id: body.id,
                        result: { tools: [] } // Simplified - would need proper tool listing
                    }), {
                        headers: { 'Content-Type': 'application/json' }
                    });
                }

                return new Response('MCP Server Running', { status: 200 });
            }
        };
    }

    static serveSSE(path: string, options?: any) {
        return {
            fetch: async (request: Request, env: Env) => {
                // Similar to serve but for SSE
                return new Response('SSE Not Implemented', { status: 501 });
            }
        };
    }
} 