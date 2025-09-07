import express from "express";
import cors from "cors";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { ListToolsRequestSchema, CallToolRequestSchema, ListResourcesRequestSchema, ReadResourceRequestSchema, ErrorCode, McpError, } from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";
import { setupOAuthRoutes } from "./auth/auth.js";
import { GraphTools } from "./tools/graphTools.js";
import { TokenManager } from "./auth/tokenManager.js";
import { GraphService } from "./services/graphService.js";
import { logger } from "./utils/logger.js";
dotenv.config();
const app = express();
const log = logger("main");
async function getAccessTokenForSession(sessionId) {
    const tokenData = await tokenManager.getToken(sessionId);
    if (!tokenData) {
        throw new McpError(ErrorCode.InvalidRequest, "User not authenticated");
    }
    if (tokenManager.isTokenExpired(tokenData)) {
        if (!tokenData.refreshToken) {
            throw new McpError(ErrorCode.InvalidRequest, "Please re-authenticate");
        }
        await tokenManager.refreshToken(sessionId, tokenData.refreshToken);
    }
    // fetch latest data after potential refresh
    const latest = await tokenManager.getToken(sessionId);
    if (!latest) {
        throw new McpError(ErrorCode.InvalidRequest, "Re-authentication required");
    }
    return latest.accessToken;
}
// Middleware
app.use(express.json());
app.use(cors({
    origin: process.env.BASE_URL || "*",
    methods: ["GET", "POST", "OPTIONS"],
    exposedHeaders: ["Mcp-Session-Id", "WWW-Authenticate"],
    allowedHeaders: ["Content-Type", "Authorization", "Mcp-Session-Id"],
}));
// Health & discovery endpoints
app.get("/health", (_, res) => res.json({ status: "ok", timestamp: new Date().toISOString() }));
app.get("/.well-known/oauth-protected-resource", (_, res) => res.json({
    resource: `https://pbm-ai.ddns.net:${process.env.PORT}`,
    authorization_servers: [
        `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0`,
    ],
    scopes_supported: [
        "User.Read",
        "Mail.Read",
        "Calendars.Read",
        "Files.Read",
    ],
    bearer_methods_supported: ["header"],
}));
app.get("/mcp/.well-known", (_, res) => res.json({
    name: "msgraph-mcp-server",
    version: "1.0.0",
    description: "Microsoft Graph MCP Server with OAuth2 support",
    transports: ["streamable-http"],
}));
// OAuth routes
const tokenManager = new TokenManager();
setupOAuthRoutes(app, tokenManager);
// MCP SDK server and tools
const server = new Server({ name: "msgraph-mcp-server", version: "1.0.0" }, { capabilities: { tools: {}, resources: {} } });
const graphTools = new GraphTools();
// Utility: extract session ID header or throw
function getSessionId(req) {
    const sid = req.header("Mcp-Session-Id");
    if (!sid) {
        throw new McpError(ErrorCode.InvalidRequest, "Missing Mcp-Session-Id header");
    }
    return sid;
}
// Main MCP endpoint – one long-lived HTTP POST per session
app.post("/mcp", (req, res) => {
    // Do NOT call res.json or res.end – keep the response open
    const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => getSessionId(req),
    });
    server.connect(transport).catch((err) => {
        log.error("MCP connect error:", err);
        // Do not close res here; SDK will handle errors in-stream
    });
});
// Register MCP handlers
server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: graphTools.getToolDefinitions(),
}));
server.setRequestHandler(CallToolRequestSchema, async (request, extra) => {
    const req = extra.req;
    const sessionId = getSessionId(req);
    const accessToken = await getAccessTokenForSession(sessionId);
    const service = new GraphService(accessToken);
    try {
        const result = await graphTools.executeTool(request.params.name, request.params.arguments || {}, service);
        return {
            content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
    }
    catch (err) {
        log.error("Tool error:", err);
        throw new McpError(ErrorCode.InternalError, err.message);
    }
});
server.setRequestHandler(ListResourcesRequestSchema, async () => ({
    resources: [
        {
            uri: "graph://user/profile",
            name: "User Profile",
            description: "Current user profile information",
            mimeType: "application/json",
        },
        {
            uri: "graph://user/mail",
            name: "User Mail",
            description: "User email messages",
            mimeType: "application/json",
        },
    ],
}));
server.setRequestHandler(ReadResourceRequestSchema, async (request, extra) => {
    const req = extra.req;
    const sessionId = getSessionId(req);
    const accessToken = await getAccessTokenForSession(sessionId);
    const service = new GraphService(accessToken);
    let data;
    switch (request.params.uri) {
        case "graph://user/profile":
            data = await service.getUserProfile();
            break;
        case "graph://user/mail":
            data = await service.getMessages();
            break;
        default:
            throw new McpError(ErrorCode.InvalidRequest, `Unknown resource ${request.params.uri}`);
    }
    return {
        contents: [
            {
                uri: request.params.uri,
                mimeType: "application/json",
                text: JSON.stringify(data, null, 2),
            },
        ],
    };
});
// Start server
const port = Number(process.env.PORT) || 3000;
app.listen(port, () => {
    log.info(`MCP server listening on port ${port}`);
});
