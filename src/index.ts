import express from "express";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  ErrorCode,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";
import { v4 as uuidv4 } from "uuid";
import { setupOAuthRoutes } from "./auth/auth.ts";
import { GraphTools } from "./tools/graphTools.ts";
import { TokenManager } from "./auth/tokenManager.ts";
import { GraphService } from "./services/graphService.ts";
import { logger } from "./utils/logger.ts";

dotenv.config();

const app = express();
const log = logger("main");

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Basic request logger (safe headers only)
app.use((req, res, next) => {
  const requestId = uuidv4();
  (req as any).requestId = requestId;
  res.setHeader("X-Request-Id", requestId);
  const safeHeaders = {
    "mcp-session-id": req.headers["mcp-session-id"],
    "content-type": req.headers["content-type"],
    accept: req.headers["accept"],
    "user-agent": req.headers["user-agent"],
  };
  log.info("REQ", {
    id: requestId,
    method: req.method,
    url: req.originalUrl,
    headers: safeHeaders,
    query: req.query,
  });

  res.on("finish", () => {
    log.info("RES", {
      id: requestId,
      statusCode: res.statusCode,
    });
  });
  next();
});

// CORS middleware
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  // Only allow Content-Type and mcp-session-id headers
  res.header("Access-Control-Allow-Headers", "Content-Type, mcp-session-id");
  if (req.method === "OPTIONS") return res.status(200).end();
  next();
});

// Health check endpoint
app.get("/health", (req, res) => {
  log.debug("Health check ping");
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// Well-known endpoint for MCP server discovery
app.get("/mcp/.well-known", (req, res) => {
  log.debug("Serving MCP discovery metadata");
  res.json({
    name: "msgraph-mcp-server",
    version: "1.0.0",
    description: "Microsoft Graph MCP Server with OAuth2 support",
    capabilities: {
      tools: {},
      resources: {},
    },
    oauth: {
      authorization_url: "/oauth/authorize",
      callback_url: "/oauth/callback",
      token_url: "/oauth/token",
    },
    transports: ["streamable-http"],
  });
});

// Initialize token manager and graph tools
const tokenManager = new TokenManager();
const graphTools = new GraphTools();

// Create MCP server instance
const server = new Server(
  {
    name: "msgraph-mcp-server",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
      resources: {},
    },
  }
);

// Helper function to get session context from headers (mcp-session-id only)
function getSessionContext(req: express.Request) {
  const sessionId = (req.headers["mcp-session-id"] as string) || "";
  if (!sessionId) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      "Missing mcp-session-id header"
    );
  }
  log.debug("Session resolved", {
    sessionId,
    requestId: (req as any).requestId,
  });
  return { sessionId };
}

// Helper function to get access token for a session
async function getAccessTokenForSession(sessionId: string): Promise<string> {
  const tokenData = await tokenManager.getToken(sessionId);
  if (!tokenData) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      "User not authenticated. Please authenticate first."
    );
  }

  if (tokenManager.isTokenExpired(tokenData)) {
    // Refresh not implemented; require re-authentication
    throw new McpError(
      ErrorCode.InvalidRequest,
      "Token expired. Please re-authenticate."
    );
  }

  return tokenData.accessToken;
}

// Register MCP request handlers
server.setRequestHandler(ListToolsRequestSchema, async (_request: any) => {
  log.info("Listing tools");
  return {
    tools: graphTools.getToolDefinitions(),
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request: any) => {
  const { name, arguments: args } = request.params;
  log.info(`Calling tool: ${name}`, args);

  // Get user context from request metadata (if available)
  const context = request.meta?.context as any;
  if (!context?.sessionId) {
    throw new McpError(ErrorCode.InvalidRequest, "Missing session context");
  }

  try {
    const accessToken = await getAccessTokenForSession(context.sessionId);
    const graphService = new GraphService(accessToken);

    const result = await graphTools.executeTool(name, args, graphService);

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error: any) {
    log.error("Tool execution error:", error);
    if (
      error.message.includes("Authentication required") ||
      error.message.includes("Re-authentication required")
    ) {
      // Return specific error that LibreChat recognizes as needing OAuth
      throw new McpError(
        ErrorCode.InvalidRequest,
        "OAuth authentication required"
      );
    }
    throw error;
  }
});

server.setRequestHandler(ListResourcesRequestSchema, async () => {
  log.info("Listing resources");
  return {
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
  };
});

server.setRequestHandler(ReadResourceRequestSchema, async (request: any) => {
  const { uri } = request.params;
  log.info(`Reading resource: ${uri}`);

  const context = request.meta?.context as any;
  if (!context?.sessionId) {
    throw new McpError(ErrorCode.InvalidRequest, "Missing session context");
  }

  try {
    const accessToken = await getAccessTokenForSession(context.sessionId);
    const graphService = new GraphService(accessToken);

    let data;
    switch (uri) {
      case "graph://user/profile":
        data = await graphService.getUserProfile();
        break;
      case "graph://user/mail":
        data = await graphService.getMessages();
        break;
      default:
        throw new McpError(
          ErrorCode.InvalidRequest,
          `Unknown resource: ${uri}`
        );
    }

    return {
      contents: [
        {
          uri,
          mimeType: "application/json",
          text: JSON.stringify(data, null, 2),
        },
      ],
    };
  } catch (error) {
    log.error("Resource read error:", error as any);
    const message = (error as any)?.message || String(error);
    throw new McpError(
      ErrorCode.InternalError,
      `Resource read failed: ${message}`
    );
  }
});

// Setup OAuth routes
setupOAuthRoutes(app, tokenManager);

// MCP endpoint handler
app.post("/mcp", async (req, res) => {
  try {
    const { sessionId } = getSessionContext(req);
    log.info("/mcp request", {
      sessionId,
      method: req.method,
      accept: req.headers["accept"],
      requestId: (req as any).requestId,
    });

    const tokenData = await tokenManager.getToken(sessionId);
    log.debug("Token lookup", {
      sessionId,
      found: !!tokenData,
      expired: tokenData ? tokenManager.isTokenExpired(tokenData) : undefined,
    });
    if (!sessionId || !tokenData || tokenManager.isTokenExpired(tokenData)) {
      // Return 401 to trigger LibreChat's OAuth flow
      res.setHeader(
        "WWW-Authenticate",
        `Bearer resource_metadata="${process.env.BASE_URL}/.well-known/oauth-protected-resource"`
      );
      log.warn("Authentication required", { sessionId });
      return res.status(401).json({
        error: "authentication_required",
        error_description: "OAuth authentication required",
      });
    }

    // Create transport with session-only context
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => sessionId,
    });

    // Add session context to request metadata (no userId)
    if (req.method === "POST" && req.body) {
      const rpcMethod = (req.body && req.body.method) || "unknown";
      log.info("Injecting MCP meta context", { sessionId, rpcMethod });
      req.body.meta = {
        ...req.body.meta,
        context: { sessionId },
      };
    }

    await server.connect(transport);
    log.info(`MCP connection established for session: ${sessionId}`);
  } catch (error) {
    log.error("MCP connection error:", error);
    res.status(400).json({
      error: "MCP_CONNECTION_ERROR",
      message: (error as any)?.message || String(error),
    });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  log.info(`Microsoft Graph MCP Server running on port ${port}`);
  log.info(`Health check: ${process.env.BASE_URL}/health`);
  log.info(`MCP endpoint: ${process.env.BASE_URL}/mcp`);
  log.info(`OAuth authorize: ${process.env.BASE_URL}/oauth/authorize`);
});

// Graceful shutdown
process.on("SIGINT", () => {
  log.info("Shutting down server...");
  process.exit(0);
});
