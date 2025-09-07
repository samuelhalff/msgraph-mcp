import express from "express";
import cors from "cors";
import { randomUUID } from "crypto";
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
import type {
  CallToolRequest,
  ReadResourceRequest,
} from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";
import { setupOAuthRoutes } from "./auth/auth.js";
import { GraphTools } from "./tools/graphTools.js";
import { TokenManager } from "./auth/tokenManager.js";
import { GraphService } from "./services/graphService.js";
import { logger } from "./utils/logger.js";

dotenv.config();

const app = express();
const log = logger("main");

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// CORS middleware (configure appropriately for production)
app.use(
  cors({
    origin: process.env.BASE_URL || "*",
    // origin: ['https://your-remote-domain.com', 'https://your-other-remote-domain.com'],
    methods: ["GET", "POST", "OPTIONS"],
  exposedHeaders: ["mcp-session-id"],
  allowedHeaders: ["Content-Type", "Authorization", "mcp-session-id"],
  })
);

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// Well-known endpoint for MCP server discovery
app.get("/mcp/.well-known", (req, res) => {
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

// Helper to read session id from the standard MCP header
function getSessionContext(req: express.Request) {
  const sessionId = (req.headers["mcp-session-id"] || req.headers["mcp-session-id"]) as string | undefined;
  if (!sessionId) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      "Missing mcp-session-id header"
    );
  }
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
    if (!tokenData.refreshToken) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        "Token expired and no refresh token available. Please re-authenticate."
      );
    }

    await tokenManager.refreshToken(sessionId, tokenData.refreshToken);
    const latest = await tokenManager.getToken(sessionId);
    if (!latest) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        "Token refresh failed. Please re-authenticate."
      );
    }
    return latest.accessToken;
  }

  return tokenData.accessToken;
}

// Register MCP request handlers
server.setRequestHandler(ListToolsRequestSchema, async () => {
  log.info("Listing tools");
  return {
    tools: graphTools.getToolDefinitions(),
  };
});

type RequestContext = { sessionId: string };

// Narrow meta/context reading without using 'any'
type ParamsWithMeta = { _meta?: { context?: RequestContext } };
function extractRequestContext(params: unknown): RequestContext | undefined {
  if (params && typeof params === "object") {
    const p = params as ParamsWithMeta;
    return p._meta?.context;
  }
  return undefined;
}

server.setRequestHandler(
  CallToolRequestSchema,
  async (request: CallToolRequest) => {
    const { name, arguments: args } = request.params;
    log.info(`Calling tool: ${name}`, args);

    // Get user context from request metadata (if available)
    const context = extractRequestContext(request.params);
    if (!context?.sessionId) {
      throw new McpError(ErrorCode.InvalidRequest, "Missing session context");
    }

    try {
  const accessToken = await getAccessTokenForSession(context.sessionId);
      const graphService = new GraphService(accessToken);

      const result = await graphTools.executeTool(
        name,
        (args as Record<string, unknown>) || {},
        graphService
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      log.error("Tool execution error:", err);
      throw new McpError(
        ErrorCode.InternalError,
        `Tool execution failed: ${err.message}`
      );
    }
  }
);

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

server.setRequestHandler(
  ReadResourceRequestSchema,
  async (request: ReadResourceRequest) => {
    const { uri } = request.params;
    log.info(`Reading resource: ${uri}`);

    const context = extractRequestContext(request.params);
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
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      log.error("Resource read error:", err);
      throw new McpError(
        ErrorCode.InternalError,
        `Resource read failed: ${err.message}`
      );
    }
  }
);

// Setup OAuth routes
setupOAuthRoutes(app, tokenManager);

// MCP endpoint handler
app.all("/mcp", async (req, res) => {
  try {
  const { sessionId } = getSessionContext(req);

    // Create transport with user context
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
    });

    // Add user context to request metadata using MCP-compliant _meta
    if (req.method === "POST" && req.body && typeof req.body === "object") {
      // Ensure _meta exists and attach context
      req.body._meta = {
        ...(req.body._meta || {}),
        context: { sessionId },
      };
    }

    await server.connect(transport);
    // Hand off the HTTP request to the transport per Streamable HTTP pattern
    if (req.method === "POST") {
      await transport.handleRequest(
        req as unknown as import("http").IncomingMessage,
        res as unknown as import("http").ServerResponse,
        req.body
      );
    } else {
      await transport.handleRequest(
        req as unknown as import("http").IncomingMessage,
        res as unknown as import("http").ServerResponse
      );
    }
  log.info(`MCP connection established for session: ${sessionId}`);
  } catch (error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    log.error("MCP connection error:", err);
    res.status(400).json({
      error: "MCP_CONNECTION_ERROR",
      message: err.message,
    });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  log.info(`Microsoft Graph MCP Server running on port ${port}`);
  log.info(`Health check: http://localhost:${port}/health`);
  log.info(`MCP endpoint: http://localhost:${port}/mcp`);
  log.info(`OAuth authorize: http://localhost:${port}/oauth/authorize`);
});

// Graceful shutdown
process.on("SIGINT", () => {
  log.info("Shutting down server...");
  process.exit(0);
});
