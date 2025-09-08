import express from "express";
import dotenv from "dotenv";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  McpError,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { GraphTools } from "./tools/graphTools.js";
import { GraphService } from "./services/graphService.js";

dotenv.config();

const app = express();
app.use(express.json());

// Health endpoint
app.get("/health", (_req, res) => res.sendStatus(200));

function getBearerToken(_meta?: {
  headers?: { authorization?: string };
}): string | null {
  const auth = _meta?.headers?.authorization;
  if (!auth) return null;
  const [scheme, token] = auth.split(" ");
  return scheme === "Bearer" ? token : null;
}

const transports: { [sessionId: string]: StreamableHTTPServerTransport } = {};

// MCP discovery metadata
app.get("/mcp/.well-known", (_req, res) => {
  res.json({
    name: "msgraph-mcp-server",
    version: "1.0.0",
    oauth: {
      authorization_url: process.env.AUTH_URL!,
      token_url: process.env.TOKEN_URL!,
      scopes: (
        process.env.OAUTH_SCOPES || "https://graph.microsoft.com/.default"
      ).split(" "),
    },
    transports: ["streamable-http"],
  });
});

// Initialize MCP server
const mcp = new Server(
  { name: "msgraph", version: "1.0.0" },
  {
    capabilities: {
      tools: {},
      resources: {},
    },
  }
);
const graphTools = new GraphTools();

// List available tools
mcp.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: graphTools.getToolDefinitions(),
}));

// List resources
mcp.setRequestHandler(ListResourcesRequestSchema, async () => ({
  resources: [
    {
      uri: "graph://user/profile",
      name: "User Profile",
      description: "Current user profile",
      mimeType: "application/json",
    },
    {
      uri: "graph://user/mail",
      name: "User Mail",
      description: "List of user emails",
      mimeType: "application/json",
    },
  ],
}));

// Read resource
mcp.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  const token = getBearerToken(request.params._meta?.headers!);
  if (!token)
    throw new McpError(
      ErrorCode.InvalidRequest,
      "OAuth authentication required"
    );
  const service = new GraphService(token);
  switch (request.params.uri) {
    case "graph://user/profile":
      return {
        contents: [
          {
            uri: request.params.uri,
            mimeType: "application/json",
            text: JSON.stringify(await service.getUserProfile(), null, 2),
          },
        ],
      };
    case "graph://user/mail":
      return {
        contents: [
          {
            uri: request.params.uri,
            mimeType: "application/json",
            text: JSON.stringify(await service.getMessages(), null, 2),
          },
        ],
      };
    default:
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Unknown resource: ${request.params.uri}`
      );
  }
});

// Call tool
mcp.setRequestHandler(CallToolRequestSchema, async (request) => {
  const token = getBearerToken(request.params._meta?.headers!);
  if (!token)
    throw new McpError(
      ErrorCode.InvalidRequest,
      "OAuth authentication required"
    );
  const service = new GraphService(token);
  return graphTools.executeTool(
    request.params.name,
    request.params.arguments,
    service
  );
});
// MCP transport endpoint
app.post("/mcp", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({
      error: "authentication_required",
      error_description: "OAuth authentication required",
    });
  }
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: () => (req.headers["mcp-session-id"] as string) || "",
    onsessioninitialized: (sessionId) => {
      // Store the transport by session ID
      transports[sessionId] = transport;
    },
  });
  try {
    await mcp.connect(transport);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(400).json({ error: "MCP_CONNECTION_ERROR", message });
  }
});

// Start server
const port = Number(process.env.PORT || 3000);
app.listen(port, () => {
  console.log(`MCP server listening on port ${port}`);
  console.log(`Health:      http://localhost:${port}/health`);
  console.log(`Discovery:   http://localhost:${port}/mcp/.well-known`);
  console.log(`MCP endpoint:http://localhost:${port}/mcp`);
});
