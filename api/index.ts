import express, { Request, Response } from "express";
import { randomUUID } from "crypto";
import cors from "cors";
import bodyParser from "body-parser";
import { Server as MCPServerCore } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { GraphMCPServer } from "./GraphMCPServer.js";
import {
  exchangeCodeForToken,
  refreshAccessToken,
  getMicrosoftAuthEndpoint,
} from "./lib/msgraph-auth.js";
import logger from "./lib/logger.js";
import { Env } from "../types.js";

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

// Utility function to extract scopes from JWT token
function extractScopesFromToken(token: string): string[] {
  try {
    if (!token) return [];

    // Decode JWT token (basic extraction without verification for scope reading)
    const parts = token.split(".");
    if (parts.length !== 3) return [];

    const payload = JSON.parse(atob(parts[1]));

    // Extract scopes from different possible claims
    const scopes = payload.scp || payload.scope || payload.scopes || "";

    if (typeof scopes === "string") {
      return scopes.split(" ").filter((s) => s.length > 0);
    } else if (Array.isArray(scopes)) {
      return scopes;
    }

    return [];
  } catch (error) {
    logger.warn("Failed to extract scopes from token", {
      error: error instanceof Error ? error.message : String(error),
    });
    return [];
  }
}

// Default Microsoft Graph scopes for when no token is available
const DEFAULT_MSGRAPH_SCOPES = [
  process.env.GRAPH_BASE_URL
    ? `${process.env.GRAPH_BASE_URL}/.default`
    : "https://graph.microsoft.com/.default",
];

// Environment variables
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;

// Validate required environment variables
if (!TENANT_ID || !CLIENT_ID) {
  logger.error("Missing required environment variables: TENANT_ID, CLIENT_ID");
  logger.warn(
    "Server starting with placeholder values - configure .env file for production use"
  );
  // For development/testing, allow server to start with placeholder values
  // throw new Error("Missing required environment variables");
}

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: "2mb" }));

// Add comprehensive logging middleware
app.use((req: Request, res: Response, next: express.NextFunction) => {
  const start = Date.now();
  const { method, path } = req;
  const userAgent = req.header("User-Agent") || "Unknown";
  const ip = req.header("x-forwarded-for") || req.header("x-real-ip") || req.ip;
  logger.info(`[${method}] ${path} - IP: ${ip} - User-Agent: ${userAgent}`, {
    method,
    path,
    userAgent,
    ip,
    query: req.query,
    headers: req.headers,
  });
  res.on("finish", () => {
    const duration = Date.now() - start;
    logger.info(`[${method}] ${path} - ${res.statusCode} - ${duration}ms`, {
      method,
      path,
      status: res.statusCode,
      duration,
    });
  });
  next();
});

// OAuth Protected Resource Metadata (RFC9728) - REQUIRED by MCP spec
app.get(
  "/.well-known/oauth-protected-resource",
  async (req: Request, res: Response) => {
    logger.info("OAuth Protected Resource Metadata endpoint hit", {
      query: req.query,
      userAgent: req.header("User-Agent"),
      ip: req.header("x-forwarded-for") || req.header("x-real-ip"),
    });

    // Get the server's base URL for canonical resource URI
    const protocol = req.header("x-forwarded-proto") || "https";
    const host = req.header("host");
    const serverBaseUrl = `${protocol}://${host}`;

    // Try to extract scopes from Authorization header if present
    let supportedScopes = DEFAULT_MSGRAPH_SCOPES;
    const authHeader = req.header("Authorization");
    if (authHeader && authHeader.startsWith("Bearer ")) {
      const token = authHeader.replace("Bearer ", "");
      const tokenScopes = extractScopesFromToken(token);
      if (tokenScopes.length > 0) {
        // Use scopes from the actual token, ensuring we include Graph scopes
        supportedScopes = [
          ...new Set([...tokenScopes, ...DEFAULT_MSGRAPH_SCOPES]),
        ];
      }
    }

    const protectedResourceMetadata = {
      // RFC9728 Section 3.1 - Required fields
      resource: serverBaseUrl, // Canonical server URI as defined in MCP spec
      authorization_servers: [
        `${
          process.env.AUTH_BASE_URL ?? "https://login.microsoftonline.com"
        }/${TENANT_ID}/v2.0`,
      ],

      // RFC9728 Section 3.2 - Optional but recommended fields
      scopes_supported: supportedScopes,

      // MCP-specific metadata
      mcp_version: "2024-11-05",
      server_info: {
        name: "Microsoft Graph MCP Server",
        version: "1.0.0",
      },
    };

    logger.info("OAuth Protected Resource Metadata generated", {
      resource: protectedResourceMetadata.resource,
      authorizationServers: protectedResourceMetadata.authorization_servers,
      scopesCount: protectedResourceMetadata.scopes_supported.length,
      dynamicScopes: authHeader ? true : false,
    });

    return res.json(protectedResourceMetadata);
  }
);

// OAuth Authorization Server Discovery
app.get(
  "/.well-known/oauth-authorization-server",
  async (req: Request, res: Response) => {
    logger.info("OAuth discovery endpoint hit", {
      query: req.query,
      userAgent: req.header("User-Agent"),
      ip: req.header("x-forwarded-for") || req.header("x-real-ip"),
    });

    // Use Microsoft Azure endpoints as per MCP standards - from environment variables
    const tenantId = TENANT_ID;
    const clientId = CLIENT_ID;
    const redirectUri = REDIRECT_URI;
    const authBaseUrl =
      process.env.AUTH_BASE_URL ?? "https://login.microsoftonline.com";
    const graphBaseUrl =
      process.env.GRAPH_BASE_URL ?? "https://graph.microsoft.com";

    // Ensure URLs match exactly the format you specified
    const authorizationUrl = `${authBaseUrl}/${tenantId}/oauth2/v2.0/authorize`;
    const tokenUrl = `${authBaseUrl}/${tenantId}/oauth2/v2.0/token`;
    const discoveryUrl = `${authBaseUrl}/${tenantId}/v2.0/.well-known/openid-configuration`;

    const discoveryDoc = {
      issuer: `${authBaseUrl}/${tenantId}/v2.0`,
      authorization_endpoint: authorizationUrl,
      token_endpoint: tokenUrl,
      jwks_uri: `${authBaseUrl}/${tenantId}/discovery/v2.0/keys`,
      response_types_supported: [
        "code",
        "id_token",
        "code id_token",
        "id_token token",
      ],
      response_modes_supported: ["query", "fragment", "form_post"],
      grant_types_supported: [
        "authorization_code",
        "refresh_token",
        "implicit",
      ],
      token_endpoint_auth_methods_supported: [
        "client_secret_post",
        "private_key_jwt",
        "client_secret_basic",
      ],
      code_challenge_methods_supported: ["S256"],
      scopes_supported: [
        "openid",
        "profile",
        "email",
        "offline_access",
        ...DEFAULT_MSGRAPH_SCOPES,
      ],
      claims_supported: [
        "sub",
        "iss",
        "aud",
        "exp",
        "iat",
        "nbf",
        "auth_time",
        "name",
        "given_name",
        "family_name",
        "email",
        "preferred_username",
        "tid",
        "oid",
        "upn",
      ],
      id_token_signing_alg_values_supported: ["RS256"],
      userinfo_endpoint: `${graphBaseUrl}/oidc/userinfo`,
      end_session_endpoint: `${authBaseUrl}/${tenantId}/oauth2/v2.0/logout`,
      // MCP-specific metadata matching your requirements exactly
      discoveryUrl: discoveryUrl,
      client_id: clientId,
      scope: DEFAULT_MSGRAPH_SCOPES[0],
      authorization_url: authorizationUrl,
      token_url: tokenUrl,
      redirect_uri: redirectUri,
      grantType: "authorization_code",
      responseType: "code",
      useRefreshTokens: true,
      usePkce: true,
    };

    logger.info("OAuth discovery document generated", {
      issuer: discoveryDoc.issuer,
      authEndpoint: discoveryDoc.authorization_endpoint,
      tokenEndpoint: discoveryDoc.token_endpoint,
      tenantId: tenantId,
      clientId: discoveryDoc.client_id,
    });

    return res.json(discoveryDoc);
  }
);

// Dynamic Client Registration endpoint
app.post("/register", async (req: Request, res: Response) => {
  logger.info("/register endpoint hit");
  try {
    const body = req.body;

    // Validate required fields
    if (!body.client_name || !body.redirect_uris) {
      return res
        .status(400)
        .json({ error: "Missing required fields: client_name, redirect_uris" });
    }

    // Generate a client ID
    const clientId = crypto.randomUUID();

    // Store the client registration
    registeredClients.set(clientId, {
      client_id: clientId,
      client_name: body.client_name || "MCP Client",
      redirect_uris: body.redirect_uris || [],
      grant_types: body.grant_types || ["authorization_code", "refresh_token"],
      response_types: body.response_types || ["code"],
      scope: body.scope,
      token_endpoint_auth_method: "none",
      created_at: Date.now(),
    });

    // Return the client registration response
    return res.status(201).json({
      client_id: clientId,
      client_name: body.client_name || "MCP Client",
      redirect_uris: body.redirect_uris || [],
      grant_types: body.grant_types || ["authorization_code", "refresh_token"],
      response_types: body.response_types || ["code"],
      scope: body.scope,
      token_endpoint_auth_method: "none",
    });
  } catch (error) {
    logger.error("Error in client registration", {
      error: error instanceof Error ? error.message : String(error),
    });
    return res.status(400).json({ error: "Invalid request body" });
  }
});

// Authorization endpoint - redirects to Microsoft
app.get("/authorize", async (req: Request, res: Response) => {
  const url = new URL(req.protocol + "://" + req.get("host") + req.originalUrl);
  const microsoftAuthUrl = new URL(getMicrosoftAuthEndpoint("authorize"));

  // Copy all query parameters except client_id
  url.searchParams.forEach((value, key) => {
    if (key !== "client_id") {
      microsoftAuthUrl.searchParams.set(key, value);
    }
  });

  // Use our Microsoft app's client_id
  microsoftAuthUrl.searchParams.set("client_id", CLIENT_ID!);
  microsoftAuthUrl.searchParams.set("tenant", TENANT_ID!);

  // Redirect to Microsoft authorization page
  return res.redirect(microsoftAuthUrl.toString());
});

// Token exchange endpoint
app.post("/token", async (req: Request, res: Response) => {
  try {
    const body = req.body || {};

    if (body.grant_type === "authorization_code") {
      const result = await exchangeCodeForToken(
        body.code as string,
        body.redirect_uri as string,
        CLIENT_ID!,
        CLIENT_SECRET!,
        body.code_verifier as string | undefined
      );
      return res.json(result);
    } else if (body.grant_type === "refresh_token") {
      const result = await refreshAccessToken(
        TENANT_ID!,
        body.refresh_token as string,
        CLIENT_ID!,
        CLIENT_SECRET!
      );
      return res.json(result);
    }

    return res.status(400).json({ error: "unsupported_grant_type" });
  } catch (error) {
    logger.error("Error in token exchange", {
      error: error instanceof Error ? error.message : String(error),
    });
    return res.status(400).json({ error: "Token exchange failed" });
  }
});

// Microsoft Graph MCP endpoint, aligned with the official minimal example
const graphMcpServer = new GraphMCPServer(
  new MCPServerCore(
    { name: "msgraph-mcp", version: "1.0.0" },
    { capabilities: { tools: {}, logging: {} } }
  )
);

// Local transports map (backed by GraphMCPServer.transports)
const transports: { [sessionId: string]: StreamableHTTPServerTransport } =
  (graphMcpServer.transports as any);

const MCP_ENDPOINT = "/mcp";

// POST requests for client->server
app.post(MCP_ENDPOINT, async (req: Request, res: Response) => {
  const sessionId = req.headers["mcp-session-id"] as string | undefined;
  const authHeader = req.header("Authorization");
  const bearer = authHeader?.startsWith("Bearer ")
    ? authHeader.slice(7)
    : undefined;

  let transport: StreamableHTTPServerTransport;

  if (sessionId && transports[sessionId]) {
    // Reuse existing transport
    transport = transports[sessionId];
    if (bearer) graphMcpServer.setSessionAuth(sessionId, bearer);
  } else if (!sessionId && isInitializeRequest(req.body)) {
    // New init request
    transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (sid) => {
        transports[sid] = transport;
        if (bearer) graphMcpServer.setSessionAuth(sid, bearer);
      },
      // enableDnsRebindingProtection: true,
      // allowedHosts: ['127.0.0.1'],
    });
    (transport as any).onclose = () => {
      const sid = (transport as any).sessionId as string | undefined;
      if (sid) delete transports[sid];
    };

    await graphMcpServer.server.connect(transport);
  } else {
    res.status(400).json({
      jsonrpc: "2.0",
      error: {
        code: -32000,
        message: "Bad Request: No valid session ID provided",
      },
      id: null,
    });
    return;
  }

  await transport.handleRequest(req as any, res as any, req.body);
});

// Reusable handler for GET and DELETE
const handleSessionRequest = async (req: Request, res: Response) => {
  const sessionId = req.headers["mcp-session-id"] as string | undefined;
  if (!sessionId || !transports[sessionId]) {
    res.status(400).send("Invalid or missing session ID");
    return;
  }
  const transport = transports[sessionId];
  await transport.handleRequest(req as any, res as any);
};

app.get(MCP_ENDPOINT, handleSessionRequest);
app.delete(MCP_ENDPOINT, handleSessionRequest);

// Prewarm removed to follow the minimal example; tools are registered per-session on initialize.

// Health check endpoint
app.get("/", (req: Request, res: Response) => {
  logger.info("Health check endpoint hit - root path");
  res.setHeader("Content-Type", "text/plain");
  return res.send("Microsoft Graph MCP Server is running");
});

app.get("/health", (req: Request, res: Response) => {
  logger.info("Health check endpoint hit - /health path");
  return res.json({ status: "ok", service: "msgraph-mcp" });
});

const PORT = parseInt(process.env.PORT || "3001", 10);
const server = app.listen(PORT, () => {
  logger.info("🚀 Microsoft Graph MCP Server started successfully", {
    port: PORT,
    address: "0.0.0.0",
    environment: {
      NODE_ENV: process.env.NODE_ENV,
      PUBLIC_BASE_URL: process.env.PUBLIC_BASE_URL,
      hasTenantId: !!process.env.TENANT_ID,
      hasClientId: !!process.env.CLIENT_ID,
      hasClientSecret: !!process.env.CLIENT_SECRET,
    },
  });
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});

// Tune server timeouts for SSE stability
try {
  // Node defaults: keepAliveTimeout 5s in some environments; increase for SSE
  server.keepAliveTimeout = 120000; // 120s
  // Must be greater than keepAliveTimeout + headers read time
  server.headersTimeout = 130000; // 130s
  // Disable per-request automatic timeouts (SSE is long-lived)
  // 0 means no timeout for incoming requests
  // @ts-ignore - requestTimeout may not be typed on some Node types
  server.requestTimeout = 0;
  logger.info('HTTP server timeouts configured for SSE', {
    keepAliveTimeout: server.keepAliveTimeout,
    headersTimeout: server.headersTimeout,
    // @ts-ignore
    requestTimeout: server.requestTimeout,
  });
} catch (e) {
  logger.warn('Failed to adjust HTTP server timeouts', {
    error: e instanceof Error ? e.message : String(e)
  });
}
