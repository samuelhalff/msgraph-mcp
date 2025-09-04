import express, { Request, Response } from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { Server as MCPServerCore } from "@modelcontextprotocol/sdk/server/index.js";
import { GraphMCPServer } from "./GraphMCPServer.js";
import { MSGraphMCP } from "./MSGraphMCP.js";
import {
  exchangeCodeForToken,
  refreshAccessToken,
  getMicrosoftAuthEndpoint,
} from "./lib/msgraph-auth.js";
import logger from "./lib/logger.js";
import { Env, MSGraphAuthContext } from "../types.js";

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

// Microsoft Graph MCP endpoint with built-in auth logic

// Set up MCP streamable HTTP using SDK example pattern
const mcpServer = new GraphMCPServer(
  new MCPServerCore(
    { name: "mcp-server", version: "1.0.0" },
    { capabilities: { tools: {}, logging: {} } }
  )
);
const MCP_ENDPOINT = "/mcp";
app.post(MCP_ENDPOINT, async (req: Request, res: Response) => {
  const startTime = Date.now();
  const sessionId = req.headers['mcp-session-id'];
  const userAgent = req.headers['user-agent'];
  const origin = req.headers['origin'];
  const contentLength = req.headers['content-length'];
  
  logger.info("MCP POST request received", {
    sessionId: sessionId || "none",
    userAgent,
    origin,
    contentLength,
    ip: req.ip || req.connection?.remoteAddress,
    url: req.url,
    hasBody: !!req.body,
    bodyType: typeof req.body,
    contentType: req.headers['content-type']
  });

  try {
    await mcpServer.handlePostRequest(req, res);
    const duration = Date.now() - startTime;
    logger.info("MCP POST completed", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      statusCode: res.statusCode
    });
  } catch (error) {
    const duration = Date.now() - startTime;
    logger.error("MCP POST failed", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined
    });
    throw error;
  }
});

app.get(MCP_ENDPOINT, async (req: Request, res: Response) => {
  const startTime = Date.now();
  const sessionId = req.headers['mcp-session-id'];
  const userAgent = req.headers['user-agent'];
  const origin = req.headers['origin'];
  const acceptHeader = req.headers['accept'];
  
  logger.info("MCP GET (SSE) request received", {
    sessionId: sessionId || "none",
    userAgent,
    origin,
    acceptHeader,
    ip: req.ip || req.connection?.remoteAddress,
    url: req.url,
    queryParams: req.query
  });

  // Track connection state
  res.on('close', () => {
    const duration = Date.now() - startTime;
    logger.info("MCP GET connection closed", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      reason: 'client_disconnect'
    });
  });

  res.on('error', (error) => {
    const duration = Date.now() - startTime;
    logger.error("MCP GET connection error", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      error: error.message,
      stack: error.stack
    });
  });

  try {
    await mcpServer.handleGetRequest(req, res);
    const duration = Date.now() - startTime;
    logger.info("MCP GET completed", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      statusCode: res.statusCode
    });
  } catch (error) {
    const duration = Date.now() - startTime;
    logger.error("MCP GET failed", {
      sessionId: sessionId || "none",
      duration: `${duration}ms`,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined
    });
    throw error;
  }
});

// Prewarm: register tools at startup to avoid first-request latency
try {
  const prewarmEnv: Env = {
    TENANT_ID: process.env.TENANT_ID,
    CLIENT_ID: process.env.CLIENT_ID,
    CLIENT_SECRET: process.env.CLIENT_SECRET,
    REDIRECT_URI: process.env.REDIRECT_URI,
    AUTH_BASE_URL: process.env.AUTH_BASE_URL,
    GRAPH_BASE_URL: process.env.GRAPH_BASE_URL,
    USE_CLIENT_TOKEN: process.env.USE_CLIENT_TOKEN,
    USE_CERTIFICATE: process.env.USE_CERTIFICATE,
    USE_INTERACTIVE: process.env.USE_INTERACTIVE,
    USE_GRAPH_BETA: process.env.USE_GRAPH_BETA,
    CERTIFICATE_PATH: process.env.CERTIFICATE_PATH,
    CERTIFICATE_PASSWORD: process.env.CERTIFICATE_PASSWORD,
    SCOPES: process.env.SCOPES,
    PORT: process.env.PORT,
    ACCESS_TOKEN: "",
  } as Env;
  const prewarmMcp = new MSGraphMCP(prewarmEnv, { accessToken: "" });
  const tools = prewarmMcp.getAvailableTools();
  logger.info("Prewarmed MCP tool registry", { count: tools.length });
} catch (err) {
  logger.warn("Tool prewarm failed (continuing)", {
    error: err instanceof Error ? err.message : String(err),
  });
}

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
app.listen(PORT, () => {
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
