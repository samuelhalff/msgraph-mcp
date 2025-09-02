import { Hono } from "hono";
import { serve } from "@hono/node-server";
import { cors } from "hono/cors";
import { MSGraphMCP } from "./MSGraphMCP.js";
import {
  exchangeCodeForToken,
  refreshAccessToken,
  getMicrosoftAuthEndpoint,
} from "./lib/msgraph-auth.js";
import logger from "./lib/logger.js";

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

// Environment variables
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || "http://localhost:3001";
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

// Validate required environment variables
if (!TENANT_ID || !CLIENT_ID) {
  logger.error("Missing required environment variables: TENANT_ID, CLIENT_ID");
  throw new Error("Missing required environment variables");
}

const app = new Hono();

// Add comprehensive logging middleware
app.use('*', async (c, next) => {
  const start = Date.now();
  const method = c.req.method;
  const path = c.req.path;
  const userAgent = c.req.header('User-Agent') || 'Unknown';
  const ip = c.req.header('x-forwarded-for') || c.req.header('x-real-ip') || 'Unknown';

  logger.info(`[${method}] ${path} - IP: ${ip} - User-Agent: ${userAgent}`, {
    method,
    path,
    userAgent,
    ip,
    query: Object.fromEntries(new URL(c.req.url).searchParams),
    headers: Object.fromEntries(c.req.raw.headers)
  });

  try {
    await next();
    const duration = Date.now() - start;
    logger.info(`[${method}] ${path} - ${c.res.status} - ${duration}ms`, {
      method,
      path,
      status: c.res.status,
      duration
    });
  } catch (error) {
    const duration = Date.now() - start;
    logger.error(`[${method}] ${path} - ERROR - ${duration}ms`, {
      method,
      path,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
      duration
    });
    throw error;
  }
});

app.use(cors());

// OAuth Authorization Server Discovery
app.get("/.well-known/oauth-authorization-server", async (c) => {
  logger.info("OAuth discovery endpoint hit", {
    query: c.req.query(),
    userAgent: c.req.header("User-Agent"),
    ip: c.req.header("x-forwarded-for") || c.req.header("x-real-ip")
  });

  // Use public URL for OAuth discovery (required for external clients like LibreChat)
  const baseUrl = PUBLIC_BASE_URL.replace(/\/$/, "");

  const discoveryDoc = {
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/token`,
    registration_endpoint: `${baseUrl}/register`,
    jwks_uri: `${baseUrl}/.well-known/jwks.json`,
    response_types_supported: ["code"],
    response_modes_supported: ["query"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    token_endpoint_auth_methods_supported: ["client_secret_basic", "client_secret_post", "none"],
    code_challenge_methods_supported: ["S256"],
    scopes_supported: [
      "openid",
      "profile",
      "email",
      "https://graph.microsoft.com/.default",
      "https://graph.microsoft.com/User.Read",
      "https://graph.microsoft.com/User.ReadWrite",
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/Mail.ReadWrite",
      "https://graph.microsoft.com/Calendars.Read",
      "https://graph.microsoft.com/Calendars.ReadWrite",
      "https://graph.microsoft.com/Contacts.Read",
      "https://graph.microsoft.com/Contacts.ReadWrite",
      "https://graph.microsoft.com/Files.Read",
      "https://graph.microsoft.com/Files.ReadWrite",
      "https://graph.microsoft.com/Notes.Read",
      "https://graph.microsoft.com/Notes.ReadWrite",
      "https://graph.microsoft.com/Tasks.Read",
      "https://graph.microsoft.com/Tasks.ReadWrite"
    ],
    claims_supported: [
      "sub",
      "iss",
      "aud",
      "exp",
      "iat",
      "auth_time",
      "nonce",
      "acr",
      "amr",
      "azp"
    ],
    id_token_signing_alg_values_supported: ["RS256"],
    userinfo_endpoint: `${baseUrl}/userinfo`,
    end_session_endpoint: `${baseUrl}/logout`,
  };

  logger.info("OAuth discovery document generated", {
    issuer: discoveryDoc.issuer,
    authEndpoint: discoveryDoc.authorization_endpoint,
    tokenEndpoint: discoveryDoc.token_endpoint
  });

  return c.json(discoveryDoc);
});

// Dynamic Client Registration endpoint
app.post("/register", async (c) => {
  logger.info("/register endpoint hit");
  try {
    const body = await c.req.json();

    // Validate required fields
    if (!body.client_name || !body.redirect_uris) {
      return c.json(
        { error: "Missing required fields: client_name, redirect_uris" },
        400
      );
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
    return c.json(
      {
        client_id: clientId,
        client_name: body.client_name || "MCP Client",
        redirect_uris: body.redirect_uris || [],
        grant_types: body.grant_types || [
          "authorization_code",
          "refresh_token",
        ],
        response_types: body.response_types || ["code"],
        scope: body.scope,
        token_endpoint_auth_method: "none",
      },
      201
    );
  } catch (error) {
    logger.error("Error in client registration", {
      error: error instanceof Error ? error.message : String(error),
    });
    return c.json({ error: "Invalid request body" }, 400);
  }
});

// Authorization endpoint - redirects to Microsoft
app.get("/authorize", async (c) => {
  const url = new URL(c.req.url);
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
  return c.redirect(microsoftAuthUrl.toString());
});

// Token exchange endpoint
app.post("/token", async (c) => {
  try {
    const body = await c.req.parseBody();

    if (body.grant_type === "authorization_code") {
      const result = await exchangeCodeForToken(
        body.code as string,
        body.redirect_uri as string,
        CLIENT_ID!,
        CLIENT_SECRET!,
        body.code_verifier as string | undefined
      );
      return c.json(result);
    } else if (body.grant_type === "refresh_token") {
      const result = await refreshAccessToken(
        TENANT_ID!,
        body.refresh_token as string,
        CLIENT_ID!,
        CLIENT_SECRET!
      );
      return c.json(result);
    }

    return c.json({ error: "unsupported_grant_type" }, 400);
  } catch (error) {
    logger.error("Error in token exchange", {
      error: error instanceof Error ? error.message : String(error),
    });
    return c.json({ error: "Token exchange failed" }, 400);
  }
});

// Microsoft Graph MCP endpoints
app.use("/mcp/*", async (c, next) => {
  const authHeader = c.req.header("Authorization");
  logger.info("MCP middleware - checking authorization", {
    hasAuthHeader: !!authHeader,
    authHeaderPrefix: authHeader ? authHeader.substring(0, 20) + "..." : null,
    path: c.req.path,
    method: c.req.method
  });

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    logger.warn("MCP middleware - missing or invalid authorization", {
      authHeader: authHeader || "null",
      path: c.req.path
    });
    return c.json({ error: "Missing or invalid Authorization header" }, 401);
  }
  await next();
});

app.post("/mcp", async (c) => {
  const startTime = Date.now();
  logger.info("MCP endpoint hit - starting request processing", {
    contentType: c.req.header("Content-Type"),
    contentLength: c.req.header("Content-Length"),
    userAgent: c.req.header("User-Agent")
  });

  try {
    // Extract token from Authorization header
    const authHeader = c.req.header("Authorization")!;
    const token = authHeader.replace("Bearer ", "");

    logger.info("MCP endpoint - token extracted", {
      tokenLength: token.length,
      tokenPrefix: token.substring(0, 10) + "..."
    });

    // Parse JSON-RPC request
    let request;
    try {
      request = await c.req.json();
      logger.info("MCP endpoint - JSON-RPC request parsed", {
        jsonrpc: request.jsonrpc,
        id: request.id,
        method: request.method,
        hasParams: !!request.params
      });
    } catch (parseError) {
      logger.error("MCP endpoint - failed to parse JSON-RPC request", {
        error: parseError instanceof Error ? parseError.message : String(parseError),
        rawBody: await c.req.text()
      });
      return c.json({ error: "Invalid JSON-RPC request" }, 400);
    }

    // Create MSGraphMCP instance with auth context
    const env = {
      TENANT_ID,
      CLIENT_ID,
      CLIENT_SECRET,
      ACCESS_TOKEN: token,
    } as any;

    const auth = {
      accessToken: token,
    };

    logger.info("MCP endpoint - creating MSGraphMCP instance", {
      hasTenantId: !!env.TENANT_ID,
      hasClientId: !!env.CLIENT_ID,
      hasClientSecret: !!env.CLIENT_SECRET,
      hasAccessToken: !!env.ACCESS_TOKEN
    });

  } catch (error) {
    logger.error("MCP request failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return c.json({ error: "MCP request failed" }, 500);
  }
});

// Health check endpoint
app.get("/", (c) => {
  logger.info("Health check endpoint hit - root path");
  return c.text("Microsoft Graph MCP Server is running");
});

app.get("/health", (c) => {
  logger.info("Health check endpoint hit - /health path");
  return c.json({ status: "ok", service: "msgraph-mcp" });
});

serve(app, (info) => {
  logger.info("🚀 Microsoft Graph MCP Server started successfully", {
    port: info.port,
    address: info.address,
    environment: {
      NODE_ENV: process.env.NODE_ENV,
      PUBLIC_BASE_URL: process.env.PUBLIC_BASE_URL,
      hasTenantId: !!process.env.TENANT_ID,
      hasClientId: !!process.env.CLIENT_ID,
      hasClientSecret: !!process.env.CLIENT_SECRET
    }
  });
  console.log(`🚀 Server running on http://localhost:${info.port}`);
});
