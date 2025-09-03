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

// Environment variables
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || "http://localhost:3001";
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

const app = new Hono();

// Add comprehensive logging middleware
app.use("*", async (c, next) => {
  const start = Date.now();
  const method = c.req.method;
  const path = c.req.path;
  const userAgent = c.req.header("User-Agent") || "Unknown";
  const ip =
    c.req.header("x-forwarded-for") || c.req.header("x-real-ip") || "Unknown";

  logger.info(`[${method}] ${path} - IP: ${ip} - User-Agent: ${userAgent}`, {
    method,
    path,
    userAgent,
    ip,
    query: Object.fromEntries(new URL(c.req.url).searchParams),
    headers: Object.fromEntries(c.req.raw.headers),
  });

  try {
    await next();
    const duration = Date.now() - start;
    logger.info(`[${method}] ${path} - ${c.res.status} - ${duration}ms`, {
      method,
      path,
      status: c.res.status,
      duration,
    });
  } catch (error) {
    const duration = Date.now() - start;
    logger.error(`[${method}] ${path} - ERROR - ${duration}ms`, {
      method,
      path,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
      duration,
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
    ip: c.req.header("x-forwarded-for") || c.req.header("x-real-ip"),
  });

    // Use Microsoft Azure endpoints as per MCP standards - from environment variables
  const tenantId = TENANT_ID;
  const clientId = CLIENT_ID;
  const redirectUri = REDIRECT_URI;
  
  // Ensure URLs match exactly the format you specified
  const authorizationUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const discoveryUrl = `https://login.microsoftonline.com/${tenantId}/v2.0/.well-known/openid-configuration`;
  
  const discoveryDoc = {
    issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
    authorization_endpoint: authorizationUrl,
    token_endpoint: tokenUrl,
    jwks_uri: `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`,
    response_types_supported: ["code", "id_token", "code id_token", "id_token token"],
    response_modes_supported: ["query", "fragment", "form_post"],
    grant_types_supported: ["authorization_code", "refresh_token", "implicit"],
    token_endpoint_auth_methods_supported: [
      "client_secret_post",
      "private_key_jwt",
      "client_secret_basic"
    ],
    code_challenge_methods_supported: ["S256"],
    scopes_supported: [
      "openid",
      "profile",
      "email",
      "offline_access",
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
      "https://graph.microsoft.com/Tasks.ReadWrite",
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
      "upn"
    ],
    id_token_signing_alg_values_supported: ["RS256"],
    userinfo_endpoint: `https://graph.microsoft.com/oidc/userinfo`,
    end_session_endpoint: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/logout`,
    // MCP-specific metadata matching your requirements exactly
    discoveryUrl: discoveryUrl,
    client_id: clientId,
    scope: "https://graph.microsoft.com/.default",
    authorization_url: authorizationUrl,
    token_url: tokenUrl,
    redirect_uri: redirectUri,
    grantType: "authorization_code",
    responseType: "code",
    useRefreshTokens: true,
    usePkce: true
  };

  logger.info("OAuth discovery document generated", {
    issuer: discoveryDoc.issuer,
    authEndpoint: discoveryDoc.authorization_endpoint,
    tokenEndpoint: discoveryDoc.token_endpoint,
    tenantId: tenantId,
    clientId: discoveryDoc.client_id
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
    method: c.req.method,
  });

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    logger.warn("MCP middleware - missing or invalid authorization", {
      authHeader: authHeader || "null",
      path: c.req.path,
    });
    return c.json({ error: "Missing or invalid Authorization header" }, 401);
  }
  await next();
});

app.post("/mcp", async (c) => {
  logger.info("MCP endpoint hit - starting request processing", {
    contentType: c.req.header("Content-Type"),
    contentLength: c.req.header("Content-Length"),
    userAgent: c.req.header("User-Agent"),
  });

  try {
    // Extract token from Authorization header
    const authHeader = c.req.header("Authorization")!;
    const token = authHeader.replace("Bearer ", "");

    logger.info("MCP endpoint - token extracted", {
      tokenLength: token.length,
      tokenPrefix: token.substring(0, 10) + "...",
    });

    // Parse JSON-RPC request
    let request;
    try {
      request = await c.req.json();
      logger.info("MCP endpoint - JSON-RPC request parsed", {
        jsonrpc: request.jsonrpc,
        id: request.id,
        method: request.method,
        hasParams: !!request.params,
      });
    } catch (parseError) {
      logger.error("MCP endpoint - failed to parse JSON-RPC request", {
        error:
          parseError instanceof Error ? parseError.message : String(parseError),
        rawBody: await c.req.text(),
      });
      return c.json({ error: "Invalid JSON-RPC request" }, 400);
    }

    // Create MSGraphMCP instance with auth context
    const env = {
      TENANT_ID,
      CLIENT_ID,
      CLIENT_SECRET,
      ACCESS_TOKEN: token,
    } as Env;

    const auth: MSGraphAuthContext = {
      accessToken: token,
    };

    logger.info("MCP endpoint - creating MSGraphMCP instance", {
      hasTenantId: !!env.TENANT_ID,
      hasClientId: !!env.CLIENT_ID,
      hasClientSecret: !!env.CLIENT_SECRET,
      hasAccessToken: !!env.ACCESS_TOKEN,
    });

    // Create MSGraphMCP instance
    const mcp = new MSGraphMCP(env, auth);
    const mcpServer = mcp.server;

    logger.info("MCP endpoint - processing JSON-RPC request", {
      method: request.method,
      id: request.id,
      hasParams: !!request.params
    });

    // Process the JSON-RPC request directly
    let result;
    try {
      switch (request.method) {
        case 'initialize':
          // MCP initialization - return server capabilities
          result = {
            protocolVersion: "2024-11-05",
            capabilities: {
              tools: {},
              logging: {}
            },
            serverInfo: {
              name: "Microsoft Graph MCP Server",
              version: "1.0.0"
            }
          };
          break;

        case 'tools/list':
          // Return the list of registered tools
          result = {
            tools: [
              {
                name: "microsoft-graph-api",
                description: "Versatile Graph / ARM request helper."
              },
              {
                name: "microsoft-graph-profile",
                description: "Retrieves information about the current user's profile."
              },
              {
                name: "list-users",
                description: "Lists users from Microsoft Graph."
              },
              {
                name: "list-groups",
                description: "Lists groups from Microsoft Graph."
              },
              {
                name: "search-users",
                description: "Searches for users in Microsoft Graph."
              },
              {
                name: "send-mail",
                description: "Sends an email via Microsoft Graph."
              },
              {
                name: "list-calendar-events",
                description: "Lists calendar events for the current user."
              },
              {
                name: "create-calendar-event",
                description: "Creates a new calendar event."
              },
              {
                name: "search-files",
                description: "Search for files across OneDrive, SharePoint, and Teams using Microsoft Graph Search API."
              },
              {
                name: "get-schedule",
                description: "Get the free/busy availability information for users, distribution lists, or resources for a specified time period."
              }
            ]
          };
          break;

        case 'tools/call':
          // Call a specific tool using the MCP server's registered tools
          if (!request.params?.name) {
            throw new Error("Tool name is required");
          }
          
          // The MCP server should handle tool calls internally
          // For now, return a placeholder response
          logger.info("MCP endpoint - tool call requested", {
            toolName: request.params.name,
            arguments: request.params.arguments
          });
          
          result = {
            content: [
              {
                type: "text",
                text: `Tool '${request.params.name}' called successfully (placeholder response)`
              }
            ]
          };
          break;

        case 'ping':
          // Health check
          result = { status: "ok" };
          break;

        default:
          logger.error("MCP endpoint - unsupported method", {
            method: request.method,
            id: request.id
          });
          return c.json({
            jsonrpc: "2.0",
            id: request.id,
            error: {
              code: -32601,
              message: `Method not found: ${request.method}`
            }
          });
      }

      logger.info("MCP endpoint - request processed successfully", {
        method: request.method,
        id: request.id,
        resultType: typeof result
      });

      return c.json({
        jsonrpc: "2.0",
        id: request.id,
        result
      });

    } catch (methodError) {
      logger.error("MCP endpoint - method execution failed", {
        method: request.method,
        id: request.id,
        error: methodError instanceof Error ? methodError.message : String(methodError)
      });

      return c.json({
        jsonrpc: "2.0",
        id: request.id,
        error: {
          code: -32603,
          message: methodError instanceof Error ? methodError.message : "Internal error"
        }
      });
    }
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

serve(
  {
    fetch: app.fetch,
    port: parseInt(process.env.PORT || "3001"),
  },
  (info) => {
    logger.info("🚀 Microsoft Graph MCP Server started successfully", {
      port: info.port,
      address: info.address,
      environment: {
        NODE_ENV: process.env.NODE_ENV,
        PUBLIC_BASE_URL: process.env.PUBLIC_BASE_URL,
        hasTenantId: !!process.env.TENANT_ID,
        hasClientId: !!process.env.CLIENT_ID,
        hasClientSecret: !!process.env.CLIENT_SECRET,
      },
    });
    console.log(`🚀 Server running on http://localhost:${info.port}`);
  }
);
