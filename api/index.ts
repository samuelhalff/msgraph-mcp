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

// Utility function to extract scopes from JWT token
function extractScopesFromToken(token: string): string[] {
  try {
    if (!token) return [];
    
    // Decode JWT token (basic extraction without verification for scope reading)
    const parts = token.split('.');
    if (parts.length !== 3) return [];
    
    const payload = JSON.parse(atob(parts[1]));
    
    // Extract scopes from different possible claims
    const scopes = payload.scp || payload.scope || payload.scopes || "";
    
    if (typeof scopes === 'string') {
      return scopes.split(' ').filter(s => s.length > 0);
    } else if (Array.isArray(scopes)) {
      return scopes;
    }
    
    return [];
  } catch (error) {
    logger.warn("Failed to extract scopes from token", { error: error instanceof Error ? error.message : String(error) });
    return [];
  }
}

// Default Microsoft Graph scopes for when no token is available
const DEFAULT_MSGRAPH_SCOPES = [
  process.env.GRAPH_BASE_URL ? `${process.env.GRAPH_BASE_URL}/.default` : "https://graph.microsoft.com/.default"
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

// OAuth Protected Resource Metadata (RFC9728) - REQUIRED by MCP spec
app.get("/.well-known/oauth-protected-resource", async (c) => {
  logger.info("OAuth Protected Resource Metadata endpoint hit", {
    query: c.req.query(),
    userAgent: c.req.header("User-Agent"),
    ip: c.req.header("x-forwarded-for") || c.req.header("x-real-ip"),
  });

  // Get the server's base URL for canonical resource URI
  const protocol = c.req.header("x-forwarded-proto") || "https";
  const host = c.req.header("host");
  const serverBaseUrl = `${protocol}://${host}`;

  // Try to extract scopes from Authorization header if present
  let supportedScopes = DEFAULT_MSGRAPH_SCOPES;
  const authHeader = c.req.header("Authorization");
  if (authHeader && authHeader.startsWith("Bearer ")) {
    const token = authHeader.replace("Bearer ", "");
    const tokenScopes = extractScopesFromToken(token);
    if (tokenScopes.length > 0) {
      // Use scopes from the actual token, ensuring we include Graph scopes
      supportedScopes = [...new Set([...tokenScopes, ...DEFAULT_MSGRAPH_SCOPES])];
    }
  }

  const protectedResourceMetadata = {
    // RFC9728 Section 3.1 - Required fields
    resource: serverBaseUrl, // Canonical server URI as defined in MCP spec
    authorization_servers: [
      `${process.env.AUTH_BASE_URL ?? 'https://login.microsoftonline.com'}/${TENANT_ID}/v2.0`
    ],
    
    // RFC9728 Section 3.2 - Optional but recommended fields
    scopes_supported: supportedScopes,
    
    // MCP-specific metadata
    mcp_version: "2024-11-05",
    server_info: {
      name: "Microsoft Graph MCP Server",
      version: "1.0.0"
    }
  };

  logger.info("OAuth Protected Resource Metadata generated", {
    resource: protectedResourceMetadata.resource,
    authorizationServers: protectedResourceMetadata.authorization_servers,
    scopesCount: protectedResourceMetadata.scopes_supported.length,
    dynamicScopes: authHeader ? true : false
  });

  return c.json(protectedResourceMetadata);
});

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
  const authBaseUrl = process.env.AUTH_BASE_URL ?? 'https://login.microsoftonline.com';
  const graphBaseUrl = process.env.GRAPH_BASE_URL ?? 'https://graph.microsoft.com';
  
  // Ensure URLs match exactly the format you specified
  const authorizationUrl = `${authBaseUrl}/${tenantId}/oauth2/v2.0/authorize`;
  const tokenUrl = `${authBaseUrl}/${tenantId}/oauth2/v2.0/token`;
  const discoveryUrl = `${authBaseUrl}/${tenantId}/v2.0/.well-known/openid-configuration`;
  
  const discoveryDoc = {
    issuer: `${authBaseUrl}/${tenantId}/v2.0`,
    authorization_endpoint: authorizationUrl,
    token_endpoint: tokenUrl,
    jwks_uri: `${authBaseUrl}/${tenantId}/discovery/v2.0/keys`,
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
      ...DEFAULT_MSGRAPH_SCOPES
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

// Microsoft Graph MCP endpoint with built-in auth logic

app.post("/mcp", async (c) => {
  logger.info("MCP endpoint hit - starting request processing", {
    contentType: c.req.header("Content-Type"),
    contentLength: c.req.header("Content-Length"),
    userAgent: c.req.header("User-Agent"),
  });

  try {
    // Parse JSON-RPC request first
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
      });
      return c.json({ error: "Invalid JSON-RPC request" }, 400);
    }

    // Check if this method requires authentication
    const publicMethods = ["initialize", "ping", "tools/list"];
    const protectedMethods = ["tools/call"];  
    const allSupportedMethods = [...publicMethods, ...protectedMethods];
    
    // Check if method is supported first
    if (!allSupportedMethods.includes(request.method)) {
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
    
    const requiresAuth = !publicMethods.includes(request.method);
    
    let token = "";
    let env: Env;
    let auth: MSGraphAuthContext;

    if (requiresAuth) {
      // Extract token from Authorization header for authenticated methods
      const authHeader = c.req.header("Authorization");
      if (!authHeader || !authHeader.startsWith("Bearer ")) {
        logger.warn("MCP endpoint - missing auth for protected method", {
          method: request.method,
          hasAuthHeader: !!authHeader
        });
        
        // Add WWW-Authenticate header as required by RFC9728 Section 5.1
        const protocol = c.req.header("x-forwarded-proto") || "https";
        const host = c.req.header("host");
        const resourceMetadataUrl = `${protocol}://${host}/.well-known/oauth-protected-resource`;
        
        c.header("WWW-Authenticate", `Bearer realm="MCP", resource_metadata_url="${resourceMetadataUrl}"`);
        
        return c.json({
          jsonrpc: "2.0",
          id: request.id,
          error: {
            code: -32002,
            message: "Authentication required for this method"
          }
        }, 401);
      }
      
      token = authHeader.replace("Bearer ", "");
      logger.info("MCP endpoint - token extracted", {
        tokenLength: token.length,
        tokenPrefix: token.substring(0, 10) + "...",
      });

      // Create auth context for authenticated requests
      env = {
        TENANT_ID,
        CLIENT_ID,
        CLIENT_SECRET,
        ACCESS_TOKEN: token,
      } as Env;

      auth = {
        accessToken: token,
      };
    } else {
      // For public methods, create minimal auth context
      env = {
        TENANT_ID,
        CLIENT_ID,
        CLIENT_SECRET,
      } as Env;

      auth = {
        accessToken: "", // Empty for public methods
      };
    }

    logger.info("MCP endpoint - processing JSON-RPC request", {
      method: request.method,
      id: request.id,
      hasParams: !!request.params,
      requiresAuth
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

        case 'tools/list': {
          // Return the list of registered tools dynamically
          const mcp = new MSGraphMCP(env, auth);
          result = {
            tools: mcp.getAvailableTools()
          };
          break;
        }

        case 'tools/call': {
          // Call a specific tool using the MCP server's registered tools
          if (!request.params?.name) {
            throw new Error("Tool name is required");
          }
          
          // Create MSGraphMCP instance with auth context for tool calls
          const mcp = new MSGraphMCP(env, auth);
          // Access server to ensure it's initialized
          void mcp.server;
          
          // Call the tool through our direct tool handler
          logger.info("MCP endpoint - tool call requested", {
            toolName: request.params.name,
            arguments: request.params.arguments
          });
          
          try {
            // Call the tool directly using our stored handlers
            result = await mcp.callTool(request.params.name, request.params.arguments || {});
          } catch (error) {
            logger.error("Tool call failed", { error: error instanceof Error ? error.message : String(error) });
            throw error;
          }
          
          break;
        }

        case 'ping':
          // Health check
          result = { status: "ok" };
          break;
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
