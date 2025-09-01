import { MSGraphMCP } from "./MSGraphMCP";
import { exchangeCodeForToken, refreshAccessToken } from "./lib/msgraph-auth";
import { cors } from "hono/cors";
import { Hono } from "hono";
import { serve } from "@hono/node-server";
import dotenv from "dotenv";
import logger from "./lib/logger.js";

// Load environment variables
dotenv.config();

// Export the MSGraphMCP class so the Worker runtime can find it
export { MSGraphMCP };

// Helper functions
function getMSGraphAuthEndpoint(type: string): string {
  // Microsoft Graph authorization endpoints
  const tenantId = process.env.TENANT_ID || "common";
  const endpoints = {
    authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
  };
  return endpoints[type as keyof typeof endpoints] || endpoints.authorize;
}

// Simple bearer token authentication middleware for Microsoft Graph
const msGraphBearerTokenAuthMiddleware = async (c: { req: { header: (key: string) => string | undefined }, json: (data: unknown, status?: number) => Response }, next: () => Promise<void>) => {
  const authHeader = c.req.header("Authorization");
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return c.json({ error: "Missing or invalid Authorization header" }, 401);
  }
  // In a real implementation, you would validate the token here
  await next();
};

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
  // Optional Azure AD application client id to use when redirecting to Microsoft
  azure_client_id?: string;
}
const registeredClients = new Map<string, RegisteredClient>();

const app = new Hono();

// Enable CORS for all routes
app.use(
  "*",
  cors({
    origin: ["http://localhost:3000", "https://librechat.example.com"], // Add your LibreChat domain
    allowHeaders: ["Content-Type", "Authorization"],
    allowMethods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    credentials: true,
  })
);
app.get("/.well-known/oauth-authorization-server", (c) => {
  logger.info(
    `/.well-known/oauth-authorization-server endpoint hit, ${JSON.stringify(
      c.req.query()
    )}`
  );
  // Use public URL for OAuth discovery (required for external clients like LibreChat)
  const baseUrl = process.env.PUBLIC_BASE_URL || c.req.url.split("/.well-known")[0];
  return c.json({
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/token`,
    jwks_uri: `${baseUrl}/.well-known/jwks.json`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    token_endpoint_auth_methods_supported: [
  "client_secret_basic",
  "client_secret_post",
  "none",
    ],
    scopes_supported: [
      "openid",
      "profile",
      "email",
      "https://graph.microsoft.com/.default",
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
      "azp",
    ],
    id_token_signing_alg_values_supported: ["RS256"],
    userinfo_endpoint: `${baseUrl}/userinfo`,
    end_session_endpoint: `${baseUrl}/logout`,
    registration_endpoint: `${baseUrl}/register`,
  });
});

// Client registration endpoint
app.post("/register", async (c) => {
  logger.info(`/register endpoint hit, ${JSON.stringify(c.req.json())}`);
  try {
    const registration = await c.req.json();

    // Validate required fields
    if (!registration.client_name || !registration.redirect_uris) {
      return c.json(
        { error: "Missing required fields: client_name, redirect_uris" },
        400
      );
    }

    // Generate a dynamic client id for the registering client (public client)
    const client_id = crypto.randomUUID();

    const client: RegisteredClient = {
      client_id,
      client_name: registration.client_name || "MCP Client",
      redirect_uris: Array.isArray(registration.redirect_uris)
        ? registration.redirect_uris
        : [registration.redirect_uris],
      grant_types: registration.grant_types || ["authorization_code"],
      response_types: registration.response_types || ["code"],
      scope: registration.scope || "https://graph.microsoft.com/.default",
      token_endpoint_auth_method: registration.token_endpoint_auth_method || "none",
      created_at: Date.now(),
      azure_client_id: process.env.CLIENT_ID, // map to server's Azure app by default
    };

    registeredClients.set(client_id, client);

    // Return registration details for a public client (no client_secret)
    return c.json({
      client_id: client.client_id,
      client_id_issued_at: client.created_at,
      client_secret_expires_at: 0,
      redirect_uris: client.redirect_uris,
      grant_types: client.grant_types,
      response_types: client.response_types,
      scope: client.scope,
      token_endpoint_auth_method: client.token_endpoint_auth_method,
    });
  } catch (error) {
    logger.error("Client registration error:", error);
    return c.json({ error: "Invalid registration request" }, 400);
  }
});

// Authorization endpoint
app.get("/authorize", async (c) => {
  logger.info(`/authorize endpoint hit, ${JSON.stringify(c.req.query())}`);
  const { client_id, redirect_uri } = c.req.query();

  if (!client_id || !redirect_uri) {
    return c.json(
      { error: "Missing required parameters: client_id, redirect_uri" },
      400
    );
  }

  // Validate client
  const client = registeredClients.get(client_id as string);
  if (!client) {
    return c.json({ error: "Invalid client_id" }, 400);
  }

  // Validate redirect_uri
  if (!client.redirect_uris.includes(redirect_uri as string)) {
    return c.json({ error: "Invalid redirect_uri" }, 400);
  }

  // Redirect to Microsoft Graph authorization endpoint using the mapped azure app id
  const msGraphAuthUrl = new URL(getMSGraphAuthEndpoint("authorize"));
  const azureClientIdToUse = client.azure_client_id || process.env.CLIENT_ID;

  if (!azureClientIdToUse) {
    return c.json({ error: "Missing Microsoft Client ID" }, 400);
  }

  // Copy all parameters except client_id (we set the mapped Azure client id)
  const incoming = c.req.query();
  Object.keys(incoming).forEach((k) => {
    if (k !== "client_id") {
      msGraphAuthUrl.searchParams.set(k, String(incoming[k]));
    }
  });

  msGraphAuthUrl.searchParams.set("client_id", azureClientIdToUse);
  return c.redirect(msGraphAuthUrl.toString());
});

// Token endpoint
app.post("/token", async (c) => {
  logger.info(`/token endpoint hit, ${JSON.stringify(c.req.query())}`);
  try {
    const body = await c.req.parseBody();
    const grant_type = body.grant_type as string;
    const code = body.code as string;
    const redirect_uri = body.redirect_uri as string;
    const client_id = body.client_id as string;
    const refresh_token = body.refresh_token as string;

  logger.info(`/token called with grant_type=${grant_type}, client_id=${client_id}, redirect_uri=${redirect_uri}`);

    if (grant_type === "authorization_code") {
      if (!code || !redirect_uri || !client_id) {
        return c.json({ error: "Missing required parameters" }, 400);
      }

      // Exchange code for token
  // Find registered client so we can determine token endpoint auth method and mapped azure client id
  const regClient = registeredClients.get(client_id);
  logger.info(`/token: regClient lookup for client_id=${client_id} => ${regClient ? 'found' : 'not found'}`);
      const azureClientIdToUse = regClient?.azure_client_id || process.env.CLIENT_ID || "";
      const tokenAuthMethod = regClient?.token_endpoint_auth_method || "client_secret_post";

      // Only provide client secret when the registered client expects confidential auth
      const secretToSend = tokenAuthMethod === "none" ? undefined : process.env.CLIENT_SECRET || undefined;
  logger.info(`/token exchanging code for client_id=${client_id} mappedAzureClient=${azureClientIdToUse} token_endpoint_auth_method=${tokenAuthMethod} willSendClientSecret=${Boolean(secretToSend)}`);

      const tokenResponse = await exchangeCodeForToken(
        code,
        redirect_uri,
        azureClientIdToUse,
        secretToSend || "",
        // if client supplied PKCE verifier, forward it
        body.code_verifier as string | undefined
      );
      return c.json(tokenResponse);
    } else if (grant_type === "refresh_token") {
      if (!refresh_token) {
        return c.json({ error: "Missing refresh_token" }, 400);
      }

      // Refresh token
      const regClient = registeredClients.get(body.client_id as string);
      const azureClientIdToUse = regClient?.azure_client_id || process.env.CLIENT_ID || "";
      const tokenAuthMethod = regClient?.token_endpoint_auth_method || "client_secret_post";
      const secretToSend = tokenAuthMethod === "none" ? undefined : process.env.CLIENT_SECRET || undefined;
  logger.info(`/token refresh for client_id=${body.client_id} mappedAzureClient=${azureClientIdToUse} token_endpoint_auth_method=${tokenAuthMethod} willSendClientSecret=${Boolean(secretToSend)}`);

      const tokenResponse = await refreshAccessToken(
        refresh_token,
        azureClientIdToUse,
        secretToSend || ""
      );
      return c.json(tokenResponse);
    } else {
      return c.json({ error: "Unsupported grant_type" }, 400);
    }
  } catch (error) {
    logger.error("Token exchange error:", error);
    return c.json({ error: "Token exchange failed" }, 500);
  }
});

// User info endpoint
app.get("/userinfo", msGraphBearerTokenAuthMiddleware, async (c) => {
  logger.info(`/userinfo endpoint hit, ${JSON.stringify(c.req.query())}`);
  // This would typically fetch user info from Microsoft Graph
  return c.json({
    sub: "user-id",
    name: "User Name",
    email: "user@example.com",
  });
});

// Logout endpoint
app.post("/logout", (c) => {
  // Handle logout
  return c.json({ message: "Logged out successfully" });
});

app.all('/mcp', async (c) => {
    // MSGraphMCP.serve() returns our main fetch handler function.
    const mcpFetchHandler = MSGraphMCP.serve().fetch;
    // We call it with the raw Request object from the Hono context.
    return await mcpFetchHandler(c.req.raw);
});

// SSE route for streaming connections - commented out as serveSSE method doesn't exist
// app.use('/sse/*', msGraphBearerTokenAuthMiddleware)
// app.route('/sse', new Hono().mount('/', MSGraphMCP.serveSSE().fetch))

// Health check endpoint
app.get("/health", (c) => {
  logger.info(`/health endpoint hit, ${JSON.stringify(c.req.query())}`);
  return c.json({ status: "ok", service: "msgraph-mcp" });
});

export default app;

// Start the server if this file is run directly
if (import.meta.url === `file://${process.argv[1]}`) {
  const port = process.env.PORT || 3001;
  logger.info(`🚀 Starting Microsoft Graph MCP Server on port ${port}`);

  serve({
    fetch: app.fetch,
    port: Number(port),
  });
}
