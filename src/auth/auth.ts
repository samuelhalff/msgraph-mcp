import { Express, Request, Response } from "express";
import {
  ConfidentialClientApplication,
  AuthorizationCodeRequest,
} from "@azure/msal-node";
import { v4 as uuidv4 } from "uuid";
import { TokenManager } from "./tokenManager.ts";
import { logger } from "../utils/logger.ts";

const log = logger("oauth");

// Lazy MSAL initialization
let msalInstance: ConfidentialClientApplication | null = null;
function getMsalInstance(): ConfidentialClientApplication {
  if (msalInstance) return msalInstance;

  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  if (!clientId || !tenantId || !clientSecret) {
    const missing = [
      !clientId ? "CLIENT_ID" : null,
      !tenantId ? "TENANT_ID" : null,
      !clientSecret ? "CLIENT_SECRET" : null,
    ].filter(Boolean);
    log.error("MSAL config missing environment variables", { missing });
    throw new Error(`MSAL configuration missing: ${missing.join(", ")}`);
  }

  msalInstance = new ConfidentialClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}`, clientSecret },
  });
  log.info("MSAL client initialized");
  return msalInstance;
}

function getScopes(): string[] {
  const s = (process.env.OAUTH_SCOPES || "openid profile email User.Read").split(" ");
  log.debug("Using OAuth scopes", { scopes: s });
  return s;
}

// In-memory state map
const oauthStates = new Map<string, { sessionId: string; redirectUri: string }>();

export function setupOAuthRoutes(app: Express, tokenManager: TokenManager) {
  // Step 1: Redirect to Azure login
  app.get("/oauth/authorize", async (req: Request, res: Response) => {
    try {
      const sessionId =
        (req.headers["mcp-session-id"] as string) ||
        (req.query.session_id as string) ||
        "";
      const redirectUri = (req.query.redirect_uri as string) || process.env.OAUTH_REDIRECT_URI!;

      log.info("OAuth authorize request", { sessionId, requestId: (req as any).requestId });
      if (!sessionId) {
        return res.status(400).json({ error: "Missing mcp-session-id or session_id" });
      }

      const state = uuidv4();
      oauthStates.set(state, { sessionId, redirectUri });

      const authUrl = await getMsalInstance().getAuthCodeUrl({
        scopes: getScopes(),
        redirectUri,
        state,
        responseMode: "query",
      });

      res.redirect(authUrl);
    } catch (err: any) {
      log.error("OAuth authorize error", err);
      res.status(500).json({ error: "OAuth authorization failed", detail: err.message });
    }
  });

  // Step 2: Handle callback and store tokens
  app.get("/oauth/callback", async (req: Request, res: Response) => {
    try {
      const { code, state: rawState, error, error_description } =
        req.query as Record<string, string>;

      if (error) {
        return res.status(400).json({ error: "OAuth error", description: error_description });
      }
      if (!code || !rawState) {
        return res.status(400).json({ error: "Missing code or state" });
      }

      // LibreChat appends ":<serverName>" to the state
      const [state] = rawState.split(":", 2);
      const stateData = oauthStates.get(state);
      if (!stateData) {
        return res.status(400).json({ error: "Invalid state parameter" });
      }
      oauthStates.delete(state);

      const tokenResponse = await getMsalInstance().acquireTokenByCode({
        code,
        scopes: getScopes(),
        redirectUri: stateData.redirectUri,
      } as AuthorizationCodeRequest);

      if (!tokenResponse.accessToken) {
        throw new Error("No access token received");
      }

      // Extract refreshToken via any-cast
      const maybeRefresh = (tokenResponse as any).refreshToken as string | undefined;

      await tokenManager.storeToken(stateData.sessionId, {
        accessToken: tokenResponse.accessToken,
        expiresAt: tokenResponse.expiresOn!.getTime(),
        ...(maybeRefresh && { refreshToken: maybeRefresh }),
      });

      log.info("OAuth tokens stored", { sessionId: stateData.sessionId });
      res.json({ success: true, message: "Authentication successful" });
    } catch (err: any) {
      log.error("OAuth callback error", err);
      res.status(500).json({ error: "OAuth callback failed", detail: err.message });
    }
  });

  // Status check
  app.get("/oauth/status/:sessionId", async (req: Request, res: Response) => {
    try {
      const { sessionId } = req.params;
      const tokenData = await tokenManager.getToken(sessionId);
      if (!tokenData) return res.json({ authenticated: false });

      res.json({
        authenticated: !tokenManager.isTokenExpired(tokenData),
        hasRefreshToken: !!tokenData.refreshToken,
        expiresAt: tokenData.expiresAt,
      });
    } catch (err: any) {
      res.status(500).json({ error: "Failed to check token status", detail: err.message });
    }
  });
}
