import { Express, Request, Response } from "express";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { v4 as uuidv4 } from "uuid";
import { TokenManager } from "./tokenManager.ts";
import { logger } from "../utils/logger.ts";

const log = logger("oauth");

// Lazy MSAL initialization to avoid server crash when env vars are missing
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

  const msalConfig = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientSecret,
    },
  };
  msalInstance = new ConfidentialClientApplication(msalConfig);
  log.info("MSAL client initialized");
  return msalInstance;
}

function getScopes(): string[] {
  const s = (
    process.env.OAUTH_SCOPES || "openid profile email User.Read"
  ).split(" ");
  log.debug("Using OAuth scopes", { scopes: s });
  return s;
}

// Store state for OAuth flow
const oauthStates = new Map<
  string,
  { sessionId: string; redirectUri: string }
>();

export function setupOAuthRoutes(app: Express, tokenManager: TokenManager) {
  // OAuth authorization endpoint
  app.get("/oauth/authorize", async (req: Request, res: Response) => {
    try {
      // Use only mcp-session-id (header preferred, query fallback)
      const sessionId =
        (req.headers["mcp-session-id"] as string) ||
        (req.query.session_id as string);
      const redirectUri =
        (req.query.redirect_uri as string) || process.env.OAUTH_REDIRECT_URI!;
      log.info("OAuth authorize request", {
        sessionId,
        hasRedirectUri: !!redirectUri,
        requestId: (req as any).requestId,
      });

      if (!sessionId) {
        return res
          .status(400)
          .json({ error: "Missing mcp-session-id header or session_id param" });
      }

      const state = uuidv4();
      oauthStates.set(state, { sessionId, redirectUri });

      const msal = getMsalInstance();
      const authUrl = await msal.getAuthCodeUrl({
        scopes: getScopes(),
        redirectUri,
        state,
        responseMode: "query",
      });

      log.info("Redirecting to OAuth authorization", {
        sessionId,
        state,
        redirectUri,
        urlLen: authUrl.length,
      });
      res.redirect(authUrl);
    } catch (error) {
      log.error("OAuth authorize error", error as any);
      res
        .status(500)
        .json({
          error: "OAuth authorization failed",
          detail: (error as any)?.message,
        });
    }
  });

  // OAuth callback endpoint
  app.get("/oauth/callback", async (req: Request, res: Response) => {
    try {
      const { code, state, error, error_description } = req.query;

      if (error) {
        log.error("OAuth callback provider error", {
          error,
          error_description,
        });
        return res.status(400).json({
          error: "OAuth error",
          description: error_description,
        });
      }

      if (!code || !state) {
        return res
          .status(400)
          .json({ error: "Missing authorization code or state" });
      }

      log.info("OAuth callback received", {
        hasCode: !!code,
        hasState: !!state,
      });

      const stateData = oauthStates.get(state as string);
      if (!stateData) {
        return res.status(400).json({ error: "Invalid state parameter" });
      }

      oauthStates.delete(state as string);

      const msal = getMsalInstance();
      const tokenResponse = await msal.acquireTokenByCode({
        code: code as string,
        scopes: getScopes(),
        redirectUri: stateData.redirectUri,
      });

      if (!tokenResponse.accessToken) {
        throw new Error("No access token received");
      }

      // Store tokens
      await tokenManager.storeToken(stateData.sessionId, {
        accessToken: tokenResponse.accessToken!,
        refreshToken: tokenResponse.refreshToken!, // ← include refresh token
        expiresAt: tokenResponse.expiresOn!.getTime(),
      });

      log.info("OAuth tokens stored", { sessionId: stateData.sessionId });

      res.json({ success: true, message: "Authentication successful" });
    } catch (error) {
      log.error("OAuth callback error", error as any);
      res
        .status(500)
        .json({
          error: "OAuth callback failed",
          detail: (error as any)?.message,
        });
    }
  });

  // Token status endpoint (keyed by sessionId)
  app.get("/oauth/status/:sessionId", async (req: Request, res: Response) => {
    try {
      const { sessionId } = req.params;
      log.info("OAuth status check", { sessionId });
      const tokenData = await tokenManager.getToken(sessionId);

      if (!tokenData) {
        log.debug("No token data for session", { sessionId });
        return res.json({ authenticated: false });
      }

      const isExpired = tokenManager.isTokenExpired(tokenData);
      log.debug("Token status", { sessionId, isExpired });

      res.json({
        authenticated: !isExpired,
        hasRefreshToken: !!tokenData.refreshToken,
        expiresAt: tokenData.expiresAt,
      });
    } catch (error) {
      log.error("Token status error", error as any);
      res
        .status(500)
        .json({
          error: "Failed to check token status",
          detail: (error as any)?.message,
        });
    }
  });
}
