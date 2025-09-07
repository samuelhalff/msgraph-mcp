import { Express, Request, Response } from "express";
import {
  ConfidentialClientApplication,
  AuthorizationCodeRequest,
} from "@azure/msal-node";
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
  const s = (process.env.OAUTH_SCOPES || "openid profile email User.Read").split(
    " "
  );
  log.debug("Using OAuth scopes", { scopes: s });
  return s;
}

// Store state for OAuth flow
const oauthStates = new Map<string, { sessionId: string; redirectUri: string }>();

export function setupOAuthRoutes(app: Express, tokenManager: TokenManager) {
  // OAuth authorization endpoint
  app.get("/oauth/authorize", async (req: Request, res: Response) => {
    try {
      const sessionId =
        (req.headers["mcp-session-id"] as string) ||
        (req.query.session_id as string) ||
        "";
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
    } catch (error: any) {
      log.error("OAuth authorize error", error);
      res.status(500).json({
        error: "OAuth authorization failed",
        detail: error.message,
      });
    }
  });

  // OAuth callback endpoint
  app.get("/oauth/callback", async (req: Request, res: Response) => {
    try {
      const {
        code,
        state: rawState,
        error,
        error_description,
      } = req.query as Record<string, string>;

      if (error) {
        log.error("OAuth callback provider error", { error, error_description });
        return res.status(400).json({
          error: "OAuth error",
          description: error_description,
        });
      }

      if (!code || !rawState) {
        return res
          .status(400)
          .json({ error: "Missing authorization code or state" });
      }

      // LibreChat appends “:<serverName>” to state
      const [state] = rawState.split(":", 2);

      log.info("OAuth callback received", {
        hasCode: true,
        hasState: !!state,
      });

      const stateData = oauthStates.get(state);
      if (!stateData) {
        return res.status(400).json({ error: "Invalid state parameter" });
      }
      oauthStates.delete(state);

      const msal = getMsalInstance();
      const tokenRequest: AuthorizationCodeRequest = {
        code,
        scopes: getScopes(),
        redirectUri: stateData.redirectUri,
      };
      const tokenResponse = await msal.acquireTokenByCode(tokenRequest);

      if (!tokenResponse.accessToken) {
        throw new Error("No access token received");
      }

      // Store full token set under the original sessionId
      const refreshToken = (tokenResponse as any).refreshToken;
      await tokenManager.storeToken(stateData.sessionId, {
        accessToken: tokenResponse.accessToken,
        expiresAt: tokenResponse.expiresOn!.getTime(),
        ...(refreshToken && { refreshToken }),
      });

      log.info("OAuth tokens stored", { sessionId: stateData.sessionId });

      res.json({ success: true, message: "Authentication successful" });
    } catch (error: any) {
      log.error("OAuth callback error", error);
      res.status(500).json({
        error: "OAuth callback failed",
        detail: error.message,
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
    } catch (error: any) {
      log.error("Token status error", error);
      res.status(500).json({
        error: "Failed to check token status",
        detail: error.message,
      });
    }
  });
}
