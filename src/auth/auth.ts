import { Express, Request, Response } from 'express';
import { ConfidentialClientApplication, AuthenticationResult } from '@azure/msal-node';
import { v4 as uuidv4 } from 'uuid';
import { TokenManager } from './tokenManager.js';
import { logger } from '../utils/logger.js';

const log = logger('oauth');

const scopes = (process.env.OAUTH_SCOPES || 'openid profile email User.Read').split(' ');

// Store state for OAuth flow
const oauthStates = new Map<string, { userId: string; redirectUri: string }>();

export function setupOAuthRoutes(app: Express, tokenManager: TokenManager) {
  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const oauthEnabled = Boolean(clientId && tenantId && clientSecret);
  let msalInstance: ConfidentialClientApplication | undefined;

  if (oauthEnabled) {
    msalInstance = new ConfidentialClientApplication({
      auth: {
        clientId: clientId as string,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret: clientSecret as string,
      },
    });
    log.info('OAuth is enabled with provided credentials');
  } else {
    log.warn('OAuth is disabled: missing CLIENT_ID, TENANT_ID, or CLIENT_SECRET');
  }
  
  // OAuth authorization endpoint
  app.get('/oauth/authorize', async (req: Request, res: Response) => {
    try {
      if (!oauthEnabled || !msalInstance) {
        return res.status(503).json({ error: 'OAUTH_NOT_CONFIGURED', message: 'Missing CLIENT_ID, TENANT_ID, or CLIENT_SECRET' });
      }
      const userId = req.query.user_id as string || req.headers['x-librechat-user-id'] as string;
      const redirectUri = req.query.redirect_uri as string || process.env.OAUTH_REDIRECT_URI!;
      
      if (!userId) {
        return res.status(400).json({ error: 'Missing user_id parameter or header' });
      }

      const state = uuidv4();
      oauthStates.set(state, { userId, redirectUri });

  const authUrl = await msalInstance.getAuthCodeUrl({
        scopes,
        redirectUri,
        state,
        responseMode: 'query'
      });

      log.info(`Redirecting user ${userId} to OAuth authorization`);
      res.redirect(authUrl);
    } catch (error) {
      log.error('OAuth authorize error:', error);
      res.status(500).json({ error: 'OAuth authorization failed' });
    }
  });

  // OAuth callback endpoint
  app.get('/oauth/callback', async (req: Request, res: Response) => {
  try {
      const { code, state, error, error_description } = req.query;

      if (error) {
        log.error('OAuth callback error:', error, error_description);
        return res.status(400).json({ 
          error: 'OAuth error', 
          description: error_description 
        });
      }

      if (!code || !state) {
        return res.status(400).json({ error: 'Missing authorization code or state' });
      }

      const stateData = oauthStates.get(state as string);
      if (!stateData) {
        return res.status(400).json({ error: 'Invalid state parameter' });
      }

      oauthStates.delete(state as string);

      if (!oauthEnabled || !msalInstance) {
        return res.status(503).json({ error: 'OAUTH_NOT_CONFIGURED', message: 'Missing CLIENT_ID, TENANT_ID, or CLIENT_SECRET' });
      }
      const tokenResponse: AuthenticationResult = await msalInstance.acquireTokenByCode({
        code: code as string,
        scopes,
        redirectUri: stateData.redirectUri
      });

      if (!tokenResponse.accessToken) {
        throw new Error('No access token received');
      }

      // Store tokens (MSAL's acquireTokenByCode rarely returns a refresh token; omit if not provided)
      await tokenManager.storeToken(stateData.userId, {
        accessToken: tokenResponse.accessToken,
        expiresAt: tokenResponse.expiresOn ? tokenResponse.expiresOn.getTime() : (Date.now() + 3600_000),
        scope: tokenResponse.scopes?.join(' ')
      });

      log.info(`OAuth tokens stored for user: ${stateData.userId}`);
      
      res.json({
        success: true,
        message: 'Authentication successful',
        userId: stateData.userId
      });
    } catch (error) {
      log.error('OAuth callback error:', error);
      res.status(500).json({ error: 'OAuth callback failed' });
    }
  });

  // Token status endpoint
  app.get('/oauth/status/:userId', async (req: Request, res: Response) => {
    try {
      const { userId } = req.params;
      const tokenData = await tokenManager.getToken(userId);
      
      if (!tokenData) {
        return res.json({ authenticated: false });
      }
      
      const isExpired = tokenManager.isTokenExpired(tokenData);
      
      res.json({
        authenticated: !isExpired,
        hasRefreshToken: !!tokenData.refreshToken,
        expiresAt: tokenData.expiresAt
      });
    } catch (error) {
      log.error('Token status error:', error);
      res.status(500).json({ error: 'Failed to check token status' });
    }
  });
}