import { ConfidentialClientApplication } from '@azure/msal-node';
import { v4 as uuidv4 } from 'uuid';
import { logger } from '../utils/logger.js';
const log = logger('oauth');
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET
    }
};
const msalInstance = new ConfidentialClientApplication(msalConfig);
const scopes = (process.env.OAUTH_SCOPES || 'openid profile email User.Read').split(' ');
// Store state for OAuth flow
const oauthStates = new Map();
export function setupOAuthRoutes(app, tokenManager) {
    // OAuth authorization endpoint
    app.get('/oauth/authorize', async (req, res) => {
        try {
            // Use only mcp-session-id (header preferred, query fallback)
            const sessionId = req.headers['mcp-session-id'] || req.query.session_id;
            const redirectUri = req.query.redirect_uri || process.env.OAUTH_REDIRECT_URI;
            if (!sessionId) {
                return res.status(400).json({ error: 'Missing mcp-session-id header or session_id param' });
            }
            const state = uuidv4();
            oauthStates.set(state, { sessionId, redirectUri });
            const authUrl = await msalInstance.getAuthCodeUrl({
                scopes,
                redirectUri,
                state,
                responseMode: 'query'
            });
            log.info(`Redirecting session ${sessionId} to OAuth authorization`);
            res.redirect(authUrl);
        }
        catch (error) {
            log.error('OAuth authorize error:', error);
            res.status(500).json({ error: 'OAuth authorization failed' });
        }
    });
    // OAuth callback endpoint
    app.get('/oauth/callback', async (req, res) => {
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
            const stateData = oauthStates.get(state);
            if (!stateData) {
                return res.status(400).json({ error: 'Invalid state parameter' });
            }
            oauthStates.delete(state);
            const tokenResponse = await msalInstance.acquireTokenByCode({
                code: code,
                scopes,
                redirectUri: stateData.redirectUri
            });
            if (!tokenResponse.accessToken) {
                throw new Error('No access token received');
            }
            // Store tokens
            await tokenManager.storeToken(stateData.sessionId, {
                accessToken: tokenResponse.accessToken,
                expiresAt: tokenResponse.expiresOn?.getTime() || (Date.now() + 3600000)
            });
            log.info(`OAuth tokens stored for session: ${stateData.sessionId}`);
            res.json({ success: true, message: 'Authentication successful' });
        }
        catch (error) {
            log.error('OAuth callback error:', error);
            res.status(500).json({ error: 'OAuth callback failed' });
        }
    });
    // Token status endpoint (keyed by sessionId)
    app.get('/oauth/status/:sessionId', async (req, res) => {
        try {
            const { sessionId } = req.params;
            const tokenData = await tokenManager.getToken(sessionId);
            if (!tokenData) {
                return res.json({ authenticated: false });
            }
            const isExpired = tokenManager.isTokenExpired(tokenData);
            res.json({
                authenticated: !isExpired,
                hasRefreshToken: !!tokenData.refreshToken,
                expiresAt: tokenData.expiresAt
            });
        }
        catch (error) {
            log.error('Token status error:', error);
            res.status(500).json({ error: 'Failed to check token status' });
        }
    });
}
