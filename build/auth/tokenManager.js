import { logger } from '../utils/logger.js';
const log = logger('tokenManager');
export class TokenManager {
    tokens = new Map();
    async storeToken(userId, tokenData) {
        log.info(`Storing token for user: ${userId}`);
        this.tokens.set(userId, tokenData);
    }
    async getToken(userId) {
        return this.tokens.get(userId) || null;
    }
    async removeToken(userId) {
        log.info(`Removing token for user: ${userId}`);
        this.tokens.delete(userId);
    }
    isTokenExpired(tokenData) {
        return Date.now() >= tokenData.expiresAt;
    }
    async refreshToken(userId, _refreshToken) {
        log.info(`Refreshing token for user: ${userId}`);
        // This would typically call Microsoft Graph's token refresh endpoint
        // For now, throwing an error to indicate refresh is needed
        throw new Error('Token refresh not implemented - please re-authenticate');
    }
}
// ESM export already provided above
