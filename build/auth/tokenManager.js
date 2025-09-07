import { logger } from '../utils/logger.js';
const log = logger('tokenManager');
export class TokenManager {
    tokens = new Map();
    async storeToken(key, tokenData) {
        log.info(`Storing token for key: ${key}`);
        this.tokens.set(key, tokenData);
    }
    async getToken(key) {
        return this.tokens.get(key) || null;
    }
    async removeToken(key) {
        log.info(`Removing token for key: ${key}`);
        this.tokens.delete(key);
    }
    isTokenExpired(tokenData) {
        return Date.now() >= tokenData.expiresAt;
    }
    async refreshToken(_key, _refreshToken) {
        log.info(`Refresh requested but not implemented`);
        throw new Error('Token refresh not implemented - please re-authenticate');
    }
}
