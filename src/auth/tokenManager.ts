import { logger } from '../utils/logger.js';
const log = logger('tokenManager');
export interface TokenData {
    accessToken: string;
    refreshToken?: string;
    expiresAt: number;
    scope?: string;
}

export class TokenManager {
    tokens = new Map<string, TokenData>();
    async storeToken(userId: string, tokenData: TokenData): Promise<void> {
        log.info(`Storing token for user: ${userId}`);
        this.tokens.set(userId, tokenData);
    }
    async getToken(userId: string): Promise<TokenData | null> {
        return this.tokens.get(userId) || null;
    }
    async removeToken(userId: string): Promise<void> {
        log.info(`Removing token for user: ${userId}`);
        this.tokens.delete(userId);
    }
    isTokenExpired(tokenData: TokenData): boolean {
        return Date.now() >= tokenData.expiresAt;
    }
    async refreshToken(userId: string, _refreshToken: string): Promise<void> {
        log.info(`Refreshing token for user: ${userId}`);
        // This would typically call Microsoft Graph's token refresh endpoint
        // For now, throwing an error to indicate refresh is needed
        throw new Error('Token refresh not implemented - please re-authenticate');
    }
}
// ESM export already provided above
