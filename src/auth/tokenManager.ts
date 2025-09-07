import { logger } from '../utils/logger.ts';

const log = logger('tokenManager');

export interface TokenData {
    accessToken: string;
    refreshToken?: string;
    expiresAt: number; // epoch ms
}

export class TokenManager {
    private tokens = new Map<string, TokenData>();

    async storeToken(key: string, tokenData: TokenData): Promise<void> {
        log.info(`Storing token for key: ${key}`);
        this.tokens.set(key, tokenData);
    }

    async getToken(key: string): Promise<TokenData | null> {
        return this.tokens.get(key) || null;
    }

    async removeToken(key: string): Promise<void> {
        log.info(`Removing token for key: ${key}`);
        this.tokens.delete(key);
    }

    isTokenExpired(tokenData: TokenData): boolean {
        return Date.now() >= tokenData.expiresAt;
    }

    async refreshToken(_key: string, _refreshToken: string): Promise<never> {
        log.info(`Refresh requested but not implemented`);
        throw new Error('Token refresh not implemented - please re-authenticate');
    }
}
