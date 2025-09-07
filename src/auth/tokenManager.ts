import { logger } from "../utils/logger.ts";

const log = logger("tokenManager");

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
    // Dump entire map after storing
    log.info("Current tokens map:", Array.from(this.tokens.entries()));
  }

  async getToken(key: string): Promise<TokenData | null> {
    log.info(`Getting token for key: ${key}`);
    log.info(
      `Token for key ${key} ${this.tokens.has(key) ? "exists" : "not found"}`
    );
    // Dump entire map on every lookup
    log.info("Current tokens map:", Array.from(this.tokens.entries()));
    return this.tokens.get(key) || null;
  }

  async removeToken(key: string): Promise<void> {
    log.info(`Removing token for key: ${key}`);
    this.tokens.delete(key);
    log.info("Current tokens map:", Array.from(this.tokens.entries()));
  }

  isTokenExpired(tokenData: TokenData): boolean {
    return Date.now() >= tokenData.expiresAt;
  }

  async refreshToken(_key: string, _refreshToken: string): Promise<never> {
    log.info(`Refresh requested but not implemented`);
    throw new Error("Token refresh not implemented - please re-authenticate");
  }

  public dumpTokens(): void {
    const entries = Array.from(this.tokens.entries()).map(([k, v]) => ({
      sessionId: k,
      expiresAt: v.expiresAt,
    }));
    log.info("All stored tokens:", entries);
  }
}
