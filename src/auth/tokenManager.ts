import RedisDefault from "ioredis";
import { logger } from "../utils/logger.ts";

const log = logger("tokenManager");
// Initialize Redis client once
const Redis = (RedisDefault as any).default || RedisDefault;
const redis = new Redis(process.env.REDIS_URL!);

redis.on("error", (err: Error) => {
  console.error("[ioredis] Error event:", err.message);
});

export interface TokenData {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number; // epoch ms
}

export class TokenManager {
  private prefix = "mcp:token:";

  /** Persist tokens in Redis under mcp:token:<sessionId> */
  async storeToken(key: string, tokenData: TokenData): Promise<void> {
    log.info(`Storing token in Redis for key: ${key}`);
    await redis.set(this.prefix + key, JSON.stringify(tokenData));
  }

  /** Retrieve tokens from Redis */
  async getToken(key: string): Promise<TokenData | null> {
    log.info(`Fetching token from Redis for key: ${key}`);
    const json = await redis.get(this.prefix + key);
    if (!json) {
      log.info(`No token found in Redis for key: ${key}`);
      return null;
    }
    try {
      const data = JSON.parse(json) as TokenData;
      log.info(`Token retrieved for key: ${key}`, { expiresAt: data.expiresAt });
      return data;
    } catch (err: any) {
      log.error("Error parsing token JSON from Redis", err);
      return null;
    }
  }

  /** Remove tokens from Redis */
  async removeToken(key: string): Promise<void> {
    log.info(`Removing token from Redis for key: ${key}`);
    await redis.del(this.prefix + key);
  }

  /** Check expiry */
  isTokenExpired(tokenData: TokenData): boolean {
    return Date.now() >= tokenData.expiresAt;
  }
}
