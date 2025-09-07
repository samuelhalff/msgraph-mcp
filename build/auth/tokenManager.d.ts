export interface TokenData {
    accessToken: string;
    refreshToken?: string;
    expiresAt: number;
    scope?: string;
}
export declare class TokenManager {
    tokens: Map<string, TokenData>;
    storeToken(userId: string, tokenData: TokenData): Promise<void>;
    getToken(userId: string): Promise<TokenData | null>;
    removeToken(userId: string): Promise<void>;
    isTokenExpired(tokenData: TokenData): boolean;
    refreshToken(userId: string, _refreshToken: string): Promise<void>;
}
