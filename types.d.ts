// Environment variables and bindings for Node.js
export interface Env {
    // Microsoft Graph / Azure configuration
    TENANT_ID?: string
    CLIENT_ID?: string
    CLIENT_SECRET?: string
    ACCESS_TOKEN?: string
    REDIRECT_URI?: string
    CERTIFICATE_PATH?: string
    CERTIFICATE_PASSWORD?: string
    MS_GRAPH_CLIENT_ID?: string
    OAUTH_SCOPES?: string
    USE_GRAPH_BETA?: string
    USE_INTERACTIVE?: string
    USE_CLIENT_TOKEN?: string
    USE_CERTIFICATE?: string
}

export type Todo = {
    id: string;
    text: string;
    completed: boolean;
}

// Context from the auth process, extracted from the Stytch auth token JWT
// and provided to the MCP Server as this.props
type AuthenticationContext = {
    claims: {
        "iss": string,
        "scope": string,
        "sub": string,
        "aud": string[],
        "client_id": string,
        "exp": number,
        "iat": number,
        "nbf": number,
        "jti": string,
    },
    accessToken: string
}

// Context from the Microsoft Graph OAuth process
export type MSGraphAuthContext = {
    accessToken: string
    refreshToken?: string
    expiresIn?: number
    tokenType?: string
    scope?: string
}
