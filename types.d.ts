export type KeyResult = {
    id: string;
    text: string;
    attainment: number;
}

export type Objective = {
    id: string;
    objectiveText: string;
    keyResults: KeyResult[];
}

export type Permissions = {
    objective: 'create' | 'read' | 'update' | 'delete';
    key_result: 'create' | 'read' | 'update' | 'delete';
};

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
        'https://stytch.com/organization': {
            organization_id: string,
        },

    },
    accessToken: string
}
