import {createRemoteJWKSet, jwtVerify, JWTVerifyResult} from "jose";
import {createMiddleware} from "hono/factory";
import {HTTPException} from "hono/http-exception";
import {getCookie} from "hono/cookie";


type AuthMiddlewareParams = {
    resource_id: 'objective' | 'key_result',
    action: 'create' | 'read' | 'update' | 'delete',
}
/**
 * stytchAuthMiddleware is a Hono middleware that validates that the user is logged in
 * It checks for the stytch_session_jwt cookie set by the Stytch FE SDK
 */
export const stytchSessionAuthMiddleware = ({resource_id, action}: AuthMiddlewareParams) => createMiddleware<{
    Variables: {
        memberID: string,
        organizationID: string,
    },
    Bindings: Env,
}>(async (c, next) => {
    const sessionCookie = getCookie(c, 'stytch_session_jwt');

    type Payload = {
        sub: string,
        'https://stytch.com/organization': {
            organization_id: string,
        },
        'https://stytch.com/session' :{
            roles: string[]
        }
    }
    let verifyResult: JWTVerifyResult<Payload>;


    try {
        verifyResult = await validateStytchJWT<Payload>(sessionCookie ?? '', c.env)
        console.log(verifyResult.payload);
        c.set('memberID', verifyResult.payload.sub!);
        c.set('organizationID', verifyResult.payload['https://stytch.com/organization'].organization_id);
    } catch (error) {
        console.error(error);
        throw new HTTPException(401, {message: 'Unauthenticated'})
    }

    if (!hasRBACAccess(resource_id, action, verifyResult.payload['https://stytch.com/session'].roles)) {
        console.error(verifyResult.payload['https://stytch.com/session'].roles, 'could not access', action, 'on', resource_id);
        throw new HTTPException(401, {message: 'Unauthenticated'})
    }

    await next()
})

/**
 * stytchBearerTokenAuthMiddleware is a Hono middleware that validates that the request has a Stytch-issued bearer token
 * Tokens are issued to clients at the end of a successful OAuth flow
 */
export const stytchBearerTokenAuthMiddleware = createMiddleware<{
    Bindings: Env,
}>(async (c, next) => {
    const authHeader = c.req.header('Authorization')

    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        throw new HTTPException(401, {message: 'Missing or invalid access token'})
    }
    const accessToken = authHeader.substring(7);

    try {
        const verifyResult = await validateStytchJWT(accessToken, c.env)
        // @ts-expect-error Props go brr
        c.executionCtx.props = {
            claims: verifyResult.payload,
            accessToken,
        }
    } catch (error) {
        console.error(error);
        throw new HTTPException(401, {message: 'Unauthenticated'})
    }

    await next()
})

let jwks: ReturnType<typeof createRemoteJWKSet> | null = null;

async function validateStytchJWT<T>(token: string, env: Env) {
    if (!jwks) {
        jwks = createRemoteJWKSet(new URL(getStytchOAuthEndpointUrl(env, '.well-known/jwks.json')))
    }

    return await jwtVerify<T>(token, jwks, {
        audience: env.STYTCH_PROJECT_ID,
        issuer: [`stytch.com/${env.STYTCH_PROJECT_ID}`],
        typ: "JWT",
        algorithms: ['RS256'],
    })
}

export function getStytchOAuthEndpointUrl(env: Env, endpoint: string): string {
    const baseURL = env.STYTCH_PROJECT_ID.includes('test') ?
        'https://test.stytch.com/v1/public' :
        'https://api.stytch.com/v1/public';

    return `${baseURL}/${env.STYTCH_PROJECT_ID}/${endpoint}`
}

// TODO: Terribly terribly sorry - vite can't bundle the Stytch Node SDK for Reasons Unknown
// as a placeholder, we are hardcoding the RBAC logic
// in a real example, we would pull this dynamically using the SDK, and all this logic would be handled for you
// Please avert your eyes, WIP
export function hasRBACAccess(resourceID: string, action: string, roles: string[]): boolean {

    if (resourceID === 'objective') {
        switch (action) {
            case 'create':
                return roles.includes('stytch_admin')
            case 'read':
                return roles.includes('stytch_admin') || roles.includes('stytch_member')
            case 'update':
                return roles.includes('stytch_admin')
            case 'delete':
                return roles.includes('stytch_admin')
            default:
                return false
        }
    }
    if (resourceID === 'key_result') {
        switch (action) {
            case 'create':
                return roles.includes('stytch_admin') || roles.includes('manager')
            case 'read':
                return roles.includes('stytch_admin') || roles.includes('manager') || roles.includes('stytch_member')
            case 'update':
                return roles.includes('stytch_admin') || roles.includes('manager') || roles.includes('stytch_member')
            case 'delete':
                return roles.includes('stytch_admin') || roles.includes('manager')
            default:
                return false
        }
    }
    return false
}