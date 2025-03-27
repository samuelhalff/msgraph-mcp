import {
    AuthFlowType,
    B2BOAuthProviders,
    B2BProducts,
    StytchB2BUIConfig,
    StytchEvent,
} from "@stytch/vanilla-js";
import {useEffect, useMemo, useState} from "react";
import {useStytchB2BClient, useStytchMember, StytchB2B, B2BIdentityProvider} from "@stytch/react/b2b";
import {
    AdminPortalB2BProducts,
    AdminPortalMemberManagement,
    AdminPortalOrgSettings,
    AdminPortalSSO
} from '@stytch/react/b2b/adminPortal';
import {NavLink, useLocation} from "react-router-dom";

/**
 * A higher-order component that enforces a login requirement for the wrapped component.
 * If the user is not logged in, the user is redirected to the login page and the
 * current URL is stored in localStorage to enable return after authentication.
 */
export const withLoginRequired = <P extends object>(
    Component: React.ComponentType<P>
) => {
    const WrappedComponent: React.FC<P> = (props) => {
        const {member} = useStytchMember()
        useEffect(() => {
            if (!member) {
                localStorage.setItem('returnTo', window.location.href);
                window.location.href = '/login';
            }
        }, [member])

        if (!member) {
            return null
        }
        return <Component {...props}  />;
    };

    WrappedComponent.displayName = `withLoginRequired(${Component.displayName ?? Component.name})`;

    return WrappedComponent;
};

/**
 * The other half of the withLoginRequired flow
 * Redirects the user to a specified URL stored in local storage or a default location.
 * Behavior:
 * - Checks for a `returnTo` entry in local storage to determine the redirection target.
 * - If `returnTo` exists, clears its value from local storage and navigates to the specified URL.
 * - If `returnTo` does not exist, redirects the user to the default '/okrs' location.
 */
const onLoginComplete = () => {
    const returnTo = localStorage.getItem('returnTo')
    if (returnTo) {
        localStorage.setItem('returnTo', '');
        window.location.href = returnTo;
    } else {
        window.location.href = '/okrs';
    }
}

/**
 * The Login page implementation. Wraps the StytchLogin UI component.
 * View all configuration options at https://stytch.com/docs/sdks/ui-configuration
 */
export function Login() {
    const loginConfig = useMemo<StytchB2BUIConfig>(() => ({
        authFlowType: AuthFlowType.Discovery,
        products: [B2BProducts.oauth, B2BProducts.emailOtp],
        sessionOptions: {sessionDurationMinutes: 60},
        oauthOptions: {
            providers: [{type: B2BOAuthProviders.Google}],
            discoveryRedirectURL: window.location.origin + '/authenticate',
        },
    }), [])

    const handleOnLoginComplete = (evt: StytchEvent) => {
        if (evt.type !== "AUTHENTICATE_FLOW_COMPLETE") return;
        // Let them savor the success screen
        setTimeout(onLoginComplete, 300);
    }

    return (
        <>
            <h1>OKR Manager MCP Demo</h1>
            <StytchB2B config={loginConfig} callbacks={{onEvent: handleOnLoginComplete}}/>
        </>
    )
}

/**
 * The OAuth Authorization page implementation. Wraps the Stytch B2BIdentityProvider UI component.
 * View all configuration options at https://stytch.com/docs/sdks/idp-ui-configuration
 */
export const Authorize = withLoginRequired(function () {
    const [initialized, setInitialized] = useState(false)
    // HACK! MCP doesn't support "scope discovery"
    // so there are no custom scopes being requested
    // we need to fake them
    useEffect(() => {
        const url = new URL(window.location.href);
        url.searchParams.set('scope', 'openid email profile read:okrs manage:okrs manage:krs report_kr_status');
        window.history.pushState(null, '', url.toString());
        setInitialized(true)
    }, []);

    return initialized && <B2BIdentityProvider/>
})

type Role = {
    role_id: string;
    description: string;
}
const adminPortalConfig = {
    allowedAuthMethods: [
        AdminPortalB2BProducts.emailMagicLinks,
        AdminPortalB2BProducts.oauthGoogle
    ],
    getRoleDescription: (role: Role) => {
        if (role.role_id == 'stytch_admin') {
            return 'The Big Cheese. Full access. Unlimited power.'
        } else if (role.role_id == 'manager') {
            return 'Defines Key Results for Employees to implement.'
        } else if (role.role_id == 'stytch_member') {
            return 'Gives status reports.'
        } else {
            return role.description;
        }
    },
    getRoleDisplayName: (role: Role) => {
        if (role.role_id == 'stytch_admin') {
            return 'CEO'
        } else if (role.role_id == 'manager') {
            return 'Manager'
        } else if (role.role_id == 'stytch_member') {
            return 'Employee'
        } else {
            return role.role_id
        }
    }
}

const adminPortalStyles = {
    fontFamily: `'IBM Plex Sans', monospace;`,
    container: {
        backgroundColor: 'rgb(251, 250, 249)',
        borderWidth: 0,
    }
}

export const SSOSettings = withLoginRequired(() => {
    return (<AdminPortalSSO styles={adminPortalStyles}/>)
})

export const OrgSettings = withLoginRequired(() => {
    return (<AdminPortalOrgSettings styles={adminPortalStyles}/>)
})

export const MemberSettings = withLoginRequired(() => {
    return (<AdminPortalMemberManagement styles={adminPortalStyles} config={adminPortalConfig}/>)
})

export const Nav = () => {
    const stytch = useStytchB2BClient()
    useLocation()
    const {member} = useStytchMember()

    if (!member) return null;

    return (
        <nav>
            <NavLink className={location.pathname === "/okrs" ? "active" : ""} to="/okrs">
                OKR Editor
            </NavLink>
            <NavLink className={location.pathname === "/settings/sso" ? "active" : ""} to="/settings/sso">
                SSO Configuration
            </NavLink>
            <NavLink className={location.pathname === "/settings/organization" ? "active" : ""}
                     to="/settings/organization">
                Organization Settings
            </NavLink>
            <NavLink className={location.pathname === "/settings/members" ? "active" : ""} to="/settings/members">
                Member Management
            </NavLink>
            <button className="primary" onClick={() => stytch.session.revoke()}> Log Out</button>
        </nav>
    )
}