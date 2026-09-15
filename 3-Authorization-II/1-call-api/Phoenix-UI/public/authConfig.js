/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 */
const msalConfig = {
    auth: {
        clientId: '7df9ecfb-32cd-4a38-ada0-d796349cac68', // This is the ONLY mandatory field that you need to supply. This is the BlueFlames-AuthWeb app registration.
        authority: 'https://login.microsoftonline.com/e5d00c4d-eacb-4b8c-bd6e-4e57a780ce5d', // Replace the placeholder with your tenant name
        redirectUri: '/', // You must register this URI on Azure Portal/App Registration. Defaults to window.location.href e.g. http://localhost:3000/,
        postLogoutRedirectUri: '/', // Indicates the page to navigate after logout.
    },
    cache: {
        // "sessionStorage", not "localStorage": these BlueFlames apps run on shared
        // frontline-worker tablets where the browser is reused by many people back
        // to back. sessionStorage is cleared automatically when the tab/browser
        // closes, so it can't leak one person's tokens into the next person's
        // session the way localStorage (which persists indefinitely) could.
        // Titan and Phoenix are genuinely separate origins/ports, so each has its
        // own sessionStorage — cross-app SSO comes from the Entra ID session
        // (via ssoSilent(), see authRedirect.js), not from sharing this cache.
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false, // set this to true if you have to support IE
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case msal.LogLevel.Error:
                        console.error(message);
                        return;
                    case msal.LogLevel.Info:
                        console.info(message);
                        return;
                    case msal.LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case msal.LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }
            },
        },
    },
};

/**
 * Add here the endpoints and scopes when obtaining an access token for protected web APIs. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
const protectedResources = {
    // ── BlueFlames demo resources ──────────────────────────────────
    // Each spoke app's API is its own Entra ID app registration (own
    // audience/scope, real per-app client IDs below) — but the UI side (this
    // file, identical in shape across every BlueFlames UI app) all uses the
    // SAME AuthWeb client ID above, just from a separate MSAL instance per
    // origin (Titan-UI's own, Phoenix-UI's own, Falcon-UI's own, ...).
    // Keys here must match BLUEFLAMES_APPS[...].resourceKey in apps.js.
    graphMe: {
        endpoint: 'https://graph.microsoft.com/v1.0/me',
        scopes: ['User.Read'],
    },
    phoenixApp: {
        // Port 5050, not 5000 - macOS AirPlay Receiver squats on 5000 by default.
        endpoint: 'http://localhost:5050/api/data',
        scopes: ['api://58d35cac-7505-4ab4-82e0-babaa572cb21/access_as_user'],
    },
    titanApp: {
        // Port 6060, not 6000 - Chrome/Firefox block 6000 outright (old X11 port).
        endpoint: 'http://localhost:6060/api/data',
        scopes: ['api://69271258-d8bd-42d0-bed7-7079829e91a2/access_as_user'],
    },
    falconApp: {
        endpoint: 'http://localhost:7070/api/data',
        scopes: ['api://07b9051c-9b0c-4370-b53b-89c430ad1c9e/access_as_user'],
    },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
const loginRequest = {
    // THIS_APP (apps.js) says which protectedResources entry is "this app's
    // own" — its scope is the PRIMARY one, so the access token that comes
    // back directly from sign-in (response.accessToken in authRedirect.js)
    // is already usable against this app's own API — no separate
    // acquireTokenSilent() round trip needed just to get a first token.
    scopes: [...protectedResources[THIS_APP.resourceKey].scopes],
    // extraScopesToConsent pre-consents the user to every OTHER resource
    // (sibling apps' APIs + Graph) during the FIRST sign-in (on whichever app
    // the user opens first) — so that when they later open another app, its
    // silent sign-in (see authRedirect.js) doesn't hit an incremental-consent
    // prompt either. Derived from protectedResources so adding a 4th app
    // needs no change here — just a new entry above and in apps.js.
    extraScopesToConsent: Object.keys(protectedResources)
        .filter((key) => key !== THIS_APP.resourceKey)
        .flatMap((key) => protectedResources[key].scopes),
};

/**
 * An optional silentRequest object can be used to achieve silent SSO
 * between applications by providing a "login_hint" property.
 */

// const silentRequest = {
//   scopes: ["openid", "profile"],
//   loginHint: "example@domain.net"
// };

// exporting config object for jest
if (typeof exports !== 'undefined') {
    module.exports = {
        msalConfig,
        loginRequest,
        protectedResources,
    };
}
