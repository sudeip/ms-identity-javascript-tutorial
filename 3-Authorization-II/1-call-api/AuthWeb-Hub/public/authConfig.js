/**
 * authConfig.js — AuthWeb Hub's configuration.
 *
 * THIS is the single source of truth for every BlueFlames app's Entra ID
 * identity: the one app registration's client ID/tenant, and every spoke's
 * resource scope. Titan-UI/Phoenix-UI/Falcon-UI hold NONE of this anymore —
 * they only know their own small numeric id (apps.js) and AuthWeb Hub's URL,
 * and ask AuthWeb to resolve a token for them. Adding a 4th BlueFlames app
 * means one new entry here (+ one in apps.js) — nothing on the spoke side.
 */
const msalConfig = {
    auth: {
        clientId: '7df9ecfb-32cd-4a38-ada0-d796349cac68', // BlueFlames-AuthWeb app registration — the ONLY client ID any BlueFlames app ever uses.
        authority: 'https://login.microsoftonline.com/e5d00c4d-eacb-4b8c-bd6e-4e57a780ce5d',
        // '/', not '/redirect': redirectUri must be a page that actually calls
        // handleRedirectPromise() to process the response (authRedirect.js,
        // loaded by index.html) — redirect.html is deliberately blank and
        // loads no scripts at all, so it can't be the target here. Only
        // AuthWeb Hub's own origin needs to be registered as a redirect URI
        // in Entra at all — spokes never talk to Entra directly.
        redirectUri: '/',
        postLogoutRedirectUri: '/',
    },
    cache: {
        // Same kiosk rationale as before (shared frontline-worker tablets,
        // reused browser between people) — now it only needs to apply here,
        // since AuthWeb Hub is the only origin that ever holds a real MSAL
        // cache. Spokes keep nothing longer-lived than the current tab either
        // (see their spoke.js), so "sign out" here is enough to clear both.
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
 * Every BlueFlames resource's Entra scope, keyed by the same `resourceKey`
 * apps.js uses to say "this is Titan/Phoenix/Falcon's own API". AuthWeb never
 * calls any of these APIs itself — it only ever brokers a token for a spoke
 * — so no endpoint URL is needed here, just the scope.
 */
const protectedResources = {
    graphMe: {
        scopes: ['User.Read'],
    },
    phoenixApp: {
        scopes: ['api://58d35cac-7505-4ab4-82e0-babaa572cb21/access_as_user'],
    },
    titanApp: {
        scopes: ['api://69271258-d8bd-42d0-bed7-7079829e91a2/access_as_user'],
    },
    falconApp: {
        scopes: ['api://07b9051c-9b0c-4370-b53b-89c430ad1c9e/access_as_user'],
    },
};

/**
 * AuthWeb itself never calls a protected API directly — its whole job is
 * brokering tokens for spokes — so its own sign-in only needs the OIDC
 * baseline. Every resource scope goes into extraScopesToConsent so the
 * very FIRST sign-in (wherever/whenever it happens) consents to all of them
 * in one shot — every later broker request for any spoke is then silent.
 */
const loginRequest = {
    scopes: ['User.Read'],
    extraScopesToConsent: Object.keys(protectedResources)
        .filter((key) => key !== 'graphMe')
        .flatMap((key) => protectedResources[key].scopes),
};

// exporting config object for jest
if (typeof exports !== 'undefined') {
    module.exports = {
        msalConfig,
        loginRequest,
        protectedResources,
    };
}
