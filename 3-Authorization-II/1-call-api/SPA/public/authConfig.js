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
        // "localStorage" is required (over "sessionStorage") for the BlueFlames app-switcher demo below:
        // acquireTokenSilent needs to find the cached account/tokens even though Phoenix/Titan
        // are "different apps" sharing the same AuthWeb MSAL instance.
        cacheLocation: 'localStorage',
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
    // Each BlueFlames app is its own Entra ID app registration (own audience/scope),
    // but all are called from the single shared AuthWeb MSAL instance below.
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
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
const loginRequest = {
    scopes: [...protectedResources.graphMe.scopes],
    // extraScopesToConsent pre-consents the user to the BlueFlames Phoenix and Titan APIs
    // during the FIRST sign-in, so later switching between apps is a silent
    // acquireTokenSilent() call — no extra consent prompts. See authWeb.js.
    extraScopesToConsent: [
        ...protectedResources.phoenixApp.scopes,
        ...protectedResources.titanApp.scopes,
    ],
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
