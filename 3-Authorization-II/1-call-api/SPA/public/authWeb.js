/**
 * authWeb.js — BlueFlames shared AuthWeb library (demo)
 *
 * Simulates a shared MSAL library used by every BlueFlames app. In production
 * this would be a versioned npm package or CDN-hosted script shared across
 * BlueFlames apps; here it just builds on top of the single `myMSALObj`
 * instance created in authRedirect.js so the demo stays self-contained.
 *
 * The point being demonstrated: one sign-in (with extraScopesToConsent, see
 * authConfig.js) is enough for every BlueFlames app to silently obtain its
 * own correctly-scoped token via getTokenForApp(scopes) below — no separate
 * sign-in or consent per app.
 */

/**
 * Central token acquisition for any BlueFlames app.
 * Tries acquireTokenSilent first (cache or hidden iframe refresh) and only
 * falls back to a full-page redirect if interaction is actually required
 * (e.g. consent was revoked, or MFA is being stepped up). This is a
 * redirect, not a popup, to match the app's redirect-based login flow —
 * popups can be unreliable or outright blocked on locked-down kiosk browsers.
 *
 * Note: unlike the silent path, the redirect fallback does NOT resolve with
 * a token here — it navigates the whole page away to Entra ID and back.
 * On return, handleRedirectPromise() (authRedirect.js) picks up the result;
 * whatever button click triggered this call will need to be pressed again
 * once the page reloads signed in. In this demo that path is rare, since
 * extraScopesToConsent (authConfig.js) pre-consents every app's scope up
 * front at the initial sign-in.
 *
 * @param {string[]} scopes - the scopes for the target app's API,
 *   e.g. ["api://<phoenix-client-id>/access_as_user"]
 * @returns {Promise<string>} access token
 */
async function getTokenForApp(scopes) {
    const account = myMSALObj.getAccountByUsername(username);
    if (!account) {
        throw new Error('No active account. Please sign in first.');
    }

    const tokenRequest = { scopes, account };

    try {
        const result = await myMSALObj.acquireTokenSilent(tokenRequest);
        console.log(`[AuthWeb] Silent token acquired for scope(s): ${scopes}`);
        return result.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            console.warn(`[AuthWeb] Silent acquisition failed, falling back to redirect: ${error.message}`);
            await myMSALObj.acquireTokenRedirect(tokenRequest); // navigates away; does not return here
            return;
        }

        throw error;
    }
}

/**
 * Decodes a base64url string (the encoding used in JWT segments, which uses
 * '-'/'_' instead of '+'/'/' and omits padding — plain atob() rejects it).
 * @param {string} base64Url
 * @returns {string}
 */
function base64UrlDecode(base64Url) {
    let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const paddingNeeded = base64.length % 4;
    if (paddingNeeded) {
        base64 += '='.repeat(4 - paddingNeeded);
    }
    return atob(base64);
}

/**
 * Decode and log the key claims from an access token.
 * Useful for the demo — shows aud (which app the token is for), scp, and user.
 *
 * Note: this only works for JWT-format access tokens, which is what Azure AD
 * issues for APIs you register yourself (Phoenix, Titan). Microsoft Graph
 * issues OPAQUE access tokens to public client (SPA) apps by default — those
 * aren't decodable client-side, so this returns null for them instead of
 * throwing; see the caller for how that's handled in the UI.
 *
 * @param {string} token - raw access token
 * @param {string} label - label for console output
 * @returns {Object|null} decoded token payload, or null if it isn't a decodable JWT
 */
function inspectToken(token, label = 'Token') {
    try {
        const payload = JSON.parse(base64UrlDecode(token.split('.')[1]));
        console.group(`[AuthWeb] ${label}`);
        console.log('aud (audience):', payload.aud);
        console.log('scp (scopes):  ', payload.scp);
        console.log('upn (user):    ', payload.upn || payload.preferred_username);
        console.log('exp (expires): ', new Date(payload.exp * 1000).toLocaleTimeString());
        console.groupEnd();
        return payload;
    } catch (error) {
        console.group(`[AuthWeb] ${label}`);
        console.log('(opaque access token — not a decodable JWT; expected for Microsoft Graph on a public client app)');
        console.groupEnd();
        return null;
    }
}
