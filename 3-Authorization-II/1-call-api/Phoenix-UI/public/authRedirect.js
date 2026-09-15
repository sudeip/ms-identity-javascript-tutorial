// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = '';

// Marks that this tab already tried the prompt:'none' redirect fallback
// (see attemptCrossAppSso below) and it came back with no session — stops
// that fallback from retrying itself forever when there's genuinely no
// session yet. Cleared on a successful sign-in.
const SSO_ATTEMPT_KEY = 'blueflames.ssoAttempted';

/**
 * Network log demo hook (networkLog.js): the authorize/logout redirects
 * aren't fetch() calls MSAL makes — they're full browser navigations — so
 * they can't be caught by wrapping fetch(). This custom NavigationClient
 * captures the exact URL MSAL builds (client_id, scope, redirect_uri,
 * code_challenge, logout_hint, etc.) right before navigating, then performs
 * the identical default navigation (window.location.assign/replace) MSAL's
 * own NavigationClient would have done anyway — behavior is unchanged.
 */
if (typeof logNetworkEntry === 'function') {
    myMSALObj.setNavigationClient({
        navigateInternal: (url, options) => defaultMsalNavigate(url, options),
        navigateExternal: (url, options) => {
            logNetworkEntry({ type: 'redirect', method: 'GET', url });
            return defaultMsalNavigate(url, options);
        },
    });
}

function defaultMsalNavigate(url, options) {
    if (options.noHistory) {
        window.location.replace(url);
    } else {
        window.location.assign(url);
    }
    return new Promise((resolve) => setTimeout(() => resolve(true), options.timeout));
}

/**
 * A promise handler needs to be registered for handling the
 * response returned from redirect flow. For more information, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/initialization.md#redirect-apis
 */
myMSALObj
    .handleRedirectPromise()
    .then(handleResponse)
    .catch((error) => {
        // attemptCrossAppSso()'s prompt:'none' fallback below can legitimately
        // land back here with an error (login_required/interaction_required)
        // when there's truly no session yet — that's expected, not a real
        // failure. Fall through to selectAccount(), which will see
        // SSO_ATTEMPT_KEY already set and show the Sign-in button instead of
        // retrying the redirect again.
        console.error(error);
        selectAccount();
    });

function selectAccount() {
    /**
     * See here for more info on account retrieval:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const currentAccounts = myMSALObj.getAllAccounts();

    if (!currentAccounts || currentAccounts.length === 0) {
        // No account cached in THIS origin's own storage — but the user may
        // already have signed into another BlueFlames app in this browser.
        // Try silently before giving up and showing the Sign-in button.
        attemptCrossAppSso();
    } else if (currentAccounts.length > 1) {
        // Add your account choosing logic here
        console.warn('Multiple accounts detected.');
    } else if (currentAccounts.length === 1) {
        completeSignIn(currentAccounts[0]);
    }
}

/**
 * Cross-app SSO: every BlueFlames app is a separate origin with its own
 * browser storage, so there's no cached account here even if the user just
 * signed into another app moments ago. Two silent mechanisms are tried, in
 * order, before ever showing the Sign-in button:
 *
 *   1. ssoSilent() — a hidden iframe to Entra ID, no visible navigation at
 *      all. Fastest and fully invisible, but browsers increasingly block
 *      third-party cookies inside cross-site iframes, which can make this
 *      fail even with a genuinely live session.
 *
 *   2. loginRedirect({ prompt: 'none' }) — a REAL top-level navigation to
 *      Entra ID and back. This is NOT subject to third-party cookie
 *      blocking (the browser's address bar actually goes there, so the
 *      session cookie is sent as an ordinary first-party request), so it's
 *      more reliable — at the cost of a brief visible flash while it
 *      navigates away and immediately back, rather than being invisible.
 *      If a session exists, Entra ID redirects straight back with no login
 *      form ever shown. If not, it comes back with an error, caught above.
 *
 * SSO_ATTEMPT_KEY stops step 2 from retrying itself forever when there's
 * genuinely no session — only ever attempted once per tab.
 */
function attemptCrossAppSso() {
    myMSALObj
        .ssoSilent(loginRequest)
        .then((result) => {
            console.log('[AuthWeb] Cross-app SSO succeeded via ssoSilent (hidden iframe, fully invisible).');
            completeSignIn(result.account, result.idToken);
        })
        .catch((error) => {
            console.log('[AuthWeb] ssoSilent failed (often third-party cookie blocking on the hidden iframe):', error.errorCode || error.message);

            if (sessionStorage.getItem(SSO_ATTEMPT_KEY)) {
                console.log('[AuthWeb] Already tried the silent redirect fallback this session — no existing Entra ID session. Showing Sign-in button.');
                return;
            }

            console.log('[AuthWeb] Falling back to a silent top-level redirect (prompt: none)...');
            sessionStorage.setItem(SSO_ATTEMPT_KEY, 'true');
            myMSALObj.loginRedirect({ ...loginRequest, prompt: 'none' });
        });
}

/** Shared by a fresh redirect response, an already-cached account, and a freshly ssoSilent()'d account. */
function completeSignIn(account, idToken) {
    sessionStorage.removeItem(SSO_ATTEMPT_KEY);
    username = account.username;
    welcomeUser(username);

    if (idToken) {
        updateTable(account, idToken);
        return;
    }

    // The cached AccountInfo only carries decoded ID token claims, not the
    // raw JWT string — refresh silently (scoped to this app's own API, so
    // this doubles as a fresh token for it too) to get one to show.
    myMSALObj
        .acquireTokenSilent({ scopes: protectedResources[THIS_APP.resourceKey].scopes, account })
        .then((result) => updateTable(result.account, result.idToken))
        .catch((error) => {
            console.warn('Could not refresh ID token for display:', error);
            updateTable(account, null);
        });
}

function handleResponse(response) {
    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */

    if (response !== null) {
        completeSignIn(response.account, response.idToken);
    } else {
        selectAccount();
    }
}

function signIn() {
    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    sessionStorage.removeItem(SSO_ATTEMPT_KEY);
    myMSALObj.loginRedirect(loginRequest);
}

function getTokenRedirect(request) {
    /**
     * See here for more info on account retrieval:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    request.account = myMSALObj.getAccountByUsername(username);
    return myMSALObj.acquireTokenSilent(request).catch((error) => {
        console.error(error);
        console.warn('silent token acquisition fails. acquiring token using popup');
        if (error instanceof msal.InteractionRequiredAuthError) {
            // fallback to interaction when silent call fails
            return myMSALObj.acquireTokenRedirect(request);
        } else {
            console.error(error);
        }
    });
}

function signOut() {
    /**
     * Kiosk / shared-device hardening: this app runs on tablets frontline workers
     * share back-to-back, so sign-out here means "End Shift" — it must fully
     * clear this browser's session, not just this app's view of it. A plain
     * per-app logout risks leaving the Entra ID session cookie alive, which
     * would let the NEXT person get silently signed back in via SSO.
     *
     * logoutRedirect() navigates the whole page through Entra ID's own
     * end_session endpoint, clearing the AAD session cookie itself — not just
     * this app's local token cache. MSAL clears its own cache internally as
     * part of that call, so no manual clear is needed here.
     *
     * logoutHint is set explicitly to skip Entra ID's "choose an account to
     * sign out of" picker. MSAL normally derives this itself from the
     * account's ID token 'login_hint' claim — but Entra ID only includes
     * that claim if it's been added as an optional claim in the app
     * registration's manifest, which this one hasn't. Passing it ourselves
     * (any value from the account works — MSAL skips its own derivation
     * once logoutHint is already set) avoids relying on that manifest change.
     */
    const account = myMSALObj.getAccountByUsername(username);
    const logoutRequest = {
        account,
        postLogoutRedirectUri: '/',
        logoutHint: account ? account.username : undefined,
    };

    sessionStorage.removeItem(SSO_ATTEMPT_KEY);
    myMSALObj.logoutRedirect(logoutRequest);
}
