/**
 * authRedirect.js — AuthWeb Hub's MSAL plumbing.
 *
 * This is the ONE MSAL instance in the whole BlueFlames demo — no spoke app
 * (Titan-UI/Phoenix-UI/Falcon-UI) creates one of its own anymore. On load,
 * this figures out whether the browser just arrived here as:
 *   1. a fresh broker request from a spoke (?app=<id>|resource=<key>&return=<url>)
 *   2. a redirect BACK from Entra ID (mid-sign-in, or mid-broker-request that
 *      needed interaction)
 *   3. a plain, direct visit to AuthWeb's own landing page
 * and routes to the right handler in broker.js.
 */

const myMSALObj = new msal.PublicClientApplication(msalConfig);

/**
 * Network log demo hook (networkLog.js): the authorize/logout redirects
 * aren't fetch() calls MSAL makes — they're full browser navigations — so
 * they can't be caught by wrapping fetch(). This custom NavigationClient
 * captures the exact URL MSAL builds right before navigating, then performs
 * the identical default navigation MSAL's own NavigationClient would have
 * done anyway — behavior is unchanged.
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

// A fresh arrival carrying a broker request (?app=/?resource=&return=) is
// stashed immediately, before any redirect can happen — query params don't
// survive a loginRedirect() round trip to Entra ID and back, sessionStorage
// does. An explicit ?signout=1 (from a spoke's "Sign-out" button/idle
// timeout) is handled the same way: remembered, URL cleaned up, acted on
// once handleRedirectPromise() below has had a chance to settle first.
captureIncomingRequest();

myMSALObj
    .handleRedirectPromise()
    .then((response) => {
        if (takeSignoutRequest()) {
            signOut();
            return;
        }

        const pending = takePendingBrokerRequest();
        if (pending) {
            // Either a fresh loginRedirect (first-ever sign-in) or an
            // acquireTokenRedirect (interaction-required fallback) just
            // completed — resume exactly the broker request that triggered it.
            resolveBrokerRequest(pending.resourceKey, pending.returnUrl);
            return;
        }

        renderHubLanding({ account: (response && response.account) || myMSALObj.getAllAccounts()[0] });
    })
    .catch((error) => {
        console.error(error);
        const pending = takePendingBrokerRequest();
        if (pending) {
            sendErrorBack(pending.returnUrl, error.errorMessage || error.message || 'Sign-in failed');
            return;
        }
        renderHubLanding({});
    });

function signIn() {
    myMSALObj.loginRedirect(loginRequest);
}

function signOut() {
    /**
     * Kiosk / shared-device hardening carried over unchanged from the
     * per-spoke design: this navigates the whole page through Entra ID's own
     * end_session endpoint, clearing the AAD session cookie itself — not
     * just AuthWeb's own token cache. Since AuthWeb is now the ONLY place
     * that ever holds a real MSAL cache, signing out here is enough to end
     * the session everywhere — no spoke needs (or is able) to do its own
     * separate sign-out against Entra ID.
     */
    broadcastAuthWebSignedOut(); // see below — tells any sibling AuthWeb tab before this one navigates away
    const account = myMSALObj.getAllAccounts()[0];
    myMSALObj.logoutRedirect({
        account,
        postLogoutRedirectUri: '/',
        logoutHint: account ? account.username : undefined,
    });
}

/**
 * Cross-tab reactivity via BroadcastChannel, not the 'storage' event: with
 * sessionStorage (see authConfig.js), each tab's cache is its own — a
 * 'storage' event never fires between separate tabs for sessionStorage, only
 * for localStorage. BroadcastChannel is a plain live pub/sub signal between
 * browsing contexts of the SAME origin, independent of storage type, so it's
 * used here purely as a "sign-out just happened" notice — no session data
 * rides along, sibling tabs just re-check their own (now-empty) account list.
 * Only reaches tabs on AuthWeb's OWN origin — a spoke's tab is a different
 * origin entirely and has its own separate channel (see its spoke.js).
 */
const authWebChannel = 'BroadcastChannel' in window ? new BroadcastChannel('blueflames-authweb') : null;

function broadcastAuthWebSignedOut() {
    if (authWebChannel) authWebChannel.postMessage({ type: 'signed-out' });
}

if (authWebChannel) {
    authWebChannel.onmessage = (event) => {
        if (event.data && event.data.type !== 'signed-out') return;

        // This tab's OWN sessionStorage is untouched by the OTHER tab's
        // sign-out (that's the whole reason sessionStorage needs this
        // broadcast at all) — it may still show a cached account. Re-running
        // the real sign-out here (rather than just re-rendering) actually
        // clears THIS tab's cache too, via the same proven logoutRedirect()
        // path, instead of reaching into MSAL's internal storage format by
        // hand. Entra ID's session is already gone, so this completes fast
        // with no visible login form either way.
        if (myMSALObj.getAllAccounts().length > 0) {
            signOut();
        }
    };
}
