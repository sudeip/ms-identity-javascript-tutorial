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
    const account = myMSALObj.getAllAccounts()[0];
    myMSALObj.logoutRedirect({
        account,
        postLogoutRedirectUri: '/',
        logoutHint: account ? account.username : undefined,
    });
}
