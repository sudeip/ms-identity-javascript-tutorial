/**
 * broker.js — AuthWeb Hub's actual job: given a spoke's request (either its
 * numeric appId, or a named resource like Graph) and a return URL, resolve a
 * token using AuthWeb's OWN signed-in session/cache and hand it back.
 *
 * A spoke never holds Entra config of its own — every time it needs a token
 * it sends the whole browser here:
 *
 *   GET /?app=<id>&return=<spoke's own origin>       (a spoke's own API)
 *   GET /?resource=graphMe&return=<spoke's own origin>  (Microsoft Graph)
 *
 * This page then:
 *   1. tries to resolve that request SILENTLY against AuthWeb's own session
 *      (acquireTokenSilent) — falling back to a real sign-in only if nobody
 *      is signed in yet, or Entra ID genuinely requires interaction.
 *   2. redirects straight back to the spoke with the token in the URL
 *      FRAGMENT (#authweb_token=...). Fragments are never sent over the
 *      network and never logged by any server — a deliberate choice over
 *      putting the token in a query string (see the demo plan's Architecture
 *      section for the alternative — a one-time code + server exchange —
 *      and why this simpler option is fine for AuthWeb Hub's job here).
 */

const BROKER_PENDING_KEY = 'authweb.broker.pending'; // { resourceKey, returnUrl }
const BROKER_SIGNOUT_KEY = 'authweb.broker.signout'; // '1' | absent

/** Reads ?app=/?resource=&return= off the CURRENT URL (not sessionStorage) — only meaningful on a fresh, non-redirect arrival. */
function parseIncomingBrokerRequest() {
    const params = new URLSearchParams(window.location.search);
    const returnUrl = params.get('return');
    if (!returnUrl) return null;

    const resourceParam = params.get('resource');
    if (resourceParam) return { resourceKey: resourceParam, returnUrl };

    const appId = Number(params.get('app'));
    const entry = BLUEFLAMES_APPS_BY_ID[appId];
    return entry ? { resourceKey: entry.resourceKey, returnUrl } : null;
}

/**
 * Stashes a fresh arrival's broker/signout request (if any) into
 * sessionStorage and strips it from the visible URL — called once, before
 * handleRedirectPromise() runs, so the request survives a possible
 * loginRedirect() round trip to Entra ID and back.
 */
function captureIncomingRequest() {
    const params = new URLSearchParams(window.location.search);
    let handled = false;

    const request = parseIncomingBrokerRequest();
    if (request) {
        sessionStorage.setItem(BROKER_PENDING_KEY, JSON.stringify(request));
        handled = true;
    }

    if (params.get('signout') === '1') {
        sessionStorage.setItem(BROKER_SIGNOUT_KEY, '1');
        handled = true;
    }

    if (handled) {
        history.replaceState({}, '', window.location.pathname);
    }
}

function takePendingBrokerRequest() {
    const raw = sessionStorage.getItem(BROKER_PENDING_KEY);
    if (!raw) return null;
    sessionStorage.removeItem(BROKER_PENDING_KEY);
    try {
        return JSON.parse(raw);
    } catch (error) {
        return null;
    }
}

function takeSignoutRequest() {
    if (sessionStorage.getItem(BROKER_SIGNOUT_KEY) !== '1') return false;
    sessionStorage.removeItem(BROKER_SIGNOUT_KEY);
    return true;
}

function sendTokenBack(returnUrl, resourceKey, result) {
    const fragment = new URLSearchParams({
        authweb_token: result.accessToken,
        resource: resourceKey,
        idToken: result.idToken || '',
        username: (result.account && result.account.username) || '',
        expiresOn: result.expiresOn ? String(result.expiresOn.getTime()) : '',
    });
    const url = new URL(returnUrl);
    url.hash = fragment.toString();
    console.log(`[AuthWeb] Handing "${resourceKey}" token back to ${url.origin}`);
    window.location.replace(url.toString());
}

function sendErrorBack(returnUrl, message) {
    const url = new URL(returnUrl);
    url.hash = new URLSearchParams({ authweb_error: message }).toString();
    window.location.replace(url.toString());
}

/**
 * Resolves one broker request end to end. Either:
 *   - hands a token straight back (silent success), or
 *   - navigates away (to sign in, or to step up for consent/MFA) — in which
 *     case authRedirect.js resumes this same call after the round trip, or
 *   - shows AuthWeb's own landing page with a Sign-in button, if nobody has
 *     ever signed in on this origin in this browser session yet.
 */
async function resolveBrokerRequest(resourceKey, returnUrl) {
    const resource = protectedResources[resourceKey];
    if (!resource) {
        sendErrorBack(returnUrl, `AuthWeb doesn't know resource "${resourceKey}"`);
        return;
    }

    const account = myMSALObj.getAllAccounts()[0];
    if (!account) {
        sessionStorage.setItem(BROKER_PENDING_KEY, JSON.stringify({ resourceKey, returnUrl }));
        renderHubLanding({ brokeringFor: resourceKey });
        return;
    }

    try {
        const result = await myMSALObj.acquireTokenSilent({ scopes: resource.scopes, account });
        console.log(`[AuthWeb] Silent token acquired for "${resourceKey}".`);
        sendTokenBack(returnUrl, resourceKey, result);
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            console.warn(`[AuthWeb] Silent acquisition needs interaction — redirecting to Entra ID: ${error.errorMessage}`);
            sessionStorage.setItem(BROKER_PENDING_KEY, JSON.stringify({ resourceKey, returnUrl }));
            myMSALObj.acquireTokenRedirect({ scopes: resource.scopes, account });
            return;
        }
        sendErrorBack(returnUrl, error.errorMessage || error.message);
    }
}

/** Renders AuthWeb Hub's own landing page — direct visits, and the "please sign in" state mid-broker-request. */
function renderHubLanding({ account, brokeringFor }) {
    const signInBtn = document.getElementById('signIn');
    const signOutBtn = document.getElementById('signOut');
    const welcomeEl = document.getElementById('welcome-div');
    const titleEl = document.getElementById('title-div');
    const brokerNoteEl = document.getElementById('broker-note');
    const idTokenPanel = document.getElementById('id-token-panel');

    const activeAccount = account || myMSALObj.getAllAccounts()[0];

    if (brokeringFor) {
        const app = Object.values(BLUEFLAMES_APPS).find((a) => a.resourceKey === brokeringFor);
        brokerNoteEl.textContent = `A BlueFlames app (${app ? app.name : brokeringFor}) is waiting for a token — sign in to continue.`;
        brokerNoteEl.classList.remove('d-none');
    } else {
        brokerNoteEl.classList.add('d-none');
    }

    if (activeAccount) {
        signInBtn.classList.add('d-none');
        signOutBtn.classList.remove('d-none');
        titleEl.classList.add('d-none');
        welcomeEl.classList.remove('d-none');
        welcomeEl.innerHTML = `Welcome ${activeAccount.username}! AuthWeb is holding your session — every BlueFlames app borrows a token from here.`;
        idTokenPanel.classList.remove('d-none');
        renderIdToken(activeAccount);
    } else {
        signInBtn.classList.remove('d-none');
        signOutBtn.classList.add('d-none');
        titleEl.classList.remove('d-none');
        welcomeEl.classList.add('d-none');
        idTokenPanel.classList.add('d-none');
    }
}

function renderIdToken(account) {
    const tableBody = document.getElementById('table-body-div');
    const tableDiv = document.getElementById('table-div');
    tableBody.innerHTML = '';
    tableDiv.classList.remove('d-none');

    const claims = createClaimsTable(account.idTokenClaims || {});
    Object.keys(claims).forEach((key) => {
        const row = tableBody.insertRow(0);
        row.insertCell(0).innerHTML = claims[key][0];
        row.insertCell(1).innerHTML = claims[key][1];
        row.insertCell(2).innerHTML = claims[key][2];
    });
}
