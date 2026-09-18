/**
 * spoke.js — a BlueFlames spoke app's entire auth client (identical across
 * Titan-UI, Phoenix-UI, Falcon-UI, ...). No MSAL SDK, no client ID, no
 * scopes — a spoke only knows its own small numeric id (apps.js) and
 * AuthWeb Hub's URL (authConfig.js).
 *
 * Every time this app needs a token — first load, or a cached one expired —
 * it sends the WHOLE BROWSER to AuthWeb: /?app=<id>&return=<here>. AuthWeb
 * checks its own cache (or signs the user in if truly nobody has yet), then
 * bounces straight back with the token in the URL fragment (#authweb_token=
 * ...) — never sent to any server, never logged, gone from the address bar
 * the instant this script reads it.
 *
 * A short-lived copy is kept in THIS tab's own sessionStorage — cleared the
 * moment the tab closes, so a shared kiosk tablet can't hand a still-cached
 * token to the next person just because a tab was left open. AuthWeb's own
 * cache remains the real source of truth for whether the user is actually
 * signed in; this is purely a local speed-up so a click/re-render doesn't
 * force a fresh round trip every time.
 *
 * Because sessionStorage isn't shared between tabs, two tabs of THIS SAME
 * spoke app are otherwise unaware of each other — signOutOfSpoke() below
 * also broadcasts over a BroadcastChannel so a sibling tab clears itself
 * immediately too, instead of carrying on with a technically-still-valid
 * token until it naturally expires. That channel is per-origin like every
 * other browser API here, so it can't reach a DIFFERENT spoke's tab
 * (Titan-UI vs. Phoenix-UI) — that gap is accepted, not fixed: the other
 * spoke's tab still discovers the real sign-out eventually, just via its own
 * next broker round trip once its cached token expires, by which point
 * AuthWeb has no session left to silently renew it from either.
 */

const SPOKE_TOKEN_KEY = 'blueflames.spoke.tokens'; // { [resourceKey]: { accessToken, idToken, username, expiresOn } }
const SPOKE_PENDING_TAB_KEY = 'blueflames.spoke.pendingTab'; // 'own' | 'graph' — which tab to resume into after a round trip

function loadSpokeTokens() {
    try {
        return JSON.parse(sessionStorage.getItem(SPOKE_TOKEN_KEY)) || {};
    } catch (error) {
        return {};
    }
}

function saveSpokeTokens(tokens) {
    sessionStorage.setItem(SPOKE_TOKEN_KEY, JSON.stringify(tokens));
}

function getCachedToken(resourceKey) {
    const entry = loadSpokeTokens()[resourceKey];
    if (!entry) return null;
    // 60s safety margin before expiry, same idea as MSAL's own default token refresh buffer.
    if (entry.expiresOn && Date.now() > entry.expiresOn - 60000) return null;
    return entry;
}

function cacheToken(resourceKey, entry) {
    const tokens = loadSpokeTokens();
    tokens[resourceKey] = entry;
    saveSpokeTokens(tokens);
}

function isSignedIn() {
    return Object.keys(loadSpokeTokens()).length > 0;
}

/** Sends the browser to AuthWeb Hub to resolve a token for `resourceKey`, remembering which tab to resume into on return. */
function goToAuthWebForToken(resourceKey, resumeTab) {
    sessionStorage.setItem(SPOKE_PENDING_TAB_KEY, resumeTab);
    const returnUrl = window.location.origin + '/';
    const requestParam = resourceKey === THIS_APP.resourceKey ? `app=${THIS_APP.id}` : `resource=${encodeURIComponent(resourceKey)}`;
    if (typeof logNetworkEntry === 'function') {
        logNetworkEntry({ type: 'redirect', method: 'GET', url: `${AUTHWEB_HUB_URL}/?${requestParam}&return=${returnUrl}` });
    }
    window.location.assign(`${AUTHWEB_HUB_URL}/?${requestParam}&return=${encodeURIComponent(returnUrl)}`);
}

/**
 * Reads a token AuthWeb just handed back in the URL fragment, if any —
 * caches it and scrubs the fragment from the visible URL either way.
 * @returns {{resourceKey:string, entry:Object}|{error:string}|null}
 */
function consumeIncomingToken() {
    if (!window.location.hash) return null;

    const hash = new URLSearchParams(window.location.hash.slice(1));
    history.replaceState({}, '', window.location.pathname + window.location.search);

    const error = hash.get('authweb_error');
    if (error) {
        console.error('[AuthWeb] Broker returned an error:', error);
        return { error };
    }

    const accessToken = hash.get('authweb_token');
    if (!accessToken) return null;

    const resourceKey = hash.get('resource');
    const entry = {
        accessToken,
        idToken: hash.get('idToken') || null,
        username: hash.get('username') || null,
        expiresOn: Number(hash.get('expiresOn')) || null,
    };
    cacheToken(resourceKey, entry);
    return { resourceKey, entry };
}

/**
 * Gets a token for the given resource: this tab's cache if still fresh,
 * otherwise sends the browser to AuthWeb (which never returns here — the
 * page navigates away) and returns null.
 * @returns {{accessToken:string, idToken:string|null, username:string|null}|null}
 */
function getTokenForResource(resourceKey, resumeTab) {
    const cached = getCachedToken(resourceKey);
    if (cached) return cached;
    goToAuthWebForToken(resourceKey, resumeTab);
    return null;
}

function signOutOfSpoke() {
    sessionStorage.removeItem(SPOKE_TOKEN_KEY);
    sessionStorage.removeItem(SPOKE_PENDING_TAB_KEY);
    broadcastSpokeSignedOut(); // tells any sibling tab of THIS spoke app before navigating away — see below
    // Routes through AuthWeb's own sign-out, which reaches the Entra ID
    // session cookie itself (not just this tab's copy) — same "End Shift"
    // guarantee as before, now centralized in the one place that actually
    // holds the session. Always lands back on AuthWeb's own page, not here —
    // that keeps postLogoutRedirectUri registered to ONE origin (AuthWeb's),
    // not every spoke's.
    window.location.assign(`${AUTHWEB_HUB_URL}/?signout=1`);
}

/**
 * Live cross-tab signal for THIS spoke app's own sibling tabs (see the big
 * comment at the top of this file for why sessionStorage alone can't do
 * this, and why this can't reach a DIFFERENT spoke's tab either). Since
 * SPOKE_TOKEN_KEY is a plain key this code fully controls — unlike AuthWeb
 * Hub's MSAL-internal cache — a sibling tab can just clear its own copy
 * directly on receipt, no need to redo any actual sign-out action.
 */
const spokeChannel = 'BroadcastChannel' in window ? new BroadcastChannel('blueflames-spoke') : null;

function broadcastSpokeSignedOut() {
    if (spokeChannel) spokeChannel.postMessage({ type: 'signed-out' });
}

if (spokeChannel) {
    spokeChannel.onmessage = (event) => {
        if (event.data && event.data.type === 'signed-out') {
            sessionStorage.removeItem(SPOKE_TOKEN_KEY);
            if (typeof handleCrossTabTokenChange === 'function') {
                handleCrossTabTokenChange();
            }
        }
    };
}
