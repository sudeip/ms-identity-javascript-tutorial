/**
 * networkLog.js — records the actual HTTP requests this app and MSAL make,
 * so "what was sent to Entra ID" can be shown live without opening DevTools'
 * Network tab.
 *
 * Two capture mechanisms, because MSAL's authorize/logout calls are full
 * browser navigations, not fetch() calls, while everything else is:
 *
 *   1. A custom MSAL NavigationClient (installed in authRedirect.js /
 *      authPopup.js via setNavigationClient) captures the exact authorize/
 *      logout URL MSAL builds — including client_id, scope, redirect_uri,
 *      response_type, code_challenge, state, logout_hint, etc. — then
 *      performs the identical default navigation MSAL would have anyway.
 *
 *   2. A global fetch() wrapper captures everything else: MSAL's own
 *      token-endpoint POST and OIDC-metadata/JWKS GETs (MSAL uses fetch()
 *      internally by default — see @azure/msal-browser's FetchClient.js —
 *      so this catches those with no MSAL-specific hook needed), plus this
 *      app's own calls to the Phoenix/Titan/Graph APIs (fetch.js's callApi()).
 *
 * Entries persist in sessionStorage so the log survives the full-page
 * redirect round-trip for sign-in/out — otherwise the entry logged right
 * before navigating away to Entra ID would be lost when the page reloads.
 *
 * Demo-only note: this deliberately shows full request/response bodies,
 * including tokens — appropriate for a live demo (this app already shows
 * raw tokens elsewhere), not something you'd ship in a real product.
 */

const NETWORK_LOG_STORAGE_KEY = 'blueflames.networkLog';
const NETWORK_LOG_MAX_ENTRIES = 50;

function loadNetworkLog() {
    try {
        return JSON.parse(sessionStorage.getItem(NETWORK_LOG_STORAGE_KEY)) || [];
    } catch (error) {
        return [];
    }
}

function saveNetworkLog(entries) {
    try {
        sessionStorage.setItem(NETWORK_LOG_STORAGE_KEY, JSON.stringify(entries));
    } catch (error) {
        // best effort only — if storage is full/blocked, the log still renders for this page view
    }
}

let networkLogEntries = loadNetworkLog();

function labelForUrl(url) {
    // Match against the parsed pathname/hostname, not the raw URL string —
    // MSAL's discovery/instance URL embeds the authorize endpoint as an
    // (unencoded) query VALUE, e.g. "...instance?...&authorization_endpoint=
    // https://.../oauth2/v2.0/authorize" — a plain url.includes(...) check
    // would wrongly match that substring and mislabel this as "Authorize".
    let pathname = url;
    let hostname = '';
    try {
        const parsed = new URL(url);
        pathname = parsed.pathname;
        hostname = parsed.hostname;
    } catch (error) {
        // not an absolute URL we can parse — fall back to matching the raw string
    }

    if (pathname.endsWith('/oauth2/v2.0/authorize')) return 'Entra ID — Authorize (redirect)';
    if (pathname.endsWith('/oauth2/v2.0/logout')) return 'Entra ID — Logout (redirect)';
    if (pathname.endsWith('/oauth2/v2.0/token')) return 'Entra ID — Token endpoint';
    if (pathname.includes('/common/discovery/instance')) return 'Entra ID — instance discovery (authority validation)';
    if (pathname.endsWith('/.well-known/openid-configuration')) return 'Entra ID — OIDC discovery';
    if (pathname.endsWith('/discovery/v2.0/keys')) return 'Entra ID — JWKS (signing keys)';
    if (hostname === 'graph.microsoft.com') return 'Microsoft Graph';
    if (typeof protectedResources !== 'undefined') {
        for (const key of Object.keys(protectedResources)) {
            if (url.indexOf(protectedResources[key].endpoint) === 0) {
                return `${key} API`;
            }
        }
    }
    return 'Other request';
}

/**
 * @param {Object} entry - { type: 'redirect'|'fetch', method, url, requestBody?, status?, responseBody? }
 */
function logNetworkEntry(entry) {
    networkLogEntries.push({
        time: new Date().toLocaleTimeString(),
        label: labelForUrl(entry.url),
        ...entry,
    });
    if (networkLogEntries.length > NETWORK_LOG_MAX_ENTRIES) {
        networkLogEntries = networkLogEntries.slice(-NETWORK_LOG_MAX_ENTRIES);
    }
    saveNetworkLog(networkLogEntries);
    renderNetworkLog();
}

function clearNetworkLog() {
    networkLogEntries = [];
    saveNetworkLog(networkLogEntries);
    renderNetworkLog();
}

function formatQueryParams(url) {
    try {
        const parsed = new URL(url);
        const pairs = [...parsed.searchParams.entries()];
        if (!pairs.length) return null;
        return pairs.map(([k, v]) => `${k} = ${v}`).join('\n');
    } catch (error) {
        return null;
    }
}

function formatFormBody(body) {
    if (!body || typeof body !== 'string') return null;
    try {
        const pairs = [...new URLSearchParams(body).entries()];
        if (!pairs.length) return null;
        return pairs.map(([k, v]) => `${k} = ${v}`).join('\n');
    } catch (error) {
        return null;
    }
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function renderNetworkLog() {
    const container = document.getElementById('network-log-entries');
    if (!container) return;

    if (networkLogEntries.length === 0) {
        container.innerHTML = '<p class="text-muted mb-0">(no requests logged yet — sign in to see the authorize/token exchange)</p>';
        return;
    }

    container.innerHTML = networkLogEntries
        .slice()
        .reverse()
        .map((entry) => {
            const query = entry.type === 'redirect' ? formatQueryParams(entry.url) : null;
            const formBody = entry.type === 'fetch' ? formatFormBody(entry.requestBody) : null;
            const urlBase = entry.url.split('?')[0];
            const statusClass = entry.status === 'ERROR' ? 'status-ERROR' : `status-${Math.floor(entry.status / 100)}xx`;
            const responseText =
                entry.responseBody === undefined || entry.responseBody === null
                    ? null
                    : typeof entry.responseBody === 'string'
                        ? entry.responseBody
                        : JSON.stringify(entry.responseBody, null, 2);

            return `
                <details class="network-log-entry">
                    <summary>
                        <span class="network-log-time">${escapeHtml(entry.time)}</span>
                        <span class="network-log-method method-${escapeHtml(entry.method)}">${escapeHtml(entry.method)}</span>
                        <span class="network-log-label">${escapeHtml(entry.label)}</span>
                        ${entry.status !== undefined ? `<span class="network-log-status ${statusClass}">${escapeHtml(entry.status)}</span>` : ''}
                    </summary>
                    <div class="network-log-detail">
                        <div><strong>URL:</strong> <code>${escapeHtml(urlBase)}</code></div>
                        ${query ? `<div><strong>Query params sent:</strong><pre>${escapeHtml(query)}</pre></div>` : ''}
                        ${formBody ? `<div><strong>POST body sent:</strong><pre>${escapeHtml(formBody)}</pre></div>` : ''}
                        ${responseText ? `<div><strong>Response received:</strong><pre>${escapeHtml(responseText)}</pre></div>` : ''}
                    </div>
                </details>`;
        })
        .join('');
}

// ── Capture every fetch() call — MSAL's own (token endpoint, OIDC discovery,
// JWKS) and this app's own (Phoenix/Titan/Graph, via fetch.js's callApi()) ──
const originalFetch = window.fetch.bind(window);
window.fetch = async function (input, init) {
    const url = typeof input === 'string' ? input : input.url;
    const method = (init && init.method) || 'GET';
    const requestBody = init && init.body;

    try {
        const response = await originalFetch(input, init);
        let responseBody = null;
        try {
            responseBody = await response.clone().json();
        } catch (error) {
            // not JSON (or empty body) — fine, just don't capture a response body for this one
        }
        logNetworkEntry({ type: 'fetch', method, url, status: response.status, requestBody, responseBody });
        return response;
    } catch (error) {
        logNetworkEntry({ type: 'fetch', method, url, status: 'ERROR', requestBody, responseBody: error.message });
        throw error;
    }
};

renderNetworkLog();
