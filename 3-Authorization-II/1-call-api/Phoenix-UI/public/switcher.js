/**
 * switcher.js — spoke-app UI wiring (identical across every BlueFlames UI app).
 *
 * This app only ever calls its own API (via THIS_APP.resourceKey, apps.js)
 * and Microsoft Graph — it doesn't know or care about sibling apps beyond
 * the cross-app links ui-identity.js renders. It holds no Entra ID
 * configuration of its own at all: every token comes from AuthWeb Hub via
 * spoke.js's getTokenForResource(), which is a full-page round trip to
 * AuthWeb, not an in-page async call — see the page-load bootstrap at the
 * bottom of this file for how a click or a fresh page load resumes correctly
 * once that round trip lands back here.
 */

const APPS = {
    own: { resourceKey: THIS_APP.resourceKey, label: `${THIS_APP.name} App`, tabBtnId: 'own-tab-btn' },
    graph: { resourceKey: 'graphMe', label: 'Microsoft Graph (me)', tabBtnId: 'graph-tab-btn' },
};

const blueFlamesSection = document.getElementById('blueflames-section');
const brandBanner = document.getElementById('brand-banner');
const brandIcon = document.getElementById('brand-icon');
const brandName = document.getElementById('brand-name');
const appButtons = document.querySelectorAll('.app-btn');
const activeAppLabel = document.getElementById('active-app-label');

// Guards against a slow response rendering after a newer one has already
// landed (e.g. clicking Graph right after this app's own button, before the
// first resolves).
let requestSeq = 0;

function showBlueFlamesSection() {
    blueFlamesSection.classList.remove('d-none');
}

function setBrand(kind) {
    if (kind === 'graph') {
        brandBanner.style.background = 'linear-gradient(90deg, #0ca678, #12b886)';
        brandIcon.textContent = '🧑‍💼';
        brandName.textContent = 'Microsoft Graph';
    } else if (kind === 'error') {
        brandBanner.style.background = 'linear-gradient(90deg, #c92a2a, #e03131)';
        brandIcon.textContent = '⚠️';
        brandName.textContent = 'Error';
    } else {
        // 'own' — this app's own brand, from apps.js
        brandBanner.style.background = `linear-gradient(90deg, ${THIS_APP.colorStart}, ${THIS_APP.colorEnd})`;
        brandIcon.textContent = THIS_APP.icon;
        brandName.textContent = THIS_APP.name;
    }
}

function setActiveButton(app) {
    appButtons.forEach((btn) => {
        btn.classList.toggle('active-app-btn', btn.dataset.app === app);
    });
}

/** Switches the visible tab to the given app, using Bootstrap's Tab component. */
function activateTab(app) {
    const tabBtn = document.getElementById(APPS[app].tabBtnId);
    if (tabBtn && window.bootstrap) {
        bootstrap.Tab.getOrCreateInstance(tabBtn).show();
    }
}

/**
 * Renders an access token's claims into a container as a table, reusing the
 * same createClaimsTable() helper (claimUtils.js) used for the ID token
 * table below, so both look consistent. Microsoft Graph issues opaque
 * (non-JWT) tokens to SPA apps, so payload is null there — show a note
 * instead of a table in that case.
 */
function renderClaimsInto(containerEl, payload) {
    if (!payload) {
        containerEl.classList.add('text-muted');
        containerEl.textContent = '(Opaque access token — Microsoft Graph does not issue decodable JWTs to SPA apps.)';
        return;
    }

    containerEl.classList.remove('text-muted');
    const claims = createClaimsTable(payload);
    const rows = Object.keys(claims)
        .map((key) => {
            const [claim, value, description] = claims[key];
            return `<tr><td>${claim}</td><td>${value}</td><td>${description}</td></tr>`;
        })
        .join('');
    containerEl.innerHTML =
        `<table class="table table-striped table-sm claims-table">
            <thead><tr><th>Claim Type</th><th>Value</th><th>Description</th></tr></thead>
            <tbody>${rows}</tbody>
        </table>`;
}

function renderLoading(app, label) {
    setBrand(app);
    setActiveButton(app);
    activateTab(app);
    activeAppLabel.textContent = `Acquiring token for ${label}…`;

    const responseEl = document.getElementById(`${app}-response`);
    responseEl.classList.remove('is-error');
    responseEl.textContent = 'Loading…';

    const claimsEl = document.getElementById(`${app}-claims`);
    claimsEl.classList.add('text-muted');
    claimsEl.textContent = 'Loading…';

    document.getElementById(`${app}-token-raw`).textContent = 'Loading…';
}

function renderSuccess(app, label, data, tokenPayload, rawToken) {
    setBrand(app);
    activeAppLabel.textContent = `Active: ${label}`;

    const responseEl = document.getElementById(`${app}-response`);
    responseEl.classList.remove('is-error');
    responseEl.textContent = JSON.stringify(data, null, 2);

    renderClaimsInto(document.getElementById(`${app}-claims`), tokenPayload);

    document.getElementById(`${app}-token-raw`).textContent = formatJwt(rawToken);
}

function renderError(app, label, error) {
    setBrand('error');
    activeAppLabel.textContent = `${label} — failed`;

    const responseEl = document.getElementById(`${app}-response`);
    responseEl.classList.add('is-error');
    responseEl.textContent = `Error: ${error.message}`;

    const claimsEl = document.getElementById(`${app}-claims`);
    claimsEl.classList.add('text-muted');
    claimsEl.textContent = '(no token acquired)';

    document.getElementById(`${app}-token-raw`).textContent = '(no token acquired)';
}

/** Where a resource's data actually lives — the one bit of non-Entra "config" a spoke still keeps locally. */
function endpointForResource(resourceKey) {
    return resourceKey === 'graphMe' ? 'https://graph.microsoft.com/v1.0/me' : THIS_APP.apiEndpoint;
}

/**
 * Gets a token for the given resource (from this tab's short-lived cache, or
 * by sending the browser to AuthWeb Hub — see spoke.js) and, once one is in
 * hand, calls that resource's API and renders both into its tab.
 * @param {string} app - 'own' or 'graph' (key into APPS, also the data-app value)
 */
async function callBlueFlamesApp(app) {
    const { resourceKey, label } = APPS[app];
    const seq = ++requestSeq;
    renderLoading(app, label);

    const tokenEntry = getTokenForResource(resourceKey, app);
    if (!tokenEntry) return; // no cached token — spoke.js just sent the browser to AuthWeb Hub; this page is unloading

    try {
        const payload = inspectToken(tokenEntry.accessToken, `${label} Token`);
        const data = await callApi('GET', endpointForResource(resourceKey), tokenEntry.accessToken);

        if (seq !== requestSeq) return; // a newer click already superseded this one
        renderSuccess(app, label, data, payload, tokenEntry.accessToken);
    } catch (error) {
        console.error(error);
        if (seq !== requestSeq) return;
        renderError(app, label, error);
    }
}

document.getElementById('own-btn').addEventListener('click', () => callBlueFlamesApp('own'));
document.getElementById('graph-btn').addEventListener('click', () => callBlueFlamesApp('graph'));

/**
 * Page-load bootstrap — resumes from a round trip to AuthWeb Hub, or kicks
 * one off. Every spoke page load funnels through here:
 *
 *   1. A fresh token just arrived in the URL fragment (#authweb_token=...) —
 *      cache it, sign the UI in, and render whichever tab ('own'/'graph')
 *      was pending when this tab sent the browser to AuthWeb.
 *   2. No incoming fragment, but this tab already has a still-fresh cached
 *      token — sign the UI in and refresh the 'own' tab's data.
 *   3. Neither — genuinely nothing cached yet. Send the browser to AuthWeb
 *      for THIS_APP's own resource; if AuthWeb has a live session this
 *      bounces straight back with no visible interaction beyond a brief
 *      page flash (mirrors the old per-app ssoSilent()/prompt:'none' cross-
 *      app SSO attempt, just delegated to AuthWeb now instead of Entra ID
 *      directly). If nobody has ever signed in, AuthWeb shows its own
 *      Sign-in button instead of forcing a login prompt here.
 */
(function initSpokeOnLoad() {
    const incoming = consumeIncomingToken();

    if (incoming && incoming.error) {
        renderError('own', 'Sign-in', new Error(incoming.error));
        showBlueFlamesSection();
        return;
    }

    if (incoming) {
        const resumeTab = sessionStorage.getItem(SPOKE_PENDING_TAB_KEY) || 'own';
        sessionStorage.removeItem(SPOKE_PENDING_TAB_KEY);

        welcomeUser(incoming.entry.username || 'BlueFlames user');
        const decodedIdToken = incoming.entry.idToken ? decodeJwtParts(incoming.entry.idToken) : null;
        updateTable(decodedIdToken ? decodedIdToken.payload : {}, incoming.entry.idToken);
        callBlueFlamesApp(resumeTab);
        return;
    }

    const cachedOwn = getCachedToken(THIS_APP.resourceKey);
    if (cachedOwn) {
        welcomeUser(cachedOwn.username || 'BlueFlames user');
        const decodedIdToken = cachedOwn.idToken ? decodeJwtParts(cachedOwn.idToken) : null;
        updateTable(decodedIdToken ? decodedIdToken.payload : {}, cachedOwn.idToken);
        callBlueFlamesApp('own');
        return;
    }

    getTokenForResource(THIS_APP.resourceKey, 'own'); // navigates to AuthWeb Hub — nothing more happens on this page
})();
