/**
 * switcher.js — spoke-app UI wiring (identical across every BlueFlames UI app).
 *
 * This app only ever calls its own API (via THIS_APP.resourceKey, apps.js)
 * and Microsoft Graph — it doesn't know or care about sibling apps beyond
 * the cross-app links ui-identity.js renders. It uses AuthWeb's shared MSAL
 * client ID (authConfig.js) with its own MSAL instance (authRedirect.js), so
 * signing in here shares an Entra ID session with siblings without sharing
 * any browser storage — see authRedirect.js's attemptCrossAppSso() for how
 * that plays out on page load, and ui.js's welcomeUser() for how signing in
 * (including silently, via cross-app SSO) auto-fetches this app's own data.
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

/**
 * Acquires a token for the given resource via AuthWeb, calls its API, and
 * renders both into that resource's tab.
 * @param {string} app - 'own' or 'graph' (key into APPS, also the data-app value)
 */
async function callBlueFlamesApp(app) {
    const { resourceKey, label } = APPS[app];
    const seq = ++requestSeq;
    renderLoading(app, label);

    try {
        const resource = protectedResources[resourceKey];
        const token = await getTokenForApp(resource.scopes);
        const payload = inspectToken(token, `${label} Token`);
        const data = await callApi('GET', resource.endpoint, token);

        if (seq !== requestSeq) return; // a newer click already superseded this one
        renderSuccess(app, label, data, payload, token);
    } catch (error) {
        console.error(error);
        if (seq !== requestSeq) return;
        renderError(app, label, error);
    }
}

document.getElementById('own-btn').addEventListener('click', () => callBlueFlamesApp('own'));
document.getElementById('graph-btn').addEventListener('click', () => callBlueFlamesApp('graph'));

// Reveal the panel immediately if the user is already signed in on page load.
if (myMSALObj.getAllAccounts().length > 0) {
    showBlueFlamesSection();
}
