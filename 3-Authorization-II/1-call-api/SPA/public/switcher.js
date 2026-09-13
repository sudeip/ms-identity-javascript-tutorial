/**
 * switcher.js — BlueFlames App Switcher demo UI wiring.
 *
 * Uses the shared AuthWeb library (authWeb.js) to switch between BlueFlames
 * apps (Phoenix, Titan) and Microsoft Graph, calling getTokenForApp(scopes)
 * for each. Each app has its own tab so a previous result stays visible
 * (and comparable) even after you've switched to a different app — clicking
 * a button fetches fresh and jumps to that app's tab; clicking a tab just
 * browses whatever was last fetched for it. The banner at the top re-skins
 * to each app's brand on switch.
 */

const BRANDS = {
    blueflames: { label: 'BlueFlames AuthWeb', icon: '🔷', cls: 'brand-blueflames' },
    phoenix: { label: 'Phoenix', icon: '🔥', cls: 'brand-phoenix' },
    titan: { label: 'Titan', icon: '🛡️', cls: 'brand-titan' },
    graph: { label: 'Microsoft Graph', icon: '🧑‍💼', cls: 'brand-graph' },
    error: { label: 'Error', icon: '⚠️', cls: 'brand-error' },
};

const APPS = {
    phoenix: { resourceKey: 'phoenixApp', label: 'Phoenix App', tabBtnId: 'phoenix-tab-btn' },
    titan: { resourceKey: 'titanApp', label: 'Titan App', tabBtnId: 'titan-tab-btn' },
    graph: { resourceKey: 'graphMe', label: 'Microsoft Graph (me)', tabBtnId: 'graph-tab-btn' },
};

const blueFlamesSection = document.getElementById('blueflames-section');
const brandBanner = document.getElementById('brand-banner');
const brandIcon = document.getElementById('brand-icon');
const brandName = document.getElementById('brand-name');
const appButtons = document.querySelectorAll('.app-btn');
const activeAppLabel = document.getElementById('active-app-label');

// Guards against a slow response rendering after a newer one has already
// landed (e.g. clicking Phoenix then Titan before the first call resolves).
let requestSeq = 0;

function showBlueFlamesSection() {
    blueFlamesSection.classList.remove('d-none');
}

function setBrand(key) {
    const brand = BRANDS[key] || BRANDS.blueflames;
    brandBanner.className = `brand-banner ${brand.cls}`;
    brandIcon.textContent = brand.icon;
    brandName.textContent = brand.label;
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
        `<table class="table table-striped table-sm">
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

    document.getElementById(`${app}-token-raw`).textContent = rawToken;
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
 * Acquires a token for the given BlueFlames app via the shared AuthWeb
 * library, calls its API, and renders both into that app's tab.
 * @param {string} app - key into APPS/BRANDS (also the data-app value)
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

document.getElementById('phoenix-btn').addEventListener('click', () => callBlueFlamesApp('phoenix'));
document.getElementById('titan-btn').addEventListener('click', () => callBlueFlamesApp('titan'));
document.getElementById('graph-btn').addEventListener('click', () => callBlueFlamesApp('graph'));

// Reveal the switcher immediately if the user is already signed in on page load.
if (myMSALObj.getAllAccounts().length > 0) {
    showBlueFlamesSection();
}
