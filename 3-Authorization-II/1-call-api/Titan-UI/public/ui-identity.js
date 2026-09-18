/**
 * ui-identity.js — applies apps.js's registry to the page.
 *
 * This is what makes index.html/switcher.js truly identical across every
 * BlueFlames UI app: this file reads THIS_APP/OTHER_APPS (apps.js) and fills
 * in the name/icon/color/links, so nothing else needs per-app editing.
 */

const appGradient = `linear-gradient(90deg, ${THIS_APP.colorStart}, ${THIS_APP.colorEnd})`;

document.title = THIS_APP.name;
document.getElementById('app-navbar').style.background = appGradient;
document.getElementById('app-navbar-brand').textContent = `${THIS_APP.icon} ${THIS_APP.name}`;
document.getElementById('title-div').textContent = `${THIS_APP.name} — a BlueFlames spoke app, secured via AuthWeb Hub`;

document.getElementById('kiosk-note').innerHTML =
    `${THIS_APP.name} is a genuinely independent app on its own origin (<code>localhost:${THIS_APP.port}</code>) — ` +
    `it holds no Entra ID configuration of its own at all, only its small app id (<code>${THIS_APP.id}</code>). Every ` +
    `token comes from <a href="${AUTHWEB_HUB_URL}/" target="_blank">AuthWeb Hub</a>, which owns the one app ` +
    `registration and every BlueFlames app's scope. This app's own copy of a token lives in <code>localStorage</code>, ` +
    `shared by every ${THIS_APP.name} tab open on this device — sign out in one and the others notice immediately, ` +
    `no reload needed. <strong>Sign-out (End Shift)</strong> routes through AuthWeb to clear its session and sign out of ` +
    `Entra ID itself, not just this tab's copy — but only an explicit sign-out (or the idle timeout below) does that; ` +
    `just closing a tab leaves the session usable from a fresh one, so don't skip End Shift on a shared tablet.`;

// One "Open <App> App" link per sibling, inserted before the Sign-in button.
const crossAppLinksEl = document.getElementById('cross-app-links');
OTHER_APPS.forEach((app) => {
    const link = document.createElement('a');
    link.href = `http://localhost:${app.port}/`;
    link.className = 'btn btn-outline-light me-2';
    link.textContent = `${app.icon} Open ${app.name} App →`;
    crossAppLinksEl.appendChild(link);
});

// Default (pre-interaction) brand banner + button/tab labels for "this app".
document.getElementById('brand-icon').textContent = THIS_APP.icon;
document.getElementById('brand-name').textContent = THIS_APP.name;
document.getElementById('brand-banner').style.background = appGradient;
document.getElementById('own-tab-btn').textContent = `${THIS_APP.icon} ${THIS_APP.name}`;

// "Get <App> Data" button: same gradient as the banner, not Bootstrap's
// generic blue btn-primary — this is the app's own action, so it gets the
// app's own color. Border matches too so it doesn't show a mismatched blue
// outline on hover/focus. The Graph button stays a neutral outline, since
// it represents a different identity (Microsoft Graph), not "this app".
const ownBtn = document.getElementById('own-btn');
ownBtn.textContent = `${THIS_APP.icon} Get ${THIS_APP.name} Data`;
ownBtn.style.background = appGradient;
ownBtn.style.borderColor = THIS_APP.colorStart;
