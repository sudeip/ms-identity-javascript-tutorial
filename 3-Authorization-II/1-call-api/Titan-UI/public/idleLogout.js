/**
 * idleLogout.js — Kiosk inactivity auto sign-out.
 *
 * Frontline workers on a shared tablet won't always remember to tap
 * "Sign-out (End Shift)" before walking away. This is defense-in-depth for
 * that: after a period of no interaction, show a warning, then automatically
 * run the real signOut() flow (authRedirect.js/authPopup.js) if nobody
 * responds. This DOES reach Entra ID's own session cookie, not just this
 * app, because it runs while the tab is still open and can complete the
 * redirect.
 *
 * This file used to also clear storage on 'pagehide' (tab closing /
 * navigating away) as a belt-and-braces safety net. That was removed: a
 * browser has no reliable way to distinguish "the tab is actually closing"
 * from "the page is just reloading or navigating to another in-app link" —
 * 'pagehide' fires for both — so it was wiping the session (looking like an
 * unwanted logout) on a plain refresh or even a same-page link click.
 * sessionStorage still clears itself on a *genuine* tab/browser close per
 * the browser's own spec, with no help needed from this file; what it can't
 * do is reach the Entra ID session cookie in that case — only an explicit
 * sign-out (this file's timeout, or the button) can do that.
 *
 * Total idle timeout is IDLE_WARNING_MS + IDLE_LOGOUT_GRACE_MS = 15 minutes
 * (14 minutes before the warning appears, then a 1-minute countdown to
 * sign-out) — a realistic value for actual use, not just a live demo. Tune
 * both to whatever balances a typical task's length against how long you're
 * willing to leave a session exposed.
 */

const IDLE_WARNING_MS = 14 * 60 * 1000; // inactivity before showing the warning
const IDLE_LOGOUT_GRACE_MS = 60 * 1000; // time the warning stays up before auto sign-out

let idleWarningTimer = null;
let idleLogoutTimer = null;
let idleCountdownInterval = null;

function createIdleWarningBanner() {
    const banner = document.createElement('div');
    banner.id = 'idle-warning-banner';
    banner.className = 'idle-warning-banner d-none';
    banner.innerHTML = `⏱️ You'll be signed out due to inactivity in <strong id="idle-countdown"></strong>s — tap anywhere to stay signed in.`;
    document.body.appendChild(banner);
    return banner;
}

const idleWarningBanner = createIdleWarningBanner();
const idleCountdownEl = document.getElementById('idle-countdown');

function hideIdleWarning() {
    idleWarningBanner.classList.add('d-none');
    if (idleCountdownInterval) {
        clearInterval(idleCountdownInterval);
        idleCountdownInterval = null;
    }
}

function showIdleWarning() {
    let secondsLeft = Math.round(IDLE_LOGOUT_GRACE_MS / 1000);
    idleCountdownEl.textContent = secondsLeft;
    idleWarningBanner.classList.remove('d-none');

    idleCountdownInterval = setInterval(() => {
        secondsLeft -= 1;
        idleCountdownEl.textContent = Math.max(secondsLeft, 0);
        if (secondsLeft <= 0) {
            clearInterval(idleCountdownInterval);
            idleCountdownInterval = null;
        }
    }, 1000);
}

function triggerIdleLogout() {
    hideIdleWarning();
    // Only actually sign out if someone is currently signed in — no point
    // "logging out" a device already sitting idle at the sign-in screen.
    if (myMSALObj.getAllAccounts().length > 0) {
        console.warn('[Kiosk] Signing out automatically due to inactivity.');
        signOut();
    }
}

function resetIdleTimers() {
    hideIdleWarning();
    clearTimeout(idleWarningTimer);
    clearTimeout(idleLogoutTimer);

    if (myMSALObj.getAllAccounts().length === 0) return;

    idleWarningTimer = setTimeout(() => {
        showIdleWarning();
        idleLogoutTimer = setTimeout(triggerIdleLogout, IDLE_LOGOUT_GRACE_MS);
    }, IDLE_WARNING_MS);
}

['mousedown', 'mousemove', 'keydown', 'touchstart', 'scroll', 'click'].forEach((eventName) => {
    document.addEventListener(eventName, resetIdleTimers, { passive: true });
});

// Arms the timers if a session was already restored on page load (e.g. a
// popup-based sign-in that happened moments ago). welcomeUser() (ui.js) also
// calls this right after a fresh sign-in, so the timer doesn't depend on
// incidental mouse activity to start counting.
resetIdleTimers();
