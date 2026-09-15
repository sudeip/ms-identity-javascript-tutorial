/**
 * apps.js — BlueFlames spoke-app registry.
 *
 * This file is IDENTICAL across every UI app (Titan-UI, Phoenix-UI,
 * Falcon-UI, ...) — copy it verbatim when you add a new one. To rename an
 * app, change its icon, or recolor its banner, edit the one entry below;
 * every app picks up the change immediately since they all load the same
 * file. Nothing else in the codebase hardcodes a display name/icon/color.
 *
 * "Which app am I" is inferred from window.location.port matching an entry
 * below, not from anything per-app-edited in this file — that's what lets
 * every app share this exact same file unmodified.
 *
 * resourceKey must match a key in authConfig.js's protectedResources for
 * that entry's own API (each app's authConfig.js still holds the real
 * per-app Entra data — client ID/scope/API endpoint — since that's tied to
 * an actual app registration, not something to genericize away).
 */
const BLUEFLAMES_APPS = {
    titan: {
        name: 'Titan',
        icon: '🛡️',
        colorStart: '#1864ab',
        colorEnd: '#0c8599',
        port: 3000,
        resourceKey: 'titanApp',
    },
    phoenix: {
        name: 'Phoenix',
        icon: '🔥',
        colorStart: '#d9480f',
        colorEnd: '#f76707',
        port: 4000,
        resourceKey: 'phoenixApp',
    },
    falcon: {
        name: 'Falcon',
        icon: '🦅',
        colorStart: '#5f3dc4',
        colorEnd: '#7048e8',
        port: 4500,
        resourceKey: 'falconApp',
    },
};

const THIS_APP_KEY =
    Object.keys(BLUEFLAMES_APPS).find((key) => String(BLUEFLAMES_APPS[key].port) === window.location.port) ||
    Object.keys(BLUEFLAMES_APPS)[0];

const THIS_APP = BLUEFLAMES_APPS[THIS_APP_KEY];
const OTHER_APPS = Object.keys(BLUEFLAMES_APPS)
    .filter((key) => key !== THIS_APP_KEY)
    .map((key) => BLUEFLAMES_APPS[key]);

// exporting for jest
if (typeof exports !== 'undefined') {
    module.exports = { BLUEFLAMES_APPS, THIS_APP_KEY, THIS_APP, OTHER_APPS };
}
