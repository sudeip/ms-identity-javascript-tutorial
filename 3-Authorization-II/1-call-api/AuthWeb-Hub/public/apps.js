/**
 * apps.js — BlueFlames app registry.
 *
 * This file is IDENTICAL across every BlueFlames app — AuthWeb-Hub,
 * Titan-UI, Phoenix-UI, Falcon-UI, ... — copy it verbatim when you add a new
 * spoke. To rename an app, change its icon/color, or move its API to a new
 * port, edit the one entry below; every app picks up the change immediately
 * since they all load the same file. Nothing else in the codebase hardcodes
 * a display name/icon/color/port/endpoint.
 *
 * `id` is the small number (1, 2, 3, ...) a spoke passes to AuthWeb Hub as
 * `?app=<id>` when it needs a token — AuthWeb translates that id into the
 * right scope via `resourceKey` (see AuthWeb-Hub/public/authConfig.js, the
 * ONLY place actual Entra ID client IDs/scopes live now). `apiEndpoint` is
 * just where that app's own backend lives — ordinary application config, not
 * an Entra secret, so it's fine for every app to know every other app's here.
 *
 * "Which app am I" (spokes only) is inferred from window.location.port
 * matching an entry below, not from anything per-app-edited in this file —
 * that's what lets every spoke share this exact same file unmodified.
 */
const BLUEFLAMES_APPS = {
    titan: {
        id: 1,
        name: 'Titan',
        icon: '🛡️',
        colorStart: '#1864ab',
        colorEnd: '#0c8599',
        port: 3000,
        // Port 6060, not 6000 - Chrome/Firefox block 6000 outright (old X11 port).
        apiEndpoint: 'http://localhost:6060/api/data',
        resourceKey: 'titanApp',
    },
    phoenix: {
        id: 2,
        name: 'Phoenix',
        icon: '🔥',
        colorStart: '#d9480f',
        colorEnd: '#f76707',
        port: 4000,
        // Port 5050, not 5000 - macOS AirPlay Receiver squats on 5000 by default.
        apiEndpoint: 'http://localhost:5050/api/data',
        resourceKey: 'phoenixApp',
    },
    falcon: {
        id: 3,
        name: 'Falcon',
        icon: '🦅',
        colorStart: '#5f3dc4',
        colorEnd: '#7048e8',
        port: 4500,
        apiEndpoint: 'http://localhost:7070/api/data',
        resourceKey: 'falconApp',
    },
};

const BLUEFLAMES_APPS_BY_ID = Object.fromEntries(Object.values(BLUEFLAMES_APPS).map((app) => [app.id, app]));

// Spokes only: which of the apps above is "this app". AuthWeb-Hub isn't in
// the registry (it isn't a spoke), so this resolves to undefined there —
// AuthWeb's own scripts never reference THIS_APP.
const THIS_APP_KEY = Object.keys(BLUEFLAMES_APPS).find((key) => String(BLUEFLAMES_APPS[key].port) === window.location.port);

const THIS_APP = THIS_APP_KEY ? BLUEFLAMES_APPS[THIS_APP_KEY] : null;
const OTHER_APPS = THIS_APP_KEY
    ? Object.keys(BLUEFLAMES_APPS)
          .filter((key) => key !== THIS_APP_KEY)
          .map((key) => BLUEFLAMES_APPS[key])
    : Object.values(BLUEFLAMES_APPS);

// exporting for jest
if (typeof exports !== 'undefined') {
    module.exports = { BLUEFLAMES_APPS, BLUEFLAMES_APPS_BY_ID, THIS_APP_KEY, THIS_APP, OTHER_APPS };
}
