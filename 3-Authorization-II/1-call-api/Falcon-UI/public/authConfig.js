/**
 * authConfig.js — a spoke app's ENTIRE Entra ID footprint: none. Just where
 * to find AuthWeb Hub, which holds the real client ID/tenant/scopes for
 * every BlueFlames app (see AuthWeb-Hub/public/authConfig.js). This file is
 * identical across every spoke (Titan-UI, Phoenix-UI, Falcon-UI, ...).
 */
const AUTHWEB_HUB_URL = 'http://localhost:3900';

// exporting for jest
if (typeof exports !== 'undefined') {
    module.exports = { AUTHWEB_HUB_URL };
}
