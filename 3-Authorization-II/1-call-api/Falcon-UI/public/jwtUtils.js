/**
 * jwtUtils.js — pure, client-side-only JWT decode/display helpers. No auth
 * logic lives here anymore (that used to be this file's other half, under
 * the name authWeb.js — it moved to AuthWeb Hub's actual broker.js/
 * authRedirect.js, which is now a REAL separate app rather than a shared
 * filename copied into every spoke). What's left is just formatting: turning
 * a raw token string into the "JWT format" / claims-table views this demo
 * shows. Identical across every spoke.
 */

/**
 * Decodes a base64url string (the encoding used in JWT segments, which uses
 * '-'/'_' instead of '+'/'/' and omits padding — plain atob() rejects it).
 * @param {string} base64Url
 * @returns {string}
 */
function base64UrlDecode(base64Url) {
    let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const paddingNeeded = base64.length % 4;
    if (paddingNeeded) {
        base64 += '='.repeat(4 - paddingNeeded);
    }
    return atob(base64);
}

/**
 * Decodes a JWT into its header and payload objects — the classic "JWT
 * format" view (as e.g. jwt.ms shows it), not signature-verified (this is a
 * client-side demo helper, not a security boundary). Works for any JWT, not
 * just BlueFlames access tokens — also used for the ID token.
 *
 * @param {string} token - raw token string
 * @returns {{header: Object, payload: Object}|null} null if it isn't a decodable JWT
 *   (e.g. Microsoft Graph's opaque access tokens for public client apps)
 */
function decodeJwtParts(token) {
    try {
        const [headerSegment, payloadSegment] = token.split('.');
        return {
            header: JSON.parse(base64UrlDecode(headerSegment)),
            payload: JSON.parse(base64UrlDecode(payloadSegment)),
        };
    } catch (error) {
        return null;
    }
}

/**
 * Renders a token's header + payload as the text shown in a "JWT format"
 * panel, or a note explaining why it can't be shown for an opaque token.
 * @param {string} token
 * @returns {string}
 */
function formatJwt(token) {
    const decoded = decodeJwtParts(token);
    if (!decoded) {
        return '(Opaque access token — not a decodable JWT; expected for Microsoft Graph on a public client app.)';
    }
    return `Header:\n${JSON.stringify(decoded.header, null, 2)}\n\nPayload:\n${JSON.stringify(decoded.payload, null, 2)}`;
}

/**
 * Decode and log the key claims from an access token.
 * Useful for the demo — shows aud (which app the token is for), scp, and user.
 *
 * Note: this only works for JWT-format access tokens, which is what Azure AD
 * issues for APIs you register yourself (Titan/Phoenix/Falcon). Microsoft
 * Graph issues OPAQUE access tokens to public client (SPA) apps by default —
 * those aren't decodable client-side, so this returns null for them instead
 * of throwing; see the caller for how that's handled in the UI.
 *
 * @param {string} token - raw access token
 * @param {string} label - label for console output
 * @returns {Object|null} decoded token payload, or null if it isn't a decodable JWT
 */
function inspectToken(token, label = 'Token') {
    const decoded = decodeJwtParts(token);
    if (!decoded) {
        console.group(`[AuthWeb] ${label}`);
        console.log('(opaque access token — not a decodable JWT; expected for Microsoft Graph on a public client app)');
        console.groupEnd();
        return null;
    }
    const payload = decoded.payload;
    console.group(`[AuthWeb] ${label}`);
    console.log('aud (audience):', payload.aud);
    console.log('scp (scopes):  ', payload.scp);
    console.log('upn (user):    ', payload.upn || payload.preferred_username);
    console.log('exp (expires): ', new Date(payload.exp * 1000).toLocaleTimeString());
    console.groupEnd();
    return payload;
}
