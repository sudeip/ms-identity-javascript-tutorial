# BlueFlames — MSAL Multi-App SSO Demo Plan
## Entra ID Migration · Shared AuthWeb Library · VS Code Walkthrough

> **Goal:** Demo how a single shared JavaScript auth library (AuthWeb) enables
> seamless token switching across multiple BlueFlames apps using one Entra ID
> app registration, `extraScopesToConsent`, and `acquireTokenSilent`.

---

## Architecture

```
┌──────────────────────────────────────────────────────────┐
│               AuthWeb  (shared JS library)               │
│   msalInstance · acquireTokenSilent · scope routing      │
│           One Entra app registration: AuthWeb            │
└────────────────┬──────────────────────────┬─────────────┘
                 │ extraScopesToConsent      │ silent token
                 │ at login (all apps)       │ per app switch
                 ▼                           ▼
       ┌─────────────────┐         ┌─────────────────┐
       │  App: Phoenix   │         │  App: Titan     │
       │  BFF API reg    │         │  BFF API reg    │
       │  (App Reg 2)    │         │  (App Reg 3)    │
       └─────────────────┘         └─────────────────┘
```

**Three Entra ID registrations:**
| Registration | Type | Purpose |
|---|---|---|
| AuthWeb | SPA | Shared MSAL client. Holds all delegated permissions. |
| Phoenix API | Resource (API) | Exposes `api://phoenix/access_as_user` scope |
| Titan API | Resource (API) | Exposes `api://titan/access_as_user` scope |

**Key flow:**
1. User signs in once via AuthWeb → consents to Graph + Phoenix + Titan in one shot
2. User opens Phoenix → `acquireTokenSilent` returns token with `aud=api://phoenix`
3. User switches to Titan → `acquireTokenSilent` returns token with `aud=api://titan`
4. Zero additional prompts — all from MSAL cache or hidden iframe refresh

---

## Prerequisites

- [ ] Node.js 18+ installed → `node -v`
- [ ] VS Code installed
- [ ] Git installed → `git -v`
- [ ] Microsoft 365 / Entra ID tenant (work/school account with admin or app registration rights)
- [ ] VS Code extensions installed (see below)

### Install VS Code extensions
Open VS Code → Extensions (Ctrl+Shift+X) → install:
```
ms-vscode.vscode-js-debug
humao.rest-client
esbenp.prettier-vscode
ms-vscode.azure-account
```

---

## Step 1 — Fork and Clone the Base Repo

### 1a. Fork on GitHub
1. Go to: https://github.com/Azure-Samples/ms-identity-javascript-tutorial
2. Click **Fork** (top right)
3. Fork to your own GitHub account

### 1b. Clone your fork
```bash
git clone https://github.com/<your-username>/ms-identity-javascript-tutorial.git
cd ms-identity-javascript-tutorial
```

### 1c. Open in VS Code
```bash
code .
```

### 1d. Navigate to the base sample
We are building on Chapter 3, Sample 1 — SPA calling a protected API:
```bash
cd 3-Authorization-II/1-call-api
```

Open this folder as your workspace root:
- File → Open Folder → select `3-Authorization-II/1-call-api`

### 1e. Install dependencies
```bash
# In VS Code terminal (Ctrl+`)
cd SPA && npm install
cd ../API && npm install
cd ..
```

### 1f. Verify the base sample runs
```bash
# Terminal 1 — start the API
cd API && npm start

# Terminal 2 — start the SPA
cd SPA && npm start
```
Open http://localhost:3000 — you should see the sample login page.
Stop both servers (Ctrl+C) when confirmed working.

---

## Step 2 — Register Apps in Entra ID

Go to: https://entra.microsoft.com → App registrations

### 2a. Register: AuthWeb (the SPA)

1. New registration
   - Name: `BlueFlames-AuthWeb`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: Platform = **Single-page application** → `http://localhost:3000`
2. Click **Register**
3. Copy and save the **Application (client) ID** — you will need this
4. Copy and save the **Directory (tenant) ID**
5. Go to **Authentication** → add second redirect URI (SPA): `http://localhost:3000/redirect`
6. Go to **API permissions** → Add permission:
   - Microsoft Graph → Delegated → `User.Read` → Add
   - (Phoenix and Titan will be added after steps 2b and 2c)

### 2b. Register: Phoenix API

1. New registration
   - Name: `BlueFlames-Phoenix-API`
   - Supported account types: **This directory only**
   - No redirect URI needed
2. Click **Register**
3. Copy the **Application (client) ID** → save as `PHOENIX_CLIENT_ID`
4. Go to **Expose an API**
   - Application ID URI → Set → accept default `api://<phoenix-client-id>`
   - Add a scope:
     - Scope name: `access_as_user`
     - Who can consent: **Admins and users**
     - Admin consent display name: `Access Phoenix API as user`
     - Admin consent description: `Allows BlueFlames apps to call Phoenix API on behalf of the signed-in user`
     - State: **Enabled**
     - Click **Add scope**
5. Copy the full scope URI: `api://<phoenix-client-id>/access_as_user`

### 2c. Register: Titan API

Repeat step 2b with:
- Name: `BlueFlames-Titan-API`
- Scope name: `access_as_user`
- Admin consent display name: `Access Titan API as user`
- Copy the scope URI: `api://<titan-client-id>/access_as_user`

### 2d. Grant permissions from AuthWeb to Phoenix and Titan

Back in **BlueFlames-AuthWeb** app registration:
1. API permissions → **Add a permission**
2. My APIs → select **BlueFlames-Phoenix-API**
   - Delegated permissions → check `access_as_user` → Add
3. Add again → My APIs → select **BlueFlames-Titan-API**
   - Delegated permissions → check `access_as_user` → Add
4. Click **Grant admin consent for [your tenant]** → Yes

Your final permissions list in AuthWeb should show:
- Microsoft Graph / User.Read ✓
- BlueFlames-Phoenix-API / access_as_user ✓
- BlueFlames-Titan-API / access_as_user ✓

---

## Step 3 — Configure the SPA

### 3a. Update authConfig.js

Open `SPA/src/authConfig.js` and replace entirely:

```javascript
// authConfig.js — BlueFlames AuthWeb configuration

export const msalConfig = {
    auth: {
        clientId:    "PASTE_AUTHWEB_CLIENT_ID_HERE",
        authority:   "https://login.microsoftonline.com/PASTE_TENANT_ID_HERE",
        redirectUri: "http://localhost:3000",
    },
    cache: {
        cacheLocation:      "localStorage",  // persist across page refresh
        storeAuthStateInCookie: false,
    }
};

// First login — acquire consent for all BlueFlames apps up front
// so switching between apps is silent (no prompts)
export const loginRequest = {
    scopes: ["User.Read"],
    extraScopesToConsent: [
        "api://PASTE_PHOENIX_CLIENT_ID_HERE/access_as_user",
        "api://PASTE_TITAN_CLIENT_ID_HERE/access_as_user"
    ]
};

// Resource map — each BlueFlames app references its own entry
export const protectedResources = {
    graphMe: {
        endpoint: "https://graph.microsoft.com/v1.0/me",
        scopes:   ["User.Read"]
    },
    phoenixApp: {
        endpoint: "http://localhost:5000/api/data",
        scopes:   ["api://PASTE_PHOENIX_CLIENT_ID_HERE/access_as_user"]
    },
    titanApp: {
        endpoint: "http://localhost:6000/api/data",
        scopes:   ["api://PASTE_TITAN_CLIENT_ID_HERE/access_as_user"]
    }
};
```

### 3b. Create authWeb.js — the shared library

Create new file: `SPA/src/authWeb.js`

```javascript
/**
 * authWeb.js
 * Simulates the shared BlueFlames AuthWeb MSAL library.
 *
 * In production this would be a versioned npm package or
 * CDN-hosted JS file shared across all BlueFlames apps.
 * Here it is a local module to make the demo self-contained.
 */

import {
    PublicClientApplication,
    InteractionRequiredAuthError
} from "@azure/msal-browser";

import { msalConfig, loginRequest } from "./authConfig.js";

// ── Single MSAL instance shared across all BlueFlames apps ──────
const msalInstance = new PublicClientApplication(msalConfig);
await msalInstance.initialize();

// Handle redirect response if returning from /authorize
await msalInstance.handleRedirectPromise();

/**
 * Sign the user in.
 * Uses loginPopup with extraScopesToConsent so the user consents
 * to ALL BlueFlames app scopes in a single interaction.
 */
export async function signIn() {
    const result = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(result.account);
    return result.account;
}

/**
 * Sign the user out.
 */
export async function signOut() {
    await msalInstance.logoutPopup({
        account: msalInstance.getActiveAccount()
    });
}

/**
 * Get the active account (null if not signed in).
 */
export function getActiveAccount() {
    return msalInstance.getActiveAccount();
}

/**
 * Central token acquisition for any BlueFlames app.
 * Tries silent first → falls back to popup if interaction required.
 *
 * @param {string[]} scopes - API-specific scopes
 *   e.g. ["api://phoenix-client-id/access_as_user"]
 * @returns {Promise<string>} access token
 */
export async function getTokenForApp(scopes) {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error("No active account. Call signIn() first.");
    }

    const tokenRequest = { scopes, account };

    try {
        // 1. Try silent — returns from cache or uses refresh token
        const result = await msalInstance.acquireTokenSilent(tokenRequest);
        console.log(`[AuthWeb] Silent token acquired for scope: ${scopes}`);
        return result.accessToken;

    } catch (error) {

        if (error instanceof InteractionRequiredAuthError) {
            // 2. Silent failed (expired, needs MFA, consent revoked)
            //    Fall back to popup — user sees a brief interaction
            console.warn(`[AuthWeb] Silent failed, falling back to popup: ${error.message}`);
            const result = await msalInstance.acquireTokenPopup(tokenRequest);
            return result.accessToken;
        }

        throw error;
    }
}

/**
 * Decode and log the key claims from an access token.
 * Useful for demo/debugging — shows aud, scp, upn.
 *
 * @param {string} token  - raw JWT access token
 * @param {string} label  - label for console output
 */
export function inspectToken(token, label = "Token") {
    const payload = JSON.parse(atob(token.split(".")[1]));
    console.group(`[AuthWeb] ${label}`);
    console.log("aud (audience):", payload.aud);
    console.log("scp (scopes):  ", payload.scp);
    console.log("upn (user):    ", payload.upn || payload.preferred_username);
    console.log("exp (expires): ", new Date(payload.exp * 1000).toLocaleTimeString());
    console.groupEnd();
    return payload;
}

export { msalInstance };
```

### 3c. Update index.html — add app switcher UI

Open `SPA/src/index.html` and add inside `<body>`:

```html
<!-- BlueFlames App Switcher Demo -->
<div id="auth-section">
    <button id="sign-in-btn">Sign In (AuthWeb)</button>
    <button id="sign-out-btn" style="display:none">Sign Out</button>
    <span id="user-display"></span>
</div>

<div id="app-section" style="display:none">
    <h2>BlueFlames App Switcher</h2>
    <p>Switch between apps — watch the console for token audience changes.</p>

    <button id="phoenix-btn">Open Phoenix App</button>
    <button id="titan-btn">Open Titan App</button>
    <button id="graph-btn">Get My Profile (Graph)</button>

    <div id="active-app-label"></div>
    <pre id="api-response"></pre>
    <pre id="token-info"></pre>
</div>
```

### 3d. Create app.js — wire up the buttons

Create `SPA/src/app.js`:

```javascript
import {
    signIn,
    signOut,
    getActiveAccount,
    getTokenForApp,
    inspectToken
} from "./authWeb.js";

import { protectedResources } from "./authConfig.js";

// ── UI helpers ──────────────────────────────────────────────────
function showAppSection(account) {
    document.getElementById("sign-in-btn").style.display  = "none";
    document.getElementById("sign-out-btn").style.display = "inline";
    document.getElementById("user-display").textContent   = `Signed in: ${account.username}`;
    document.getElementById("app-section").style.display  = "block";
}

function showApiResponse(label, data, tokenPayload) {
    document.getElementById("active-app-label").textContent = `Active: ${label}`;
    document.getElementById("api-response").textContent     = JSON.stringify(data, null, 2);
    document.getElementById("token-info").textContent =
        `Token audience : ${tokenPayload.aud}\n` +
        `Token scopes   : ${tokenPayload.scp}\n` +
        `User           : ${tokenPayload.upn || tokenPayload.preferred_username}\n` +
        `Expires        : ${new Date(tokenPayload.exp * 1000).toLocaleTimeString()}`;
}

// ── Call a protected API with a bearer token ────────────────────
async function callApi(endpoint, token) {
    const response = await fetch(endpoint, {
        headers: { Authorization: `Bearer ${token}` }
    });
    if (!response.ok) throw new Error(`API error: ${response.status}`);
    return response.json();
}

// ── Button handlers ─────────────────────────────────────────────
document.getElementById("sign-in-btn").onclick = async () => {
    const account = await signIn();
    showAppSection(account);
};

document.getElementById("sign-out-btn").onclick = async () => {
    await signOut();
    location.reload();
};

// Phoenix app — acquires token for Phoenix API scope
document.getElementById("phoenix-btn").onclick = async () => {
    const token   = await getTokenForApp(protectedResources.phoenixApp.scopes);
    const payload = inspectToken(token, "Phoenix Token");
    const data    = await callApi(protectedResources.phoenixApp.endpoint, token);
    showApiResponse("Phoenix App", data, payload);
};

// Titan app — acquires token for Titan API scope
document.getElementById("titan-btn").onclick = async () => {
    const token   = await getTokenForApp(protectedResources.titanApp.scopes);
    const payload = inspectToken(token, "Titan Token");
    const data    = await callApi(protectedResources.titanApp.endpoint, token);
    showApiResponse("Titan App", data, payload);
};

// Graph — acquires token for Microsoft Graph
document.getElementById("graph-btn").onclick = async () => {
    const token   = await getTokenForApp(protectedResources.graphMe.scopes);
    const payload = inspectToken(token, "Graph Token");
    const data    = await callApi(protectedResources.graphMe.endpoint, token);
    showApiResponse("Microsoft Graph (me)", data, payload);
};

// ── On page load — restore session if user was already signed in ─
const account = getActiveAccount();
if (account) showAppSection(account);
```

---

## Step 4 — Set Up Two API Instances (Phoenix & Titan)

### 4a. Copy the base API folder

```bash
# From the 3-Authorization-II/1-call-api folder
cp -r API API-Phoenix
cp -r API API-Titan
```

### 4b. Configure Phoenix API

Open `API-Phoenix/authConfig.json`:
```json
{
    "credentials": {
        "tenantID":  "PASTE_TENANT_ID_HERE",
        "clientID":  "PASTE_PHOENIX_CLIENT_ID_HERE"
    },
    "settings": {
        "audience":  "api://PASTE_PHOENIX_CLIENT_ID_HERE",
        "issuer":    "https://login.microsoftonline.com/PASTE_TENANT_ID_HERE/v2.0"
    }
}
```

Open `API-Phoenix/app.js` and change the port:
```javascript
const PORT = process.env.PORT || 5000;

// Add a demo endpoint
app.get("/api/data", (req, res) => {
    res.json({
        app:     "Phoenix",
        message: "Hello from Phoenix API",
        user:    req.authInfo.preferred_username || req.authInfo.upn,
        time:    new Date().toISOString()
    });
});
```

### 4c. Configure Titan API

Same as 4b but:
- `API-Titan/authConfig.json` → use Titan client ID
- Port: `6000`
- Response: `app: "Titan"`, `message: "Hello from Titan API"`

### 4d. Enable CORS on both APIs

In both `API-Phoenix/app.js` and `API-Titan/app.js`, add:
```javascript
const cors = require("cors");
app.use(cors({ origin: "http://localhost:3000" }));
```

---

## Step 5 — VS Code Launch Configuration

Create `.vscode/launch.json` in your workspace root:

```json
{
    "version": "0.2.0",
    "compounds": [
        {
            "name": "BlueFlames — Full Stack",
            "configurations": ["SPA", "Phoenix API", "Titan API"],
            "stopAll": true
        }
    ],
    "configurations": [
        {
            "name": "SPA",
            "type": "node",
            "request": "launch",
            "program": "${workspaceFolder}/SPA/server.js",
            "env": { "PORT": "3000" },
            "console": "integratedTerminal"
        },
        {
            "name": "Phoenix API",
            "type": "node",
            "request": "launch",
            "program": "${workspaceFolder}/API-Phoenix/app.js",
            "env": { "PORT": "5000" },
            "console": "integratedTerminal"
        },
        {
            "name": "Titan API",
            "type": "node",
            "request": "launch",
            "program": "${workspaceFolder}/API-Titan/app.js",
            "env": { "PORT": "6000" },
            "console": "integratedTerminal"
        }
    ]
}
```

Press **F5** → select **BlueFlames — Full Stack** → all three servers start.

---

## Step 6 — Run and Verify

### 6a. Start everything
- Press F5 → BlueFlames — Full Stack
- Open http://localhost:3000

### 6b. Sign in
1. Click **Sign In (AuthWeb)**
2. Consent screen appears — should show permissions for Graph, Phoenix, and Titan
3. Sign in with your work account
4. Consent once → never asked again for these scopes

### 6c. Switch between apps
1. Click **Open Phoenix App**
   - Check VS Code console: `[AuthWeb] Silent token acquired`
   - Check `token-info` on page: `aud = api://phoenix-...`
2. Click **Open Titan App**
   - Check console: still silent — no popup
   - Check `token-info`: `aud = api://titan-...`  ← different audience, same user
3. Click **Get My Profile (Graph)**
   - Console: silent again
   - Response: your Entra ID profile from Graph

### 6d. Inspect tokens in jwt.ms
Copy any token from the browser's localStorage → paste into https://jwt.ms
- `aud` claim → confirms which API the token is for
- `scp` claim → shows `access_as_user`
- `oid` → same user across all three tokens

---

## Step 7 — Demo Talking Points

When presenting this to the BlueFlames teams:

| What you show | What it proves |
|---|---|
| Single sign-in → consent to 3 resources | `extraScopesToConsent` pre-consents all BlueFlames apps in one shot |
| Switch Phoenix → Titan, zero prompts | `acquireTokenSilent` + shared MSAL cache does the work |
| Different `aud` claim per app | Each app gets a correctly scoped token — not a shared credential |
| Console logs from `authWeb.js` | All token logic centralised — apps call `getTokenForApp(scopes)`, nothing else |
| No client secret in browser | PKCE + SPA registration → no secret needed |
| Network tab: no `/authorize` call on switch | Token served from cache — SSO is real, not simulated |

---

## Folder Structure When Done

```
3-Authorization-II/1-call-api/
├── .vscode/
│   └── launch.json          ← F5 starts everything
├── SPA/
│   ├── src/
│   │   ├── authConfig.js    ← MSAL config + scope map  (you edit)
│   │   ├── authWeb.js       ← shared auth library      (you create)
│   │   ├── app.js           ← app switcher UI logic    (you create)
│   │   └── index.html       ← updated with buttons     (you edit)
│   └── server.js            ← static file server (unchanged)
├── API-Phoenix/
│   ├── authConfig.json      ← Phoenix client ID + audience
│   └── app.js               ← port 5000 (you copy + edit)
└── API-Titan/
    ├── authConfig.json      ← Titan client ID + audience
    └── app.js               ← port 6000 (you copy + edit)
```

---

## Checklist Before Demo

- [ ] All 3 redirect URIs registered in AuthWeb: `http://localhost:3000` and `http://localhost:3000/redirect`
- [ ] Admin consent granted on Phoenix and Titan permissions in AuthWeb registration
- [ ] Phoenix and Titan client IDs pasted into all config files
- [ ] Tenant ID pasted into all config files
- [ ] CORS enabled on both API servers
- [ ] All three servers start without errors (F5)
- [ ] Sign-in works and consent screen shows all three resources
- [ ] Switching apps shows different `aud` claims, zero prompts
- [ ] Tokens visible in localStorage (DevTools → Application → Local Storage)

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| `AADSTS700054: response_type 'token' is not enabled` | App not registered as SPA | Change redirect URI type to SPA in Authentication blade |
| Consent screen doesn't show Phoenix/Titan | `extraScopesToConsent` not set, or permissions not granted | Check loginRequest + admin consent |
| `acquireTokenSilent` always falls back to popup | localStorage not set, or private browsing | Set `cacheLocation: "localStorage"` in msalConfig |
| API returns 401 | Token audience mismatch | Verify `audience` in API authConfig matches `api://<client-id>` |
| CORS error on API call | CORS not enabled on API | Add `app.use(cors({ origin: "http://localhost:3000" }))` |
| Phoenix and Titan return same token | Both scopes passed together | Pass only one app's scope per `getTokenForApp()` call |

---

*IAM Modernization Program · Platform Engineering · BlueFlames Demo Guide*
