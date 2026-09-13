// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = '';

function selectAccount() {
    /**
     * See here for more info on account retrieval:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const currentAccounts = myMSALObj.getAllAccounts();
    if (!currentAccounts || currentAccounts.length < 1) {
        return;
    } else if (currentAccounts.length > 1) {
        // Add your account choosing logic here
        console.warn('Multiple accounts detected.');
    } else if (currentAccounts.length === 1) {
        const account = currentAccounts[0];
        username = account.username;
        welcomeUser(username);

        // The cached AccountInfo only carries decoded ID token claims, not
        // the raw JWT string — refresh silently to get a fresh one to show.
        myMSALObj
            .acquireTokenSilent({ scopes: ['User.Read'], account })
            .then((result) => updateTable(result.account, result.idToken))
            .catch((error) => {
                console.warn('Could not refresh ID token for display:', error);
                updateTable(account, null);
            });
    }
}

function handleResponse(response) {
    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */

    if (response !== null) {
        username = response.account.username;
        welcomeUser(username);
        updateTable(response.account, response.idToken);
    } else {
        selectAccount();
    }
}

function signIn() {
    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    myMSALObj
        .loginPopup({
            ...loginRequest,
            redirectUri: '/redirect',
        })
        .then(handleResponse)
        .catch((error) => {
            console.log(error);
        });
}


function getTokenPopup(request) {
    /**
     * See here for more information on account retrieval:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    request.account = myMSALObj.getAccountByUsername(username);
    return myMSALObj.acquireTokenSilent(request).catch((error) => {
        console.warn(error);
        console.warn('silent token acquisition fails. acquiring token using popup');
        if (error instanceof msal.InteractionRequiredAuthError) {
            // fallback to interaction when silent call fails
            return myMSALObj
                .acquireTokenPopup(request)
                .then((response) => {
                    return response;
                })
                .catch((error) => {
                    console.error(error);
                });
        } else {
            console.warn(error);
        }
    });
}

function signOut() {
    /**
     * Kiosk / shared-device hardening: this app runs on tablets frontline workers
     * share back-to-back, so sign-out here means "End Shift" — it must fully
     * clear this browser's session, not just this app's view of it. A plain
     * per-app logout risks leaving the Entra ID session cookie alive, which
     * would let the NEXT person get silently signed back in via SSO.
     *
     * logoutRedirect() (not logoutPopup()) navigates the whole page through
     * Entra ID's own end_session endpoint, clearing the AAD session cookie
     * itself — not just this app's local token cache. MSAL clears its own
     * cache internally as part of that call, so no manual clear is needed here.
     *
     * logoutHint is set explicitly to skip Entra ID's "choose an account to
     * sign out of" picker. MSAL normally derives this itself from the
     * account's ID token 'login_hint' claim — but Entra ID only includes
     * that claim if it's been added as an optional claim in the app
     * registration's manifest, which this one hasn't. Passing it ourselves
     * (any value from the account works — MSAL skips its own derivation
     * once logoutHint is already set) avoids relying on that manifest change.
     */
    const account = myMSALObj.getAccountByUsername(username);
    const logoutRequest = {
        account,
        postLogoutRedirectUri: '/',
        logoutHint: account ? account.username : undefined,
    };

    myMSALObj.logoutRedirect(logoutRequest);
}

selectAccount();
