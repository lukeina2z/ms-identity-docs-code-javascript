// Browser check variables
// If you support IE, our recommendation is that you sign-in using Redirect APIs
// If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check
const ua = window.navigator.userAgent;
const msie = ua.indexOf("MSIE ");
const msie11 = ua.indexOf("Trident/");
const msedge = ua.indexOf("Edge/");
const isIE = msie > 0 || msie11 > 0;
const isEdge = msedge > 0;

let signInType;
let accountId = "";

// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

myMSALObj.initialize().then(() => {
    return myMSALObj.handleRedirectPromise();
}).then(handleResponse).catch(err => {
    console.error(err);
});

function selectAccount() {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const activeAccount = myMSALObj.getActiveAccount();
    if (activeAccount) {
        return activeAccount;
    }

    const currentAccounts = myMSALObj.getAllAccounts();

    if (currentAccounts.length > 1) {
        // Add your account choosing logic here
        console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
        return currentAccounts[0];
    }
}

async function handleResponse(resp) {
    const account = resp ? resp.account : selectAccount();
    if (!account) {
        return;
    }

    accountId = account.homeAccountId;
    myMSALObj.setActiveAccount(account);
    showWelcomeMessage(account);
    await loadUserProfile(account);
}

async function loadUserProfile(account) {
    showProfileLoading();

    try {
        const tokenResponse = await getTokenRedirect(loginRequest, account);
        if (!tokenResponse) {
            return; // Interactive token acquisition is navigating away.
        }
        if (!tokenResponse.accessToken) {
            throw new Error("MSAL did not return a Microsoft Graph access token.");
        }

        const response = await fetch("https://graph.microsoft.com/v1.0/me", {
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`Microsoft Graph /me failed (${response.status}): ${await response.text()}`);
        }

        showUserProfile(await response.json());
    } catch (error) {
        console.error("Unable to load Microsoft Graph profile.", error);
        showProfileError(error);
    }
}

async function signIn(method) {
    signInType = isIE ? "redirect" : method;
    if (signInType === "popup") {
        return myMSALObj.loginPopup({
            ...loginRequest,
            redirectUri: "/redirect"
        }).then(handleResponse).catch(function (error) {
            console.log(error);
        });
    } else if (signInType === "redirect") {
        return myMSALObj.loginRedirect(loginRequest)
    }
}

function signOut(interactionType) {
    const logoutRequest = {
        account: myMSALObj.getAccount({ homeAccountId: accountId })
    };

    if (interactionType === "popup") {
        myMSALObj.logoutPopup(logoutRequest).then(() => {
            window.location.reload();
        });
    } else {
        myMSALObj.logoutRedirect(logoutRequest);
    }
}

async function getTokenRedirect(request, account) {
    const tokenRequest = { ...request, account };

    try {
        return await myMSALObj.acquireTokenSilent(tokenRequest);
    } catch (error) {
        if (!(error instanceof msal.InteractionRequiredAuthError)) {
            throw error;
        }

        await myMSALObj.acquireTokenRedirect(tokenRequest);
        return null;
    }
}