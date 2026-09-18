console.log("🔥  word-auth.js LOADED SUCCESSFULLY!");
// Ensure MSAL configuration and library are loaded before instantiation
window.msalInstance = null;
if (window.msal && window.msalConfig) {
  window.msalInstance = new window.msal.PublicClientApplication(window.msalConfig);
}

window.isMsalInitialized = false;

/**
 * Initializes the MSAL PublicClientApplication instance
 */
window.initMsal = async function() {
  if (!window.msalInstance && window.msal && window.msalConfig) {
    window.msalInstance = new window.msal.PublicClientApplication(window.msalConfig);
  }

  if (window.msalInstance && !window.isMsalInitialized) {
    await window.msalInstance.initialize();
    window.isMsalInitialized = true;
  }
};

/**
 * Opens the authentication dialog using the Office Dialog API
 */
window.openLoginDialog = function() {
  return new Promise((resolve, reject) => {
    const config = window.CONFIG;

    if (!config || !config.clientId || !config.tenantId) {
      reject(new Error("window.CONFIG.clientId or window.CONFIG.tenantId is missing."));
      return;
    }

    // Persist to localStorage so values survive the round-trip to Microsoft login
    localStorage.setItem("entra_clientId", config.clientId);
    localStorage.setItem("entra_tenantId", config.tenantId);

    const params = new URLSearchParams({
      clientId: config.clientId,
      tenantId: config.tenantId
    });

    const dialogUrl = new URL(`auth-redirect.html?${params.toString()}`, window.location.href).href;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 70, width: 60, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Failed to open dialog: ${asyncResult.error.message}`));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const response = JSON.parse(arg.message);
            if (response.status === "success") {
              resolve(response.token);
            } else {
              reject(new Error(response.error || "Authentication failed"));
            }
          } catch (e) {
            reject(new Error("Malformed dialog response"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          if (arg.error === 12006) {
            reject(new Error("Login window was closed by the user."));
          }
        });
      }
    );
  });
};/**
 * Primary token retrieval flow: attempts silent acquisition, falls back to interactive dialog
 */
window.getAccessToken = async function() {
  await window.initMsal();

  const loginRequest = {
    scopes: ["https://cognitiveservices.azure.com/.default"]
  };

  // Attempt silent token retrieval if cached account session exists
  if (window.msalInstance) {
    try {
      const accounts = window.msalInstance.getAllAccounts();

      if (accounts.length > 0) {
        const silentResponse = await window.msalInstance.acquireTokenSilent({
          ...loginRequest,
          account: accounts[0]
        });
        return silentResponse.accessToken;
      }
    } catch (silentError) {
      console.warn("Silent token acquisition failed. Opening dialog...", silentError);
    }
  }

  // Fall back to Office interactive login dialog
  return await window.openLoginDialog();
};