
async function initMsal() {
  if (!isMsalInitialized) {
    await msalInstance.initialize();
    isMsalInitialized = true;
  }
}

/**
 * פתיחת חלון אימות דרך Office Dialog API
 */
function openLoginDialog() {
  return new Promise((resolve, reject) => {
    // שמירת המזהים ישירות ב-localStorage המשותף
    if (!CONFIG || !CONFIG.clientId || !CONFIG.tenantId) {
      reject(new Error("CONFIG.clientId or CONFIG.tenantId is missing in taskpane.js"));
      return;
    }
    localStorage.setItem("entra_clientId", CONFIG.clientId);
    localStorage.setItem("entra_tenantId", CONFIG.tenantId);

    // נתיב הדיאלוג נשאר נקי וקבוע
    const dialogUrl = new URL("auth-redirect.html", window.location.href).href;

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
}

/**
 * פונקציית האימות הראשית
 */
async function getAccessToken() {
  await initMsal();

  const loginRequest = {
    scopes: ["https://cognitiveservices.azure.com/.default"]
  };

  try {
    const accounts = msalInstance.getAllAccounts();

    // אם קיים חשבון שמור, מנסים להוציא טוקן באופן שקט ללא שום חלון
    if (accounts.length > 0) {
      const silentResponse = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0]
      });
      return silentResponse.accessToken;
    }
  } catch (silentError) {
    console.warn("Silent token acquisition failed. Opening dialog...", silentError);
  }

  // אם אין חשבון או שהטוקן השקט נכשל – פותחים דיאלוג ייעודי של אופיס
  return await openLoginDialog();
}
