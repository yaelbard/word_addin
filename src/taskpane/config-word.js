const CONFIG = {
  deployment: "gpt-5.6-sol",
  endpoint: "https://gpt-agents-1972-foundry.openai.azure.com",
  apiVersion: "2025-01-01-preview",
  clientId: "2603315b-9e9f-4b43-b2fd-3de9ff9c41bd", 
  tenantId: "73ea3442-65e1-4556-a609-904f5d2e45ab"
};

const redirectPageUri = new URL("auth-redirect.html", window.location.href).href;

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: redirectPageUri,
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};
