console.log("🔥 config-word.js LOADED SUCCESSFULLY!");

// Global Application Configuration
window.CONFIG = {
  deployment: "gpt-5.6-sol",
  endpoint: "https://gpt-agents-1972-foundry.openai.azure.com",
  apiVersion: "2025-01-01-preview",
  clientId: "2603315b-9e9f-4b43-b2fd-3de9ff9c41bd", 
  tenantId: "73ea3442-65e1-4556-a609-904f5d2e45ab"
};

// Compute fully-qualified redirect URL based on current host/path
window.redirectPageUri = new URL("auth-redirect.html", window.location.href).href;

// Global MSAL Browser Configuration
window.msalConfig = {
  auth: {
    clientId: window.CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${window.CONFIG.tenantId}`,
    redirectUri: window.redirectPageUri,
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};