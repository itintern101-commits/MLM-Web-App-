// ================================
// 🔐 LOGIN.JS
// ================================

let msalInstance;

// Fetch config from backend
fetch("/config")
  .then(res => res.json())
  .then(async (config) => {

    const msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: window.location.origin + "/login.html"
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    };

    msalInstance = new msal.PublicClientApplication(msalConfig);

    // 🔥 ONLY HERE handle redirect
    const response = await msalInstance.handleRedirectPromise();

    if (response && response.account) {
      console.log("Login success:", response.account);

      sessionStorage.setItem("user", JSON.stringify(response.account));

      window.location.href = "/dashboard.html";
      return;
    }

    initUI();

  })
  .catch(err => {
    console.error("Config load error:", err);
    updateStatus("Unable to load authentication config.");
  });


// ================================
// 🔘 INIT UI
// ================================
function initUI() {
  checkExistingLogin();

  const btn = document.getElementById("loginButton");
  if (btn) {
    btn.addEventListener("click", handleLogin);
  }
}


// ================================
// 🔘 LOGIN BUTTON
// ================================
function handleLogin() {
  const btn = document.getElementById("loginButton");

  btn.disabled = true;
  btn.textContent = "Signing in...";

  msalInstance.loginRedirect({ scopes: ["user.read"] });
}


// ================================
// 🔍 CHECK EXISTING LOGIN
// ================================
function checkExistingLogin() {
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    sessionStorage.setItem("user", JSON.stringify(accounts[0]));
    window.location.href = "/dashboard.html";
  }
}


// ================================
// 💬 STATUS
// ================================
function updateStatus(message) {
  const statusDiv = document.getElementById("status");
  if (statusDiv) statusDiv.textContent = message;
}