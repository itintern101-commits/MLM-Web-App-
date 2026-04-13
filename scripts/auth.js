// ================================
// 🔐 AUTH.JS (PROTECTED PAGES)
// ================================

let msalInstance;

// Fetch config
// This is done to keep sensitive info like clientId and tenantId out of the public code
fetch("/config")
    .then(res => res.json())
    .then((config) => {

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

        checkLogin();

    })
    .catch(err => {
        console.error("Config error:", err);
        window.location.href = "/login.html";
    });


// ================================
// 🔍 CHECK LOGIN ONLY
// ================================
function checkLogin() {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        // ❌ Not logged in
        window.location.href = "/login.html";
    } else {
        // ✅ Logged in
        sessionStorage.setItem("user", JSON.stringify(accounts[0]));
        document.body.style.display = "block";
    }
}


// ================================
// 🚪 LOGOUT
// ================================
// auth.js
function logout() {
    sessionStorage.clear();
    window.location.href = "/login.html";
}

// Attach listener to the document
document.addEventListener("click", (e) => {
    if (e.target && e.target.id === "logoutBtn") {
        e.preventDefault();
        logout();
    }
});