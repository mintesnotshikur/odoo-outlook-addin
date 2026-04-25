const ODOO_URL = "https://dablo.grace-erp-consultancy.com";
const DB_NAME = "dablo_DB"; // Ensure this is your correct DB name
const JSON_RPC_PATH = "/jsonrpc";
const ODOO_PROXY_URL = ""; // Optional: point this to your own relay endpoint to avoid browser CORS issues.
const AUTH_STORAGE_KEY = "odooBridgeAuth";

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("btnPush").onclick = runPush;
        document.getElementById("btnResetAuth").onclick = clearSavedAuth;
        restoreSavedAuth();
    }
});

async function runPush() {
    const status = document.getElementById("status");
    const button = document.getElementById("btnPush");
    const auth = getAuthValues();
    const user = auth.username;
    const pass = auth.password;
    const email_from = document.getElementById("email_from").value;

    if (!user || !pass) {
        status.innerText = "Enter your Odoo email and password or API key.";
        return;
    }

    status.innerText = "Authenticating...";
    button.disabled = true;

    try {
        const uid = await odooRpc("common", "authenticate", [DB_NAME, user, pass, {}]);

        if (!uid) {
            status.innerText = "Error: Invalid Credentials.";
            return;
        }

        const item = Office.context.mailbox.item;
        const senderEmail = item.from?.emailAddress || item.sender?.emailAddress || "Unknown sender";
        await saveAuth(user, pass);
        showSavedAuth(user);

        status.innerText = "Creating record...";
        const newId = await odooRpc("object", "execute_kw", [
            DB_NAME, uid, pass,
            "crm.lead", "create",
            [{
                name: `Email: ${item.subject || "No subject"}`,
                description: `From: ${senderEmail}`,
                email_from: email_from // Replace with your real Odoo custom field name.
            }]
        ]);

        status.innerText = `Success! Record ID: ${newId}`;
    } catch (err) {
        status.innerText = `Error: ${err.message}`;
        console.error(err);
    } finally {
        button.disabled = false;
    }
}

async function odooRpc(service, method, args) {
    const payload = {
        jsonrpc: "2.0",
        method: "call",
        params: {
            service,
            method,
            args
        },
        id: Date.now()
    };

    const requestUrl = ODOO_PROXY_URL || `${ODOO_URL}${JSON_RPC_PATH}`;
    const requestOptions = ODOO_PROXY_URL
        ? {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        }
        : {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        };

    let response;

    try {
        response = await fetch(requestUrl, requestOptions);
    } catch (err) {
        if (err instanceof TypeError) {
            throw new Error(
                "Browser blocked the request to Odoo. Enable CORS on the Odoo server or send requests through your own backend proxy."
            );
        }
        throw err;
    }

    if (!response.ok) {
        throw new Error(`Odoo request failed with status ${response.status}.`);
    }

    const result = await response.json();

    if (result.error) {
        const message = result.error.data?.message || result.error.message || "Unknown Odoo error.";
        throw new Error(message);
    }

    return result.result;
}

function getAuthValues() {
    const savedAuth = loadSavedAuth();

    if (savedAuth?.username && savedAuth?.password) {
        return savedAuth;
    }

    return {
        username: document.getElementById("username").value.trim(),
        password: document.getElementById("password").value
    };
}

function restoreSavedAuth() {
    const savedAuth = loadSavedAuth();

    if (!savedAuth?.username || !savedAuth?.password) {
        showManualAuth();
        return;
    }

    document.getElementById("username").value = savedAuth.username;
    document.getElementById("password").value = savedAuth.password;
    showSavedAuth(savedAuth.username);
    document.getElementById("status").innerText = `Signed in as ${savedAuth.username}.`;
}

async function saveAuth(username, password) {
    const payload = JSON.stringify({ username, password });
    const roamingSettings = Office.context?.roamingSettings;

    if (roamingSettings) {
        roamingSettings.set(AUTH_STORAGE_KEY, payload);
        await saveRoamingSettings(roamingSettings);
        return;
    }

    window.localStorage.setItem(AUTH_STORAGE_KEY, payload);
}

function loadSavedAuth() {
    const roamingSettings = Office.context?.roamingSettings;
    const rawValue = roamingSettings
        ? roamingSettings.get(AUTH_STORAGE_KEY)
        : window.localStorage.getItem(AUTH_STORAGE_KEY);

    if (!rawValue) {
        return null;
    }

    try {
        return JSON.parse(rawValue);
    } catch (err) {
        console.error("Failed to parse saved auth.", err);
        return null;
    }
}

async function clearSavedAuth() {
    const status = document.getElementById("status");
    const roamingSettings = Office.context?.roamingSettings;

    if (roamingSettings) {
        roamingSettings.remove(AUTH_STORAGE_KEY);
        await saveRoamingSettings(roamingSettings);
    } else {
        window.localStorage.removeItem(AUTH_STORAGE_KEY);
    }

    document.getElementById("username").value = "";
    document.getElementById("password").value = "";
    showManualAuth();
    status.innerText = "Saved sign-in removed. Enter another Odoo account.";
}

function showSavedAuth(username) {
    const savedAuth = document.getElementById("savedAuth");
    const authFields = document.getElementById("authFields");
    const resetButton = document.getElementById("btnResetAuth");

    savedAuth.innerText = `Using saved Odoo sign-in for ${username}.`;
    savedAuth.style.display = "block";
    authFields.classList.add("hidden");
    resetButton.classList.remove("hidden");
}

function showManualAuth() {
    const savedAuth = document.getElementById("savedAuth");
    const authFields = document.getElementById("authFields");
    const resetButton = document.getElementById("btnResetAuth");

    savedAuth.style.display = "none";
    savedAuth.innerText = "";
    authFields.classList.remove("hidden");
    resetButton.classList.add("hidden");
}

function saveRoamingSettings(roamingSettings) {
    return new Promise((resolve, reject) => {
        roamingSettings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
                return;
            }

            reject(new Error(result.error?.message || "Failed to save Outlook roaming settings."));
        });
    });
}
