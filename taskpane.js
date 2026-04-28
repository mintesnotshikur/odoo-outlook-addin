const ODOO_URL = "https://dablo.grace-erp-consultancy.com";
const DB_NAME = "dablo_DB"; // Ensure this is your correct DB name
const JSON_RPC_PATH = "/jsonrpc";
const ODOO_PROXY_URL = ""; // Optional: point this to your own relay endpoint to avoid browser CORS issues.
const AUTH_STORAGE_KEY = "odooBridgeAuth";

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("btnLogin").onclick = loginToOdoo;
        document.getElementById("btnPush").onclick = runPush;
        document.getElementById("btnResetAuth").onclick = clearSavedAuth;
        restoreSavedAuth();
    }
});

async function loginToOdoo() {
    const status = document.getElementById("status");
    const button = document.getElementById("btnLogin");
    const user = document.getElementById("username").value.trim();
    const pass = document.getElementById("password").value;

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

        await saveAuth(user, pass);
        showPipelineView(user);
        prefillLeadEmail();
        status.innerText = `Signed in as ${user}.`;
    } catch (err) {
        status.innerText = `Error: ${err.message}`;
        console.error(err);
    } finally {
        button.disabled = false;
    }
}

async function runPush() {
    const status = document.getElementById("status");
    const button = document.getElementById("btnPush");
    const auth = getAuthValues();
    const user = auth.username;
    const pass = auth.password;
    const item = Office.context.mailbox.item;
    const senderName = getSenderName(item);
    const senderEmail = getSenderEmail(item);
    const email_from = document.getElementById("email_from").value.trim() || senderEmail;

    if (!user || !pass) {
        showLoginView();
        status.innerText = "Sign in to Odoo first.";
        return;
    }

    if (!email_from) {
        status.innerText = "No sender email found. Enter an email before creating the lead.";
        return;
    }

    status.innerText = "Checking client...";
    button.disabled = true;

    try {
        const uid = await odooRpc("common", "authenticate", [DB_NAME, user, pass, {}]);

        if (!uid) {
            showLoginView();
            status.innerText = "Saved sign-in is no longer valid. Please sign in again.";
            return;
        }

        showPipelineView(user);
        document.getElementById("email_from").value = email_from;

        const partner = await ensurePartner(uid, pass, senderName, email_from);
        const emailBody = await getEmailBodyText(item);
        const description = buildLeadDescription(senderEmail, emailBody);

        status.innerText = partner.created
            ? `Created client ${partner.name}. Creating lead...`
            : `Found client ${partner.name}. Creating lead...`;
        const newId = await odooRpc("object", "execute_kw", [
            DB_NAME, uid, pass,
            "crm.lead", "create",
            [{
                name: `Email: ${item.subject || "No subject"}`,
                description,
                email_from,
                partner_id: partner.id,
                partner_name: partner.name
            }]
        ]);

        status.innerText = `Success! Lead ${newId} linked to ${partner.name}.`;
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
        showLoginView();
        return;
    }

    document.getElementById("username").value = savedAuth.username;
    document.getElementById("password").value = savedAuth.password;
    showPipelineView(savedAuth.username);
    prefillLeadEmail();
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
    document.getElementById("email_from").value = "";
    showLoginView();
    status.innerText = "Saved sign-in removed. Enter another Odoo account.";
}

function showPipelineView(username) {
    const savedAuth = document.getElementById("savedAuth");
    const loginView = document.getElementById("loginView");
    const pipelineView = document.getElementById("pipelineView");

    savedAuth.innerText = `Using saved Odoo sign-in for ${username}.`;
    loginView.classList.add("hidden");
    pipelineView.classList.remove("hidden");
}

function showLoginView() {
    const loginView = document.getElementById("loginView");
    const pipelineView = document.getElementById("pipelineView");

    loginView.classList.remove("hidden");
    pipelineView.classList.add("hidden");
}

function prefillLeadEmail() {
    const emailInput = document.getElementById("email_from");
    const item = Office.context?.mailbox?.item;
    const senderEmail = getSenderEmail(item);

    if (!emailInput.value && senderEmail) {
        emailInput.value = senderEmail;
    }
}

async function ensurePartner(uid, pass, senderName, email) {
    const existingPartners = await odooRpc("object", "execute_kw", [
        DB_NAME, uid, pass,
        "res.partner", "search_read",
        [[["email", "=ilike", email]]],
        {
            fields: ["id", "name", "email"],
            limit: 1
        }
    ]);

    if (existingPartners.length > 0) {
        return {
            ...existingPartners[0],
            created: false
        };
    }

    const partnerName = senderName || email;
    const partnerId = await odooRpc("object", "execute_kw", [
        DB_NAME, uid, pass,
        "res.partner", "create",
        [{
            name: partnerName,
            email
        }]
    ]);

    return {
        id: partnerId,
        name: partnerName,
        email,
        created: true
    };
}

function getSenderEmail(item) {
    return item?.from?.emailAddress || item?.sender?.emailAddress || "";
}

function getSenderName(item) {
    return item?.from?.displayName || item?.sender?.displayName || "";
}

function getEmailBodyText(item) {
    return new Promise((resolve) => {
        if (!item?.body?.getAsync) {
            resolve("");
            return;
        }

        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve((result.value || "").trim());
                return;
            }

            console.error("Failed to read email body.", result.error);
            resolve("");
        });
    });
}

function buildLeadDescription(senderEmail, emailBody) {
    const sections = [`From: ${senderEmail || "Unknown sender"}`];

    if (emailBody) {
        sections.push("", "Body:", emailBody);
    }

    return sections.join("\n");
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
