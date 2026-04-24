const ODOO_URL = "https://dablo.grace-erp-consultancy.com";
const DB_NAME = "dablo_DB"; // Ensure this is your correct DB name
const JSON_RPC_PATH = "/jsonrpc";
const ODOO_PROXY_URL = ""; // Optional: point this to your own relay endpoint to avoid browser CORS issues.

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("btnPush").onclick = runPush;
    }
});

async function runPush() {
    const status = document.getElementById("status");
    const button = document.getElementById("btnPush");
    const user = document.getElementById("username").value;
    const pass = document.getElementById("password").value;
    const email = document.getElementById("email").value;

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
        const bodyPreview = item.body?.preview || "No preview available";

        status.innerText = "Creating record...";
        const newId = await odooRpc("object", "execute_kw", [
            DB_NAME, uid, pass,
            "crm.lead", "create",
            [{
                name: `Email: ${item.subject || "No subject"}`,
                description: `From: ${senderEmail}`,
                email_from: email // Replace with your real Odoo custom field name.
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
