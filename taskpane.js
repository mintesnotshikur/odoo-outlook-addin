const ODOO_URL = "https://dablo.grace-erp-consultancy.com";
const DB_NAME = "dablo"; // Ensure this is your correct DB name

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("btnPush").onclick = runPush;
    }
});

async function runPush() {
    const status = document.getElementById("status");
    const user = document.getElementById("username").value;
    const pass = document.getElementById("password").value;
    const customVal = document.getElementById("custom_field").value;

    status.innerText = "Authenticating...";

    try {
        // 1. Authenticate to get UID
        const uid = await odooRpc("/xmlrpc/2/common", "authenticate", [DB_NAME, user, pass, {}]);
        
        if (!uid) {
            status.innerText = "Error: Invalid Credentials.";
            return;
        }

        // 2. Get Email Data from Outlook
        const item = Office.context.mailbox.item;
        
        // 3. Create Record in Odoo (Example: CRM Lead)
        status.innerText = "Creating record...";
        const newId = await odooRpc("/xmlrpc/2/object", "execute_kw", [
            DB_NAME, uid, pass,
            'crm.lead', 'create', 
            [{
                'name': `Email: ${item.subject}`,
                'description': `From: ${item.from.emailAddress}`,
                'x_custom_field': customVal // REPLACE THIS with your actual Odoo custom field name
            }]
        ]);

        status.innerText = `Success! Record ID: ${newId}`;
    } catch (err) {
        status.innerText = "Error: " + err.message;
        console.error(err);
    }
}

async function odooRpc(path, method, params) {
    const response = await fetch(`${ODOO_URL}${path}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            jsonrpc: "2.0",
            method: "call",
            params: {
                service: path.includes("common") ? "common" : "object",
                method: method,
                args: params
            }
        })
    });
    const result = await response.json();
    if (result.error) throw new Error(result.error.data.message);
    return result.result;
}
