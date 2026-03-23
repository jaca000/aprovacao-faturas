    const msalConfig = {
    auth: {
        clientId: "81693cb9-9ffb-41d0-b0ff-41dbf29990eb",
        authority: "https://login.microsoftonline.com/ee417351-ea90-41e0-9147-5ea6ab38ea49",
        redirectUri: "https://jaca000.github.io/aprovacao-faturas/pages/dashboard.html"
    },
    cache: {
        cacheLocation: "localStorage", // 🔥 mantém sessão
        storeAuthStateInCookie: false
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let contaAtual = null;

/* ================= LOGIN ================= */

async function login() {
    try {

        // 🔥 limpar sessões antigas
        const contas = msalInstance.getAllAccounts();

        if (contas.length > 0) {
            await msalInstance.logoutPopup({
                account: contas[0]
            });
        }

        const response = await msalInstance.loginPopup({
            scopes: ["User.Read"],
            prompt: "select_account"
        });

        contaAtual = response.account;

        console.log("Login OK:", contaAtual);

        window.location.href = "/aprovacao-faturas/pages/dashboard.html";

    } catch (error) {
        console.error("Erro no login:", error);
    }
}

/* ================= TOKEN ================= */

async function getAccessToken() {

    const contas = msalInstance.getAllAccounts();

    if (contas.length === 0) {
        throw new Error("Sem utilizador autenticado");
    }

    const request = {
        scopes: ["User.Read", "Sites.ReadWrite.All"],
        account: contas[0]
    };

    try {

        const response = await msalInstance.acquireTokenSilent(request);

        return response.accessToken;

    } catch (error) {

        console.warn("Token silencioso falhou, a pedir popup...");

        const response = await msalInstance.acquireTokenPopup(request);

        return response.accessToken;
    }
}

/* ================= UTIL ================= */

async function testarGraph() {

    const token = await getAccessToken();

    const resp = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
            Authorization: "Bearer " + token
        }
    });

    const data = await resp.json();

    return data;
}
