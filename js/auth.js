const msalConfig = {
    auth: {
        clientId: "81693cb9-9ffb-41d0-b0ff-41dbf29990eb",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
    scopes: ["User.Read"]
};

async function iniciarLogin() {

    const contas = msalInstance.getAllAccounts();

    if (contas.length === 0) {

        console.log("Sem sessão. Redirecionar para login...");

        await msalInstance.loginRedirect(loginRequest);

    } else {

        const utilizador = contas[0];

        console.log("Utilizador autenticado:", utilizador);

        sessionStorage.setItem("userName", utilizador.name);
        sessionStorage.setItem("userEmail", utilizador.username);

    }

}
