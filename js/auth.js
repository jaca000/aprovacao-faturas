const msalConfig = {
    auth: {
        clientId: "81693cb9-9ffb-41d0-b0ff-41dbf29990eb",
        authority: "https://login.microsoftonline.com/ee417351-ea90-41e0-9147-5ea6ab38ea49",
        redirectUri: "https://jaca000.github.io/aprovacao-faturas/"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
    scopes: ["User.Read"]
};

async function iniciarLogin() {

    await msalInstance.handleRedirectPromise();

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
