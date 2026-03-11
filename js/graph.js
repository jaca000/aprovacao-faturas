async function getAccessToken() {

    const account = msalInstance.getAllAccounts()[0];

    const request = {
        scopes: ["User.Read", "Sites.ReadWrite.All"],
        account: account
    };

    try {

        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;

    } catch (error) {

        console.log("Token silencioso falhou, pedir novamente...");

        const response = await msalInstance.acquireTokenPopup(request);
        return response.accessToken;

    }

}
