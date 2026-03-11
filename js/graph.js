async function getAccessToken() {

    const account = msalInstance.getAllAccounts()[0];

    const request = {
        scopes: ["User.Read", "Sites.Read.All"],
        account: account
    };

    try {

        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;

    } catch (error) {

        console.log("Silent token falhou");

        const response = await msalInstance.acquireTokenPopup(request);
        return response.accessToken;

    }

}


async function testarGraph(){

    const token = await getAccessToken();

    const resposta = await fetch(
        "https://graph.microsoft.com/v1.0/me",
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
