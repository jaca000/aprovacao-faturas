async function getAccessToken() {

    const account = msalInstance.getAllAccounts()[0];

    const request = {
        scopes: ["User.Read"],
        account: account
    };

    const response = await msalInstance.acquireTokenSilent(request);

    return response.accessToken;

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
async function obterSiteApp(){

    const token = await getAccessToken();

    const resposta = await fetch(
        "https://graph.microsoft.com/v1.0/sites/montedopastopt.sharepoint.com:/sites/AppRegistoFaturas",
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPedidos(){

    const token = await getAccessToken();

    const siteId = "montedopastopt.sharepoint.com,309b2348-8df0-4dbe-9d3b2348-8df0-4dbe-945126c5bec7,3a90922f-7a65-44d9-ae1e-ef11c749a820";

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/PedidosAprovacao/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterListas(){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPedidosFaturas(){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPerfilUtilizador(){

    const token = await getAccessToken();

    const utilizador = await testarGraph();

    const email = utilizador.mail || utilizador.userPrincipalName;

    const site = await obterSiteApp();

    const siteId = site.id;

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/UtilizadoresApp/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    const lista = dados.value;

    const encontrado = lista.find(u => u.fields.Email === email);

    if(encontrado){
        return encontrado.fields.Perfil;
    }

    return "Utilizador";

}
async function uploadPdfSharePoint(file){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const nomeFicheiro = file.name;

    const urlUpload =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/Documents/root:/DocumentosAprovacao/${nomeFicheiro}:/content`;

    const resposta = await fetch(urlUpload,{
        method:"PUT",
        headers:{
            Authorization:"Bearer " + token,
            "Content-Type":"application/pdf"
        },
        body:file
    });

    const dados = await resposta.json();

    return dados;

}
