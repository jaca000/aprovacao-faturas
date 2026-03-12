console.log("App iniciada");

const appDiv = document.getElementById("app");

if(appDiv){
    appDiv.innerHTML = `
    <button onclick="login()">Iniciar sessão Microsoft</button>
    `;
}

function mostrarDashboard(utilizador){

    document.getElementById("app").innerHTML = `

        <h2>Dashboard</h2>

        <p>Bem-vindo ${utilizador.name} | Monte do Pasto</p>

        <button onclick="testarLigacao()">Testar ligação</button>

    `;

}



async function testarLigacao(){

    console.log("A testar ligação ao Microsoft Graph");

    const utilizador = await testarGraph();
 
    console.log("Utilizador Graph:", utilizador);
document.getElementById("utilizador").innerText =
"Bem-vindo " + utilizador.displayName;
    const site = await obterSiteApp();
    console.log("Site da aplicação:", site);

    const listas = await obterListas();
    console.log("Listas do site:", listas);

    const pedidos = await obterPedidosFaturas();
    console.log("Pedidos da lista:", pedidos);

}
async function carregarDashboard(){
const perfil = await obterPerfilUtilizador();

console.log("Perfil do utilizador:", perfil);
    const btnFatura = document.getElementById("btnNovaFatura");
    btnFatura.onclick = () => {

    window.location.href = "nova-fatura.html";

};
const btnDespesa = document.getElementById("btnNovaDespesa");
const btnAdmin = document.getElementById("btnAdmin");

if(perfil === "Admin"){

    btnFatura.style.display = "inline-block";
    btnDespesa.style.display = "inline-block";
    btnAdmin.style.display = "inline-block";

}

else if(perfil === "GestorFaturas"){

    btnFatura.style.display = "inline-block";
    btnDespesa.style.display = "inline-block";
    btnAdmin.style.display = "none";

}

else{

    btnFatura.style.display = "none";
    btnDespesa.style.display = "inline-block";
    btnAdmin.style.display = "none";

}
    const pedidos = await obterPedidosFaturas();

    const lista = pedidos.value;

    const total = lista.length;

    let pendentes = 0;
    let aprovados = 0;
    let rejeitados = 0;

    lista.forEach(p => {

        const estado = p.fields.EstadoPedido;

        if(estado === "Pendente") pendentes++;
        if(estado === "Aprovado") aprovados++;
        if(estado === "Rejeitado") rejeitados++;

    });

    document.getElementById("totalPedidos").innerText = total;
    document.getElementById("pendentes").innerText = pendentes;
    document.getElementById("aprovados").innerText = aprovados;
    document.getElementById("rejeitados").innerText = rejeitados;
const tabela = document.getElementById("listaPedidos");

tabela.innerHTML = "";

lista.forEach(p => {

    const f = p.fields;

    const linha = document.createElement("tr");

    linha.innerHTML = `
        <td>${p.id}</td>
        <td>${f.Fornecedor || ""}</td>
        <td>${f.NumeroFatura || ""}</td>
        <td>${f.Valor || ""}</td>
        <td>${f.Estado || ""}</td>
    `;

    tabela.appendChild(linha);

});
}
window.addEventListener("load", () => {

    if(window.location.pathname.includes("dashboard.html")){

        carregarDashboard();

        const btnFatura = document.getElementById("btnNovaFatura");

        if(btnFatura){
            btnFatura.onclick = () => {
                window.location.href = "nova-fatura.html";
            };
        }

    }
window.guardarFatura = async function guardarFatura(){

    const fornecedor = document.getElementById("fornecedor").value;
    const numeroFatura = document.getElementById("numeroFatura").value;
    const valor = document.getElementById("valor").value;
    const dataDocumento = document.getElementById("dataFatura").value;
    const dataVencimento = document.getElementById("dataVencimento").value;
    const ficheiro = document.getElementById("ficheiroPDF").files[0];
    let pdfUrl = "";
let pdfNome = "";

if(ficheiro){

    const upload = await uploadPdfSharePoint(ficheiro);

    pdfUrl = upload.webUrl;
    pdfNome = upload.name;

}
    const utilizador = await testarGraph();

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const body = {
        fields: {

            Title: Fornecedor,

            TipoDocumento: "Fatura",

            Fornecedor: fornecedor,

            NumeroFaturaOriginal: numeroFatura,

            NumeroFaturaNormalizado: numeroFatura.toUpperCase(),

            ValorDocumento: parseFloat(valor),

            Moeda: "EUR",

            DataDocumento: dataDocumento,

            DataVencimento: dataVencimento,

            CriadoPorNome: utilizador.displayName,

            CriadoPorEmail: utilizador.mail || utilizador.userPrincipalName,

            DataCriacaoPedido: new Date().toISOString(),

            EstadoPedido: "Pendente",
PdfUrl: pdfUrl,
PdfNomeFicheiro: pdfNome
        }
    };

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items`,
        {
            method: "POST",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        }
    );

    const resultado = await resposta.json();

console.log("Resposta Graph:", resultado);

if(!resposta.ok){
    alert("Erro ao gravar no SharePoint");
    return;
}

alert("Fatura registada com sucesso");

    window.location.href = "dashboard.html";

};

});
