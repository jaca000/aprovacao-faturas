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
lista.sort((a,b)=>{

const estadoA = a.fields.EstadoPedido;
const estadoB = b.fields.EstadoPedido;

if(estadoA === "Pendente" && estadoB !== "Pendente") return -1;
if(estadoA !== "Pendente" && estadoB === "Pendente") return 1;

return 0;

});
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
linha.onclick = () => {
    window.location.href = `ver-pedido.html?id=${p.id}`;
};
    linha.innerHTML = `

<td>
<input type="checkbox" class="checkPedido" value="${p.id}">
</td>

<td>${f.NumeroInterno || ""}</td>
<td>${f.Fornecedor || ""}</td>
<td>${f.NumeroFaturaOriginal || ""}</td>
<td>${f.ValorDocumento || ""}</td>
<td>${badgeEstado(f.EstadoPedido)}</td>

`;

    tabela.appendChild(linha);

});
}
window.addEventListener("load", () => {

    const tabela = document.getElementById("listaPedidos");

    if(tabela){
        carregarDashboard();
    }

});
window.guardarFatura = async function guardarFatura(){

    const fornecedor = document.getElementById("fornecedor").value;
    const numeroFatura = document.getElementById("numeroFatura").value;
    const numeroNormalizado = numeroFatura.toUpperCase().trim();
    const numeroInterno = await gerarNumeroInterno();

const duplicado = await verificarFaturaDuplicada(numeroNormalizado);

if(duplicado){

    alert("Esta fatura já existe no sistema.");

    return;

}
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

            Title: fornecedor,
            
            NumeroInterno: numeroInterno,

            TipoDocumento: "Fatura",

            Fornecedor: fornecedor,

            NumeroFaturaOriginal: numeroFatura,

            NumeroFaturaNormalizado: numeroNormalizado,

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
async function carregarPedido(){

    const params = new URLSearchParams(window.location.search);
    const id = params.get("id");

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const dados = await resp.json();

    const f = dados.fields;

    document.getElementById("dadosPedido").innerHTML = `
        <p><b>Nº Interno:</b> ${f.NumeroInterno}</p>
        <p><b>Fornecedor:</b> ${f.Fornecedor}</p>
        <p><b>Nº Fatura:</b> ${f.NumeroFaturaOriginal}</p>
        <p><b>Valor:</b> ${f.ValorDocumento}</p>
        <p><b>Estado:</b> ${f.EstadoPedido}</p>
    `;

    document.getElementById("dadosPedido").innerHTML += `
<br>
<button onclick="abrirPdf('${f.PdfUrl}')">
Ver PDF da fatura
</button>
`;

}
if(window.location.pathname.includes("ver-pedido.html")){
    carregarPedido();
}
async function atualizarEstadoPedido(novoEstado){

    const params = new URLSearchParams(window.location.search);
    const id = params.get("id");

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const body = {
        fields:{
            EstadoPedido: novoEstado
        }
    };

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}/fields`,
        {
            method: "PATCH",
            headers:{
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        }
    );

    if(!resp.ok){
        alert("Erro ao atualizar o estado");
        return;
    }

    alert("Estado atualizado para: " + novoEstado);

    window.location.href = "dashboard.html";

}

function aprovarPedido(){
    atualizarEstadoPedido("Aprovado");
}

function rejeitarPedido(){
    atualizarEstadoPedido("Rejeitado");
}
function abrirPdf(url){
    window.open(url, "_blank");
}
function ordenarTabela(coluna){

const tabela = document.getElementById("tabelaPedidos");

const linhas = Array.from(tabela.rows).slice(1);

const asc = tabela.classList.toggle("asc");

linhas.sort((a,b)=>{

let A = a.cells[coluna].innerText;
let B = b.cells[coluna].innerText;

if(!isNaN(A) && !isNaN(B)){
return asc ? A-B : B-A;
}

return asc ? A.localeCompare(B) : B.localeCompare(A);

});

linhas.forEach(l => tabela.appendChild(l));

}
const pesquisa = document.getElementById("pesquisa");

if(pesquisa){

pesquisa.addEventListener("keyup",function(){

const termo = this.value.toLowerCase();

document.querySelectorAll("#listaPedidos tr").forEach(linha=>{

linha.style.display =
linha.innerText.toLowerCase().includes(termo)
? ""
: "none";

});

});

}
function badgeEstado(estado){

if(estado === "Pendente"){
return '<span class="badge pendente">Pendente</span>';
}

if(estado === "Aprovado"){
return '<span class="badge aprovado">Aprovado</span>';
}

if(estado === "Rejeitado"){
return '<span class="badge rejeitado">Rejeitado</span>';
}

return estado;

}
async function aprovarSelecionados(){

const selecionados = document.querySelectorAll(".checkPedido:checked");

for(const check of selecionados){

await atualizarEstadoPedidoId(check.value,"Aprovado");

}

alert("Pedidos aprovados");

location.reload();

}
async function rejeitarSelecionados(){

const selecionados = document.querySelectorAll(".checkPedido:checked");

for(const check of selecionados){

await atualizarEstadoPedidoId(check.value,"Rejeitado");

}

alert("Pedidos rejeitados");

location.reload();

}
