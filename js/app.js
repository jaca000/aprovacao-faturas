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

    const utilizador = await testarGraph();

const email = utilizador.mail || utilizador.userPrincipalName;

const lista = pedidos.value;
const paraMim = [];
const outros = [];

lista.forEach(p => {

const f = p.fields;

if(
f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
)
){
paraMim.push(p);
}else{
outros.push(p);
}

});

const meusPendentes = lista.filter(p => {

const f = p.fields;

return f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
);

});
let valorAprovar = 0;

meusPendentes.forEach(p => {

const v = parseFloat(p.fields.ValorDocumento || 0);

valorAprovar += v;

});
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
    document.getElementById("meusPendentes").innerText = meusPendentes.length;
    document.getElementById("valorAprovar").innerText =
valorAprovar.toLocaleString("pt-PT",{style:"currency",currency:"EUR"});
const tabela = document.getElementById("listaPedidos");

tabela.innerHTML = "";

lista.forEach(p => {

    const f = p.fields;

    const linha = document.createElement("tr");

if(
f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
)
){
linha.classList.add("linha-para-aprovar");
}
linha.onclick = () => {
window.location.href = `ver-pedido.html?id=${p.id}`;
};

const dias = diasParaVencimento(f.DataVencimento);

let alerta = "";

if(dias !== null){

if(dias <= 3){
alerta = '<span class="vencimento-urgente">⚠ vence em '+dias+' dias</span>';
}

else if(dias <= 8){
alerta = '<span class="vencimento-alerta">vence em '+dias+' dias</span>';
}

}

linha.innerHTML = `
<td>
<input type="checkbox" class="checkPedido" value="${p.id}" onclick="event.stopPropagation()">
</td>
<td>${f.NumeroInterno || ""}</td>
<td>${f.Fornecedor || ""}</td>
<td>${f.NumeroFaturaOriginal || ""}</td>
<td>${f.ValorDocumento || ""}</td>
<td>${badgeEstado(f.EstadoPedido)} ${alerta}</td>
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
/* TIMELINE */

const timeline = document.getElementById("timelinePedido");

if(timeline){

timeline.innerHTML = "";

/* criado */

timeline.innerHTML += `
<li class="timeline-criado">
Pedido criado por ${f.CriadoPorNome || "utilizador"}
</li>
`;

/* enviado */

timeline.innerHTML += `
<li>
Enviado para aprovação
</li>
`;

/* aprovado */

if(f.EstadoPedido === "Aprovado"){

timeline.innerHTML += `
<li class="timeline-aprovado">
Aprovado
</li>
`;

}

/* rejeitado */

if(f.EstadoPedido === "Rejeitado"){

timeline.innerHTML += `
<li class="timeline-rejeitado">
Rejeitado
</li>
`;

if(f.ComentarioRejeicao){

timeline.innerHTML += `
<li>
Motivo: ${f.ComentarioRejeicao}
</li>
`;

}

}

}
   document.getElementById("dadosPedido").innerHTML = `

<h3>${f.Fornecedor}</h3>

<p><b>Nº Interno:</b> ${f.NumeroInterno || "-"}</p>

<p><b>Nº Fatura:</b> ${f.NumeroFaturaOriginal}</p>

<p><b>Valor:</b> ${f.ValorDocumento} €</p>

<p><b>Estado:</b> ${badgeEstado(f.EstadoPedido)}</p>

`;

    document.getElementById("dadosPedido").innerHTML += `
<br>
<button onclick="abrirPdf('${f.PdfUrl}')">
Ver PDF da fatura
</button>
`;

}
if(window.location.pathname.includes("ver-pedido.html")){

testarGraph().then(()=>{
carregarPedido();
});

}
async function atualizarEstadoPedido(novoEstado, comentario=""){

const params = new URLSearchParams(window.location.search);
const id = params.get("id");

const token = await getAccessToken();

const site = await obterSiteApp();
const siteId = site.id;

const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

/* buscar pedido */

const respPedido = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`,
{
headers:{ Authorization:"Bearer "+token }
}
);

const dados = await respPedido.json();
const f = dados.fields;

/* carimbar PDF */

const pdfCarimbado = await carimbarPdf(f.PdfUrl, novoEstado);

/* substituir PDF no SharePoint */

const ficheiro = new Blob([pdfCarimbado],{type:"application/pdf"});

await uploadPdfSharePoint(ficheiro);

/* atualizar estado */

const body = {
fields:{
EstadoPedido: novoEstado,
ComentarioRejeicao: comentario
}
};

const resp = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}/fields`,
{
method:"PATCH",
headers:{
Authorization:"Bearer "+token,
"Content-Type":"application/json"
},
body:JSON.stringify(body)
}
);

if(!resp.ok){
alert("Erro ao atualizar estado");
return;
}

alert("Pedido "+novoEstado);

window.location.href="dashboard.html";

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
async function atualizarEstadoPedidoId(id,estado){

const token = await getAccessToken();

const site = await obterSiteApp();
const siteId = site.id;

const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

const body = {

fields:{
EstadoPedido: estado
}

};

await fetch(

`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}/fields`,

{
method:"PATCH",
headers:{
Authorization:"Bearer "+token,
"Content-Type":"application/json"
},
body:JSON.stringify(body)
}

);

}
const checkTodos = document.getElementById("checkTodos");

if(checkTodos){

checkTodos.addEventListener("change", function(){

const checks = document.querySelectorAll(".checkPedido");

checks.forEach(c => {

c.checked = this.checked;

});

});

}
function diasParaVencimento(data){

if(!data) return null;

const hoje = new Date();
const venc = new Date(data);

const diff = venc - hoje;

return Math.ceil(diff / (1000*60*60*24));

}
async function carimbarPdf(urlPdf, estado){

const { PDFDocument, rgb, StandardFonts } = PDFLib;

const token = await getAccessToken();

const resp = await fetch(urlPdf,{
headers:{ Authorization:"Bearer "+token }
});

const bytes = await resp.arrayBuffer();

const pdfDoc = await PDFDocument.load(bytes);

const pages = pdfDoc.getPages();

const page = pages[0];

const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

const utilizador = await testarGraph();

const texto =
estado + "\n" +
utilizador.displayName + "\n" +
new Date().toLocaleString("pt-PT");

page.drawText(texto,{
x:50,
y:100,
size:20,
font:font,
color: estado==="Aprovado" ? rgb(0,0.6,0) : rgb(0.8,0,0)
});

const pdfFinal = await pdfDoc.save();

return pdfFinal;

}
function atualizarDataHora(){

const agora = new Date();

const data = agora.toLocaleDateString("pt-PT",{
weekday:"short",
day:"2-digit",
month:"short",
year:"numeric"
});

const hora = agora.toLocaleTimeString("pt-PT",{
hour:"2-digit",
minute:"2-digit"
});

const el = document.getElementById("dataHora");

if(el){
el.innerText = data + " • " + hora;
}

}

setInterval(atualizarDataHora,1000);
atualizarDataHora();
if(typeof lucide !== "undefined"){
lucide.createIcons();
}
function mostrarAreaRejeicao(){

document.getElementById("areaRejeicao").style.display = "flex";

}
async function confirmarRejeicao(){

const comentario = document
.getElementById("comentarioRejeicao")
.value
.trim();

if(!comentario){

alert("Tem de escrever um comentário para rejeitar.");

return;

}

await atualizarEstadoPedido("Rejeitado", comentario);

}
function rejeitarPedido(){

mostrarAreaRejeicao();

}
async function aprovarPedido(){

await atualizarEstadoPedido("Aprovado");

}
function mostrarSegundoAprovador(){
  document.getElementById("segundoAprovadorBox").style.display = "block";
}
