console.log("App iniciada");

document.getElementById("app").innerHTML = `
<button onclick="login()">Iniciar sessão Microsoft</button>
`;

function mostrarDashboard(utilizador){

    document.getElementById("app").innerHTML = `

        <h2>Dashboard</h2>

        <p>Bem-vindo ${utilizador.name} | Monte do Pasto</p>

        <button onclick="testarLigacao()">Testar ligação</button>

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

    const pedidos = await obterPedidosFaturas();

    const lista = pedidos.value;

    const total = lista.length;

    let pendentes = 0;
    let aprovados = 0;
    let rejeitados = 0;

    lista.forEach(p => {

        const estado = p.fields.Estado;

        if(estado === "Pendente") pendentes++;
        if(estado === "Aprovado") aprovados++;
        if(estado === "Rejeitado") rejeitados++;

    });

    document.getElementById("totalPedidos").innerText = total;
    document.getElementById("pendentes").innerText = pendentes;
    document.getElementById("aprovados").innerText = aprovados;
    document.getElementById("rejeitados").innerText = rejeitados;

}
