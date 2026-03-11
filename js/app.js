console.log("App iniciada");

document.getElementById("app").innerHTML = `
<button onclick="login()">Iniciar sessão Microsoft</button>
`;

function mostrarDashboard(utilizador){

    document.getElementById("app").innerHTML = `

        <h2>Dashboard</h2>

        <p>Bem-vindo ${utilizador.name}</p>

        <button onclick="testar()">Testar ligação</button>

    `;

}

function testar(){

    console.log("Dashboard ativo");

}
