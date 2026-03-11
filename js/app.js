console.log("Aplicação iniciada");

async function iniciarApp() {

    await iniciarLogin();

    const email = sessionStorage.getItem("userEmail");

    if (email) {

        console.log("Utilizador autenticado:", email);

        carregarDashboard();

    }

}

function carregarDashboard(){

    document.getElementById("conteudo").innerHTML = `
    
    <h2>Dashboard</h2>
    
    <p>Aplicação em construção</p>
    
    <button onclick="testarGraph()">Testar ligação Graph</button>

    `;

}

iniciarApp();
