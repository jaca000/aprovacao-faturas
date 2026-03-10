console.log("Aplicação iniciada");

async function iniciarApp() {

    await iniciarLogin();

    const email = sessionStorage.getItem("userEmail");

    if (email) {

        console.log("Utilizador autenticado:", email);

        window.location.href = "pages/dashboard.html";

    } else {

        console.log("A aguardar login...");

    }

}

iniciarApp();
