console.log("Aplicação iniciada");

async function iniciarApp() {

    await iniciarLogin();

    const email = sessionStorage.getItem("userEmail");

    if (email) {

        console.log("Utilizador:", email);

        window.location.href = "pages/dashboard.html";

    }

}

iniciarApp();
