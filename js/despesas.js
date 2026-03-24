/* =============================
   TABELA KM
============================= */

function addLinhaKM(){

    const tbody = document.getElementById("linhasKM");

    const tr = document.createElement("tr");

    tr.innerHTML = `
    <td><input type="date" class="data"></td>
    <td><input type="text" class="origem"></td>
    <td><input type="text" class="destino"></td>
    <td><input type="text" class="justificacao"></td>
    <td><input type="number" class="kms" oninput="calcularKM()"></td>
    <td><button onclick="removerLinha(this)">X</button></td>
    `;

    tbody.appendChild(tr);

}


/* =============================
   REMOVER LINHA
============================= */

function removerLinha(btn){
    btn.closest("tr").remove();
    calcularKM();
}


/* =============================
   CALCULAR TOTAIS
============================= */

function calcularKM(){

    let totalKMs = 0;

    document.querySelectorAll(".kms").forEach(input => {
        totalKMs += Number(input.value) || 0;
    });

    const elTotal = document.getElementById("totalKMs");
    if(elTotal){
        elTotal.innerText = totalKMs;
    }

    const valorKM = Number(document.getElementById("valorKM")?.value) || 0;

    const totalFinal = totalKMs * valorKM;

    const elFinal = document.getElementById("totalFinalKM");
    if(elFinal){
        elFinal.innerText = totalFinal.toFixed(2) + " €";
    }

}
/* =============================
   GUARDAR DESPESA KM
============================= */

async function guardarDespesaKM(){

    const utilizador = await testarGraph();
    const token = await getAccessToken();
    const site = await obterSiteApp();

    const siteId = site.id;

    const linhas = [];

    const rows = document.querySelectorAll("#linhasKM tr");

for(const tr of rows){

    const data = tr.querySelector(".data")?.value || "";
    const origem = tr.querySelector(".origem")?.value.trim() || "";
    const destino = tr.querySelector(".destino")?.value.trim() || "";
    const justificacao = tr.querySelector(".justificacao")?.value.trim() || "";
    const kms = Number(tr.querySelector(".kms")?.value) || 0;

    if(!data || !origem || !destino || !justificacao || kms <= 0){
        alert("Preencha todos os campos corretamente em todas as linhas.");
        return; // 👈 aqui sim pára tudo
    }

    linhas.push({
        data,
        origem,
        destino,
        justificacao,
        kms
    });

}

    if(linhas.length === 0){
        alert("Tem de inserir pelo menos uma linha.");
        return;
    }

    /* totais */
    let totalKMs = 0;
    linhas.forEach(l => totalKMs += l.kms);

    const valorKM = Number(document.getElementById("valorKM").value) || 0;
    const totalRecebido = totalKMs * valorKM;

    /* JSON */
    const linhasJSON = JSON.stringify(linhas);

    const listaNome = "NotasDespesa";

    const body = {
        fields: {
            Title: "Nota KM - " + new Date().toLocaleDateString("pt-PT"),
            TipoDocumento: "KMS",
            CriadoPorNome: utilizador.displayName,
            CriadoPorEmail: utilizador.mail || utilizador.userPrincipalName,
            DataCriacao: new Date().toISOString(),
            TotalKMs: totalKMs,
            ValorPorKM: valorKM,
            TotalRecebido: totalRecebido,
            LinhasJSON: linhasJSON,
            Estado: "Pendente"
        }
    };

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items`,
        {
            method: "POST",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        }
    );

    const data = await resp.json();

    console.log("Resposta SharePoint:", data);

    if(!resp.ok){
        alert("Erro ao guardar nota de despesa");
        return;
    }

    alert("Nota de despesa guardada com sucesso!");

    window.location.href = "dashboard.html";

}
async function carregarAprovadoresDespesa(){

    const aprovadores = await obterAprovadores();

    const select1 = document.getElementById("aprovador1");
    const select2 = document.getElementById("aprovador2");

    if(!select1) return;

    select1.innerHTML = `<option value="">Selecionar</option>`;

    if(select2){
        select2.innerHTML = `<option value="">Selecionar</option>`;
    }

    aprovadores.forEach(a => {

        const opt1 = document.createElement("option");
        opt1.value = a.email;
        opt1.textContent = a.nome;
        select1.appendChild(opt1);

        if(select2){
            const opt2 = document.createElement("option");
            opt2.value = a.email;
            opt2.textContent = a.nome;
            select2.appendChild(opt2);
        }

    });

}
