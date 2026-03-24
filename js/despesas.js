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

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";
const aprovador1 = document.getElementById("aprovador1")?.value || "";
const aprovador2 = document.getElementById("aprovador2")?.value || "";

if(!aprovador1){
    alert("Tem de selecionar um aprovador.");
    return;
}
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
            Estado: "Pendente",
Aprovador1Email: aprovador1,
Aprovador2Email: aprovador2,
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
async function obterAprovadores(){

    const token = await getAccessToken();
    const site = await obterSiteApp();

    const siteId = site.id;

    const listaNome = "AprovadoresApp";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();

    console.log("APROVADORES RAW:", data); // 👈 IMPORTANTE

    return data.value.map(item => ({
        nome: item.fields.NomeAprovador,
        email: item.fields.EmailAprovador
    }));

}
async function gerarPdfKM(linhas, totalKMs, valorKM, totalRecebido, utilizador){

    const { PDFDocument, StandardFonts, rgb } = PDFLib;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage();

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const { width, height } = page.getSize();

    let y = height - 50;

    // Título
    page.drawText("Nota de Despesa - Deslocações", {
        x: 50,
        y,
        size: 18,
        font: bold
    });

    y -= 30;

    // Info utilizador
    page.drawText("Colaborador: " + utilizador.displayName, {
        x: 50,
        y,
        size: 10,
        font
    });

    y -= 20;

    page.drawText("Data: " + new Date().toLocaleDateString("pt-PT"), {
        x: 50,
        y,
        size: 10,
        font
    });

    y -= 30;

    // Cabeçalho tabela
    const headers = ["Data", "Origem", "Destino", "Justificação", "KMs"];

    let x = 50;

    headers.forEach(h => {
        page.drawText(h, { x, y, size: 10, font: bold });
        x += 100;
    });

    y -= 20;

    // Linhas
    linhas.forEach(l => {

        let x = 50;

        page.drawText(l.data || "", { x, y, size: 9, font }); x += 100;
        page.drawText(l.origem || "", { x, y, size: 9, font }); x += 100;
        page.drawText(l.destino || "", { x, y, size: 9, font }); x += 100;
        page.drawText(l.justificacao || "", { x, y, size: 9, font }); x += 100;
        page.drawText(String(l.kms), { x, y, size: 9, font });

        y -= 15;

    });

    y -= 20;

    // Totais
    page.drawText("Total KMs: " + totalKMs, {
        x: 50,
        y,
        size: 11,
        font: bold
    });

    y -= 15;

    page.drawText("Valor por KM: " + valorKM + " €", {
        x: 50,
        y,
        size: 11,
        font: bold
    });

    y -= 15;

    page.drawText("Total a receber: " + totalRecebido.toFixed(2) + " €", {
        x: 50,
        y,
        size: 12,
        font: bold,
        color: rgb(0,0.6,0)
    });

    const pdfBytes = await pdfDoc.save();

    return pdfBytes;
}
async function verPdfKM(id){

    const token = await getAccessToken();
    const site = await obterSiteApp();
    const siteId = site.id;

    const listaNome = "NotasDespesa";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items/${id}?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();
    const f = data.fields;

    const linhas = JSON.parse(f.LinhasJSON || "[]");

    const utilizador = {
        displayName: f.CriadoPorNome
    };

    const pdfBytes = await gerarPdfKM(
        linhas,
        f.TotalKMs,
        f.ValorPorKM,
        f.TotalRecebido,
        utilizador
    );

    // 🔥 abrir no browser
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    window.open(url, "_blank");
}
