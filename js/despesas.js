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
