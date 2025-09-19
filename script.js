let baseDados = [];
let fila = [];

const uploadInput = document.getElementById('uploadExcel');
const customBtn = document.getElementById('customUploadBtn');

customBtn.addEventListener('click', () => {
  uploadInput.click();
});

uploadInput.addEventListener('change', (e) => {
  if (uploadInput.files.length > 0) {
    alert('Arquivo selecionado: ' + uploadInput.files[0].name);
  }
});

document.getElementById("uploadExcel").addEventListener("change", function (e) {
  let file = e.target.files[0];
  if (!file) return;

  let reader = new FileReader();
  reader.onload = function (e) {
    let data = new Uint8Array(e.target.result);
    let workbook = XLSX.read(data, { type: "array" });
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    baseDados = XLSX.utils.sheet_to_json(sheet);
    alert("Planilha carregada com sucesso!");
    console.log("Base carregada:", baseDados);
  };
  reader.readAsArrayBuffer(file);
});

function buscarZUI(codigo) {
  codigo = String(codigo).trim();
  const material = baseDados.find(item => String(item.Material).trim() === codigo);
  if (!material) {
    alert("Material não encontrado!");
    return null;
  }
  return material.ZUI;
}

function adicionarFila() {
  let material = document.getElementById("material").value;
  let nome = document.getElementById("nome").value;
  let lote = document.getElementById("lote").value;
  let validade = document.getElementById("validade").value;
  let hora = document.getElementById("hora").value;
  let quantidade = parseInt(document.getElementById("quantidade").value) || 1;

  let zui = buscarZUI(material);
  if (!zui) return;

  for (let i = 0; i < quantidade; i++) {
    fila.push({ material, nome, lote, validade, hora, zui, quantidade: 1 });
  }
  renderFila();
}

function renderFila() {
  let filaDiv = document.getElementById("fila");
  filaDiv.innerHTML = "";

  fila.forEach((etq, i) => {
    let div = document.createElement("div");
    div.className = "etiqueta";
    div.innerHTML = `
      <b>${etq.nome}</b><br>
      COD: ${etq.material}<br>
      L: ${etq.lote} &nbsp; V: ${etq.validade} ${etq.hora}<br>
      ZUI: ${etq.zui}
      <button class="remover" onclick="removerEtiqueta(${i})">X</button>
    `;
    filaDiv.appendChild(div);
  });
}

function removerEtiqueta(index) {
  fila.splice(index, 1);
  renderFila();
}

// Ajusta nome e retorna array de linhas para quebrar automaticamente
function quebrarNome(nome, maxCharsPorLinha) {
  let palavras = nome.split(" ");
  let linhas = [];
  let linhaAtual = "";

  palavras.forEach(palavra => {
    if ((linhaAtual + " " + palavra).trim().length <= maxCharsPorLinha) {
      linhaAtual = (linhaAtual + " " + palavra).trim();
    } else {
      if (linhaAtual) linhas.push(linhaAtual);
      linhaAtual = palavra;
    }
  });

  if (linhaAtual) linhas.push(linhaAtual);

  return linhas;
}

// Imprimir fila em ZPL
function imprimirFila() {
  if (fila.length === 0) {
    alert("Fila vazia!");
    return;
  }

  let zplCompleto = "";

  fila.forEach(etq => {
    // Quebrar nome em 1 ou 2 linhas (máx 18 chars por linha, ajuste conforme necessário)
    let linhasNome = quebrarNome(etq.nome, 18);
    let alturaLinha = 21; // TAMANHO DA FONTE DO NOME
    let yNome = 27;       // posição inicial do nome
    let totalAlturaNome = alturaLinha * linhasNome.length;

    // Posição vertical das infos COD, L, V (logo abaixo do nome)
    let yCOD = yNome + totalAlturaNome + 5; // +5 dots de espaçamento
    let yL = yCOD + 30;
    let yV = yL + 30;

    // Montar linhas do nome
    let zplNome = "";
    linhasNome.forEach((linha, idx) => {
      zplNome += `^FO370,${yNome + (alturaLinha*idx)}^A0N,20,20^FD${linha}^FS\n`; //REPETIR AJUSTES DA FONTE DO NOME
    });

    zplCompleto += `
^XA
^MMT
^PW559
^LL160
^LS0

^BY3,2,92^FT58,118^BEN,,Y,N
^FH\\^FD${etq.zui}^FS            ; Código de barras dinâmico

${zplNome}
^FO365,${yCOD}^A0N,25,25^FD COD: ${etq.material}^FS
^FO365,${yL}^A0N,25,25^FD L: ${etq.lote}^FS
^FO365,${yV}^A0N,25,25^FD V: ${etq.validade} ${etq.hora}^FS

^PQ1,0,1,Y
^XZ
    `;
  });

  fetch("IP IMPRESSORA", {
    method: "POST",
    body: zplCompleto
  })
  .then(() => alert("Impressão enviada!"))
  .catch(err => console.error("Erro ao imprimir:", err));
}
