const PASTA_ORIGEM_ID = "1VH0jbAxuo1N9OFEEFkHV3CaseBjb4FIs";
const PLANILHA_ALVO_ID = "1Beq9Gwq3l1o19r1t-yfX-_IVMTDBi8FPUsALGc58YtM";
const PLANILHA_HISTORICO_ID = "1TNJA__LaYgi0PLIjYRn3o-dgX9zcrVN4NdnM3LahS_c";
const EMAIL_ALERTA = "rafael.guerra.unia@gmail.com";
const TIMEZONE = "America/Sao_Paulo";
const ABA_ME2W = "ME2W";
const ABA_STORE = "Confirmacoes_Store";
const COLUNAS_MANUAIS = [
  "Confirmação Smarthub",
  "Deliv Date - Confirmação Planejamento (SMART HUB)",
  "Qtd - Confirmação Planejamento (SMART HUB)",
  "Prioridade Smarthub",
  "Causa de Desvio"
];

const CHAVE_ME2W = ["Purchasing Document", "Item", "Schedule Line"];

// Fotografados no ato da confirmacao para detectar reprogramacao do SAP.
const COL_SAP_DATA = "Delivery Date";
const COL_SAP_QTD = "Order Quantity";

const STORE_HEADERS = ["Chave", "Purchasing Document", "Item", "Schedule Line"]
  .concat(COLUNAS_MANUAIS)
  .concat(["SAP Delivery Date (no ato)", "SAP Order Quantity (no ato)",
           "Atualizado em", "Atualizado por", "Visto por último no SAP", "Status"]);

const ST_CHAVE = 0, ST_DOC = 1, ST_ITEM = 2, ST_SCHED = 3;
const ST_MANUAIS = 4;
const ST_SAP_DATA = 9, ST_SAP_QTD = 10;
const ST_ATUALIZADO_EM = 11, ST_ATUALIZADO_POR = 12, ST_VISTO_EM = 13, ST_STATUS = 14;

// ====================================================================
// CATALOGO DAS BASES
// ====================================================================

// Duas rotas, escolhidas pelo que cada base tem a perder:
//
//   "critica" — so a ME2W. Carrega as confirmacoes manuais, entao paga
//               validacao do export, lock e restore do store antes de escrever.
//
//   "rapida"  — as demais. Sao insumos reconstruiveis a cada sync: nao tem dado
//               manual, ninguem escreve nelas pelo portal e um export ruim custa
//               so o proximo gatilho. Nao ha o que verificar, entao vao pelo
//               caminho de copia pura (copiarRapido), sem passar linha a linha.
//
// A acumulacao no historico saiu do sync: roda em sincronizarHistoricos(), com
// gatilho proprio, para nao disputar o orcamento de 6 min com a copia das bases.
const BASES = [
  {
    aba: ABA_ME2W,
    arquivo: "STO-ME2W.xlsx",
    modo: "critica",
    historico: ["Purchasing Document", "Item", "Material", "Order Quantity",
                "Delivery Date", "Qty Delivered", "Schedule Line"]
  },
  {
    aba: "RESB",
    arquivo: "RESB_TRANS.xlsx",
    modo: "rapida",
    historico: ["Material", "ReqmtsDate", "Reqmnt qty", "Plnd Ord.",
                "Reserv.no.", "Pegged Requirement"]
  },
  { aba: "ME5A", arquivo: "STO-ME5A.xlsx", modo: "rapida" },
  { aba: "Stock Control BR14 BR10 BR12", arquivo: "Stock_BR14_BR12_BR10.xlsx", modo: "rapida" },
  { aba: "ME2N", arquivo: "STO-ME2N.xlsx", modo: "rapida" }
];

// O gatilho morre em 6 min. Parar por conta propria antes disso e a diferenca
// entre "o proximo gatilho continua de onde parou" e "o proximo gatilho refaz
// tudo do zero e estoura de novo" — o ciclo que travava a sincronizacao.
const ORCAMENTO_MS = 4.5 * 60 * 1000;

function chaveSync(base) { return "SYNC_" + base.aba.replace(/\s+/g, "_"); }
function chaveHist(base) { return "HIST_PEND_" + base.aba.replace(/\s+/g, "_"); }
function dentroDoOrcamento(inicio) { return (Date.now() - inicio) < ORCAMENTO_MS; }

// ====================================================================
// CHAVE
// ====================================================================

// Normaliza um componente de chave. Os dois lados da comparacao passam pela
// planilha, mas nao necessariamente com o mesmo tipo: o Item pode voltar como
// numero 10 de um lado e texto "00010" do outro.
function normalizarChave(v) {
  if (v instanceof Date) return String(v.getTime());
  var s = String(v == null ? "" : v).trim();
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, "");
  return s;
}

function montarChave(doc, item, sched) {
  return normalizarChave(doc) + "|" + normalizarChave(item) + "|" + normalizarChave(sched);
}

// ====================================================================
// STORE DE CONFIRMACOES — fonte da verdade, fora do ciclo destrutivo
// ====================================================================

function obterAbaStore() {
  const ss = SpreadsheetApp.openById(PLANILHA_ALVO_ID);
  let aba = ss.getSheetByName(ABA_STORE);
  if (!aba) {
    aba = ss.insertSheet(ABA_STORE);
    aba.getRange(1, 1, 1, STORE_HEADERS.length).setValues([STORE_HEADERS]);
    aba.setFrozenRows(1);
    console.log("Aba " + ABA_STORE + " criada.");
  }
  return aba;
}

// O store nunca remove linhas, entao as chaves ocupam um bloco contiguo a partir
// da linha 2 — o que permite reescrever tudo de uma vez sem embaralhar nada.
function lerStore() {
  const aba = obterAbaStore();
  const ultimaLinha = aba.getLastRow();
  const mapa = new Map();
  if (ultimaLinha < 2) return { aba: aba, mapa: mapa };

  const dados = aba.getRange(2, 1, ultimaLinha - 1, STORE_HEADERS.length).getValues();
  for (let i = 0; i < dados.length; i++) {
    const chave = String(dados[i][ST_CHAVE]).trim();
    if (chave) mapa.set(chave, { linha: i + 2, valores: dados[i] });
  }
  return { aba: aba, mapa: mapa };
}

function gravarStore(store) {
  const existentes = [], novos = [];
  store.mapa.forEach(reg => (reg.linha ? existentes.push(reg) : novos.push(reg)));
  existentes.sort((a, b) => a.linha - b.linha);

  if (existentes.length > 0) {
    store.aba.getRange(2, 1, existentes.length, STORE_HEADERS.length)
             .setValues(existentes.map(r => r.valores));
  }
  if (novos.length > 0) {
    const inicio = 2 + existentes.length;
    store.aba.getRange(inicio, 1, novos.length, STORE_HEADERS.length)
             .setValues(novos.map(r => r.valores));
    novos.forEach((r, k) => { r.linha = inicio + k; });
  }
}

// Grava so os registros tocados. O portal salva poucas linhas por vez e nao
// pode pagar a reescrita do store inteiro a cada clique.
function gravarStoreParcial(store, regs) {
  const novos = [];
  regs.forEach(reg => {
    if (reg.linha) store.aba.getRange(reg.linha, 1, 1, STORE_HEADERS.length).setValues([reg.valores]);
    else novos.push(reg);
  });
  if (novos.length > 0) {
    const inicio = Math.max(store.aba.getLastRow(), 1) + 1;
    store.aba.getRange(inicio, 1, novos.length, STORE_HEADERS.length)
             .setValues(novos.map(r => r.valores));
    novos.forEach((r, k) => { r.linha = inicio + k; });
  }
}

function upsertStore(store, doc, item, sched, manuais, sapData, sapQtd, usuario) {
  const chave = montarChave(doc, item, sched);
  const agora = new Date();
  let reg = store.mapa.get(chave);

  if (!reg) {
    const linha = new Array(STORE_HEADERS.length).fill("");
    linha[ST_CHAVE] = chave;
    linha[ST_DOC] = doc;
    linha[ST_ITEM] = item;
    linha[ST_SCHED] = sched;
    linha[ST_VISTO_EM] = agora;
    linha[ST_STATUS] = "ATIVA";
    reg = { linha: null, valores: linha };
    store.mapa.set(chave, reg);
  }

  for (let k = 0; k < COLUNAS_MANUAIS.length; k++) reg.valores[ST_MANUAIS + k] = manuais[k];
  if (sapData !== undefined) reg.valores[ST_SAP_DATA] = sapData;
  if (sapQtd !== undefined) reg.valores[ST_SAP_QTD] = sapQtd;
  reg.valores[ST_ATUALIZADO_EM] = agora;
  reg.valores[ST_ATUALIZADO_POR] = usuario || "";
  return reg;
}

// A data vem do <input type="date"> do portal como "YYYY-MM-DD". Montar a Date
// por componentes evita o new Date(string), que interpreta conforme o locale.
function parseDataPortal(v) {
  if (v instanceof Date) return v;
  const s = String(v == null ? "" : v).trim();
  if (s === "") return "";
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return v;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// Popula o store a partir das confirmacoes que ja existem na aba ME2W viva.
// Roda uma unica vez, antes do primeiro clearContents sob a nova logica —
// sem isso, a migracao perderia todo o historico de confirmacoes.
function semearStoreDaMe2w(targetSS, store) {
  const aba = targetSS.getSheetByName(ABA_ME2W);
  if (!aba || aba.getLastRow() < 2) return 0;

  const dados = aba.getDataRange().getValues();
  const headers = dados[0].map(h => String(h).trim());
  const iDoc = headers.indexOf("Purchasing Document");
  const iItem = headers.indexOf("Item");
  const iSched = headers.indexOf("Schedule Line");
  if (iDoc === -1 || iItem === -1 || iSched === -1) {
    throw new Error("Semeadura abortada: a aba ME2W atual nao tem as colunas-chave. Cabecalho: " + headers.join(" | "));
  }

  const iSapData = headers.indexOf(COL_SAP_DATA);
  const iSapQtd = headers.indexOf(COL_SAP_QTD);
  const idxManuais = COLUNAS_MANUAIS.map(c => headers.indexOf(c));

  let n = 0;
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const manuais = idxManuais.map(idx => idx === -1 ? "" : linha[idx]);
    if (manuais.every(v => v === "" || v === null)) continue;

    upsertStore(store, linha[iDoc], linha[iItem], linha[iSched], manuais,
                iSapData !== -1 ? linha[iSapData] : "",
                iSapQtd !== -1 ? linha[iSapQtd] : "",
                "migracao");
    n++;
  }
  console.log("Store semeado com " + n + " confirmacoes vindas da aba ME2W.");
  return n;
}

// ====================================================================
// COMPARACAO SAP (deteccao de reprogramacao)
// ====================================================================

function tipoDeValor(v) {
  if (v instanceof Date) return "data";
  if (typeof v === "number") return "numero";
  return "texto";
}

function valorComparavel(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, "yyyy-MM-dd");
  if (typeof v === "number") return String(v);
  return String(v == null ? "" : v).trim();
}

// Compara o valor do SAP fotografado no ato da confirmacao com o do export atual.
// Se os dois lados vierem com tipos diferentes a comparacao nao tem sentido —
// avisa e nao acusa mudanca, para nao gerar alarme falso em massa.
function sapMudou(antes, agora, rotulo) {
  if (antes === "" || antes === null || antes === undefined) return false;
  if (tipoDeValor(antes) !== tipoDeValor(agora)) {
    console.warn("Comparacao de " + rotulo + " ignorada: tipos diferentes (" +
                 tipoDeValor(antes) + " vs " + tipoDeValor(agora) + ").");
    return false;
  }
  return valorComparavel(antes) !== valorComparavel(agora);
}

// ====================================================================
// VALIDACAO DO EXPORT DA ME2W
// ====================================================================

// A ME2W e o unico lugar onde vivem as ordens que serao transferidas. Um export
// ruim aqui apaga confirmacoes de forma irreversivel, entao ela aborta em vez de
// sobrescrever. As bases de analise nao tem esse risco e seguem normalmente.
function validarExportMe2w(dados, abaAtual) {
  if (!dados || dados.length < 2) {
    throw new Error("export sem linhas de dados (" + (dados ? dados.length : 0) + " linha(s) recebida(s)).");
  }

  const headers = dados[0].map(h => String(h).trim());
  const faltando = CHAVE_ME2W.filter(c => headers.indexOf(c) === -1);
  if (faltando.length > 0) {
    throw new Error("colunas-chave ausentes no export: " + faltando.join(", ") +
                    ". Cabecalho recebido: " + headers.join(" | "));
  }

  const linhasAtuais = abaAtual ? Math.max(abaAtual.getLastRow() - 1, 0) : 0;
  const linhasNovas = dados.length - 1;
  if (linhasAtuais > 10 && linhasNovas < linhasAtuais * 0.5) {
    throw new Error("export suspeito — " + linhasNovas + " linhas contra " + linhasAtuais +
                    " atuais (queda de mais de 50%).");
  }
  return headers;
}

// ====================================================================
// RESTORE: store -> export novo da ME2W
// ====================================================================

function aplicarStoreNaMe2w(dados, store) {
  const headers = dados[0].map(h => String(h).trim());
  const iDoc = headers.indexOf("Purchasing Document");
  const iItem = headers.indexOf("Item");
  const iSched = headers.indexOf("Schedule Line");
  const iSapData = headers.indexOf(COL_SAP_DATA);
  const iSapQtd = headers.indexOf(COL_SAP_QTD);

  // As 5 manuais sempre no fim, nesta ordem — a Pagina Transferencia depende disso.
  const base = headers.length;
  const headersFinais = headers.concat(COLUNAS_MANUAIS);
  dados[0] = headersFinais;

  const vistas = new Set();
  const agora = new Date();
  let restauradas = 0, reapareceram = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    while (linha.length < headersFinais.length) linha.push("");

    const chave = montarChave(linha[iDoc], linha[iItem], linha[iSched]);
    vistas.add(chave);

    const reg = store.mapa.get(chave);
    if (!reg) continue;

    for (let k = 0; k < COLUNAS_MANUAIS.length; k++) {
      linha[base + k] = reg.valores[ST_MANUAIS + k];
    }
    restauradas++;

    if (String(reg.valores[ST_STATUS]).trim() === "AUSENTE") {
      reapareceram++;
      const mudouData = sapMudou(reg.valores[ST_SAP_DATA], iSapData !== -1 ? linha[iSapData] : "", COL_SAP_DATA);
      const mudouQtd = sapMudou(reg.valores[ST_SAP_QTD], iSapQtd !== -1 ? linha[iSapQtd] : "", COL_SAP_QTD);

      if (mudouData || mudouQtd) {
        const oQue = mudouData && mudouQtd ? "data e quantidade" : (mudouData ? "data" : "quantidade");
        const aviso = "⚠️ Reapareceu no SAP com " + oQue + " diferente — revalidar";
        const causaAtual = String(linha[base + 4] || "").trim();
        linha[base + 4] = causaAtual ? aviso + " | " + causaAtual : aviso;
      }
      reg.valores[ST_STATUS] = "REAPARECEU";
    } else {
      reg.valores[ST_STATUS] = "ATIVA";
    }
    reg.valores[ST_VISTO_EM] = agora;
  }

  // Ordens do store que nao vieram no export: marcadas ausentes, nunca apagadas.
  let ausentes = 0;
  store.mapa.forEach((reg, chave) => {
    if (vistas.has(chave)) return;
    if (String(reg.valores[ST_STATUS]).trim() !== "AUSENTE") {
      reg.valores[ST_STATUS] = "AUSENTE";
      ausentes++;
    }
  });

  return { restauradas: restauradas, reapareceram: reapareceram, ausentes: ausentes };
}

// ====================================================================
// LEITURA DAS BASES
// ====================================================================

function buscarArquivoPorNome(pasta, nomeArquivo) {
  if (!pasta) {
    throw new Error("Erro: A função 'buscarArquivoPorNome' não pode ser rodada diretamente. Execute a função 'sincronizarNovasBases'.");
  }
  const arquivos = pasta.getFilesByName(nomeArquivo);
  if (arquivos.hasNext()) {
    return arquivos.next();
  }
  throw new Error("Arquivo não encontrado na pasta: " + nomeArquivo);
}

function obterOuCriarAba(planilha, nomeAba) {
  return planilha.getSheetByName(nomeAba) || planilha.insertSheet(nomeAba);
}

function atualizarAba(targetSS, nomeAba, dados) {
  if (!dados || dados.length === 0 || (dados.length === 1 && dados[0][0] === "")) return;
  const aba = obterOuCriarAba(targetSS, nomeAba);
  aba.clearContents();
  aba.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
}

function converterParaSheets(fileObj, nomeTemp) {
  return Drive.Files.create({ name: nomeTemp, mimeType: MimeType.GOOGLE_SHEETS }, fileObj.getBlob()).id;
}

// Leitura pelo SpreadsheetApp: devolve Date de verdade nas colunas de data, que
// e o que o store da ME2W compara. So a rota critica passa por aqui.
function lerValoresDoTemp(tempId) {
  const aba = SpreadsheetApp.openById(tempId).getSheets()[0];
  return aba.getRange(1, 1, Math.max(aba.getLastRow(), 1), Math.max(aba.getLastColumn(), 1)).getValues();
}

// ====================================================================
// CAMINHO RAPIDO — copia pura, sem verificacao
// ====================================================================

// Nome de aba em notacao A1. Sem as aspas, "Stock Control BR14 BR10 BR12"
// vira um intervalo invalido.
function faixaA1(nome) {
  return "'" + String(nome).replace(/'/g, "''") + "'";
}

function lerMetadadosAlvo() {
  const info = Sheets.Spreadsheets.get(PLANILHA_ALVO_ID, {
    fields: "sheets.properties(sheetId,title,gridProperties)"
  });
  const mapa = new Map();
  info.sheets.forEach(s => mapa.set(s.properties.title, {
    sheetId: s.properties.sheetId,
    linhas: s.properties.gridProperties.rowCount,
    colunas: s.properties.gridProperties.columnCount
  }));
  return mapa;
}

// So cresce a grade, nunca encolhe: apagar linha ou coluna de uma aba que a
// Pagina Transferencia referencia por intervalo produz #REF! irreversivel —
// renomear ou recriar a aba depois nao desfaz.
function garantirGrade(meta, nomeAba, linhas, colunas) {
  let alvo = meta.get(nomeAba);

  if (!alvo) {
    const r = Sheets.Spreadsheets.batchUpdate({
      requests: [{
        addSheet: {
          properties: {
            title: nomeAba,
            gridProperties: { rowCount: Math.max(linhas, 1000), columnCount: Math.max(colunas, 26) }
          }
        }
      }]
    }, PLANILHA_ALVO_ID);
    const p = r.replies[0].addSheet.properties;
    alvo = { sheetId: p.sheetId, linhas: p.gridProperties.rowCount, colunas: p.gridProperties.columnCount };
    meta.set(nomeAba, alvo);
    return alvo;
  }

  const requests = [];
  if (linhas > alvo.linhas) {
    requests.push({ appendDimension: { sheetId: alvo.sheetId, dimension: "ROWS", length: linhas - alvo.linhas } });
  }
  if (colunas > alvo.colunas) {
    requests.push({ appendDimension: { sheetId: alvo.sheetId, dimension: "COLUMNS", length: colunas - alvo.colunas } });
  }
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, PLANILHA_ALVO_ID);
    alvo.linhas = Math.max(alvo.linhas, linhas);
    alvo.colunas = Math.max(alvo.colunas, colunas);
  }
  return alvo;
}

// A copia rapida transporta data como numero de serie (nao como Date), porque
// serial nao depende de locale para ser lido nem escrito. O preco e que a coluna
// apareceria como 45123 se o formato numerico nao viesse junto — entao ele vem,
// copiado da primeira linha de dados da origem.
function replicarFormatoNumerico(tempId, tituloOrigem, alvo, colunas) {
  const info = Sheets.Spreadsheets.get(tempId, {
    ranges: [faixaA1(tituloOrigem) + "!A2:2"],
    includeGridData: true,
    fields: "sheets(data(rowData(values(userEnteredFormat(numberFormat)))))"
  });

  const dados = info.sheets && info.sheets[0] && info.sheets[0].data;
  const rowData = dados && dados[0] && dados[0].rowData;
  const celulas = rowData && rowData[0] && rowData[0].values;
  if (!celulas) return;

  const requests = [];
  for (let c = 0; c < Math.min(celulas.length, colunas); c++) {
    const fmt = celulas[c] && celulas[c].userEnteredFormat && celulas[c].userEnteredFormat.numberFormat;
    if (!fmt) continue;
    requests.push({
      repeatCell: {
        range: { sheetId: alvo.sheetId, startRowIndex: 1, startColumnIndex: c, endColumnIndex: c + 1 },
        cell: { userEnteredFormat: { numberFormat: fmt } },
        fields: "userEnteredFormat.numberFormat"
      }
    });
  }
  if (requests.length > 0) Sheets.Spreadsheets.batchUpdate({ requests: requests }, PLANILHA_ALVO_ID);
}

// Copia pura: le e escreve em blocos, sem percorrer linha a linha, sem montar
// chave e sem construir objeto Date — e o que sobra quando a base nao tem nada
// a verificar. Vai em blocos porque a base inteira num unico corpo de requisicao
// esbarra no limite de payload justamente nos arquivos que mais doem (RESB,
// ME2N, Stock).
const LIMITE_CELULAS = 400000;

function copiarRapido(nomeAba, tempId, meta) {
  const tempInfo = Sheets.Spreadsheets.get(tempId, { fields: "sheets.properties(title,gridProperties)" });
  const origem = tempInfo.sheets[0].properties;
  const tituloOrigem = origem.title;
  const totalLinhas = origem.gridProperties.rowCount;
  const totalColunas = Math.max(origem.gridProperties.columnCount, 1);
  const passo = Math.max(1, Math.floor(LIMITE_CELULAS / totalColunas));

  let alvo = null;
  let escritas = 0;
  let colunasVistas = 0;

  for (let primeira = 1; primeira <= totalLinhas; primeira += passo) {
    const ultima = Math.min(primeira + passo - 1, totalLinhas);
    const resp = Sheets.Spreadsheets.Values.get(tempId,
      faixaA1(tituloOrigem) + "!" + primeira + ":" + ultima, {
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
      });

    const valores = resp.values || [];
    if (valores.length === 0) break;

    for (let i = 0; i < valores.length; i++) {
      // Linha vazia no meio do bloco volta como [] e desalinharia a escrita.
      if (valores[i].length === 0) valores[i] = [""];
      else if (valores[i].length > colunasVistas) colunasVistas = valores[i].length;
    }

    // A grade cresce pelo que foi mesmo lido, nao pelo grid da origem: XLSX
    // convertido costuma vir com folga de linhas vazias no fim, e a grade daqui
    // nunca encolhe de volta.
    const primeiroBloco = (alvo === null);
    alvo = garantirGrade(meta, nomeAba, escritas + valores.length, totalColunas);

    // Limpa uma vez so, e so depois de ter dado em maos: um export vazio nao
    // pode zerar a aba que a Pagina Transferencia le.
    if (primeiroBloco) Sheets.Spreadsheets.Values.clear({}, PLANILHA_ALVO_ID, faixaA1(nomeAba));

    Sheets.Spreadsheets.Values.update({ values: valores }, PLANILHA_ALVO_ID,
      faixaA1(nomeAba) + "!A" + (escritas + 1), { valueInputOption: "RAW" });
    escritas += valores.length;

    // Bloco mais curto que o pedido = so restava linha vazia daqui pra frente.
    if (valores.length < (ultima - primeira + 1)) break;
  }

  if (alvo === null) throw new Error("export vazio — aba preservada.");
  replicarFormatoNumerico(tempId, tituloOrigem, alvo, colunasVistas);

  return escritas - 1;
}

// ====================================================================
// SYNC
// ====================================================================

function processarMe2w(tempId) {
  const targetSS = SpreadsheetApp.openById(PLANILHA_ALVO_ID);
  const dados = lerValoresDoTemp(tempId);
  validarExportMe2w(dados, targetSS.getSheetByName(ABA_ME2W));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error("lock ocupado por 60s (portal escrevendo) — adiado para o próximo gatilho.");
  }
  try {
    const store = lerStore();
    if (store.mapa.size === 0) semearStoreDaMe2w(targetSS, store);

    const stats = aplicarStoreNaMe2w(dados, store);
    atualizarAba(targetSS, ABA_ME2W, dados);
    gravarStore(store);
    SpreadsheetApp.flush();

    console.log("ME2W: " + (dados.length - 1) + " linhas | " +
                stats.restauradas + " confirmações restauradas | " +
                stats.reapareceram + " reapareceram | " +
                stats.ausentes + " sumiram do export (preservadas no store).");
  } finally {
    lock.releaseLock();
  }
}

function sincronizarNovasBases() {
  const inicio = Date.now();
  const now = new Date();

  if (now.getDay() === 0 || now.getHours() < 1) {
    console.log("Fora da janela de execução (" + Utilities.formatDate(now, TIMEZONE, "EEE HH:mm") + ") — sync pulado.");
    return;
  }

  const pasta = DriveApp.getFolderById(PASTA_ORIGEM_ID);
  const props = PropertiesService.getScriptProperties();
  const erros = [];
  const adiadas = [];
  let meta = null;
  let escreveu = 0;

  // A rota rapida depende do servico avancado Sheets. Habilitar pelo manifesto
  // so vale se o appsscript.json chegou ao projeto — colar apenas o Sync.gs no
  // editor deixa o servico desligado, e o ReferenceError saia uma vez por base
  // sem dizer o que fazer. Falha uma vez so, com o passo a passo, e sem gastar
  // conversao de XLSX que seria descartada.
  const temSheets = (typeof Sheets !== "undefined");
  if (!temSheets) {
    erros.push("serviço avançado 'Sheets' não habilitado no projeto — as bases de cópia rápida (" +
               BASES.filter(b => b.modo === "rapida").map(b => b.aba).join(", ") +
               ") foram puladas, sem tocar nas abas. Habilite em Editor > Serviços (+) > " +
               "Google Sheets API (identificador 'Sheets', versão v4) e rode a sincronização " +
               "uma vez na mão para reautorizar.");
  }

  // Uma base por vez: converte, escreve, apaga o temporario e so entao carimba.
  // Converter as cinco de uma vez antes de escrever gastava o orcamento inteiro
  // antes da primeira linha entrar na planilha — e um estouro no meio jogava
  // fora as cinco conversoes.
  for (let i = 0; i < BASES.length; i++) {
    const base = BASES[i];
    let tempId = null;

    try {
      const arquivo = buscarArquivoPorNome(pasta, base.arquivo);
      const carimbo = String(arquivo.getLastUpdated().getTime());

      // Carimbo por arquivo. O carimbo unico dos cinco fazia um export novo de
      // ME2W arrastar RESB, ME2N e Stock inteiros junto, sem nada ter mudado.
      if (props.getProperty(chaveSync(base)) === carimbo) {
        console.log(base.aba + ": arquivo inalterado — pulado.");
        continue;
      }

      if (base.modo === "rapida" && !temSheets) continue;

      if (!dentroDoOrcamento(inicio)) {
        adiadas.push(base.aba);
        continue;
      }

      console.log("Convertendo " + base.arquivo + "...");
      tempId = converterParaSheets(arquivo, "Temp_XLSX_" + base.aba);

      if (base.modo === "critica") {
        processarMe2w(tempId);
      } else {
        if (!meta) meta = lerMetadadosAlvo();
        const linhas = copiarRapido(base.aba, tempId, meta);
        console.log(base.aba + ": " + linhas + " linhas copiadas (sem verificação).");
      }

      props.setProperty(chaveSync(base), carimbo);
      if (base.historico) props.setProperty(chaveHist(base), carimbo);
      escreveu++;
    } catch (e) {
      erros.push(base.aba + ": " + e.message);
      console.error(base.aba + " falhou: " + e.message);
    } finally {
      if (tempId) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {} }
    }
  }

  if (adiadas.length > 0) {
    console.log("Adiadas por orçamento de tempo: " + adiadas.join(", ") +
                " — o próximo gatilho continua daqui (as concluídas não se repetem).");
  }

  if (erros.length > 0) {
    console.error("Sincronização concluída COM FALHAS: " + erros.join(" || "));
    notificarFalha(erros.join("\n"));
  } else if (escreveu > 0) {
    console.log("Sincronização concluída: " + escreveu + " base(s) atualizada(s) em " +
                Math.round((Date.now() - inicio) / 1000) + "s.");
  } else {
    console.log("Nenhum arquivo mudou — nada a fazer.");
  }
}

// ====================================================================
// HISTORICO — gatilho proprio, fora do orcamento do sync
// ====================================================================

// Mesma normalizacao dos dois lados da comparacao: data vira epoch, o resto vai
// como esta. Sem isso a mesma linha entraria de novo a cada execucao.
function normalizarHistorico(v) {
  return v instanceof Date ? v.getTime() : v;
}

function acumularHistorico(targetSS, histSS, base) {
  const aba = targetSS.getSheetByName(base.aba);
  if (!aba || aba.getLastRow() < 2) return 0;

  const dados = aba.getDataRange().getValues();
  const cabecalho = dados[0].map(h => String(h).trim());
  const idxOrigem = base.historico.map(c => cabecalho.indexOf(c));

  const nomeHist = base.aba + "-Historico";
  let abaHist = histSS.getSheetByName(nomeHist);
  const vistos = new Set();

  if (!abaHist) {
    abaHist = histSS.insertSheet(nomeHist);
    const cabecalhoHist = ["Data Cópia Histórico"].concat(cabecalho);
    abaHist.getRange(1, 1, 1, cabecalhoHist.length).setValues([cabecalhoHist]);
  } else {
    const ultima = abaHist.getLastRow();
    if (ultima > 1) {
      const cabecalhoHist = abaHist.getRange(1, 1, 1, abaHist.getLastColumn())
                                   .getValues()[0].map(h => String(h).trim());
      const idxHist = base.historico.map(c => cabecalhoHist.indexOf(c));

      // So as colunas-chave. O getDataRange().getValues() desta aba lia todas as
      // colunas de um historico que so cresce — era o custo que aumentava
      // sozinho a cada sync ate estourar o tempo.
      const colunas = idxHist.map(idx =>
        idx === -1 ? null : abaHist.getRange(2, idx + 1, ultima - 1, 1).getValues());

      for (let i = 0; i < ultima - 1; i++) {
        vistos.add(colunas.map(col => col === null ? "" : normalizarHistorico(col[i][0])).join("_"));
      }
    }
  }

  const novas = [];
  const agora = new Date();
  for (let i = 1; i < dados.length; i++) {
    const chave = idxOrigem.map(idx => idx === -1 ? "" : normalizarHistorico(dados[i][idx])).join("_");
    if (vistos.has(chave)) continue;
    vistos.add(chave);
    novas.push([agora].concat(dados[i]));
  }

  if (novas.length > 0) {
    abaHist.getRange(Math.max(abaHist.getLastRow(), 1) + 1, 1, novas.length, novas[0].length)
           .setValues(novas);
  }
  return novas.length;
}

// Le a aba ja sincronizada na planilha alvo — nao reconverte o XLSX. Roda em
// gatilho separado justamente para que o custo do dedup, que cresce com o
// historico, nunca mais dispute os 6 min da copia das bases.
function sincronizarHistoricos() {
  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  const pendentes = BASES.filter(b => b.historico && props.getProperty(chaveHist(b)));

  if (pendentes.length === 0) {
    console.log("Nenhum histórico pendente.");
    return;
  }

  const targetSS = SpreadsheetApp.openById(PLANILHA_ALVO_ID);
  const histSS = SpreadsheetApp.openById(PLANILHA_HISTORICO_ID);
  const erros = [];

  for (let i = 0; i < pendentes.length; i++) {
    const base = pendentes[i];
    if (!dentroDoOrcamento(inicio)) {
      console.log("Histórico de " + base.aba + " adiado por orçamento de tempo — segue no próximo gatilho.");
      continue;
    }
    try {
      const n = acumularHistorico(targetSS, histSS, base);
      props.deleteProperty(chaveHist(base));
      console.log(base.aba + "-Historico: " + n + " linha(s) nova(s).");
    } catch (e) {
      erros.push(base.aba + "-Historico: " + e.message);
      console.error(base.aba + "-Historico falhou: " + e.message);
    }
  }

  if (erros.length > 0) notificarFalha(erros.join("\n"));
}

// ====================================================================
// GATILHOS
// ====================================================================

// Recria os dois gatilhos do zero. O historico roda numa frequencia menor: ele
// so precisa alcancar o sync, nao acompanha-lo.
function instalarGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "sincronizarNovasBases" || fn === "sincronizarHistoricos") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("sincronizarNovasBases").timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger("sincronizarHistoricos").timeBased().everyHours(1).create();
  console.log("Gatilhos instalados: sync a cada 15 min, histórico a cada 1 h.");
}

// Forca o proximo sync a reimportar tudo, ignorando os carimbos por arquivo.
// Util depois de mexer manualmente numa aba de base.
function forcarRessincronizacao() {
  const props = PropertiesService.getScriptProperties();
  BASES.forEach(b => props.deleteProperty(chaveSync(b)));
  console.log("Carimbos limpos — o próximo gatilho reimporta todas as bases.");
}

function notificarFalha(detalhe) {
  try {
    MailApp.sendEmail({
      to: EMAIL_ALERTA,
      subject: "[Portal de Transferência] Falha na sincronização",
      body: "A sincronização falhou em " + Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm") +
            ".\n\nDetalhe:\n" + detalhe +
            "\n\nAs confirmações manuais estão preservadas na aba " + ABA_STORE + "."
    });
  } catch (e) {
    console.error("Falha ao enviar e-mail de alerta: " + e);
  }
}