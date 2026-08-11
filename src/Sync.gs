const PASTA_ORIGEM_ID = "1VH0jbAxuo1N9OFEEFkHV3CaseBjb4FIs";
const PLANILHA_ALVO_ID = "1Beq9Gwq3l1o19r1t-yfX-_IVMTDBi8FPUsALGc58YtM";
const PLANILHA_HISTORICO_ID = "1TNJA__LaYgi0PLIjYRn3o-dgX9zcrVN4NdnM3LahS_c";
const EMAIL_ALERTA = "rafael.guerra.unia@gmail.com";
const TIMEZONE = "America/Sao_Paulo";
const ABA_ME2W = "ME2W";
const ABA_STORE = "Confirmacoes_Store";

// A Pagina Transferencia le a ME2W por VLOOKUP com indice fixo. A aba precisa
// terminar com exatamente estas 5 colunas, nesta ordem: mudar a largura da ME2W
// quebra todas as formulas de uma vez.
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

function atualizarAbaComHistorico(targetSS, histSS, nomeAba, dados, chavesUnicas) {
  if (!dados || dados.length === 0 || (dados.length === 1 && dados[0][0] === "")) return;
  atualizarAba(targetSS, nomeAba, dados);

  const cabecalhoOrigem = dados[0].map(h => String(h).trim());
  const idxChavesOrigem = chavesUnicas.map(colName => cabecalhoOrigem.indexOf(colName));
  const nomeAbaHist = `${nomeAba}-Historico`;
  let abaHist = histSS.getSheetByName(nomeAbaHist);
  let mapExistente = new Map();

  if (!abaHist) {
    abaHist = histSS.insertSheet(nomeAbaHist);
    const cabecalhoHist = ["Data Cópia Histórico"].concat(cabecalhoOrigem);
    abaHist.getRange(1, 1, 1, cabecalhoHist.length).setValues([cabecalhoHist]);
  } else {
    const dadosHist = abaHist.getDataRange().getValues();
    if (dadosHist.length > 1) {
      const cabecalhoHistLocal = dadosHist[0].map(h => String(h).trim());
      const idxChavesHist = chavesUnicas.map(colName => cabecalhoHistLocal.indexOf(colName));

      for (let i = 1; i < dadosHist.length; i++) {
        const linhaHist = dadosHist[i];
        const arrChave = idxChavesHist.map(idx => {
          if (idx === -1) return "";
          return linhaHist[idx] instanceof Date ? linhaHist[idx].getTime() : linhaHist[idx];
        });
        mapExistente.set(arrChave.join("_"), true);
      }
    }
  }

  const linhasParaInserir = [];
  const timestampAgora = new Date();

  for (let i = 1; i < dados.length; i++) {
    const linhaOrigem = dados[i];
    const arrChaveAtual = idxChavesOrigem.map(idx => {
      if (idx === -1) return "";
      return linhaOrigem[idx] instanceof Date ? linhaOrigem[idx].getTime() : linhaOrigem[idx];
    });
    const chaveComposta = arrChaveAtual.join("_");

    if (!mapExistente.has(chaveComposta)) {
      linhasParaInserir.push([timestampAgora].concat(linhaOrigem));
    }
  }

  if (linhasParaInserir.length > 0) {
    const ultimaLinha = Math.max(abaHist.getLastRow(), 1);
    abaHist.getRange(ultimaLinha + 1, 1, linhasParaInserir.length, linhasParaInserir[0].length).setValues(linhasParaInserir);
  }
}

// ====================================================================
// SYNC
// ====================================================================

function sincronizarNovasBases() {
  const now = new Date();
  const hour = now.getHours();
  const day = now.getDay();

  if (day === 0 || day === 7 || hour < 1 || hour >= 24) {
    console.log("Fora da janela de execução (" + Utilities.formatDate(now, TIMEZONE, "EEE HH:mm") + ") — sync pulado.");
    return;
  }

  const pasta = DriveApp.getFolderById(PASTA_ORIGEM_ID);

  const fileResb  = buscarArquivoPorNome(pasta, "RESB_TRANS.xlsx");
  const fileMe5a  = buscarArquivoPorNome(pasta, "STO-ME5A.xlsx");
  const fileMe2w  = buscarArquivoPorNome(pasta, "STO-ME2W.xlsx");
  const fileStock = buscarArquivoPorNome(pasta, "Stock_BR14_BR12_BR10.xlsx");
  const fileMe2n  = buscarArquivoPorNome(pasta, "STO-ME2N.xlsx");

  const updatedTime = fileResb.getLastUpdated().getTime().toString() + "_" +
                      fileMe5a.getLastUpdated().getTime().toString() + "_" +
                      fileMe2w.getLastUpdated().getTime().toString() + "_" +
                      fileStock.getLastUpdated().getTime().toString() + "_" +
                      fileMe2n.getLastUpdated().getTime().toString();

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_SYNC_NOVAS_BASES') === updatedTime) {
    console.log("Nenhum dos arquivos Excel sofreu modificações. Sincronização abortada.");
    return;
  }

  console.log("Alterações detectadas. Iniciando conversão...");
  const tempIds = [];
  const erros = [];

  const extrairDadosExcel = (fileObj, tempName) => {
    const tempFile = Drive.Files.create({ name: tempName, mimeType: MimeType.GOOGLE_SHEETS }, fileObj.getBlob());
    tempIds.push(tempFile.id);
    const tempSS = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSS.getSheets()[0];
    return tempSheet.getRange(1, 1, Math.max(tempSheet.getLastRow(), 1), Math.max(tempSheet.getLastColumn(), 1)).getValues();
  };

  try {
    const targetSS = SpreadsheetApp.openById(PLANILHA_ALVO_ID);
    const histSS = SpreadsheetApp.openById(PLANILHA_HISTORICO_ID);

    // 1) Conversao dos XLSX — a parte lenta, deliberadamente fora do lock.
    console.log("Convertendo bases...");
    const dadosResb  = extrairDadosExcel(fileResb, "Temp_XLSX_RESB");
    const dadosMe2w  = extrairDadosExcel(fileMe2w, "Temp_XLSX_ME2W");
    const dadosMe5a  = extrairDadosExcel(fileMe5a, "Temp_XLSX_ME5A");
    const dadosStock = extrairDadosExcel(fileStock, "Temp_XLSX_Stock");
    const dadosMe2n  = extrairDadosExcel(fileMe2n, "Temp_XLSX_ME2N");

    // 2) ME2W: unica base com dado insubstituivel. Valida antes de destruir e
    //    escreve sob lock, para nao competir com os saves do portal.
    try {
      console.log("Processando ME2W...");
      validarExportMe2w(dadosMe2w, targetSS.getSheetByName(ABA_ME2W));

      const lock = LockService.getScriptLock();
      if (!lock.tryLock(60000)) {
        throw new Error("lock ocupado por 60s (portal escrevendo) — adiado para o próximo gatilho.");
      }
      try {
        const store = lerStore();
        if (store.mapa.size === 0) semearStoreDaMe2w(targetSS, store);

        const stats = aplicarStoreNaMe2w(dadosMe2w, store);
        atualizarAba(targetSS, ABA_ME2W, dadosMe2w);
        gravarStore(store);
        SpreadsheetApp.flush();

        console.log("ME2W: " + (dadosMe2w.length - 1) + " linhas | " +
                    stats.restauradas + " confirmações restauradas | " +
                    stats.reapareceram + " reapareceram | " +
                    stats.ausentes + " sumiram do export (preservadas no store).");
      } finally {
        lock.releaseLock();
      }

      atualizarAbaComHistorico(targetSS, histSS, "ME2W", dadosMe2w,
        ["Purchasing Document", "Item", "Material", "Order Quantity", "Delivery Date", "Qty Delivered", "Schedule Line"]);
    } catch (e) {
      erros.push("ME2W: " + e.message);
      console.error("ME2W abortada (dado preservado): " + e.message);
    }

    // 3) Bases de analise: insumos reconstruiveis, sem dado manual. Cada uma
    //    falha por conta propria sem derrubar as demais.
    const analise = [
      { nome: "RESB", dados: dadosResb, chaves: ["Material", "ReqmtsDate", "Reqmnt qty", "Plnd Ord.", "Reserv.no.", "Pegged Requirement"] },
      { nome: "ME5A", dados: dadosMe5a },
      { nome: "Stock Control BR14 BR10 BR12", dados: dadosStock },
      { nome: "ME2N", dados: dadosMe2n }
    ];

    analise.forEach(base => {
      try {
        console.log("Processando " + base.nome + "...");
        if (base.chaves) atualizarAbaComHistorico(targetSS, histSS, base.nome, base.dados, base.chaves);
        else atualizarAba(targetSS, base.nome, base.dados);
      } catch (e) {
        erros.push(base.nome + ": " + e.message);
        console.error(base.nome + " falhou: " + e.message);
      }
    });

    SpreadsheetApp.flush();

    if (erros.length === 0) {
      props.setProperty('LAST_SYNC_NOVAS_BASES', updatedTime);
      console.log("Sincronização concluída com sucesso!");
    } else {
      console.error("Sincronização concluída COM FALHAS: " + erros.join(" || "));
      notificarFalha(erros.join("\n"));
    }

  } catch (err) {
    console.error("Erro Crítico no processamento: " + (err && err.stack ? err.stack : err));
    notificarFalha(String(err && err.stack ? err.stack : err));
    throw err;
  } finally {
    tempIds.forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
  }
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
