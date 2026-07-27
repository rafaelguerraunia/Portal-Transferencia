const SPREADSHEET_ID = "1Beq9Gwq3l1o19r1t-yfX-_IVMTDBi8FPUsALGc58YtM";
const SHEET_NAME = "Pagina Transferência";
const SHEET_TOKENS = "Tokens_Link";

function doGet(e) {
  if (!e || !e.parameter || !e.parameter.token) {
    return ContentService.createTextOutput("Erro: Link inválido. Token de acesso não fornecido.");
  }
  
  const isValid = validateToken(e.parameter.token);
  if (!isValid) {
    return ContentService.createTextOutput("Erro: Acesso negado. Token expirado ou inexistente.");
  }

  const template = HtmlService.createTemplateFromFile('Sto-Frontend');
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Portal de Transferência de Material");
}

function validateToken(tokenInput) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TOKENS);
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === tokenInput) return true;
  }
  return false;
}

function formatCustomDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return dateObj || "";
  const months = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  const day = String(dateObj.getDate()).padStart(2, '0');
  const month = months[dateObj.getMonth()];
  const year = dateObj.getFullYear();
  return `${day}/${month}/${year}`;
}

function formatIsoDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "";
  const d = String(dateObj.getDate()).padStart(2, '0');
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const y = dateObj.getFullYear();
  return `${y}-${m}-${d}`;
}

function getTransferData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];

  const headers = data[0].map(h => String(h).trim());
  const colConfDate = headers.indexOf("Deliv Date - Confirmação Planejamento (SMART HUB)");
  const colConfQty = headers.indexOf("Qtd - Confirmação Planejamento (SMART HUB)");
  const colFlagConf = headers.indexOf("Confirmação Smarthub");
  const colPrioridade = headers.indexOf("Prioridade Smarthub"); 
  const colCausa = headers.indexOf("Causa de Desvio");
  const colOrigem = headers.indexOf("Origem");
  const colDocumento = headers.indexOf("Documento");
  const colItem = headers.indexOf("Item");
  const colSched = headers.indexOf("Schedule Line");
  const colPlantaDestino = headers.indexOf("Planta_ Destino");
  const colDelivDate = headers.indexOf("Deliv Date");
  const colDocDate = headers.indexOf("Document Date");
  const colMaterial = headers.indexOf("Material");
  const colShortText = headers.indexOf("Short Text");
  const colOrderUnit = headers.indexOf("Order Unit");
  const colOrderQty = headers.indexOf("Order Quantity");
  const colDeletada = headers.indexOf("Deletada?");
  const colCompleta = headers.indexOf("Completa?");
  const colReqSemana = headers.indexOf("Req Semana");
  const colEstqDestino = headers.indexOf("Estoque na Planta Destino");
  const colStatusEstoque = headers.indexOf("Há Estoque para Semana?");
  const colDiasEstoque = headers.indexOf("Dias Disponiveis em Estoque");
  const colEstqBR14 = headers.indexOf("Estoque BR14 Destinado para Planta");
  const colEstqConjunto = headers.indexOf("Estoque Conjunto de Plantas");
  const colEstqBR14Sub = headers.indexOf("Estoque BR14 Sub Destinado para Planta");
  const colDiasDispBR14 = headers.indexOf("Dias Disponiveis na BR14");
  const colQtdPedidos = headers.indexOf("Qtd Pedidos Abertos");
  const colListaPedidos = headers.indexOf("Lista de Pedidos");
  const colFornPedido = headers.indexOf("Fornecedor Pedido");
  const colDelivPedido = headers.indexOf("Deliver Date Pedido");
  const colQtdPedidoItem = headers.indexOf("Qtd. Pedido");
  const colPlantaPedido = headers.indexOf("Planta Pedido");
  const colInspQualidade = headers.indexOf("Insp Qualidade"); 
  
  // COLUNAS NOVAS - LÓGICA DE STATUS DO FLUXO
  let colPreAgendado = headers.indexOf("Pré Agendado?");
  if (colPreAgendado === -1) colPreAgendado = headers.indexOf("Pré-Agendado?");
  if (colPreAgendado === -1) colPreAgendado = headers.indexOf("Pré Agendado");
  if (colPreAgendado === -1) colPreAgendado = headers.indexOf("Pré-Agendado");
  let colSeparado = headers.indexOf("Separado?");
  if (colSeparado === -1) colSeparado = headers.indexOf("Separado");
  const colStatusTransporte = headers.indexOf("Status Transporte");
  const colIssuedQty = headers.indexOf("Issued Quantity");
  const colQtyDelivered = headers.indexOf("Qty Delivered");
  const colQtyReceived = headers.indexOf("Quantity Received");

  let portalData = [];
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  const limite7d = new Date(hoje.getTime() + 7 * 24 * 60 * 60 * 1000);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const confDate = colConfDate !== -1 ? row[colConfDate] : "";
    const confQty = colConfQty !== -1 ? row[colConfQty] : "";
    const confFlag = colFlagConf !== -1 ? row[colFlagConf] : "";
    let prioridade = colPrioridade !== -1 ? row[colPrioridade] : "";
    let causaDesvio = colCausa !== -1 ? row[colCausa] : "";
    const origemRaw = colOrigem !== -1 ? row[colOrigem] : "";
    const documento = colDocumento !== -1 ? row[colDocumento] : "";
    const itemNum = colItem !== -1 ? row[colItem] : "";
    const schedNum = colSched !== -1 ? row[colSched] : "";
    const plantaDestino = colPlantaDestino !== -1 ? row[colPlantaDestino] : "";
    const delivDate = colDelivDate !== -1 ? row[colDelivDate] : "";
    const docDate = colDocDate !== -1 ? row[colDocDate] : "";
    const material = colMaterial !== -1 ? row[colMaterial] : "";
    const shortText = colShortText !== -1 ? row[colShortText] : "";
    const orderUnit = colOrderUnit !== -1 ? row[colOrderUnit] : "";
    const orderQty = colOrderQty !== -1 ? row[colOrderQty] : "";
    const deletada = colDeletada !== -1 ? row[colDeletada] : "";
    const completa = colCompleta !== -1 ? row[colCompleta] : "";
    const reqSemana = colReqSemana !== -1 ? row[colReqSemana] : "";
    const estqDestino = colEstqDestino !== -1 ? row[colEstqDestino] : "";
    const diasEstoque = colDiasEstoque !== -1 ? row[colDiasEstoque] : "";
    const estqBR14 = colEstqBR14 !== -1 ? row[colEstqBR14] : "";
    const estqConjunto = colEstqConjunto !== -1 ? row[colEstqConjunto] : "";
    const estqBR14Sub = colEstqBR14Sub !== -1 ? row[colEstqBR14Sub] : "";
    const diasDispBR14 = colDiasDispBR14 !== -1 ? row[colDiasDispBR14] : "";
    const qtdPedidosAbertos = colQtdPedidos !== -1 ? row[colQtdPedidos] : 0;
    const listaPedidos = colListaPedidos !== -1 ? row[colListaPedidos] : "";
    const fornPedido = colFornPedido !== -1 ? row[colFornPedido] : "";
    const delivPedido = colDelivPedido !== -1 ? row[colDelivPedido] : "";
    const qtdPedidoItem = colQtdPedidoItem !== -1 ? row[colQtdPedidoItem] : "";
    const plantaPedido = colPlantaPedido !== -1 ? row[colPlantaPedido] : "";
    const inspQualidade = colInspQualidade !== -1 ? row[colInspQualidade] : "";

    if (String(documento).trim() === "") continue;
    let origemFinal = origemRaw === "ME5A" ? "REQ" : (origemRaw === "ME2W" ? "STO" : "OUTRO");
    const isDeletedOrCompleted = (deletada !== "" || completa !== "");
    const hasConfirmation = (confDate !== "" || confQty !== "" || confFlag !== "");
    if (isDeletedOrCompleted && !hasConfirmation) continue;

    // --- CÁLCULO DAS DUAS LÓGICAS DO STATUS DO FLUXO (NOVO) ---
    let preAgendado = colPreAgendado !== -1 ? String(row[colPreAgendado]).trim().toLowerCase() : "";
    let separado = colSeparado !== -1 ? String(row[colSeparado]).trim().toLowerCase() : "";
    let statusTransporte = colStatusTransporte !== -1 ? String(row[colStatusTransporte]).trim() : "";

    let statusFluxo1 = "";
    if (statusTransporte !== "") {
        statusFluxo1 = statusTransporte;                 // estágio mais avançado
    } else if (preAgendado === "sim" && separado === "sim") {
        statusFluxo1 = "Separado";
    } else if (preAgendado === "sim") {
        statusFluxo1 = "Pré-Agendado";
    }

    let sq = Number(String(orderQty).replace(/,/g, '')) || 0;
    let iq = colIssuedQty !== -1 ? Number(String(row[colIssuedQty]).replace(/,/g, '')) || 0 : 0;
    let dq = colQtyDelivered !== -1 ? Number(String(row[colQtyDelivered]).replace(/,/g, '')) || 0 : 0;
    let rq = colQtyReceived !== -1 ? Number(String(row[colQtyReceived]).replace(/,/g, '')) || 0 : 0;

    let statusFluxo2 = "-";
    if (sq === iq && iq === dq && dq === rq && sq > 0) {
        statusFluxo2 = "Fluxo Concluído com Sucesso";
    } else if (sq > 0 && iq === 0) {
        statusFluxo2 = "Pendente de Separação";
    } else if (iq === sq && dq === 0 && rq === 0 && sq > 0) {
        statusFluxo2 = "Trânsito Físico (Full)";
    } else if (iq > 0 && dq === iq && rq === 0) {
        statusFluxo2 = "Trânsito Documentado";
    } else if (rq > dq || rq > iq) {
        statusFluxo2 = "Divergência (Sobra Física)";
    } else if (rq > 0 && (rq < dq || rq < iq)) {
        statusFluxo2 = "Divergência (Falta Física)";
    } else if (rq > 0 && rq < sq) {
        statusFluxo2 = "Recebimento Parcial";
    } else if (iq > 0 && iq < sq) {
        statusFluxo2 = "Envio Parcial";
    }

    let statusFluxoFinal = statusFluxo2;
    if (statusFluxo1 !== "") {
        statusFluxoFinal = statusFluxo1 + " | " + statusFluxo2;
    }
    // -----------------------------------------------------------

    // --- ALERTA DE PLANEJAMENTO ---
    let alertaStatus = "Firme";
    const dConfIso = formatIsoDate(confDate);
    const dSapIso = formatIsoDate(delivDate);
    const qConfNum = Number(String(confQty).replace(/,/g, ''));
    const qSapNum = Number(String(orderQty).replace(/,/g, ''));

    if (isDeletedOrCompleted && hasConfirmation) {
      alertaStatus = "Revisão Urgente(Completo/Cancelado)";
    } else {
      if (origemFinal === "REQ") alertaStatus = "Pendente Criação STO";
      else {
         if (confDate === "" && confQty === "") alertaStatus = "Aguardando Confirmação";
         else if (dConfIso !== dSapIso || qConfNum !== qSapNum) alertaStatus = "Solicitar ajuste";
         else alertaStatus = "Firme";
      }
    }

    // --- LÓGICA: STATUS ESTOQUE UNIFICADO ---
    let req = Number(String(reqSemana).replace(/,/g, '')) || 0;
    let estqPlanta = Number(String(estqDestino).replace(/,/g, '')) || 0;
    let estqBR14Num = Number(String(estqBR14).replace(/,/g, '')) || 0;
    let rompePlanta = estqPlanta < req;
    let rompeGlobal = (estqPlanta + estqBR14Num) < req;

    let sumBRX = 0;
    let sumBR14 = 0;
    let hasAtrasado = false;
    let hasMaior7d = false;
    let arrPlanta = String(plantaPedido).split('|').map(s=>s.trim());
    let arrDeliv = String(delivPedido).split('|').map(s=>s.trim());
    let arrQtd = String(qtdPedidoItem).split('|').map(s=>s.trim());

    for(let j = 0; j < arrQtd.length; j++){
      if(!arrQtd[j]) continue;
      let q = Number(arrQtd[j].replace(/,/g, '')) || 0;
      let dP = new Date(arrDeliv[j]);
      if(dP < hoje) hasAtrasado = true;
      if(dP > limite7d) hasMaior7d = true;

      if(dP <= limite7d) { 
        if(arrPlanta[j] === plantaDestino) sumBRX += q;
        else sumBR14 += q; 
      }
    }

    let statusEstoqueCalc = "";
    let isSalvoBRX = rompePlanta && ((estqPlanta + sumBRX) >= req);
    let isSalvoBR14 = rompePlanta && !isSalvoBRX && ((estqPlanta + sumBRX + estqBR14Num + sumBR14) >= req);

    if (!rompePlanta) {
        statusEstoqueCalc = "✅ Suficiente (Físico)";
    } else if (isSalvoBRX) {
        statusEstoqueCalc = `📦 Salvo por Pedidos ${plantaDestino}`;
    } else if (isSalvoBR14) {
        statusEstoqueCalc = "📦 Salvo por Pedidos BR14";
    } else {
        let abertos = Number(qtdPedidosAbertos) || 0;
        if (abertos === 0) {
            if (rompeGlobal) statusEstoqueCalc = "🚨 R.R.T";
            else statusEstoqueCalc = "⚠️ R.R.P. s/ Pedidos"; 
        } else {
            if (sumBRX === 0 && sumBR14 === 0 && hasMaior7d) {
                statusEstoqueCalc = "🚨 R.R.P. c/ Pedidos Atrasados";
            } else {
                if ((estqPlanta + sumBRX + estqBR14Num + sumBR14) < req) statusEstoqueCalc = "🚨 R.R.T";
                else statusEstoqueCalc = "⚠️ R.R.P. c/ Pedidos";
            }
        }
    }

    // --- LÓGICA: PRIORIDADE AUTOMÁTICA ---
    if (!prioridade) {
        let dEstoque = 0;
        let diasEstqRaw = String(diasEstoque);
        if (diasEstqRaw.includes("Sem Consumo") || diasEstqRaw.includes("Superior")) dEstoque = 999;
        else dEstoque = parseFloat(diasEstqRaw) || 0;
        if (dEstoque <= 1) prioridade = "0-Crítico";
        else if (dEstoque <= 4) prioridade = "1-Urgente";
        else if (dEstoque <= 7) prioridade = "2-Alto";
        else if (dEstoque <= 15) prioridade = "3-Normal";
        else prioridade = "4-Baixo";
    }

    if (!causaDesvio) {
        let autoCausas = [];
        if (confQty !== "" && qConfNum !== qSapNum) autoCausas.push("Quantidade alterada");
        if ((Number(inspQualidade)||0) > 0) autoCausas.push("Material em QI");
        if (statusEstoqueCalc === "🚨 R.R.T" && (Number(qtdPedidosAbertos)||0) === 0) autoCausas.push("Falta de Pedido");
        if (statusEstoqueCalc === "🚨 R.R.T") autoCausas.push("Falta de estoque");
        if (hasAtrasado) autoCausas.push("Pedido Atrasado");
        
        if (confDate && delivDate) {
             let cD = new Date(confDate).getTime();
             let sD = new Date(delivDate).getTime();
             if (cD < sD) autoCausas.push("Antecipação de transferência");
             if (cD > sD) autoCausas.push("Postergação de transferência");
        }
        if (typeof docDate !== 'undefined' && docDate && delivDate) {
             let docD = new Date(docDate).getTime();
             let sD = new Date(delivDate).getTime();
             let diffDias = Math.abs(sD - docD) / (1000 * 60 * 60 * 24);
             if (diffDias <= 1) autoCausas.push("FastTrack");
        }
        
        causaDesvio = autoCausas.length > 0 ? autoCausas.join(" | ") : "";
    }

    let leadTimeTxt = "-";
    if (delivDate instanceof Date && docDate instanceof Date) {
      const diasCorridos = Math.round((delivDate - docDate) / (1000 * 60 * 60 * 24));
      const diasUteis = calcBusinessDays(docDate, delivDate);
      leadTimeTxt = `${diasCorridos} d. corridos / ${diasUteis} d. úteis`;
    }

    portalData.push({
      rowIndex: i + 1,
      alerta: alertaStatus,
      statusEstqUnificado: statusEstoqueCalc,
      statusFluxo: statusFluxoFinal, // INCLUÍDO NO PAYLOAD DO FRONT
      planta: plantaDestino,
      origem: origemFinal,
      doc: String(documento).trim(),
      item: String(itemNum).trim(),
      sched: String(schedNum).trim(),
      matCod: material,
      matDesc: shortText,
      qtySap: orderQty,
      unitSap: orderUnit,
      dataSap: formatCustomDate(delivDate), 
      dataSapIso: formatIsoDate(delivDate), 
      docDateIso: formatIsoDate(docDate), 
      hasAtrasado: hasAtrasado,
      leadTime: leadTimeTxt,
      reqSemana: reqSemana,
      estqDestino: estqDestino,
      diasEstq: diasEstoque,
      estqBR14: estqBR14,
      estqConjunto: estqConjunto,
      confDateIso: formatIsoDate(confDate), 
      confQty: confQty,
      confFlag: confFlag, 
      prioridade: prioridade, 
      causaDesvio: causaDesvio, 
      qtdPedidosAbertos: qtdPedidosAbertos,
      listaPedidos: listaPedidos,
      fornPedido: fornPedido,
      delivPedido: delivPedido,
      qtdPedidoItem: qtdPedidoItem,
      plantaPedido: plantaPedido,
      estqBR14Sub: estqBR14Sub,
      diasDispBR14: diasDispBR14,
      inspQualidade: inspQualidade 
    });
  }

  portalData.sort((a, b) => {
    let dateA = a.dataSapIso || "9999-99-99"; 
    let dateB = b.dataSapIso || "9999-99-99";
    return dateA.localeCompare(dateB);
  });
  return portalData;
}

// Resolve a aba ME2W e as colunas que o portal escreve. As 5 colunas Smarthub
// sao criadas pelo Sync; se nao existirem, a aba nao passou por uma sincronizacao
// valida e escrever aqui produziria um layout que as formulas da Pagina
// Transferencia (VLOOKUP por indice fixo) nao enxergam.
function abrirMe2wParaEscrita() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("ME2W");
  if (!sheet) throw new Error("Aba ME2W não encontrada.");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const colDoc = headers.indexOf("Purchasing Document");
  const colItem = headers.indexOf("Item");
  const colSched = headers.indexOf("Schedule Line");
  if (colDoc === -1 || colItem === -1 || colSched === -1) {
    throw new Error("Aba ME2W sem colunas-chave. Cabeçalho: " + headers.join(" | "));
  }

  const idxManuais = COLUNAS_MANUAIS.map(c => headers.indexOf(c));
  if (idxManuais.some(i => i === -1)) {
    const faltando = COLUNAS_MANUAIS.filter((c, k) => idxManuais[k] === -1);
    throw new Error("Aba ME2W sem as colunas Smarthub (" + faltando.join(", ") +
                    "). Rode a sincronização antes de confirmar.");
  }

  const linhas = new Map();
  for (let i = 1; i < data.length; i++) {
    linhas.set(montarChave(data[i][colDoc], data[i][colItem], data[i][colSched]), i);
  }

  return {
    sheet: sheet, data: data, headers: headers, linhas: linhas, idxManuais: idxManuais,
    colSapData: headers.indexOf(COL_SAP_DATA),
    colSapQtd: headers.indexOf(COL_SAP_QTD),
    primeira: Math.min.apply(null, idxManuais),
    ultima: Math.max.apply(null, idxManuais)
  };
}

function escreverManuaisNaMe2w(ctx, i, manuais) {
  const largura = ctx.ultima - ctx.primeira + 1;
  const bloco = ctx.data[i].slice(ctx.primeira, ctx.ultima + 1);
  for (let k = 0; k < ctx.idxManuais.length; k++) bloco[ctx.idxManuais[k] - ctx.primeira] = manuais[k];
  ctx.sheet.getRange(i + 1, ctx.primeira + 1, 1, largura).setValues([bloco]);
}

function saveConfirmation(doc, item, sched, dateVal, qtyVal, flagVal, priorityVal, causaVal) {
  const r = saveMultipleConfirmations([{
    doc: doc, item: item, sched: sched, dtVal: dateVal, qtVal: qtyVal,
    flagVal: flagVal, priVal: priorityVal, causaVal: causaVal
  }]);
  return r.ok === 1;
}

// Escreve nos dois lugares sob o mesmo lock do Sync: na ME2W (para o VLOOKUP da
// Pagina Transferencia enxergar agora) e no store (fonte da verdade, que
// sobrevive ao clearContents e ao sumico da ordem no export do SAP).
function saveMultipleConfirmations(updatesArray) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("As bases estão sendo sincronizadas neste momento. Aguarde alguns segundos e salve novamente.");
  }
  try {
    const ctx = abrirMe2wParaEscrita();
    const store = lerStore();
    const usuario = Session.getActiveUser().getEmail() || "portal";
    const tocados = [];
    const ok = [], falhas = [];

    updatesArray.forEach(u => {
      const rotulo = u.doc + "/" + u.item + "/" + u.sched;

      // Uma confirmacao sem data envenena o store de forma permanente: o Sync
      // passaria a restaura-la para sempre. Falhar alto e melhor que gravar isso.
      if (!u.dtVal) { falhas.push(rotulo + " (data vazia)"); return; }

      const i = ctx.linhas.get(montarChave(u.doc, u.item, u.sched));
      if (i === undefined) { falhas.push(rotulo + " (não encontrada na ME2W)"); return; }

      const manuais = [u.flagVal, parseDataPortal(u.dtVal), u.qtVal, u.priVal, u.causaVal];
      escreverManuaisNaMe2w(ctx, i, manuais);

      const reg = upsertStore(store, u.doc, u.item, u.sched, manuais,
                              ctx.colSapData !== -1 ? ctx.data[i][ctx.colSapData] : "",
                              ctx.colSapQtd !== -1 ? ctx.data[i][ctx.colSapQtd] : "",
                              usuario);
      tocados.push(reg);
      ok.push(rotulo);
    });

    if (tocados.length > 0) gravarStoreParcial(store, tocados);
    if (falhas.length > 0) console.warn("Confirmações não salvas: " + falhas.join(" | "));
    return { ok: ok.length, falhas: falhas };
  } finally {
    lock.releaseLock();
  }
}

function clearConfirmation(doc, item, sched) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("As bases estão sendo sincronizadas neste momento. Aguarde alguns segundos e tente novamente.");
  }
  try {
    const ctx = abrirMe2wParaEscrita();
    const i = ctx.linhas.get(montarChave(doc, item, sched));
    if (i === undefined) return false;

    const vazios = COLUNAS_MANUAIS.map(() => "");
    escreverManuaisNaMe2w(ctx, i, vazios);

    // Limpar no store tambem: sem isso o proximo Sync restauraria a confirmacao.
    const store = lerStore();
    const reg = upsertStore(store, doc, item, sched, vazios, "", "",
                            Session.getActiveUser().getEmail() || "portal");
    gravarStoreParcial(store, [reg]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function calcBusinessDays(startDate, endDate) {
  let count = 0;
  let curDate = new Date(startDate.getTime());
  while (curDate <= endDate) {
    const dayOfWeek = curDate.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) count++;
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
}

function getOrCreateToken(userName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TOKENS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TOKENS);
    sheet.appendRow(["Usuario/Planta", "Token"]);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userName) return data[i][1];
  }
  const newToken = Utilities.getUuid();
  sheet.appendRow([userName, newToken]);
  return newToken;
}