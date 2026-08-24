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

// ====================================================================
// CACHE DO PORTAL
// ====================================================================
//
// Abrir esta planilha custa caro. A "Pagina Transferencia" e montada inteira
// por formula — VLOOKUP nas bases do SAP e em duas planilhas externas — e o
// openById so devolve o controle depois que o servidor termina de recalcular
// tudo. No diagnostico de 24/08/2026 isso levou 123,5 s; tudo o que vem depois
// (getValues, getDisplayValues, o getTransferData inteiro) somou menos de 4 s.
//
// Como cada requisicao do Web App e uma execucao nova, sem cache o usuario
// pagava esse recalculo duas vezes: uma no doGet, so para validar o token, e
// outra no getTransferData. Sao mais de quatro minutos de tela branca — o
// portal "nao abria".
//
// O cache tira as duas esperas da frente do usuario. Quem paga o recalculo e o
// atualizarCachePortal(), chamado pelo Sync depois de escrever as bases (a
// planilha ja esta quente ali) e por um gatilho de seguranca a cada 30 min.

const CACHE_TOKENS = "PORTAL_TOKENS_V1";
const CACHE_PAYLOAD = "PORTAL_PAYLOAD_V1";

const CACHE_TTL_TOKENS = 1800;    // 30 min

// Intervalo minimo entre duas releituras da aba de tokens motivadas por token
// desconhecido. O Web App e ANYONE_ANONYMOUS: sem esse teto, um link errado —
// ou um robo batendo no /exec — dispararia uma abertura de planilha de dois
// minutos por requisicao, e a cota diaria de execucao acabaria sozinha.
const CACHE_TTL_RELEITURA = 60;   // 1 min
const CACHE_TTL_PAYLOAD = 21600;  // 6 h — teto do CacheService

// Acima disso o aquecimento claramente parou de rodar: servir dado com esse
// atraso e pior do que fazer o usuario esperar o recalculo uma vez.
const IDADE_MAX_PAYLOAD_MS = 45 * 60 * 1000;

// O CacheService corta em 100 KB por chave e o payload passa de 300 KB, entao
// o texto vai fatiado. 32768 caracteres cabem em 100 KB mesmo no pior caso de
// UTF-8 (3 bytes por caractere), que os acentos e os emojis de status trazem.
const CACHE_FATIA = 32 * 1024;

function cacheGravarTexto(chave, texto, ttl) {
  const cache = CacheService.getScriptCache();
  const fatias = {};
  let n = 0;
  for (let i = 0; i < texto.length; i += CACHE_FATIA) {
    fatias[chave + "_" + n] = texto.substring(i, i + CACHE_FATIA);
    n++;
  }
  // As fatias antes do indice: enquanto o indice nao existir, o leitor trata
  // como cache frio em vez de montar um texto pela metade.
  cache.putAll(fatias, ttl);
  cache.put(chave, String(n), ttl);
}

function cacheLerTexto(chave) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(chave));
  if (!n) return null;

  const chaves = [];
  for (let i = 0; i < n; i++) chaves.push(chave + "_" + i);
  const fatias = cache.getAll(chaves);

  let texto = "";
  for (let i = 0; i < n; i++) {
    const parte = fatias[chave + "_" + i];
    // Uma fatia pode expirar sozinha; devolver o texto truncado daria um JSON
    // invalido ou, pior, uma lista de linhas cortada no meio.
    if (parte === undefined || parte === null) return null;
    texto += parte;
  }
  return texto;
}

function cacheRemover(chave) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(chave)) || 0;
  // A guarda de releitura sai junto: sem isso um token recem-emitido esperaria
  // ate um minuto para ser aceito.
  const chaves = [chave, chave + "_RELIDO"];
  for (let i = 0; i < n; i++) chaves.push(chave + "_" + i);
  cache.removeAll(chaves);
}

// Apaga token e payload. Rodar na mao depois de revogar um token (o cache de
// tokens tem 30 min de validade) ou para forcar a releitura da planilha.
function limparCachePortal() {
  cacheRemover(CACHE_TOKENS);
  cacheRemover(CACHE_PAYLOAD);
  console.log("Cache do portal limpo — a próxima abertura relê a planilha.");
}

// --------------------------------------------------------------------
// Tokens
// --------------------------------------------------------------------

function lerTokensDaPlanilha() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_TOKENS);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const tokens = [];
  for (let i = 1; i < data.length; i++) {
    const t = String(data[i][1] === null || data[i][1] === undefined ? "" : data[i][1]).trim();
    if (t) tokens.push(t);
  }
  cacheGravarTexto(CACHE_TOKENS, JSON.stringify(tokens), CACHE_TTL_TOKENS);
  return tokens;
}

function validateToken(tokenInput) {
  const alvo = String(tokenInput === null || tokenInput === undefined ? "" : tokenInput).trim();
  if (!alvo) return false;

  const emCache = cacheLerTexto(CACHE_TOKENS);
  if (emCache !== null) {
    try {
      if (JSON.parse(emCache).indexOf(alvo) !== -1) return true;
    } catch (e) {
      // Cache corrompido cai na releitura abaixo.
    }
  }

  // Token desconhecido pelo cache ainda pode ser um recem-emitido. So nesse
  // caso vale abrir a planilha — o caminho feliz nunca chega aqui.
  const cache = CacheService.getScriptCache();
  const guarda = CACHE_TOKENS + "_RELIDO";
  if (emCache !== null && cache.get(guarda)) return false;

  cache.put(guarda, "1", CACHE_TTL_RELEITURA);
  return lerTokensDaPlanilha().indexOf(alvo) !== -1;
}

// --------------------------------------------------------------------
// Payload do portal
// --------------------------------------------------------------------

function lerPayloadDoCache() {
  const texto = cacheLerTexto(CACHE_PAYLOAD);
  if (texto === null) return null;

  let envelope;
  try {
    envelope = JSON.parse(texto);
  } catch (e) {
    return null;
  }
  if (!envelope || !Array.isArray(envelope.linhas)) return null;
  if (Date.now() - Number(envelope.geradoEmMs || 0) > IDADE_MAX_PAYLOAD_MS) return null;
  return envelope;
}

function gravarPayloadNoCache(envelope) {
  const texto = JSON.stringify(envelope);
  cacheGravarTexto(CACHE_PAYLOAD, texto, CACHE_TTL_PAYLOAD);
  // Devolve o que foi gravado, e nao o objeto original: assim a resposta do
  // cache e a resposta do recalculo passam pelo mesmo JSON e chegam ao
  // navegador com os mesmos tipos.
  return JSON.parse(texto);
}

// Paga o recalculo da planilha e deixa a resposta pronta. Chamado pelo Sync
// (planilha ja quente) e pelo gatilho de aquecimento — nao pelo usuario.
function atualizarCachePortal() {
  const t0 = Date.now();
  const linhas = lerTransferDataDaPlanilha();
  const agora = new Date();

  const envelope = gravarPayloadNoCache({
    geradoEmMs: agora.getTime(),
    geradoEm: Utilities.formatDate(agora, TIMEZONE, "dd/MM HH:mm"),
    linhas: linhas
  });

  console.log("Cache do portal atualizado: " + linhas.length + " linha(s) em " +
              ((Date.now() - t0) / 1000).toFixed(1) + "s.");
  return envelope;
}

// Recria o gatilho de aquecimento. O caminho normal e o Sync atualizar o cache
// logo depois de escrever as bases; este gatilho fecha o buraco dos ciclos em
// que nenhum export mudou. Roda a cada 30 min para caber dentro dos 45 min de
// IDADE_MAX_PAYLOAD_MS — de hora em hora o payload passaria da idade servivel
// e alguem pagaria o openById frio. Aumentar o intervalo economiza cota de
// execucao, mas so a partir de 45 min o buraco reabre.
function instalarGatilhoDoPortal() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "aquecerCachePortal") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("aquecerCachePortal").timeBased().everyMinutes(30).create();
  console.log("Gatilho de aquecimento instalado: a cada 30 min.");
}

// Fora do horario o portal nao tem quem abra, e cada aquecimento gasta ate dois
// minutos da cota diaria de execucao. Domingo e madrugada ficam de fora.
// getDay()/getHours() ja saem em horario de Brasilia: o timeZone do projeto e
// America/Sao_Paulo — mesma leitura que a janela do sincronizarNovasBases usa.
function aquecerCachePortal() {
  const agora = new Date();
  if (agora.getDay() === 0 || agora.getHours() < 5 || agora.getHours() > 21) {
    console.log("Fora da janela de aquecimento (" +
                Utilities.formatDate(agora, TIMEZONE, "EEE HH:mm") +
                ") — cache mantido como está.");
    return;
  }
  atualizarCachePortal();
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

// "4.500.123.456" / "4 500 123 456" -> "4500123456". Devolve "" se nao for numero.
function somenteNumero(txt) {
  const s = String(txt === null || txt === undefined ? "" : txt).trim();
  if (/^\d+$/.test(s)) return s;
  if (/^\d{1,3}([.,\s]\d{3})+$/.test(s)) return s.replace(/[.,\s]/g, "");
  return "";
}

// Numero de serie da planilha (epoca 30/12/1899) a partir de um Date valido.
function serialDaData(dateObj) {
  const ms = dateObj.getTime();
  if (isNaN(ms)) return null;
  return Math.round((ms - new Date(1899, 11, 30).getTime()) / 86400000);
}

// Documento / Item / Schedule Line sao sempre numericos, mas basta a celula herdar
// um formato de data para o getValues() devolver um Date no lugar do numero - e,
// como documento SAP e grande demais para virar data, esse Date nasce invalido e
// chegava no portal como o texto "Invalid Date". Nesses casos o texto exibido na
// planilha ainda mostra o numero certo, entao ele vira a fonte de verdade; se nem
// ele servir, reconstruimos o numero de serie a partir da data.
function normalizarNumeroDoc(valor, exibido) {
  if (valor instanceof Date) {
    const doTexto = somenteNumero(exibido);
    if (doTexto) return doTexto;
    const serial = serialDaData(valor);
    return serial === null ? "" : String(serial);
  }
  if (typeof valor === "number") {
    if (!isFinite(valor)) return "";
    // toFixed evita que documentos longos virem notacao cientifica (4.5e+9).
    return Number.isInteger(valor) ? valor.toFixed(0) : String(valor);
  }
  const s = String(valor === null || valor === undefined ? "" : valor).trim();
  if (/^invalid date$/i.test(s)) return "";
  return somenteNumero(s) || s;
}

function lerNumeroDoc(row, displayRow, col) {
  if (col === -1) return "";
  return normalizarNumeroDoc(row[col], displayRow ? displayRow[col] : "");
}

// O que o navegador chama ao abrir o portal. Serve o cache; so cai na planilha
// quando nao ha nada guardado — e ai paga o recalculo de dois minutos, que e
// exatamente o que o aquecimento existe para evitar.
function getTransferData() {
  const doCache = lerPayloadDoCache();
  if (doCache) return doCache;

  console.warn("Cache do portal frio — lendo a planilha na requisição do usuário. " +
               "Confira se o gatilho aquecerCachePortal está instalado.");
  return atualizarCachePortal();
}

function lerTransferDataDaPlanilha() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const range = sheet.getDataRange();
  const data = range.getValues();
  // Texto exibido na celula: usado para recuperar Documento/Item/Schedule Line
  // quando a celula esta com formato de data e o getValues() devolve um Date.
  const display = range.getDisplayValues();
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
    const displayRow = display[i];
    const documento = lerNumeroDoc(row, displayRow, colDocumento);
    const itemNum = lerNumeroDoc(row, displayRow, colItem);
    const schedNum = lerNumeroDoc(row, displayRow, colSched);
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

    if (documento === "") continue;
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
      // Cru, separado do statusFluxo: o rastreio precisa do status de transporte
      // sozinho para posicionar o caminhão na régua de steps.
      statusTransporte: statusTransporte,
      planta: plantaDestino,
      origem: origemFinal,
      doc: documento,
      item: itemNum,
      sched: schedNum,
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

  const range = sheet.getDataRange();
  const data = range.getValues();
  const display = range.getDisplayValues();
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

  // Mesma normalizacao aplicada na Pagina Transferencia: as duas pontas da chave
  // precisam enxergar o documento do mesmo jeito, senao o Salvar nao acha a linha.
  const linhas = new Map();
  for (let i = 1; i < data.length; i++) {
    linhas.set(montarChave(lerNumeroDoc(data[i], display[i], colDoc),
                           lerNumeroDoc(data[i], display[i], colItem),
                           lerNumeroDoc(data[i], display[i], colSched)), i);
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

// ====================================================================
// REFLEXO DAS GRAVACOES NO CACHE
// ====================================================================
//
// Depois de gravar, o cache continua servindo — mas com a confirmacao ja
// aplicada nas linhas tocadas. Derruba-lo aqui devolveria a espera de dois
// minutos para o proximo que abrisse o portal, e a alternativa de nao mexer
// faria o proprio autor da confirmacao recarregar a pagina e nao ver o que
// acabou de salvar. O recalculo do resto vem no proximo aquecimento.

// Mesma regra do getTransferData: origem e situacao da ordem mandam mais do
// que o que o planejamento acabou de digitar.
function alertaDepoisDaConfirmacao(linha, dtVal, qtVal) {
  if (linha.origem === "REQ") return "Pendente Criação STO";
  if (String(linha.alerta).indexOf("Revisão Urgente") === 0) return linha.alerta;
  if (!dtVal && !qtVal) return "Aguardando Confirmação";

  const qConf = Number(String(qtVal).replace(/,/g, ''));
  const qSap = Number(String(linha.qtySap).replace(/,/g, ''));
  return (dtVal !== linha.dataSapIso || qConf !== qSap) ? "Solicitar ajuste" : "Firme";
}

function aplicarConfirmacoesNoCache(aplicadas) {
  if (aplicadas.length === 0) return;

  const envelope = lerPayloadDoCache();
  if (!envelope) return;

  const porChave = {};
  aplicadas.forEach(a => { porChave[montarChave(a.doc, a.item, a.sched)] = a; });

  let mexeu = false;
  envelope.linhas.forEach(linha => {
    const a = porChave[montarChave(linha.doc, linha.item, linha.sched)];
    if (!a) return;

    linha.confDateIso = a.dtVal;
    linha.confQty = a.qtVal;
    linha.confFlag = a.flagVal;
    if (a.priVal) linha.prioridade = a.priVal;
    if (a.causaVal) linha.causaDesvio = a.causaVal;
    linha.alerta = alertaDepoisDaConfirmacao(linha, a.dtVal, a.qtVal);
    mexeu = true;
  });

  if (mexeu) gravarPayloadNoCache(envelope);
}

function aplicarLimpezaNoCache(doc, item, sched) {
  const envelope = lerPayloadDoCache();
  if (!envelope) return;

  const chave = montarChave(doc, item, sched);
  const restantes = [];
  let mexeu = false;

  envelope.linhas.forEach(linha => {
    if (montarChave(linha.doc, linha.item, linha.sched) !== chave) {
      restantes.push(linha);
      return;
    }
    mexeu = true;
    // A linha deletada/completa so aparecia no portal por causa da confirmacao:
    // sem ela, o proprio getTransferData deixaria de devolve-la.
    if (String(linha.alerta).indexOf("Revisão Urgente") === 0) return;

    linha.confDateIso = "";
    linha.confQty = "";
    linha.confFlag = "";
    linha.alerta = alertaDepoisDaConfirmacao(linha, "", "");
    restantes.push(linha);
  });

  if (!mexeu) return;
  envelope.linhas = restantes;
  gravarPayloadNoCache(envelope);
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
    const aplicadas = [];
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
      aplicadas.push(u);

      const reg = upsertStore(store, u.doc, u.item, u.sched, manuais,
                              ctx.colSapData !== -1 ? ctx.data[i][ctx.colSapData] : "",
                              ctx.colSapQtd !== -1 ? ctx.data[i][ctx.colSapQtd] : "",
                              usuario);
      tocados.push(reg);
      ok.push(rotulo);
    });

    if (tocados.length > 0) gravarStoreParcial(store, tocados);
    // Best-effort: o cache e conveniencia, nunca motivo para o save falhar.
    try { aplicarConfirmacoesNoCache(aplicadas); }
    catch (e) { console.warn("Cache do portal não pôde ser atualizado: " + e.message); }
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

    try { aplicarLimpezaNoCache(doc, item, sched); }
    catch (e) { console.warn("Cache do portal não pôde ser atualizado: " + e.message); }
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
  cacheRemover(CACHE_TOKENS);
  return newToken;
}
