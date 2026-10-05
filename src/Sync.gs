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

// RASTREIO DA STO — o store e a aba de historico das STOs, uma linha por chave
// Purchasing Document | Item | Schedule Line. So a ULTIMA ocorrencia de cada coisa,
// como as colunas de historico da ME2N do Portal de Pedidos (mesmos nomes):
//
//   Entrou no sistema em        a chave apareceu no export e nao estava na lista anterior
//   Firmado em / Firmado por    o clique de confirmacao do Planejamento. Fica ate o proximo
//                               clique: o "Limpar" nao apaga (ele fica no Atualizado em/por)
//   Ult.Alteração Order Qty     o valor ANTERIOR, e Dt.Qnd.Alt.OrderQty a hora da
//   Ult.Alteração Deliver Date  sincronizacao que viu a mudanca (idem Dt.Qnd.Alt.DD)
//
// Acrescentadas no FIM, e isso nao e estilo: o store e lido por posicao (ST_*), e
// inserir no meio faria cada campo das linhas ja gravadas passar a ser lido no vizinho.
const COLUNAS_RASTREIO = [
  "Entrou no sistema em",
  "Firmado em", "Firmado por",
  "Ult.Alteração Order Qty", "Dt.Qnd.Alt.OrderQty",
  "Ult.Alteração Deliver Date", "Dt.Qnd.Alt.DD"
];

const STORE_HEADERS = ["Chave", "Purchasing Document", "Item", "Schedule Line"]
  .concat(COLUNAS_MANUAIS)
  .concat(["SAP Delivery Date (no ato)", "SAP Order Quantity (no ato)",
           "Atualizado em", "Atualizado por", "Visto por último no SAP", "Status"])
  .concat(COLUNAS_RASTREIO);

const ST_CHAVE = 0, ST_DOC = 1, ST_ITEM = 2, ST_SCHED = 3;
const ST_MANUAIS = 4;
const ST_SAP_DATA = 9, ST_SAP_QTD = 10;
const ST_ATUALIZADO_EM = 11, ST_ATUALIZADO_POR = 12, ST_VISTO_EM = 13, ST_STATUS = 14;
const ST_ENTROU = 15, ST_FIRMADO_EM = 16, ST_FIRMADO_POR = 17;
const ST_ULT_QTD = 18, ST_DT_ALT_QTD = 19, ST_ULT_DD = 20, ST_DT_ALT_DD = 21;

// ====================================================================
// CATALOGO DAS BASES
// ====================================================================

// Duas rotas, escolhidas pelo que cada base tem a perder:
//
//   "critica" — so a ME2W. Carrega as confirmacoes manuais, entao paga
//               validacao do export, lock e restore do store antes de escrever.
//
//   "direta"  — as demais. Sao insumos reconstruiveis a cada sync: nao tem dado
//               manual, ninguem escreve nelas pelo portal e um export ruim custa
//               so o proximo gatilho. Nao ha o que verificar, entao vao pelo
//               caminho mais curto que existe: uma leitura do temporario, uma
//               escrita na aba (copiarBase), sem passar linha a linha.
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
    modo: "direta",
    historico: ["Material", "ReqmtsDate", "Reqmnt qty", "Plnd Ord.",
                "Reserv.no.", "Pegged Requirement"]
  },
  { aba: "ME5A", arquivo: "STO-ME5A.xlsx", modo: "direta" },
  { aba: "Stock Control BR14 BR10 BR12", arquivo: "Stock_BR14_BR12_BR10.xlsx", modo: "direta" },
  { aba: "ME2N", arquivo: "STO-ME2N.xlsx", modo: "direta" }
];

// O gatilho morre em 6 min. Parar por conta propria antes disso e a diferenca
// entre "o proximo gatilho continua de onde parou" e "o proximo gatilho refaz
// tudo do zero e estoura de novo" — o ciclo que travava a sincronizacao.
const ORCAMENTO_MS = 4.5 * 60 * 1000;

// O corte de verdade. O ORCAMENTO_MS acima e o que a COPIA das bases pode gastar;
// o que sobra daqui ate o corte e o que a firmacao da Pagina Transferencia tem
// para esperar o recalculo. Separados porque sao dois consumidores do mesmo
// gatilho, e adiar a copia de uma base custa um ciclo — ser morto no meio da
// firmacao custa a pagina inteira voltando para formula.
const TETO_GATILHO_MS = 5.5 * 60 * 1000;

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

// Aceita a planilha ja aberta. O sync abre a alvo uma vez e passa adiante; o
// portal chama sem argumento, porque ali o store e a unica coisa que ele toca.
function obterAbaStore(ss) {
  ss = ss || SpreadsheetApp.openById(PLANILHA_ALVO_ID);
  let aba = ss.getSheetByName(ABA_STORE);
  if (!aba) {
    aba = ss.insertSheet(ABA_STORE);
    aba.getRange(1, 1, 1, STORE_HEADERS.length).setValues([STORE_HEADERS]);
    aba.setFrozenRows(1);
    console.log("Aba " + ABA_STORE + " criada.");
    return aba;
  }

  // Store anterior as colunas de rastreio: o cabecalho para em "Status". Completa no fim
  // antes de qualquer gravacao — o portal pode salvar antes da primeira sincronizacao
  // depois da publicacao, e a linha dele ja sai com a largura nova.
  if (aba.getLastColumn() < STORE_HEADERS.length) {
    if (aba.getMaxColumns() < STORE_HEADERS.length) {
      aba.insertColumnsAfter(aba.getMaxColumns(), STORE_HEADERS.length - aba.getMaxColumns());
    }
    aba.getRange(1, 1, 1, STORE_HEADERS.length).setValues([STORE_HEADERS]);
  }
  return aba;
}

// As chaves ocupam um bloco contiguo a partir da linha 2: o gravarStore reescreve
// o bloco inteiro de uma vez, compactado. `linhasNaAba` e o tamanho do bloco lido,
// para ele saber quanto sobra no fim quando o store encolhe.
function lerStore(ss) {
  const aba = obterAbaStore(ss);
  const ultimaLinha = aba.getLastRow();
  const mapa = new Map();
  if (ultimaLinha < 2) return { aba: aba, mapa: mapa, linhasNaAba: 0 };

  const dados = aba.getRange(2, 1, ultimaLinha - 1, STORE_HEADERS.length).getValues();
  for (let i = 0; i < dados.length; i++) {
    const chave = String(dados[i][ST_CHAVE]).trim();
    if (chave) mapa.set(chave, { linha: i + 2, valores: dados[i] });
  }
  return { aba: aba, mapa: mapa, linhasNaAba: dados.length };
}

// Reescreve o store inteiro, cabecalho incluso, na ordem em que estava (os novos no
// fim). O store pode ENCOLHER — o aplicarStoreNaMe2w tira a STO nao confirmada que
// saiu do export —, entao o que sobrar abaixo do bloco novo e limpo.
function gravarStore(store) {
  const existentes = [], novos = [];
  store.mapa.forEach(reg => (reg.linha ? existentes.push(reg) : novos.push(reg)));
  existentes.sort((a, b) => a.linha - b.linha);
  const todos = existentes.concat(novos);

  const dados = [STORE_HEADERS.slice()].concat(todos.map(r => r.valores));
  garantirGradeAba(store.aba, dados.length, STORE_HEADERS.length);
  escreverEmBlocos(store.aba, dados, STORE_HEADERS.length);

  const sobra = (store.linhasNaAba || 0) - todos.length;
  if (sobra > 0) {
    store.aba.getRange(2 + todos.length, 1, sobra, STORE_HEADERS.length).clearContent();
  }
  todos.forEach((r, k) => { r.linha = 2 + k; });
  store.linhasNaAba = todos.length;
}

// O PORTAL NAO LE O STORE INTEIRO. Ele tem uma linha por STO viva, e o clique de
// Salvar toca poucas — ler tudo a cada clique era pagar a base inteira por uma linha.
// So a coluna da chave, para achar a linha; o resto e escrito sem ser lido.
function localizarLinhasStore(ss) {
  const aba = obterAbaStore(ss);
  const ultima = aba.getLastRow();
  const linhas = new Map();
  if (ultima >= 2) {
    const chaves = aba.getRange(2, 1, ultima - 1, 1).getValues();
    for (let i = 0; i < chaves.length; i++) {
      const chave = String(chaves[i][0]).trim();
      if (chave && !linhas.has(chave)) linhas.set(chave, i + 2);
    }
  }
  return { aba: aba, linhas: linhas, proxima: Math.max(ultima, 1) + 1 };
}

// Grava o clique do portal no store: as 5 manuais, a fotografia do SAP e o
// Atualizado em/por — colunas contiguas (ST_MANUAIS..ST_ATUALIZADO_POR), uma escrita.
// `firmar` e o clique de confirmacao: carimba o Firmado em/por, que o Limpar
// (firmar = false) deixa como esta ate o proximo firme.
function gravarConfirmacaoStore(st, doc, item, sched, manuais, sapData, sapQtd, usuario, agora, firmar) {
  const chave = montarChave(doc, item, sched);
  const bloco = manuais.slice(0, COLUNAS_MANUAIS.length).concat([sapData, sapQtd, agora, usuario]);
  const linha = st.linhas.get(chave);

  if (linha) {
    st.aba.getRange(linha, ST_MANUAIS + 1, 1, bloco.length).setValues([bloco]);
    if (firmar) st.aba.getRange(linha, ST_FIRMADO_EM + 1, 1, 2).setValues([[agora, usuario]]);
    return;
  }

  // Chave fora do store: a sincronizacao ainda nao passou desde a publicacao. Nasce
  // sem "Entrou no sistema em" — a STO ja estava la, so nao se sabe desde quando.
  const nova = novoRegistroStore_(chave, doc, item, sched, agora).valores;
  for (let k = 0; k < bloco.length; k++) nova[ST_MANUAIS + k] = bloco[k];
  if (firmar) { nova[ST_FIRMADO_EM] = agora; nova[ST_FIRMADO_POR] = usuario; }
  st.aba.getRange(st.proxima, 1, 1, STORE_HEADERS.length).setValues([nova]);
  st.linhas.set(chave, st.proxima);
  st.proxima++;
}

function novoRegistroStore_(chave, doc, item, sched, agora) {
  const linha = new Array(STORE_HEADERS.length).fill("");
  linha[ST_CHAVE] = chave;
  linha[ST_DOC] = doc;
  linha[ST_ITEM] = item;
  linha[ST_SCHED] = sched;
  linha[ST_VISTO_EM] = agora;
  linha[ST_STATUS] = "ATIVA";
  return { linha: null, valores: linha };
}

function vazioStore_(v) {
  return v === "" || v === null || v === undefined;
}

// A mesma regra do hasConfirmation do portal: data, quantidade ou o flag confirmados.
function temConfirmacao_(valores) {
  return [0, 1, 2].some(k => !vazioStore_(valores[ST_MANUAIS + k]));
}

// Qualquer uma das 5 manuais preenchida — e o que o restore teria a devolver a ME2W.
function temDadoManual_(valores) {
  for (let k = 0; k < COLUNAS_MANUAIS.length; k++) {
    if (!vazioStore_(valores[ST_MANUAIS + k])) return true;
  }
  return false;
}

function upsertStore(store, doc, item, sched, manuais, sapData, sapQtd, usuario) {
  const chave = montarChave(doc, item, sched);
  const agora = new Date();
  let reg = store.mapa.get(chave);

  if (!reg) {
    reg = novoRegistroStore_(chave, doc, item, sched, agora);
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
// RASTREIO: a lista anterior da ME2W contra o export novo
// ====================================================================
//
// A ME2W como esta na aba, ANTES de o export novo passar por cima, e a lista
// anterior. O que entrou e o que mudou se decide contra ela — o ME2W-Historico
// cresceu demais para ser lido a cada passada. A hora gravada e a da sincronizacao
// que viu a diferenca: o SAP nao manda a hora da alteracao, e duas alteracoes entre
// um export e outro aparecem como uma.

// Devolve Map chave -> { qtd, data } ou null, quando nao ha lista anterior com que
// comparar (aba vazia ou sem as colunas-chave). Null nao e "tudo novo": sem lista
// nada recebe "Entrou no sistema em", em vez de tudo receber a hora de hoje.
//
// Uma leitura por coluna, so das cinco que interessam. A chave pelo TEXTO exibido
// (normalizarNumeroDoc, a mesma leitura do Salvar do portal): com formato de data na
// celula o getValues() devolve um Date no lugar do numero, a chave nao casaria com a
// do export e toda STO pareceria nova.
function lerMe2wAnterior_(aba) {
  if (!aba || aba.getLastRow() < 2) return null;
  const n = aba.getLastRow() - 1;
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const iDoc = headers.indexOf("Purchasing Document");
  const iItem = headers.indexOf("Item");
  const iSched = headers.indexOf("Schedule Line");
  if (iDoc === -1 || iItem === -1 || iSched === -1) return null;

  const coluna = (i, exibido) => {
    if (i === -1) return null;
    const r = aba.getRange(2, i + 1, n, 1);
    return exibido ? r.getDisplayValues() : r.getValues();
  };
  const codigo = (i) => {
    const valores = coluna(i, false), exibidos = coluna(i, true);
    return valores.map((v, k) => normalizarNumeroDoc(v[0], exibidos[k][0]));
  };
  const docs = codigo(iDoc), itens = codigo(iItem), scheds = codigo(iSched);
  const qtds = coluna(headers.indexOf(COL_SAP_QTD), false);
  const datas = coluna(headers.indexOf(COL_SAP_DATA), false);

  const mapa = new Map();
  for (let k = 0; k < n; k++) {
    if (docs[k] === "") continue;
    mapa.set(montarChave(docs[k], itens[k], scheds[k]), {
      qtd: qtds ? qtds[k][0] : "",
      data: datas ? datas[k][0] : ""
    });
  }
  return mapa;
}

// Order Quantity como numero. O export pode trazer "1,000.000" como texto — mesma
// limpeza que o portal faz. Vazio ou ilegivel e null: vazio nao e alteracao.
function qtdSap_(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = String(v == null ? "" : v).trim().replace(/,/g, "");
  if (s === "") return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}

// Delivery Date como { tipo, dia }. Date dos dois lados e o normal (as duas pontas
// passam pelo getValues). Texto nao e convertido: comparar texto com Date acusaria
// mudanca em toda linha no dia em que o export mudasse de formato.
function diaSap_(v) {
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? null : { tipo: "data", dia: Utilities.formatDate(v, TIMEZONE, "yyyy-MM-dd") };
  }
  const s = String(v == null ? "" : v).trim();
  return s === "" ? null : { tipo: "texto", dia: s };
}

// Grava no registro a ULTIMA alteracao: o valor anterior e a hora desta passada.
// So acusa com valor dos dois lados; tipos diferentes na data ficam fora (contados).
function compararComAnterior_(reg, ant, linha, iQtd, iData, agora, cont) {
  if (iQtd !== -1) {
    const a = qtdSap_(ant.qtd), n = qtdSap_(linha[iQtd]);
    if (a !== null && n !== null && Math.abs(a - n) > 1e-9) {
      reg.valores[ST_ULT_QTD] = ant.qtd;
      reg.valores[ST_DT_ALT_QTD] = agora;
      cont.qtd++;
    }
  }
  if (iData !== -1) {
    const a = diaSap_(ant.data), n = diaSap_(linha[iData]);
    if (a && n) {
      if (a.tipo !== n.tipo) {
        cont.naoComparadas++;
      } else if (a.dia !== n.dia) {
        reg.valores[ST_ULT_DD] = ant.data;
        reg.valores[ST_DT_ALT_DD] = agora;
        cont.data++;
      }
    }
  }
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

// Alem de devolver as manuais ao export, e aqui que o store vira a aba de historico:
// toda STO do export ganha linha (nao so as confirmadas), e o registro dela recebe a
// entrada e a ultima alteracao contra a lista anterior (`anterior`, ver
// lerMe2wAnterior_). Null em `anterior` desliga so o rastreio; o restore segue igual.
function aplicarStoreNaMe2w(dados, store, anterior) {
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
  const rastreio = { entraram: 0, qtd: 0, data: 0, naoComparadas: 0 };

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    while (linha.length < headersFinais.length) linha.push("");

    // Linha sem documento (vazia no meio do export) nao e STO: nao ganha registro.
    if (normalizarChave(linha[iDoc]) === "") continue;
    const chave = montarChave(linha[iDoc], linha[iItem], linha[iSched]);
    vistas.add(chave);

    const ant = anterior ? anterior.get(chave) : undefined;
    let reg = store.mapa.get(chave);
    if (!reg) {
      reg = novoRegistroStore_(chave, linha[iDoc], linha[iItem], linha[iSched], agora);
      store.mapa.set(chave, reg);
      // Fora da lista anterior e fora do store: entrou agora. Estava na lista e nao no
      // store e a STO que ja existia antes do rastreio — fica sem data, que nao se sabe.
      if (anterior && ant === undefined) {
        reg.valores[ST_ENTROU] = agora;
        rastreio.entraram++;
      }
    }
    if (ant) compararComAnterior_(reg, ant, linha, iSapQtd, iSapData, agora, rastreio);

    for (let k = 0; k < COLUNAS_MANUAIS.length; k++) {
      linha[base + k] = reg.valores[ST_MANUAIS + k];
    }
    if (temDadoManual_(reg.valores)) restauradas++;

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

  let ausentes = 0;
  const removidas = [];
  store.mapa.forEach((reg, chave) => {
    // Confirmadas antes do rastreio: o Firmado em e o ultimo clique, que e o que o
    // Atualizado em ja guarda enquanto a confirmacao estiver de pe.
    if (vazioStore_(reg.valores[ST_FIRMADO_EM]) && temConfirmacao_(reg.valores) &&
        !vazioStore_(reg.valores[ST_ATUALIZADO_EM])) {
      reg.valores[ST_FIRMADO_EM] = reg.valores[ST_ATUALIZADO_EM];
      reg.valores[ST_FIRMADO_POR] = reg.valores[ST_ATUALIZADO_POR];
    }

    if (vistas.has(chave)) return;

    // Fora do export e sem nada a restaurar: sai do store. Com uma linha por STO viva,
    // guardar as que ja sairam faria a aba crescer para sempre — e a planilha tem
    // teto de 10 milhoes de celulas. Se ela voltar, volta como STO nova.
    if (!temDadoManual_(reg.valores)) {
      removidas.push(chave);
      return;
    }
    // Com confirmacao: marcada ausente, nunca apagada — e o que a restaura se voltar.
    if (String(reg.valores[ST_STATUS]).trim() !== "AUSENTE") {
      reg.valores[ST_STATUS] = "AUSENTE";
      ausentes++;
    }
  });
  removidas.forEach(chave => store.mapa.delete(chave));

  if (rastreio.naoComparadas > 0) {
    console.warn("Rastreio: " + rastreio.naoComparadas + " Delivery Date(s) não comparada(s) — data de um " +
                 "lado e texto do outro entre a ME2W anterior e o export novo.");
  }

  return { restauradas: restauradas, reapareceram: reapareceram, ausentes: ausentes,
           removidas: removidas.length, rastreio: rastreio };
}

// ====================================================================
// LEITURA E ESCRITA DAS BASES
// ====================================================================
//
// AS DUAS ROTAS PASSAM PELO MESMO CAMINHO, que e o do Sincronizacao_Backend do
// Portal de Pedidos: converte o XLSX, le o temporario de uma vez (getValues) e
// grava na aba de uma vez (clearContents + setValues). A rota critica so
// acrescenta validacao, lock e restore do store por cima disso.
//
// O QUE SAIU DAQUI: a copia pelo servico avancado Sheets (Values.get e
// Values.update em blocos, mais Spreadsheets.get de metadados da alvo, mais
// Spreadsheets.get de metadados do temporario, mais batchUpdate de grade, mais
// Values.clear, mais um get de formato numerico e um batchUpdate de repeatCell
// por coluna). Ela existia para transportar data como numero de serie sem
// depender de locale, e cobrava caro por isso:
//
//   - 8 a 12 idas a API por base contra 2 aqui — e cada ida e uma chamada HTTP
//     inteira, que e onde o tempo do gatilho ia embora;
//   - o laco de blocos era guiado pelo rowCount do GRID do temporario, e um XLSX
//     convertido vem com folga de linhas vazias no fim: lia faixa sem dado ate um
//     bloco curto denunciar o fim. O getLastRow() para na ultima linha com
//     conteudo e nao paga essa leitura;
//   - o repeatCell de formato ia de startRowIndex 1 ate o fim da coluna, coluna
//     por coluna, so para o serial nao aparecer como 45123 na tela;
//   - e o servico Sheets precisava estar habilitado no projeto. Sem ele, quatro
//     das cinco bases NUNCA sincronizavam — a ME2W seguia atualizando sozinha e a
//     Pagina Transferencia cruzava STO nova com estoque e pedidos velhos.
//
// O getValues devolve Date de verdade nas colunas de data e o setValues escreve
// Date de verdade: some o serial, some a necessidade de replicar formato e some a
// dependencia do servico avancado. Sobra o Drive, so para converter o XLSX —
// exatamente a dependencia que a sincronizacao do Portal de Pedidos tem.

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

function converterParaSheets(fileObj, nomeTemp) {
  return Drive.Files.create({ name: nomeTemp, mimeType: MimeType.GOOGLE_SHEETS }, fileObj.getBlob()).id;
}

// Uma leitura por base. O retangulo para na ultima celula com conteudo, entao a
// folga de linhas vazias que o XLSX convertido carrega no fim nao entra nele.
function lerValoresDoTemp(tempId) {
  const aba = SpreadsheetApp.openById(tempId).getSheets()[0];
  return aba.getRange(1, 1, Math.max(aba.getLastRow(), 1), Math.max(aba.getLastColumn(), 1)).getValues();
}

function exportVazio(dados) {
  if (!dados || dados.length === 0) return true;
  return dados.length === 1 && (dados[0].length === 0 || String(dados[0][0]).trim() === "");
}

// So cresce a grade, nunca encolhe: apagar linha ou coluna de uma aba que a
// Pagina Transferencia referencia por intervalo produz #REF! irreversivel —
// renomear ou recriar a aba depois nao desfaz. O clearContents limpa o conteudo
// sem mexer na grade, entao a unica coisa a garantir e que ela caiba o que vem.
function garantirGradeAba(aba, linhas, colunas) {
  const maxLinhas = aba.getMaxRows();
  if (linhas > maxLinhas) aba.insertRowsAfter(maxLinhas, linhas - maxLinhas);
  const maxColunas = aba.getMaxColumns();
  if (colunas > maxColunas) aba.insertColumnsAfter(maxColunas, colunas - maxColunas);
}

// Teto por chamada de setValues. O caso normal cabe numa chamada so — e o que a
// sincronizacao do Portal de Pedidos faz com os 8 arquivos dela. O corte existe
// para o dia em que um export crescer a ponto de a chamada unica estourar; ele
// nao muda o total de dado escrito, so parte em pedacos previsiveis.
const LIMITE_CELULAS_ESCRITA = 500000;

function escreverEmBlocos(aba, dados, colunas) {
  const passo = Math.max(1, Math.floor(LIMITE_CELULAS_ESCRITA / Math.max(colunas, 1)));
  for (let i = 0; i < dados.length; i += passo) {
    const bloco = dados.slice(i, Math.min(i + passo, dados.length));
    aba.getRange(i + 1, 1, bloco.length, colunas).setValues(bloco);
  }
}

// Escreve so depois de ter o export inteiro em maos — um export vazio nao pode
// zerar a aba que a Pagina Transferencia le. Era isso que o "limpa so no primeiro
// bloco" da rota antiga tentava garantir; aqui sai de graca, porque o dado ja
// esta todo em memoria quando a aba e tocada.
function atualizarAba(targetSS, nomeAba, dados) {
  if (exportVazio(dados)) return 0;

  // setValues exige retangulo. A leitura do temporario ja devolve um, mas a rota
  // critica concatena as 5 colunas manuais linha a linha antes de chegar aqui.
  let colunas = 0;
  for (let i = 0; i < dados.length; i++) {
    if (dados[i].length > colunas) colunas = dados[i].length;
  }
  for (let i = 0; i < dados.length; i++) {
    while (dados[i].length < colunas) dados[i].push("");
  }

  const aba = obterOuCriarAba(targetSS, nomeAba);
  garantirGradeAba(aba, dados.length, colunas);
  aba.clearContents();
  escreverEmBlocos(aba, dados, colunas);
  return dados.length - 1;
}

// Rota direta: converter, ler, gravar. Sem percorrer linha a linha, sem montar
// chave e sem comparar nada — e o que sobra quando a base nao tem dado manual a
// perder.
function copiarBase(targetSS, nomeAba, tempId) {
  const dados = lerValoresDoTemp(tempId);
  if (exportVazio(dados)) throw new Error("export vazio — aba preservada.");
  return atualizarAba(targetSS, nomeAba, dados);
}

// ====================================================================
// SYNC
// ====================================================================

function processarMe2w(targetSS, tempId) {
  const dados = lerValoresDoTemp(tempId);
  validarExportMe2w(dados, targetSS.getSheetByName(ABA_ME2W));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error("lock ocupado por 60s (portal escrevendo) — adiado para o próximo gatilho.");
  }
  try {
    const store = lerStore(targetSS);
    if (store.mapa.size === 0) semearStoreDaMe2w(targetSS, store);

    // A lista anterior tem de ser lida AGORA: o atualizarAba logo abaixo apaga a aba.
    const anterior = lerMe2wAnterior_(targetSS.getSheetByName(ABA_ME2W));
    if (!anterior) console.warn("ME2W: sem lista anterior com que comparar — rastreio desta passada desligado.");

    const stats = aplicarStoreNaMe2w(dados, store, anterior);
    atualizarAba(targetSS, ABA_ME2W, dados);
    gravarStore(store);
    SpreadsheetApp.flush();

    console.log("ME2W: " + (dados.length - 1) + " linhas | " +
                stats.restauradas + " confirmações restauradas | " +
                stats.reapareceram + " reapareceram | " +
                stats.ausentes + " sumiram do export (preservadas no store) | " +
                stats.removidas + " sem confirmação saíram do store.");
    console.log("Rastreio: " + stats.rastreio.entraram + " STO(s) entraram | " +
                stats.rastreio.qtd + " Order Quantity alterada(s) | " +
                stats.rastreio.data + " Delivery Date alterada(s).");
  } finally {
    lock.releaseLock();
  }
}

function sincronizarNovasBases() {
  const inicio = Date.now();
  const now = new Date();


  const pasta = DriveApp.getFolderById(PASTA_ORIGEM_ID);
  const props = PropertiesService.getScriptProperties();

  // Aberta uma vez, para as cinco bases. Cada openById e uma abertura de planilha
  // inteira; a rota antiga ainda somava a isso os metadados da alvo pela API.
  const targetSS = SpreadsheetApp.openById(PLANILHA_ALVO_ID);

  const erros = [];
  const adiadas = [];
  let escreveu = 0;

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

      if (!dentroDoOrcamento(inicio)) {
        adiadas.push(base.aba);
        continue;
      }

      console.log("Convertendo " + base.arquivo + "...");
      tempId = converterParaSheets(arquivo, "Temp_XLSX_" + base.aba);

      if (base.modo === "critica") {
        processarMe2w(targetSS, tempId);
      } else {
        const linhas = copiarBase(targetSS, base.aba, tempId);
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

  // PASSO 3. Só agora, com as bases já gravadas, faz sentido firmar a página.
  //
  // Em try/catch próprio: firmar é o passo opcional. Falhar aqui não pode desfazer
  // nem mascarar uma sincronização que deu certo — e o pior caso, a página seguir
  // em fórmula, é o comportamento de sempre.
  try {
    firmarAposSync_(inicio, escreveu);
  } catch (e) {
    console.error("Pagina Transferência não foi firmada: " + e.message);
  }
}

// ====================================================================
// PASSO 3 DO FLUXO — FIRMAR
// ====================================================================
//
// O gatilho tem tres passos, e cada um so paga o custo se o anterior deu o que
// fazer:
//
//   1. VERIFICAR    o carimbo do .xlsx contra o guardado (chaveSync). Arquivo
//                   inalterado nao e convertido nem lido — nem entra no laco.
//   2. SINCRONIZAR  converte, le o temporario e grava na aba. `escreveu` conta
//                   quantas bases mudaram de fato.
//   3. FIRMAR       aqui. As contas da pagina precisam enxergar a base nova
//                   antes de virar valor, entao este passo nunca vem antes.
//
// O passo 3 tem DOIS CAMINHOS, e manter os dois separados e o que faz ele caber
// no orcamento do gatilho:
//
//   Calculo.gs   refaz a conta em JS, com indice, e grava valor direto. E o
//                caminho das 19 colunas Y..AQ. Nao recalcula a planilha e nao
//                espera derrame.
//   Firmar.gs    repoe a formula, espera estabilizar e congela o resultado. So
//                para coluna marcada que o Calculo.gs NAO sabe calcular.
//
// O `excluir` e o que mantem a divisao de pe. Sem ele o caminho generico repunha
// a formula das MESMAS colunas que o Calculo.gs tinha acabado de gravar: AB e AJ
// voltavam aos 365 SUMIFS por linha que o Calculo.gs existe para nao pagar, o
// fmAguardar_ estourava o tempo esperando o recalculo e a pagina terminava o
// ciclo em formula — como se o passo 1 nunca tivesse rodado.
//
// Os typeof cobrem o projeto que ainda nao recebeu o Firmar.gs ou o Calculo.gs:
// sem eles a sincronizacao segue exatamente como antes, sem ReferenceError.
function firmarAposSync_(inicio, escreveu) {
  if (typeof fmPrecisaRefirmar_ !== "function") return;

  if (!fmPrecisaRefirmar_(escreveu > 0)) {
    console.log("Pagina Transferência: nenhuma base nova e o dia não virou — já está firmada, passo 3 pulado.");
    return;
  }

  // Precisa caber no que sobrou do gatilho de 6 min. Menos de 30 s restantes
  // não dá para nada além de começar e ser morto no meio.
  const restante = TETO_GATILHO_MS - (Date.now() - inicio);
  if (restante < 30000) {
    console.log("Sem tempo de gatilho para firmar a Pagina Transferência — fica para o próximo.");
    return;
  }

  const temCalculo = typeof firmarColunasCalculadasTransferencia === "function";
  const temGenerico = typeof refirmarPaginaTransferencia === "function";

  // As colunas que o Calculo.gs resolve, para o caminho genérico não encostar
  // nelas. Lista do que ele CONHECE, não do que gravou nesta passada: origem
  // atrasada deixa a coluna um ciclo velha (comportamento documentado), o que é
  // muito melhor do que devolvê-la para a fórmula inviável.
  const doCalculo = (temCalculo && typeof ptLetrasCalculadas_ === "function")
    ? ptLetrasCalculadas_()
    : [];

  if (temCalculo) firmarColunasCalculadasTransferencia({});

  if (temGenerico) {
    const sobrou = TETO_GATILHO_MS - (Date.now() - inicio);
    if (sobrou >= 30000) {
      refirmarPaginaTransferencia({ excluir: doCalculo, esperaMaxMs: sobrou - 15000 });
    } else {
      console.log("Sem tempo para o caminho genérico nesta passada — fica para o próximo gatilho.");
    }
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
