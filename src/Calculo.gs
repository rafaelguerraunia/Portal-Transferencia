// ====================================================================
// AS COLUNAS CALCULADAS DA "Pagina Transferencia", EM JS
// ====================================================================
//
// Este arquivo faz pela Pagina Transferencia o que o firmarColunasCalculadasResumo
// faz pela "Resumo P.O" do Portal de Pedidos: refaz em memoria, com indice, a
// conta que a planilha refazia a cada recalculo — e grava valor.
//
// POR QUE NAO BASTAVA CONGELAR O RESULTADO DA FORMULA (Firmar.gs): duas das
// colunas nao sao caras, sao inviaveis. "Dias Disponiveis em Estoque" (AB) e
// "Dias Disponiveis na BR14" (AJ) tem um MAP DENTRO DE OUTRO MAP:
//
//     MAP( <linhas da pagina>, LAMBDA(w, n, i,
//       MATCH(TRUE, MAP(SEQUENCE(365,1,0), LAMBDA(dias,
//         SUMIFS(RESB!D:D, RESB!A:A,n, RESB!B:B,i,
//                RESB!C:C,">="&TODAY(), RESB!C:C,"<="&TODAY()+dias) >= w )), 0) - 1 ))
//
// Sao 365 SUMIFS por linha da pagina, e cada SUMIFS varre RESB!A:D inteiro. Com
// ~89 mil linhas isso da 32,6 MILHOES de SUMIFS por coluna — 65 milhoes somando
// as duas, a cada recalculo. Nenhuma espera de estabilizacao resolve isso; o que
// resolve e nao fazer a conta desse jeito.
//
// O QUE A MESMA PERGUNTA CUSTA AQUI: "em quantos dias o estoque acaba" e uma
// serie de consumo ordenada por data com soma acumulada, percorrida uma vez. O
// indice de RESB e montado uma vez por execucao (ptIndexResb_) e serve as ~89 mil
// linhas. De O(linhas x 365 x RESB) para O(linhas + RESB).
//
// O QUE MAIS MUDA DE FIGURA:
//   - Os tres IMPORTRANGE (Pre-Agendamento, Plano_Transporte) somem. As planilhas
//     sao abertas por openById, direto. E a mesma decisao do rpFontesStos_ do lado
//     de Pedidos, pelo mesmo motivo: IMPORTRANGE tem cache proprio e exibe o
//     retrato de horas atras sem um #N/A sequer para denunciar.
//   - As varreduras de coluna inteira ('Stock Control ...'!G:G, ME2N!M:M) viram
//     Map montado numa passada.
//
// O QUE NAO E TOCADO: a A2 continua formula. Ela e o VSTACK/FILTER que monta o
// corpo da pagina (A:X) a partir da ME5A e da ME2W — e e barata perto do resto.
// As colunas daqui so LEEM esse derrame.
//
// DEPENDENCIA ENTRE COLUNAS CALCULADAS: AB le Z, AD le Z e AC, AJ le AC. Por isso
// o laco calcula estoque (Z, AC) antes de consumo (Y, AA, AB, AD, AJ) — a ordem
// esta explicita em ptCalcularLinha_ e nao pode ser embaralhada.

// Planilhas de origem que nao sao esta. Os ids sao os mesmos que estavam dentro
// dos IMPORTRANGE das formulas AO/AP (Pre-Agendamento) e AQ (Plano_Transporte).
const PT_PEDIDOS_DB_ID = "1hWY9E5LzKt5xR98l_pDKumfT7sdy-Bzp1DDyLUGHPTg";
const PT_TRANSPORTES_DB_ID = "1NqwL4ePjBHl4bU8SKiAyp5cQnPr3474b8vMKyo0FlgI";

const PT_ABA_PRE_AGEND = "Pré-Agendamento";
const PT_ABA_PLANO_TRANSPORTE = "Plano_Transporte";
const PT_ABA_CONVERSAO = "Unit / Palete / S_N Empilha";
const PT_ABA_ESTOQUE = "Stock Control BR14 BR10 BR12";
const PT_ABA_CONSUMO = "RESB";
const PT_ABA_PEDIDOS = "ME2N";

// Colunas de ORIGEM dentro do derrame da A2 (0-based). Sao posicionais porque as
// formulas originais eram posicionais: quem mexer na ordem do HSTACK da A2 tem de
// mexer aqui junto.
const PT_O_DOC = 6;      // G  Documento
const PT_O_ITEM = 7;     // H  Item
const PT_O_PLANTA = 8;   // I  Planta_ Destino   (o "i" das formulas)
const PT_O_MATERIAL = 13;// N  Material          (o "n" das formulas)
const PT_O_QTD = 16;     // Q  quantidade usada no Pallet Order
const PT_O_SCHED = 20;   // U  Schedule Line
const PT_LARGURA_DERRAME = 24;   // A:X

// Colunas de DESTINO. `grupo` e a origem: e o que permite pular so as colunas
// cuja base nao chegou, mantendo o que ja estava la. Mesma ideia do
// RESUMO_COLUNAS_CALCULADAS.
const PT_COLUNAS_CALCULADAS = [
  { col: 25, letra: "Y",  nome: "Req Semana",                             grupo: "consumo" },
  { col: 26, letra: "Z",  nome: "Estoque na Planta Destino",              grupo: "estoque" },
  { col: 27, letra: "AA", nome: "Há Estoque para Semana?",                grupo: "consumo" },
  { col: 28, letra: "AB", nome: "Dias Disponiveis em Estoque",            grupo: "consumo" },
  { col: 29, letra: "AC", nome: "Estoque BR14 Destinado para Planta",     grupo: "estoque" },
  { col: 30, letra: "AD", nome: "Estoque Conjunto de Plantas",            grupo: "consumo" },
  { col: 31, letra: "AE", nome: "Qtd Pedidos Abertos",                    grupo: "pedidos" },
  { col: 32, letra: "AF", nome: "Lista de Pedidos",                       grupo: "pedidos" },
  { col: 33, letra: "AG", nome: "Fornecedor Pedido",                      grupo: "pedidos" },
  { col: 34, letra: "AH", nome: "Deliver Date Pedido",                    grupo: "pedidos" },
  { col: 35, letra: "AI", nome: "Qtd. Pedido",                            grupo: "pedidos" },
  { col: 36, letra: "AJ", nome: "Dias Disponiveis na BR14",               grupo: "consumo" },
  { col: 37, letra: "AK", nome: "Estoque BR14 Sub Destinado para Planta", grupo: "estoque" },
  { col: 38, letra: "AL", nome: "Pallet Order",                           grupo: "conversao" },
  { col: 39, letra: "AM", nome: "Planta Pedido",                          grupo: "pedidos" },
  { col: 40, letra: "AN", nome: "Insp Qualidade",                         grupo: "estoque" },
  { col: 41, letra: "AO", nome: "Pré Agendado?",                          grupo: "preAgend" },
  { col: 42, letra: "AP", nome: "Separado?",                              grupo: "preAgend" },
  { col: 43, letra: "AQ", nome: "Status Transporte",                      grupo: "transporte" }
];

// As letras que ESTE arquivo sabe calcular. O Sync passa a lista ao caminho
// generico do Firmar.gs para ele NAO repor a formula destas colunas: repor
// AB e AJ e voltar aos 365 SUMIFS por linha que este arquivo existe para nao
// pagar, e a espera pelo recalculo estoura o tempo do gatilho.
//
// A lista e o que o arquivo CONHECE, nao o que ele gravou nesta passada. A
// diferenca aparece quando uma origem nao esta pronta (RESB atrasada, por
// exemplo): a coluna fica um ciclo velha, que e o comportamento documentado, em
// vez de cair no caminho generico e arrastar a planilha inteira para o recalculo
// que ninguem consegue esperar.
function ptLetrasCalculadas_() {
  return PT_COLUNAS_CALCULADAS.map(function (c) { return c.letra; });
}

// Depositos que nao contam como saldo — a coluna E da aba de estoque, que as
// formulas excluiam repetindo 'Stock Control...'!E:E,"<>OB","<>DAOB",...
const PT_ESTOQUE_E_BLOQUEADOS = { "OB": 1, "DAOB": 1, "FD": 1, "FDT": 1, "DR": 1 };

// Horizonte da "Req Semana" e das duas colunas que decidem em cima dela.
const PT_JANELA_REQ = 7;

// Ate onde a busca de cobertura vai. E o SEQUENCE(365) das formulas AB/AJ: passou
// disso, a resposta e o rotulo fixo — que as formulas escreviam como
// "Superior a 30 dias" mesmo procurando 365. Rotulo preservado como estava: o
// STO-Backend testa por .includes("Superior") e vira 999 na prioridade.
const PT_HORIZONTE_DIAS = 365;
const PT_ROTULO_ALEM_DO_HORIZONTE = "Superior a 30 dias";
const PT_ROTULO_SEM_CONSUMO = "Sem Consumo Previsto";
const PT_ROTULO_SEM_CONVERSAO = "Sem Estoque para Convesão";
const PT_ROTULO_SEM_PEDIDO = "Nenhum pedido encontrado";

const PT_LOTE_CELULAS = 20000;

// ---------------------------------------------------------------- helpers ---

function ptTxt_(v) {
  return String(v === null || v === undefined ? "" : v).replace(/[﻿​]/g, "").trim();
}

function ptNum_(v) {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const s = ptTxt_(v).replace(/\s/g, "").replace(/,/g, "");
  if (s === "") return 0;
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

// Chave de comparacao que aguenta os dois lados virem com tipos diferentes. O
// mesmo material volta como numero 123 de um lado e texto "0000123" do outro,
// conforme o export do SAP tenha ou nao preservado os zeros a esquerda.
function ptChave_(v) {
  if (v instanceof Date) return String(v.getTime());
  let s = ptTxt_(v).toUpperCase();
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, "");
  return s;
}

function ptChaveDupla_(a, b) { return ptChave_(a) + "|" + ptChave_(b); }

// A chave "Documento / Item / Schedule Line" que o Pre-Agendamento e o
// Plano_Transporte usam. Normalizada dos DOIS lados: a planilha externa foi
// preenchida pela concatenacao do Sheets, e um zero a esquerda de diferenca
// bastava para o XLOOKUP nao achar.
function ptChaveDoc_(doc, item, sched) {
  return ptChave_(doc) + " / " + ptChave_(item) + " / " + ptChave_(sched);
}

// Dia em numero de serie da planilha (epoca 30/12/1899). Aceita as tres formas em
// que uma data pode estar na aba: Date (o que as duas rotas do Sync escrevem
// hoje — as duas leem e gravam por SpreadsheetApp), numero de serie (o que a
// antiga copia via Sheets API gravava, e que sobrevive nas abas ainda nao
// resincronizadas desde a troca) e texto.
const PT_EPOCA = new Date(1899, 11, 30).getTime();

function ptDia_(v) {
  if (v instanceof Date) {
    const t = v.getTime();
    if (isNaN(t)) return null;
    return Math.round((new Date(v.getFullYear(), v.getMonth(), v.getDate()).getTime() - PT_EPOCA) / 86400000);
  }
  if (typeof v === "number") return isFinite(v) ? Math.floor(v) : null;

  const s = ptTxt_(v);
  if (s === "") return null;

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return ptDia_(new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])));

  // Data com barra. Aqui mora a armadilha: "8/25/2026" e "25/8/2026" sao a MESMA
  // data escrita nas duas convencoes, e "3/11/2026" e duas datas diferentes sem
  // nada no valor que diga qual. Errar aqui desloca a data em silencio.
  //
  // A regra e: deixar o proprio numero decidir sempre que ele puder.
  //   - componente > 12 so pode ser DIA, e isso fixa a convencao da string;
  //   - com os dois <= 12 nao ha o que deduzir, e a escolha e M/D — que e o que
  //     esta planilha produz (conferido no export da Pagina Transferencia:
  //     "8/25/2026", "9/3/2026") e o que o ptDataTexto_ escreve de volta, para a
  //     ida e volta fechar.
  //
  // Na pratica quase nada chega aqui: o Sync entrega data como Date nas duas
  // rotas. Este ramo e a rede de seguranca.
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const a = Number(m[1]), b = Number(m[2]), ano = Number(m[3]);
    const mes = (a > 12) ? b : a;
    const dia = (a > 12) ? a : b;
    if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;
    return ptDia_(new Date(ano, mes - 1, dia));
  }

  const n = Number(s);
  return isFinite(n) && n > 0 ? Math.floor(n) : null;
}

function ptHoje_() {
  const d = new Date();
  return Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() - PT_EPOCA) / 86400000);
}

// M/D/YYYY, sem zero a esquerda — EXATAMENTE o que o TEXTJOIN da formula AH
// entrega hoje ("8/25/2026", "9/3/2026"), conferido contra o export da pagina.
//
// NAO usar ISO aqui. O STO-Backend faz new Date(arrDeliv[j]) e compara com um
// `hoje` que e meia-noite LOCAL; "2026-08-25" e parseado como meia-noite UTC, que
// em America/Sao_Paulo cai no dia 24 as 21h — toda entrega do proprio dia passaria
// a contar como atrasada. O formato M/D/YYYY e parseado como hora local e nao tem
// esse deslocamento.
//
// Fixo em M/D/YYYY em vez de seguir o locale da planilha: e o unico formato que o
// new Date() nao interpreta errado, e escrever d/M/yyyy aqui inverteria dia e mes
// silenciosamente em toda data com dia <= 12.
function ptDataTexto_(dia) {
  if (dia === null) return "";
  const d = new Date(PT_EPOCA + dia * 86400000);
  return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
}

// ROUNDUP(x, 2) do Sheets: arredonda para CIMA na segunda casa, nao para o par
// mais proximo. Math.ceil sobre o valor escalado e o unico jeito de bater.
function ptArredondarCima_(v, casas) {
  const f = Math.pow(10, casas);
  return Math.ceil(v * f - 1e-9) / f;
}

function ptLerAba_(ss, nome, largura) {
  const sh = ss.getSheetByName(nome);
  if (!sh) return { ok: false, motivo: "aba '" + nome + "' não existe", dados: [] };
  const ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return { ok: false, motivo: "aba '" + nome + "' sem linhas de dados", dados: [] };

  const cols = Math.min(Math.max(sh.getLastColumn(), 1), largura);
  const dados = sh.getRange(2, 1, ultimaLinha - 1, cols).getValues();

  // Uma aba pode existir e estar em "#N/A / Loading..." — calcular em cima disso
  // gravaria zero em coluna que tinha valor.
  for (let i = 0; i < Math.min(dados.length, 5); i++) {
    for (let j = 0; j < dados[i].length; j++) {
      const v = dados[i][j];
      if (typeof v === "string" && v.charAt(0) === "#") {
        return { ok: false, motivo: "aba '" + nome + "' em " + v + " (origem ainda carregando)", dados: [] };
      }
    }
  }
  return { ok: true, dados: dados };
}

// ---------------------------------------------------------------- indices ---

// CONSUMO (RESB): A=Material, B=Planta, C=ReqmtsDate, D=Qtd.
//
// Por chave material+planta guarda duas coisas, as duas ja resolvidas aqui para
// que o laco principal nao volte a varrer nada:
//
//   somaJanela — soma de tudo com data <= hoje+7. SEM piso inferior, porque a
//                formula da "Req Semana" tambem nao tinha: ela conta o consumo
//                ATRASADO junto. Mantido como estava.
//   serie      — eventos com data >= hoje, ordenados, com soma acumulada. E
//                sobre ela que "em quantos dias acaba" vira uma varredura unica.
function ptIndexResb_(ss, hoje) {
  const r = ptLerAba_(ss, PT_ABA_CONSUMO, 4);
  if (!r.ok) return { ok: false, motivo: r.motivo };

  const porChave = new Map();
  const limiteJanela = hoje + PT_JANELA_REQ;

  for (let i = 0; i < r.dados.length; i++) {
    const linha = r.dados[i];
    const mat = ptChave_(linha[0]);
    if (mat === "") continue;

    const chave = mat + "|" + ptChave_(linha[1]);
    const dia = ptDia_(linha[2]);
    const qtd = ptNum_(linha[3]);

    let reg = porChave.get(chave);
    if (!reg) { reg = { somaJanela: 0, eventos: [] }; porChave.set(chave, reg); }

    if (dia !== null && dia <= limiteJanela) reg.somaJanela += qtd;
    if (dia !== null && dia >= hoje) reg.eventos.push([dia, qtd]);
  }

  // Ordena uma vez e acumula uma vez. Depois disso, responder "em quantos dias o
  // estoque w acaba" e caminhar na serie ate a acumulada alcancar w.
  porChave.forEach(function (reg) {
    reg.eventos.sort(function (a, b) { return a[0] - b[0]; });
    let acumulado = 0;
    for (let k = 0; k < reg.eventos.length; k++) {
      acumulado += reg.eventos[k][1];
      reg.eventos[k][1] = acumulado;
    }
    reg.totalFuturo = acumulado;
  });

  return { ok: true, mapa: porChave };
}

// O MATCH(TRUE, MAP(SEQUENCE(365), ...)) das formulas AB/AJ, sem os 365 SUMIFS.
// A acumulada so cresce em dia de evento, entao o primeiro dia em que ela alcanca
// o estoque e o primeiro EVENTO em que ela alcanca — nao ha o que testar entre um
// evento e o seguinte.
function ptDiasCobertura_(reg, estoque, hoje) {
  if (!reg || reg.totalFuturo === 0) return PT_ROTULO_SEM_CONSUMO;

  const eventos = reg.eventos;
  for (let k = 0; k < eventos.length; k++) {
    if (eventos[k][1] >= estoque) {
      const dias = eventos[k][0] - hoje;
      return dias <= PT_HORIZONTE_DIAS - 1 ? dias : PT_ROTULO_ALEM_DO_HORIZONTE;
    }
  }
  return PT_ROTULO_ALEM_DO_HORIZONTE;
}

// ESTOQUE: A=Material, D=Planta, E=Storage Location, F=Tipo, G=Qtd.
// Quatro somas diferentes sobre a mesma aba, montadas numa passada so.
function ptIndexEstoque_(ss) {
  const r = ptLerAba_(ss, PT_ABA_ESTOQUE, 7);
  if (!r.ok) return { ok: false, motivo: r.motivo };

  const destino = new Map();   // Z  : mat|planta      livre  (F<>BL, F<>SC)
  const br14 = new Map();      // AC : mat|storageLoc  livre, com D=BR14
  const br14Sub = new Map();   // AK : mat             F=SC,  com D=BR14
  const qualidade = new Map(); // AN : mat|storageLoc  F=QI

  function somar(mapa, chave, qtd) {
    mapa.set(chave, (mapa.get(chave) || 0) + qtd);
  }

  for (let i = 0; i < r.dados.length; i++) {
    const linha = r.dados[i];
    const mat = ptChave_(linha[0]);
    if (mat === "") continue;

    const planta = ptChave_(linha[3]);
    const eLoc = ptChave_(linha[4]);
    const tipo = ptChave_(linha[5]);
    const qtd = ptNum_(linha[6]);
    if (PT_ESTOQUE_E_BLOQUEADOS[eLoc]) continue;

    const livre = (tipo !== "BL" && tipo !== "SC");

    if (livre) somar(destino, mat + "|" + planta, qtd);
    if (livre && planta === "BR14") somar(br14, mat + "|" + eLoc, qtd);
    if (tipo === "SC" && planta === "BR14") somar(br14Sub, mat, qtd);
    if (tipo === "QI") somar(qualidade, mat + "|" + eLoc, qtd);
  }

  return { ok: true, destino: destino, br14: br14, br14Sub: br14Sub, qualidade: qualidade };
}

// PEDIDOS (ME2N): A=Planta, B=Documento, D=Material, E=Storage Location,
// I=Deliver Date, M=Saldo a entregar, N=Fornecedor.
function ptIndexMe2n_(ss) {
  const r = ptLerAba_(ss, PT_ABA_PEDIDOS, 14);
  if (!r.ok) return { ok: false, motivo: r.motivo };

  const porMaterial = new Map();

  for (let i = 0; i < r.dados.length; i++) {
    const linha = r.dados[i];
    const mat = ptChave_(linha[3]);
    if (mat === "") continue;

    let lista = porMaterial.get(mat);
    if (!lista) { lista = []; porMaterial.set(mat, lista); }

    lista.push({
      planta: ptChave_(linha[0]),
      doc: ptTxt_(linha[1]),
      loc: ptChave_(linha[4]),
      dia: ptDia_(linha[8]),
      saldo: ptNum_(linha[12]),
      fornecedor: ptTxt_(linha[13])
    });
  }

  return { ok: true, mapa: porMaterial };
}

// CONVERSAO: A=Material, F=unidades por palete.
function ptIndexConversao_(ss) {
  const r = ptLerAba_(ss, PT_ABA_CONVERSAO, 6);
  if (!r.ok) return { ok: false, motivo: r.motivo };

  const m = new Map();
  for (let i = 0; i < r.dados.length; i++) {
    const mat = ptChave_(r.dados[i][0]);
    if (mat === "" || m.has(mat)) continue;   // XLOOKUP pega a primeira ocorrencia
    m.set(mat, ptNum_(r.dados[i][5]));
  }
  return { ok: true, mapa: m };
}

// PRE-AGENDAMENTO, na planilha do Portal de Pedidos: B = "PO / Item / SL",
// T = Separado. Lida por openById em vez de IMPORTRANGE — ver o cabecalho.
function ptIndexPreAgend_() {
  try {
    const ss = SpreadsheetApp.openById(PT_PEDIDOS_DB_ID);
    const r = ptLerAba_(ss, PT_ABA_PRE_AGEND, 20);
    if (!r.ok) return { ok: false, motivo: r.motivo };

    const m = new Map();
    for (let i = 0; i < r.dados.length; i++) {
      const bruto = ptTxt_(r.dados[i][1]);
      if (bruto === "") continue;
      const partes = bruto.split("/");
      const chave = partes.length === 3
        ? ptChaveDoc_(partes[0], partes[1], partes[2])
        : ptChave_(bruto);
      // MATCH pega a PRIMEIRA ocorrencia; manter a primeira e o que reproduz isso.
      if (!m.has(chave)) m.set(chave, ptChave_(r.dados[i][19]) === "SIM");
    }
    return { ok: true, mapa: m };
  } catch (e) {
    return { ok: false, motivo: "Pré-Agendamento inacessível: " + e.message };
  }
}

// PLANO DE TRANSPORTE, na planilha de Transportes: E = chave, BV = status.
function ptIndexTransporte_() {
  try {
    const ss = SpreadsheetApp.openById(PT_TRANSPORTES_DB_ID);
    const sh = ss.getSheetByName(PT_ABA_PLANO_TRANSPORTE);
    if (!sh) return { ok: false, motivo: "aba '" + PT_ABA_PLANO_TRANSPORTE + "' não existe" };
    const ultimaLinha = sh.getLastRow();
    if (ultimaLinha < 2) return { ok: false, motivo: "'" + PT_ABA_PLANO_TRANSPORTE + "' sem linhas" };

    // Só as duas colunas que interessam, não a faixa E:BV inteira: são 70 colunas
    // de diferença numa aba que cresce sozinha.
    const chaves = sh.getRange(2, 5, ultimaLinha - 1, 1).getValues();      // E
    const status = sh.getRange(2, 74, ultimaLinha - 1, 1).getValues();     // BV

    const m = new Map();
    for (let i = 0; i < chaves.length; i++) {
      const bruto = ptTxt_(chaves[i][0]);
      if (bruto === "") continue;
      const partes = bruto.split("/");
      const chave = partes.length === 3
        ? ptChaveDoc_(partes[0], partes[1], partes[2])
        : ptChave_(bruto);
      if (!m.has(chave)) m.set(chave, ptTxt_(status[i][0]));
    }
    return { ok: true, mapa: m };
  } catch (e) {
    return { ok: false, motivo: "Plano_Transporte inacessível: " + e.message };
  }
}

// ------------------------------------------------------------- a linha -----

// Calcula as 19 colunas de UMA linha da pagina. Devolve um objeto indexado pela
// letra da coluna — quem grava decide o que aproveita, conforme o grupo estar
// pronto ou nao.
//
// A ORDEM AQUI NAO E LIVRE: AB depende de Z, AD de Z e AC, AJ de AC.
function ptCalcularLinha_(linha, ix, hoje) {
  const material = ptChave_(linha[PT_O_MATERIAL]);
  const planta = ptChave_(linha[PT_O_PLANTA]);
  const out = {};

  // Material vazio = linha fora do derrame. As formulas devolviam "" em todas as
  // colunas nesse caso (IF(n="", "", ...)), e e o que se mantem.
  if (material === "") {
    PT_COLUNAS_CALCULADAS.forEach(function (c) { out[c.letra] = ""; });
    return out;
  }

  const chaveMatPlanta = material + "|" + planta;

  // --- estoque, primeiro: as colunas de consumo decidem em cima dele ---
  const estoqueDestino = ix.estoque ? (ix.estoque.destino.get(chaveMatPlanta) || 0) : "";
  const estoqueBr14 = ix.estoque ? (ix.estoque.br14.get(chaveMatPlanta) || 0) : "";
  out["Z"] = estoqueDestino;
  out["AC"] = estoqueBr14;
  out["AK"] = ix.estoque ? (ix.estoque.br14Sub.get(material) || 0) : "";
  out["AN"] = ix.estoque ? (ix.estoque.qualidade.get(chaveMatPlanta) || 0) : "";

  // --- consumo ---
  if (ix.resb) {
    const reg = ix.resb.get(chaveMatPlanta);
    const reqSemana = reg ? reg.somaJanela : 0;
    out["Y"] = reqSemana;

    const w = ptNum_(estoqueDestino);
    const z = ptNum_(estoqueBr14);
    out["AA"] = w >= reqSemana ? "✅ Suficiente" : "🚨 Risco de Ruptura";
    out["AD"] = (w + z) >= reqSemana ? "✅ Suficiente" : "🚨 Risco de Ruptura";
    out["AB"] = w <= 0 ? 0 : ptDiasCobertura_(reg, w, hoje);
    out["AJ"] = z <= 0 ? 0 : ptDiasCobertura_(reg, z, hoje);
  }

  // --- pedidos em aberto ---
  if (ix.me2n) {
    const todos = ix.me2n.get(material) || [];

    // AE somava ME2N!M SEM filtrar saldo — SUMIFS(M, D, n) e so isso.
    let soma = 0;
    for (let k = 0; k < todos.length; k++) soma += todos[k].saldo;
    out["AE"] = soma;

    // AF/AG/AH/AI: FILTER(..., M<>0, D=n). Mesma lista, mesma ordem, quatro
    // projecoes — e o que mantem as quatro colunas alinhadas posicao a posicao,
    // que e como o STO-Backend as consome (arrDeliv[j] com arrQtd[j]).
    const abertos = todos.filter(function (p) { return p.saldo !== 0; });
    if (abertos.length === 0) {
      out["AF"] = PT_ROTULO_SEM_PEDIDO;
      out["AG"] = PT_ROTULO_SEM_PEDIDO;
      out["AH"] = PT_ROTULO_SEM_PEDIDO;
      out["AI"] = PT_ROTULO_SEM_PEDIDO;
    } else {
      out["AF"] = abertos.map(function (p) { return p.doc; }).join(" | ");
      out["AG"] = abertos.map(function (p) { return p.fornecedor; }).join(" | ");
      // Mesmo formato que a formula ja produz (ver ptDataTexto_): o STO-Backend
      // parseia estas datas com new Date(), e mudar o formato aqui deslocaria as
      // comparacoes de atraso em um dia.
      out["AH"] = abertos.map(function (p) { return ptDataTexto_(p.dia); }).join(" | ");
      out["AI"] = abertos.map(function (p) { return p.saldo; }).join(" | ");
    }

    // AM tinha FILTRO PROPRIO, mais restrito que o das quatro acima — e por isso
    // produzia uma lista mais curta, desalinhando o zip do STO-Backend. Aqui a
    // lista percorrida e a MESMA (abertos); a regra de planta vira o VALOR de
    // cada posicao, e as cinco colunas passam a ter o mesmo comprimento.
    if (abertos.length === 0) {
      out["AM"] = PT_ROTULO_SEM_PEDIDO;
    } else {
      out["AM"] = abertos.map(function (p) {
        if (planta === "BR14") {
          return (p.planta === "BR14" && (p.loc === "FG" || p.loc === "PAC")) ? p.planta : "";
        }
        return (p.planta === planta || (p.planta === "BR14" && p.loc === planta)) ? p.planta : "";
      }).join(" | ");
    }
  }

  // --- conversao de palete ---
  if (ix.conversao) {
    const fator = ix.conversao.get(material);
    const q = ptNum_(linha[PT_O_QTD]);
    out["AL"] = (!fator || fator === 0) ? PT_ROTULO_SEM_CONVERSAO : ptArredondarCima_(q / fator, 2);
  }

  // --- agendamento e transporte ---
  const chaveDoc = ptChaveDoc_(linha[PT_O_DOC], linha[PT_O_ITEM], linha[PT_O_SCHED]);

  if (ix.preAgend) {
    const reg = ix.preAgend.get(chaveDoc);
    out["AO"] = reg === undefined ? "Não" : "Sim";
    out["AP"] = reg === true ? "Sim" : "Não";
  }
  if (ix.transporte) {
    out["AQ"] = ix.transporte.get(chaveDoc) || "";
  }

  return out;
}

// ------------------------------------------------------------- gravacao ----

// Grava em corridas de colunas ADJACENTES e PRONTAS. Um retangulo do minimo ao
// maximo passaria por cima de qualquer coluna que nao esteja na lista e das
// colunas cujo grupo foi pulado.
function ptGravarColunas_(sh, colunasProntas, valores, totalLinhas) {
  const ordenadas = colunasProntas.slice().sort(function (a, b) { return a - b; });
  const runs = [];
  for (let i = 0; i < ordenadas.length; i++) {
    const ultimo = runs.length ? runs[runs.length - 1] : null;
    if (ultimo && ordenadas[i] === ultimo[ultimo.length - 1] + 1) ultimo.push(ordenadas[i]);
    else runs.push([ordenadas[i]]);
  }

  const maxLinhas = sh.getMaxRows();

  runs.forEach(function (run) {
    const passo = Math.max(1, Math.floor(PT_LOTE_CELULAS / run.length));
    for (let ini = 0; ini < totalLinhas; ini += passo) {
      const n = Math.min(passo, totalLinhas - ini);
      const bloco = new Array(n);
      for (let i = 0; i < n; i++) {
        const linha = new Array(run.length);
        for (let j = 0; j < run.length; j++) linha[j] = valores.get(run[j])[ini + i];
        bloco[i] = linha;
      }
      sh.getRange(2 + ini, run[0], n, run.length).setValues(bloco);
    }
    // O derrame da A2 encolhe quando caem STOs. Sem limpar, as linhas orfas
    // ficariam com o resultado da passada anterior — dado velho sem nenhuma linha
    // que o explique.
    const sobra = maxLinhas - (1 + totalLinhas);
    if (sobra > 0) sh.getRange(2 + totalLinhas, run[0], sobra, run.length).clearContent();
  });

  return runs.length;
}

// ------------------------------------------------------------ principal ----

function firmarColunasCalculadasTransferencia(opcoes) {
  const opts = opcoes || {};
  const lock = LockService.getScriptLock();
  // tryLock(0): se outra execucao ja esta calculando, esta desiste. E o mesmo
  // lock do Sync e dos saves do portal — calcular por cima de uma confirmacao
  // que acabou de ser salva gravaria a pagina sem ela.
  if (!lock.tryLock(opts.esperaLockMs || 0)) {
    console.log("Pagina Transferência: outra execução está com o lock. Cálculo ignorado.");
    return null;
  }

  const inicio = Date.now();
  try {
    const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(MF_ABA_ALVO);
    if (!sh) {
      console.error("Aba '" + MF_ABA_ALVO + "' não encontrada. Nada a calcular.");
      return null;
    }

    // Pelo DERRAME, nao pelo getLastRow(). O getLastRow() responde pela ultima
    // celula com conteudo da aba, e abaixo do derrame sobra o que ficou de uma
    // pagina maior — no export atual, 433 linhas com A:X vazias e AO/AP/AQ
    // escritas. Dimensionar por ele fazia este laco calcular 694 linhas para
    // produzir 261 de dado e 433 de "" (ver ptCalcularLinha_, material vazio).
    //
    // As orfas nao ficam para tras: o ptGravarColunas_ limpa de 2+totalLinhas
    // ate o fim da grade, entao a primeira passada ja devolve o getLastRow() ao
    // tamanho da pagina.
    const ultimaLinha = (typeof fmUltimaLinhaDerrame_ === "function")
      ? fmUltimaLinhaDerrame_(sh)
      : sh.getLastRow();
    if (ultimaLinha < 2) {
      console.warn("Pagina Transferência sem linhas de dados. Nada a calcular.");
      return null;
    }
    const totalLinhas = ultimaLinha - 1;

    // O derrame da A2 (A:X). É a única leitura da própria página.
    const derrame = sh.getRange(2, 1, totalLinhas, PT_LARGURA_DERRAME).getValues();
    if (typeof derrame[0][0] === "string" && derrame[0][0].charAt(0) === "#") {
      console.warn("Pagina Transferência: A2 em " + derrame[0][0] +
                   " — dado ainda não chegou. Colunas mantidas como estão.");
      return null;
    }

    const hoje = ptHoje_();

    // Cada origem falha por conta própria. Uma aba atrasada custa as colunas dela
    // ficarem um ciclo velhas, não a página inteira zerada.
    const pulados = [];
    const ix = {};

    const resb = ptIndexResb_(ss, hoje);
    if (resb.ok) ix.resb = resb.mapa; else pulados.push("consumo (" + resb.motivo + ")");

    const estoque = ptIndexEstoque_(ss);
    if (estoque.ok) ix.estoque = estoque; else pulados.push("estoque (" + estoque.motivo + ")");

    const me2n = ptIndexMe2n_(ss);
    if (me2n.ok) ix.me2n = me2n.mapa; else pulados.push("pedidos (" + me2n.motivo + ")");

    const conv = ptIndexConversao_(ss);
    if (conv.ok) ix.conversao = conv.mapa; else pulados.push("conversao (" + conv.motivo + ")");

    const pre = ptIndexPreAgend_();
    if (pre.ok) ix.preAgend = pre.mapa; else pulados.push("preAgend (" + pre.motivo + ")");

    const transp = ptIndexTransporte_();
    if (transp.ok) ix.transporte = transp.mapa; else pulados.push("transporte (" + transp.motivo + ")");

    const gruposProntos = {
      consumo: !!ix.resb, estoque: !!ix.estoque, pedidos: !!ix.me2n,
      conversao: !!ix.conversao, preAgend: !!ix.preAgend, transporte: !!ix.transporte
    };

    // OPT-IN, coluna a coluna. O catalogo em codigo diz o que ESTE arquivo sabe
    // calcular; a coluna "Firmar?" da aba Mapa_Formulas diz o que a operacao
    // AUTORIZOU a trocar por valor. Sem a segunda, uma sincronizacao apagaria 19
    // formulas de uma planilha em producao sem ninguem ter pedido.
    //
    // opts.todas ignora o catalogo — para rodar a mao e conferir o resultado
    // antes de autorizar qualquer coisa.
    let permitidas = null;
    if (!opts.todas) {
      const cat = fmLerCatalogo_(ss);
      if (cat.erro) { console.log("Cálculo não configurado: " + cat.erro); return null; }
      permitidas = {};
      cat.itens.forEach(function (i) { permitidas[i.letra] = true; });
      if (cat.itens.length === 0) {
        console.log("Nenhuma coluna marcada com \"SIM\" em " + MF_ABA_MAPA + ". Nada a calcular.");
        return null;
      }
    }

    const semImplementacao = [];
    if (permitidas) {
      const conhecidas = {};
      PT_COLUNAS_CALCULADAS.forEach(function (c) { conhecidas[c.letra] = true; });
      Object.keys(permitidas).forEach(function (l) {
        if (!conhecidas[l]) semImplementacao.push(l);
      });
    }

    const alvo = PT_COLUNAS_CALCULADAS.filter(function (c) {
      return gruposProntos[c.grupo] && (permitidas === null || permitidas[c.letra]);
    });
    if (alvo.length === 0) {
      console.error("Nenhuma coluna a calcular." +
                    (pulados.length ? " Origem(ns) não pronta(s): " + pulados.join(" || ") : ""));
      return null;
    }
    if (semImplementacao.length > 0) {
      // Marcada para firmar mas sem conta em JS: quem cuida disso e o
      // refirmarPaginaTransferencia (congela o resultado da propria formula).
      console.log("Coluna(s) marcada(s) sem cálculo em JS, deixada(s) para o " +
                  "refirmarPaginaTransferencia(): " + semImplementacao.join(", ") + ".");
    }

    const valores = new Map();
    alvo.forEach(function (c) { valores.set(c.col, new Array(totalLinhas)); });

    for (let i = 0; i < totalLinhas; i++) {
      const out = ptCalcularLinha_(derrame[i], ix, hoje);
      for (let k = 0; k < alvo.length; k++) {
        const v = out[alvo[k].letra];
        valores.get(alvo[k].col)[i] = (v === undefined ? "" : v);
      }
    }

    const runs = ptGravarColunas_(sh, alvo.map(function (c) { return c.col; }), valores, totalLinhas);

    // Uniao, nao substituicao: o caminho generico do Firmar.gs escreve no mesmo
    // registro para as colunas que este arquivo nao calcula. Ver fmMarcarFirmadas_.
    fmMarcarFirmadas_(alvo.map(function (c) { return c.letra; }), totalLinhas);

    console.log("Pagina Transferência calculada: " + alvo.length + " coluna(s) x " +
                totalLinhas + " linha(s) em " + runs + " bloco(s), " +
                Math.round((Date.now() - inicio) / 1000) + "s." +
                (pulados.length ? " Grupo(s) pulado(s): " + pulados.join(" || ") : ""));

    return { colunas: alvo.length, linhas: totalLinhas, pulados: pulados,
             letras: alvo.map(function (c) { return c.letra; }) };
  } catch (e) {
    console.error("Pagina Transferência não foi calculada: " + e.message);
    return null;
  } finally {
    lock.releaseLock();
  }
}