// ====================================================================
// MAPA DE FORMULAS DA "Pagina Transferencia"
// ====================================================================
//
// POR QUE ESTE ARQUIVO EXISTE: a "Pagina Transferencia" e montada por formula
// de ponta a ponta — VLOOKUPs na ME2W, cruzamentos com RESB / ME2N / Stock,
// ARRAYFORMULAs que se estendem sozinhas. Cada uma delas e reavaliada a cada
// recalculo da planilha, e o recalculo dispara sozinho: o sync reescreve as
// bases, TODAY() vira a meia-noite, um IMPORTRANGE se atualiza por conta
// propria. O resultado e a planilha em recalculo permanente — e o
// getTransferData() do portal espera por ela em toda abertura de tela.
//
// O mesmo problema ja foi resolvido do lado do Portal de Pedidos, trocando as
// ~27 colunas de formula da "Resumo P.O" por valor firmado uma vez por
// sincronizacao (firmarColunasCalculadasResumo). Para repetir a receita aqui
// e preciso primeiro saber O QUE cada coluna calcula — e e isso que este
// arquivo entrega.
//
// O QUE ELE PRODUZ:
//   1. a aba "Mapa_Formulas" — uma linha por coluna de formula, com o custo
//      estimado, as funcoes usadas, as abas de origem e a formula em si;
//   2. opcionalmente um .txt no Drive com TODAS as formulas distintas, que e o
//      formato util para revisar fora da planilha;
//   3. a mesma coisa no console (View > Executions), para quando roda por
//      gatilho e ninguem esta olhando a planilha.
//
// A coluna "Firmar?" da aba Mapa_Formulas e o unico campo editavel a mao: e por
// ela que o Firmar.gs decide o que congela e o que continua formula. O mapa
// pode ser regerado quantas vezes for preciso — a escolha feita ali e
// preservada (ver mfFirmarAnterior_).
//
// NAO usa SpreadsheetApp.getUi() nem getActiveSpreadsheet(): este projeto e
// standalone (o Sync e o Backend abrem tudo por openById), e as duas chamadas
// lancam excecao fora de um script vinculado a planilha.

// Mesma aba do SHEET_NAME do STO-Backend.gs. Repetida como literal, e nao
// derivada dele, porque const de outro arquivo nao esta garantidamente
// inicializada na hora em que este e avaliado — mas as duas TEM de andar juntas:
// mudar o nome da aba na planilha exige mudar aqui e la.
const MF_ABA_ALVO = "Pagina Transferência";
const MF_ABA_MAPA = "Mapa_Formulas";

// Cabecalho da aba Mapa_Formulas. A ordem importa: o Firmar.gs le por indice
// fixo a partir daqui (ver MF_C_*).
const MF_MAPA_HEADERS = [
  "Coluna", "Cabeçalho", "Firmar?", "Tipo", "Nº de fórmulas", "Uniforme?",
  "Volátil?", "Peso estimado", "Funções", "Abas referenciadas",
  "Fórmula (exemplo)", "Fórmula R1C1", "Linha âncora"
];

const MF_C_LETRA = 0, MF_C_NOME = 1, MF_C_FIRMAR = 2, MF_C_TIPO = 3;
const MF_C_QTD = 4, MF_C_UNIFORME = 5, MF_C_VOLATIL = 6, MF_C_PESO = 7;
const MF_C_FUNCOES = 8, MF_C_ABAS = 9, MF_C_FORMULA = 10, MF_C_R1C1 = 11;
const MF_C_ANCORA = 12;

// Recalculam sozinhas, sem ninguem mexer em nada. Sao elas que mantem a
// planilha em recalculo permanente mesmo quando o sync nao rodou.
const MF_VOLATEIS = ["TODAY", "NOW", "RAND", "RANDBETWEEN", "RANDARRAY",
                     "INDIRECT", "OFFSET", "IMPORTRANGE", "IMPORTDATA",
                     "IMPORTHTML", "IMPORTXML", "IMPORTFEED", "GOOGLEFINANCE"];

// Custam por LINHA DA BASE varrida, nao por celula: uma coluna de VLOOKUP sobre
// a ME2W e O(linhas da pagina x linhas da ME2W). Sao as que mais pagam para
// virar valor firmado.
const MF_PESADAS = {
  VLOOKUP: 3, HLOOKUP: 3, XLOOKUP: 3, LOOKUP: 3, MATCH: 2, INDEX: 1,
  SUMIF: 3, SUMIFS: 3, COUNTIF: 3, COUNTIFS: 3, AVERAGEIF: 3, AVERAGEIFS: 3,
  MAXIFS: 3, MINIFS: 3, SUMPRODUCT: 4, DSUM: 3, DGET: 3,
  FILTER: 4, QUERY: 5, SORT: 2, SORTN: 2, UNIQUE: 2, VSTACK: 1, HSTACK: 1,
  MAP: 5, LAMBDA: 2, BYROW: 5, BYCOL: 5, SCAN: 5, REDUCE: 5, MAKEARRAY: 5,
  ARRAYFORMULA: 2, TEXTJOIN: 2, JOIN: 1, SPLIT: 1, SUBTOTAL: 2,
  IMPORTRANGE: 6, INDIRECT: 4, OFFSET: 3
};

// Teto de celulas por leitura de getFormulas(). Uma pagina larga e comprida lida
// de uma vez so estoura a memoria do runtime antes de estourar o tempo.
const MF_LOTE_CELULAS = 150000;

function mfLetraColuna_(n) {
  let letra = "";
  while (n > 0) {
    const resto = (n - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    n = Math.floor((n - 1) / 26);
  }
  return letra;
}

function mfTxt_(v) {
  return String(v === null || v === undefined ? "" : v).replace(/[﻿​]/g, "").trim();
}

// Nomes de funcao de dentro da formula. O ( logo depois do nome e o que separa
// a funcao SUM da coluna chamada "SUM" ou do texto "SUM" dentro de uma string.
function mfFuncoes_(formula) {
  const achadas = {};
  const re = /([A-Z][A-Z0-9._]*)\s*\(/g;
  let m;
  while ((m = re.exec(String(formula).toUpperCase())) !== null) achadas[m[1]] = true;
  return Object.keys(achadas).sort();
}

// Abas citadas pela formula, nas duas notacoes que o Sheets aceita:
// 'Nome Com Espaco'!A1 e NomeSimples!A1. O IMPORTRANGE entra pelo id da
// planilha, que e outra coisa — vai marcado como externo.
function mfAbasReferenciadas_(formula) {
  const s = String(formula);
  const abas = {};

  const reQuot = /'((?:[^']|'')+)'\s*!/g;
  let m;
  while ((m = reQuot.exec(s)) !== null) abas[m[1].replace(/''/g, "'")] = true;

  const reSimples = /(^|[^A-Za-z0-9_.'!])([A-Za-z_][A-Za-z0-9_.]*)\s*!/g;
  while ((m = reSimples.exec(s)) !== null) abas[m[2]] = true;

  if (/IMPORTRANGE\s*\(/i.test(s)) abas["(IMPORTRANGE — planilha externa)"] = true;
  return Object.keys(abas).sort();
}

function mfEhVolatil_(funcoes) {
  for (let i = 0; i < funcoes.length; i++) {
    if (MF_VOLATEIS.indexOf(funcoes[i]) !== -1) return true;
  }
  return false;
}

// Peso = quanto essa coluna custa por recalculo, em unidades arbitrarias mas
// comparaveis entre colunas. Serve para ordenar o mapa: o topo da lista e o que
// mais paga para virar valor firmado.
//
//   celulas de formula x soma do custo das funcoes pesadas x (2 se for volatil)
//
// Volatil dobra porque essas recalculam sem ninguem pedir — o custo delas e
// pago varias vezes por dia, nao so quando o sync roda.
function mfPeso_(qtdFormulas, funcoes, volatil) {
  let custo = 0;
  funcoes.forEach(function (f) { custo += (MF_PESADAS[f] || 0); });
  if (custo === 0) custo = 1;
  return Math.round(Math.max(qtdFormulas, 1) * custo * (volatil ? 2 : 1));
}

// Preserva a decisao ja tomada na coluna "Firmar?". Regerar o mapa depois de uma
// mudanca na planilha nao pode zerar a escolha de quem revisou coluna a coluna.
function mfFirmarAnterior_(ss) {
  const anterior = new Map();
  const aba = ss.getSheetByName(MF_ABA_MAPA);
  if (!aba || aba.getLastRow() < 2) return anterior;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, MF_MAPA_HEADERS.length).getValues();
  dados.forEach(function (linha) {
    const letra = mfTxt_(linha[MF_C_LETRA]);
    if (letra) anterior.set(letra, mfTxt_(linha[MF_C_FIRMAR]));
  });
  return anterior;
}

// ====================================================================
// VARREDURA
// ====================================================================

// Percorre a aba coluna a coluna e devolve um resumo por coluna de formula.
// Le em blocos de linhas: guardar o grid inteiro de formulas em memoria e o que
// derruba a execucao numa pagina grande, e nada aqui precisa dele inteiro.
function mfVarrer_(aba) {
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 2 || ultimaColuna < 1) return { colunas: [], linhas: 0 };

  const cabecalho = aba.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(mfTxt_);
  const totalDados = ultimaLinha - 1;

  // Uma entrada por coluna. `distintas` guarda no maximo 5 padroes R1C1 por
  // coluna: mais que isso ja diz o que precisava dizer (a coluna nao e uniforme)
  // e so ocuparia memoria.
  const acc = [];
  for (let c = 0; c < ultimaColuna; c++) {
    acc.push({ col: c + 1, nome: cabecalho[c] || "", qtd: 0, primeiraLinha: 0,
               primeiraA1: "", primeiraR1C1: "", valorAbaixo: false, distintas: new Map() });
  }

  const passo = Math.max(1, Math.floor(MF_LOTE_CELULAS / ultimaColuna));

  for (let ini = 2; ini <= ultimaLinha; ini += passo) {
    const n = Math.min(passo, ultimaLinha - ini + 1);
    const faixa = aba.getRange(ini, 1, n, ultimaColuna);
    const a1 = faixa.getFormulas();
    const r1c1 = faixa.getFormulasR1C1();
    // Os VALORES entram junto para separar dois casos que, olhando so a formula,
    // sao identicos: uma ARRAYFORMULA que derrama e uma formula que ficou parada
    // numa celula so, sem nunca ter sido arrastada. Nos dois a coluna tem
    // exatamente uma formula; a diferenca esta em haver ou nao valor abaixo dela.
    const vals = faixa.getValues();

    for (let i = 0; i < n; i++) {
      for (let c = 0; c < ultimaColuna; c++) {
        const e = acc[c];
        const f = a1[i][c];
        if (!f) {
          if (e.qtd > 0 && String(vals[i][c]) !== "") e.valorAbaixo = true;
          continue;
        }
        e.qtd++;
        if (e.primeiraLinha === 0) {
          e.primeiraLinha = ini + i;
          e.primeiraA1 = f;
          e.primeiraR1C1 = r1c1[i][c];
        }
        const chave = r1c1[i][c];
        if (e.distintas.has(chave)) e.distintas.set(chave, e.distintas.get(chave) + 1);
        else if (e.distintas.size < 5) e.distintas.set(chave, 1);
        else e.distintas.set("(outros padrões)", (e.distintas.get("(outros padrões)") || 0) + 1);
      }
    }
  }

  const colunas = [];
  acc.forEach(function (e) {
    if (e.qtd === 0) return;

    const funcoes = mfFuncoes_(e.primeiraA1);
    const volatil = mfEhVolatil_(funcoes);

    // Tres tipos, e a diferenca entre os dois primeiros so aparece nos VALORES:
    //
    //   array       uma ancora que DERRAMA — as celulas de baixo tem valor e nao
    //               tem formula (ARRAYFORMULA, MAP, FILTER...).
    //   solta       uma formula parada numa celula so, sem nada abaixo. Quase
    //               sempre e defeito: a coluna responde pela linha 2 e mais nada,
    //               e quem le a aba inteira ve a coluna vazia. Vale checar toda
    //               vez que aparecer.
    //   preenchida  o mesmo padrao repetido linha a linha (arrastado).
    //
    // O tipo decide como a formula volta no restaurarFormulas: a ancora sozinha
    // (array e solta) ou o padrao repetido em todas as linhas (preenchida).
    const unica = (e.qtd === 1 && totalDados > 1);
    const derrama = unica && e.valorAbaixo;
    const solta = unica && !e.valorAbaixo;
    const uniforme = (e.distintas.size === 1);

    colunas.push({
      col: e.col,
      letra: mfLetraColuna_(e.col),
      nome: e.nome,
      qtd: e.qtd,
      tipo: derrama ? "array" : (solta ? "solta" : "preenchida"),
      uniforme: uniforme,
      volatil: volatil,
      funcoes: funcoes,
      abas: mfAbasReferenciadas_(e.primeiraA1),
      formula: e.primeiraA1,
      r1c1: e.primeiraR1C1,
      ancora: e.primeiraLinha,
      // Uma coluna "solta" custa UMA celula, nao a coluna toda: ela nao esta
      // fazendo o trabalho que o cabecalho promete.
      peso: mfPeso_(derrama ? totalDados : e.qtd, funcoes, volatil),
      distintas: e.distintas
    });
  });

  colunas.sort(function (a, b) { return b.peso - a.peso; });
  return { colunas: colunas, linhas: totalDados };
}

// ====================================================================
// SAIDA
// ====================================================================

function mfEscreverMapa_(ss, resultado, firmarAnterior) {
  let aba = ss.getSheetByName(MF_ABA_MAPA);
  if (!aba) {
    aba = ss.insertSheet(MF_ABA_MAPA);
  } else {
    aba.clear();
  }

  const linhas = resultado.colunas.map(function (c) {
    const escolha = firmarAnterior.get(c.letra);
    const registro = new Array(MF_MAPA_HEADERS.length).fill("");
    registro[MF_C_LETRA] = c.letra;
    registro[MF_C_NOME] = c.nome;
    // Default NAO, sempre. Firmar troca formula por valor estatico numa planilha
    // em producao: e uma decisao de quem conhece a coluna, nunca do script que
    // acabou de descobrir que ela existe.
    registro[MF_C_FIRMAR] = escolha || "NÃO";
    registro[MF_C_TIPO] = c.tipo;
    registro[MF_C_QTD] = c.qtd;
    registro[MF_C_UNIFORME] = c.uniforme ? "sim" : "não";
    registro[MF_C_VOLATIL] = c.volatil ? "sim" : "não";
    registro[MF_C_PESO] = c.peso;
    registro[MF_C_FUNCOES] = c.funcoes.join(", ");
    registro[MF_C_ABAS] = c.abas.join(", ");
    // Apostrofo na frente: sem ele a planilha tenta AVALIAR a formula arquivada,
    // e o mapa vira uma copia da pagina em vez de um retrato dela.
    registro[MF_C_FORMULA] = "'" + c.formula;
    registro[MF_C_R1C1] = "'" + c.r1c1;
    registro[MF_C_ANCORA] = c.ancora;
    return registro;
  });

  aba.getRange(1, 1, 1, MF_MAPA_HEADERS.length).setValues([MF_MAPA_HEADERS]);
  aba.getRange(1, 1, 1, MF_MAPA_HEADERS.length).setFontWeight("bold");
  aba.setFrozenRows(1);
  if (linhas.length > 0) {
    aba.getRange(2, 1, linhas.length, MF_MAPA_HEADERS.length).setValues(linhas);
  }
  aba.getRange(1, MF_C_FIRMAR + 1, Math.max(linhas.length + 1, 2), 1).setBackground("#fff2cc");
  aba.autoResizeColumns(1, MF_C_FUNCOES + 1);
  return aba;
}

// Texto corrido com TODAS as formulas distintas, coluna a coluna. E o formato
// que serve para revisar fora da planilha (ou colar num chamado): o mapa mostra
// uma formula de exemplo por coluna, este mostra todos os padroes.
function mfMontarTexto_(resultado) {
  const linhas = [];
  linhas.push("MAPA DE FÓRMULAS — aba \"" + MF_ABA_ALVO + "\"");
  linhas.push("Gerado em " + Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm"));
  linhas.push(resultado.colunas.length + " coluna(s) de fórmula sobre " + resultado.linhas + " linha(s) de dados.");
  linhas.push("Ordenado por peso estimado (o topo é o que mais custa por recálculo).");
  linhas.push("");

  resultado.colunas.forEach(function (c) {
    linhas.push("=".repeat(70));
    linhas.push("Coluna " + c.letra + " — " + (c.nome || "(sem cabeçalho)"));
    const explica = { array: " (uma âncora que derrama)",
                      solta: " (⚠️ UMA célula só — a coluna responde pela linha " + c.ancora + " e mais nada)",
                      preenchida: " (padrão repetido linha a linha)" };
    linhas.push("  tipo ............. " + c.tipo + (explica[c.tipo] || ""));
    linhas.push("  células fórmula .. " + c.qtd);
    linhas.push("  uniforme ......... " + (c.uniforme ? "sim" : "NÃO — há mais de um padrão nesta coluna"));
    linhas.push("  volátil .......... " + (c.volatil ? "SIM — recalcula sozinha" : "não"));
    linhas.push("  peso estimado .... " + c.peso);
    linhas.push("  funções .......... " + (c.funcoes.join(", ") || "—"));
    linhas.push("  abas de origem ... " + (c.abas.join(", ") || "—"));
    linhas.push("  âncora ........... linha " + c.ancora);
    linhas.push("");
    c.distintas.forEach(function (vezes, padrao) {
      linhas.push("  [" + vezes + "x] " + padrao);
    });
    linhas.push("");
  });

  return linhas.join("\n");
}

// ====================================================================
// PONTOS DE ENTRADA
// ====================================================================

// Roda pelo editor. Sem argumento faz o mapa da "Pagina Transferencia" e grava
// tambem o .txt no Drive.
//
//   opcoes.aba      — outra aba (o mesmo mapa serve para qualquer uma)
//   opcoes.arquivo  — false para nao gerar o .txt no Drive
//   opcoes.ss       — planilha ja aberta, quando chamado de dentro do sync
function mapearFormulasPaginaTransferencia(opcoes) {
  const opts = opcoes || {};
  const nomeAba = opts.aba || MF_ABA_ALVO;
  const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(nomeAba);

  if (!aba) {
    throw new Error("Aba '" + nomeAba + "' não encontrada na planilha " + ss.getId() + ".");
  }

  const firmarAnterior = mfFirmarAnterior_(ss);
  const resultado = mfVarrer_(aba);

  if (resultado.colunas.length === 0) {
    console.log("Nenhuma fórmula encontrada em '" + nomeAba + "'. Nada a mapear.");
    return resultado;
  }

  mfEscreverMapa_(ss, resultado, firmarAnterior);

  const texto = mfMontarTexto_(resultado);
  console.log(texto);

  let url = null;
  if (opts.arquivo !== false) {
    try {
      const nome = "Formulas_" + nomeAba.replace(/[^\wÀ-ÿ]+/g, "_") + "_" +
                   Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd_HHmm") + ".txt";
      url = DriveApp.createFile(nome, texto, MimeType.PLAIN_TEXT).getUrl();
      console.log("Arquivo com todas as fórmulas: " + url);
    } catch (e) {
      // Cosmetico: o mapa na aba e o log ja entregaram o conteudo.
      console.warn("Não foi possível gravar o .txt no Drive: " + e.message);
    }
  }

  const soltas = resultado.colunas.filter(function (c) { return c.tipo === "solta"; });
  if (soltas.length > 0) {
    console.warn("⚠️ Coluna(s) com fórmula em UMA célula só, sem nada abaixo — respondem " +
                 "pela linha âncora e ficam vazias no resto da aba: " +
                 soltas.map(function (c) { return c.letra + " (" + c.nome + ")"; }).join(", ") +
                 ". Quem lê a aba inteira vê essas colunas em branco.");
  }

  const volateis = resultado.colunas.filter(function (c) { return c.volatil; });
  console.log("Resumo: " + resultado.colunas.length + " coluna(s) de fórmula, " +
              volateis.length + " volátil(eis) (" +
              (volateis.map(function (c) { return c.letra; }).join(", ") || "nenhuma") +
              "). Revise a coluna \"Firmar?\" da aba " + MF_ABA_MAPA +
              " e rode firmarPaginaTransferencia() para congelar as escolhidas.");

  return { colunas: resultado.colunas, linhas: resultado.linhas, arquivo: url };
}

// Atalho para quando o interesse e so a estrutura, sem gerar arquivo nenhum:
// devolve uma formula de exemplo por coluna, como o antigo listador do Portal
// de Pedidos fazia — mas lendo a coluna inteira, nao so a linha 2, e sem
// depender de UI (que nao existe em projeto standalone nem sob gatilho).
function listarFormulasPaginaTransferencia() {
  return mapearFormulasPaginaTransferencia({ arquivo: false });
}
