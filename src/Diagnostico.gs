// ====================================================================
// DIAGNOSTICO DO PORTAL — rode diagnosticarPortal() pelo editor
// ====================================================================
//
// PARA QUE SERVE: o portal abrir e ficar no "Carregando dados..." nao diz nada
// sobre a causa. Este arquivo percorre, na ORDEM EM QUE O PORTAL DEPENDE DELAS,
// cada peca que precisa estar de pe — projeto, planilha, abas, colunas, a
// leitura pesada, o payload — e diz onde para.
//
// DUAS REGRAS QUE ESTE ARQUIVO SEGUE:
//
//   1. SO LE. Nenhuma etapa escreve em celula, property ou arquivo. Um
//      diagnostico que altera o estado destroi a evidencia que veio buscar, e
//      no pior caso piora a avaria que estava investigando.
//
//   2. IMPRIME NA HORA. Cada etapa vai para o console assim que termina, em vez
//      de acumular um relatorio para o fim. Se o proprio diagnostico for morto
//      pelo limite de 6 min — que e um dos sintomas possiveis — o log de
//      Execucoes ja mostra tudo ate a etapa que travou, e a ULTIMA LINHA
//      IMPRESSA E A RESPOSTA.
//
// COMO LER O RESULTADO: procure as linhas [FALHA] e [AVISO]. No fim sai um
// resumo com o veredito. Cada etapa mostra quanto tempo levou — num sintoma de
// lentidao, e o tempo que aponta o culpado, nao a mensagem de erro.

const DG_LIMITE_MS = 5.5 * 60 * 1000;   // corta antes dos 6 min do runtime
const DG_LENTO_MS = 10000;              // acima disso, uma etapa merece aviso

// ---------------------------------------------------------------- relatorio ---

function dgNovoRelatorio_() {
  return { linhas: [], falhas: [], avisos: [], inicio: Date.now() };
}

function dgDizer_(rel, nivel, texto) {
  const linha = (nivel ? "[" + nivel + "] " : "        ") + texto;
  rel.linhas.push(linha);
  console.log(linha);                       // imprime JA — ver regra 2 no topo
  if (nivel === "FALHA") rel.falhas.push(texto);
  if (nivel === "AVISO") rel.avisos.push(texto);
}

function dgTitulo_(rel, texto) {
  rel.linhas.push("");
  console.log("");
  dgDizer_(rel, "", "=== " + texto + " ===");
}

// Roda uma etapa cronometrada que NUNCA derruba o diagnostico. A excecao vira
// resultado: e ela que estamos procurando, entao precisa aparecer no relatorio
// com mensagem e stack, nao interromper a varredura no meio.
//
// `nivelSeFalhar` existe porque nem toda etapa que quebra impede o portal de
// abrir. Marcar como FALHA uma peca da qual a tela nao depende — as planilhas
// externas, por exemplo — manda quem le investigar o lugar errado, que e o
// oposto do que um diagnostico serve para fazer.
function dgEtapa_(rel, nome, fn, nivelSeFalhar) {
  const t0 = Date.now();
  try {
    const valor = fn();
    const ms = Date.now() - t0;
    if (ms >= DG_LENTO_MS) {
      dgDizer_(rel, "AVISO", nome + " levou " + (ms / 1000).toFixed(1) + "s — LENTO.");
    }
    return { ok: true, valor: valor, ms: ms };
  } catch (e) {
    const ms = Date.now() - t0;
    dgDizer_(rel, nivelSeFalhar || "FALHA",
             nome + " lancou apos " + (ms / 1000).toFixed(1) + "s: " + e.message);
    if (e.stack) {
      String(e.stack).split("\n").slice(0, 4).forEach(function (l) {
        dgDizer_(rel, "", "    " + l.trim());
      });
    }
    return { ok: false, erro: e, ms: ms };
  }
}

function dgSemTempo_(rel) {
  if (Date.now() - rel.inicio < DG_LIMITE_MS) return false;
  dgDizer_(rel, "AVISO", "Diagnostico interrompido: chegou perto do limite de 6 min do gatilho.");
  return true;
}

function dgTamanho_(v) {
  try { return (JSON.stringify(v).length / 1024).toFixed(1) + " KB"; }
  catch (e) { return "(nao serializavel: " + e.message + ")"; }
}

// ---------------------------------------------------------------- principal ---

function diagnosticarPortal() {
  const rel = dgNovoRelatorio_();
  dgDizer_(rel, "", "DIAGNOSTICO DO PORTAL DE TRANSFERENCIA");
  dgDizer_(rel, "", new Date().toString());

  dgProjeto_(rel);
  const ss = dgPlanilha_(rel);
  if (ss) {
    dgAbas_(rel, ss);
    const sh = dgPaginaTransferencia_(rel, ss);
    dgTokens_(rel, ss);
    dgFirmacao_(rel, ss);
    if (sh && !dgSemTempo_(rel)) dgLeituraPesada_(rel, sh);
  }
  if (!dgSemTempo_(rel)) dgPlanilhasExternas_(rel);
  if (!dgSemTempo_(rel)) dgCargaDoPortal_(rel);

  return dgResumo_(rel);
}

// 1. O projeto carregou? Se um .gs nao compila, TODA chamada ao servidor falha
//    — inclusive a que o portal faz para abrir.
function dgProjeto_(rel) {
  dgTitulo_(rel, "1. ARQUIVOS DO PROJETO");

  // Os typeof sao escritos com o nome LITERAL, um a um, de proposito: com nome
  // dinamico so daria por eval, e typeof sobre identificador inexistente ja
  // devolve "undefined" sem lancar — que e exatamente o que precisamos aqui.
  [["STO-Backend.gs", "getTransferData", typeof getTransferData === "function"],
   ["Sync.gs", "sincronizarNovasBases", typeof sincronizarNovasBases === "function"],
   ["Calculo.gs", "firmarColunasCalculadasTransferencia", typeof firmarColunasCalculadasTransferencia === "function"],
   ["Firmar.gs", "refirmarPaginaTransferencia", typeof refirmarPaginaTransferencia === "function"],
   ["Formulas.gs", "mapearFormulasPaginaTransferencia", typeof mapearFormulasPaginaTransferencia === "function"]
  ].forEach(function (t) {
    dgDizer_(rel, t[2] ? "" : "AVISO",
             (t[2] ? "ok   " : "ausente ") + t[0] + " (" + t[1] + ")");
  });

  // Constantes que os arquivos compartilham. Uma indefinida quase sempre quer
  // dizer arquivo faltando ou renomeado no editor.
  [["SPREADSHEET_ID", "STO-Backend.gs", typeof SPREADSHEET_ID !== "undefined"],
   ["SHEET_NAME", "STO-Backend.gs", typeof SHEET_NAME !== "undefined"],
   ["PLANILHA_ALVO_ID", "Sync.gs", typeof PLANILHA_ALVO_ID !== "undefined"],
   ["TIMEZONE", "Sync.gs", typeof TIMEZONE !== "undefined"],
   ["MF_ABA_ALVO", "Formulas.gs", typeof MF_ABA_ALVO !== "undefined"]
  ].forEach(function (t) {
    dgDizer_(rel, t[2] ? "" : "FALHA",
             (t[2] ? "ok   " : "INDEFINIDA ") + t[0] + " (vem de " + t[1] + ")");
  });

  // As duas pontas do SPREADSHEET_ID tem de ser a MESMA planilha: o portal le
  // por uma constante e o sync escreve pela outra.
  if (typeof SPREADSHEET_ID !== "undefined" && typeof PLANILHA_ALVO_ID !== "undefined") {
    const iguais = SPREADSHEET_ID === PLANILHA_ALVO_ID;
    dgDizer_(rel, iguais ? "" : "FALHA",
             iguais ? "ok   SPREADSHEET_ID e PLANILHA_ALVO_ID apontam para a mesma planilha."
                    : "SPREADSHEET_ID e PLANILHA_ALVO_ID sao DIFERENTES — o portal le uma planilha e o sync escreve noutra.");
  }

  const svc = dgEtapa_(rel, "Servico avancado Drive", function () {
    return typeof Drive !== "undefined" && Drive.Files ? "ligado" : "DESLIGADO";
  });
  dgDizer_(rel, (svc.ok && svc.valor === "ligado") ? "" : "AVISO",
           "Servico avancado Drive: " + (svc.ok ? svc.valor : "erro") +
           " (so o sync usa; o portal abre sem ele)");
}

// 2. A planilha responde?
function dgPlanilha_(rel) {
  dgTitulo_(rel, "2. PLANILHA");
  if (typeof SPREADSHEET_ID === "undefined") {
    dgDizer_(rel, "FALHA", "Sem SPREADSHEET_ID — nada a abrir.");
    return null;
  }
  dgDizer_(rel, "", "id: " + SPREADSHEET_ID);

  const r = dgEtapa_(rel, "openById", function () {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  });
  if (!r.ok) {
    dgDizer_(rel, "FALHA", "A planilha nao abre. Sem isso o portal nao tem o que mostrar. " +
             "Confira se o id existe e se a conta que PUBLICOU o app ainda tem acesso a ela.");
    return null;
  }
  dgDizer_(rel, "", "ok   abriu em " + r.ms + "ms: \"" + r.valor.getName() + "\"");
  dgDizer_(rel, "", "     fuso da planilha: " + r.valor.getSpreadsheetTimeZone());
  return r.valor;
}

// 3. As abas que o portal e o sync esperam.
function dgAbas_(rel, ss) {
  dgTitulo_(rel, "3. ABAS");
  const nomes = ss.getSheets().map(function (s) { return s.getName(); });
  dgDizer_(rel, "", "existem " + nomes.length + ": " + nomes.join(" | "));

  const precisa = [
    [typeof SHEET_NAME !== "undefined" ? SHEET_NAME : "Pagina Transferência", "FALHA", "o portal le desta"],
    ["ME2W", "FALHA", "unica base com dado insubstituivel"],
    ["RESB", "AVISO", "consumo"],
    ["ME2N", "AVISO", "pedidos em aberto"],
    ["ME5A", "AVISO", "requisicoes"],
    ["Stock Control BR14 BR10 BR12", "AVISO", "estoque"],
    ["Confirmacoes_Store", "AVISO", "fonte da verdade das confirmacoes"],
    ["Tokens_Link", "FALHA", "sem ela o doGet nega todo mundo"],
    ["Mapa_Formulas", "AVISO", "configuracao do que e firmado"]
  ];
  precisa.forEach(function (p) {
    const tem = nomes.indexOf(p[0]) !== -1;
    dgDizer_(rel, tem ? "" : p[1],
             (tem ? "ok   " : "FALTA ") + "\"" + p[0] + "\" — " + p[2]);
  });
}

// 4. O estado da pagina que o portal le. E aqui que moram as avarias que a
//    gente ja viu: derrame vazio e celula em erro.
function dgPaginaTransferencia_(rel, ss) {
  const nome = typeof SHEET_NAME !== "undefined" ? SHEET_NAME : "Pagina Transferência";
  dgTitulo_(rel, "4. PAGINA \"" + nome + "\"");

  const sh = ss.getSheetByName(nome);
  if (!sh) { dgDizer_(rel, "FALHA", "Aba nao existe. O portal nao tem de onde ler."); return null; }

  const ultLinha = sh.getLastRow(), ultCol = sh.getLastColumn();
  const maxLinha = sh.getMaxRows(), maxCol = sh.getMaxColumns();
  dgDizer_(rel, "", "conteudo ate linha " + ultLinha + ", coluna " + ultCol);
  dgDizer_(rel, "", "grade: " + maxLinha + " x " + maxCol +
           "  (celulas lidas por getDataRange: " + (ultLinha * ultCol).toLocaleString() + ")");

  if (ultLinha < 2) {
    dgDizer_(rel, "FALHA", "A pagina nao tem linha de dado nenhuma — o portal abriria vazio.");
    return sh;
  }
  // getDataRange le um retangulo; e o produto que pesa, nao o numero de linhas.
  if (ultLinha * ultCol > 2000000) {
    dgDizer_(rel, "AVISO", "Retangulo muito grande — o getValues()+getDisplayValues() do " +
             "getTransferData le isso DUAS vezes e pode estourar o tempo da chamada.");
  }

  // Cabecalho: e por nome que o STO-Backend acha cada coluna.
  const headers = sh.getRange(1, 1, 1, ultCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const chave = ["Documento", "Item", "Schedule Line", "Material", "Planta_ Destino",
                 "Order Quantity", "Deliv Date", "Origem"];
  const faltando = chave.filter(function (c) { return headers.indexOf(c) === -1; });
  dgDizer_(rel, faltando.length ? "FALHA" : "",
           faltando.length ? "Cabecalho SEM as colunas: " + faltando.join(", ") +
                             " — o portal descarta toda linha e abre vazio."
                           : "ok   cabecalho tem as 8 colunas-chave.");

  // Amostra: primeiras 40 linhas bastam para flagrar derrame vazio e erro.
  const n = Math.min(40, ultLinha - 1);
  const amostra = sh.getRange(2, 1, n, ultCol).getValues();

  const iDoc = headers.indexOf("Documento");
  let semDoc = 0;
  if (iDoc !== -1) {
    for (let i = 0; i < n; i++) {
      const v = amostra[i][iDoc];
      if (v === "" || v === null) semDoc++;
    }
    if (semDoc === n) {
      dgDizer_(rel, "FALHA", "As " + n + " primeiras linhas estao SEM Documento. " +
               "E o sintoma do derrame da A2 morto (A:X vazias) — o portal descarta " +
               "toda linha e abre vazio. Rode restaurarDerramePaginaTransferencia().");
    } else if (semDoc > 0) {
      dgDizer_(rel, "AVISO", semDoc + " de " + n + " linhas sem Documento (serao descartadas).");
    } else {
      dgDizer_(rel, "", "ok   derrame A:X de pe (todas as " + n + " linhas tem Documento).");
    }
  }

  // Celulas em erro. Nao derrubam o portal, mas explicam campo vazio na tela.
  const erros = {};
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < amostra[i].length; j++) {
      const v = amostra[i][j];
      if (typeof v === "string" && (v.charAt(0) === "#" || v === "Loading...")) {
        const k = (headers[j] || ("col " + (j + 1))) + " = " + v;
        erros[k] = (erros[k] || 0) + 1;
      }
    }
  }
  const listaErros = Object.keys(erros);
  if (listaErros.length === 0) {
    dgDizer_(rel, "", "ok   nenhuma celula em #REF!/#N/A/Loading nas " + n + " primeiras linhas.");
  } else {
    listaErros.forEach(function (k) {
      dgDizer_(rel, "AVISO", "celula em erro: " + k + " (x" + erros[k] + ")");
    });
    dgDizer_(rel, "", "     Loading... = IMPORTRANGE ainda resolvendo. Se nao sair disso, " +
             "abra a planilha de origem e autorize o acesso uma vez, na mao.");
  }
  return sh;
}

// 5. Tokens — o doGet nega o acesso sem eles.
function dgTokens_(rel, ss) {
  dgTitulo_(rel, "5. TOKENS DE ACESSO");
  const nome = typeof SHEET_TOKENS !== "undefined" ? SHEET_TOKENS : "Tokens_Link";
  const sh = ss.getSheetByName(nome);
  if (!sh) {
    dgDizer_(rel, "FALHA", "Aba \"" + nome + "\" nao existe — o doGet responde " +
             "\"Acesso negado\" para qualquer link. Rode getOrCreateToken(\"nome\").");
    return;
  }
  const n = Math.max(sh.getLastRow() - 1, 0);
  dgDizer_(rel, n > 0 ? "" : "FALHA",
           n > 0 ? "ok   " + n + " token(s) cadastrado(s)."
                 : "Nenhum token cadastrado — nenhum link abre.");
}

// 6. O que esta firmado e o que esta marcado para firmar.
function dgFirmacao_(rel, ss) {
  dgTitulo_(rel, "6. FIRMACAO");

  const props = PropertiesService.getScriptProperties();
  const bruto = typeof FM_PROP_FIRMADA !== "undefined" ? props.getProperty(FM_PROP_FIRMADA) : null;
  if (!bruto) {
    dgDizer_(rel, "", "nenhum registro de firmacao — a pagina esta inteira em formula.");
  } else {
    try {
      const est = JSON.parse(bruto);
      dgDizer_(rel, "", "firmado em " + est.em + ": " + (est.colunas || []).join(", ") +
               " x " + est.linhas + " linha(s).");
    } catch (e) {
      dgDizer_(rel, "AVISO", "registro de firmacao ilegivel: " + bruto);
    }
  }

  const mapa = ss.getSheetByName(typeof MF_ABA_MAPA !== "undefined" ? MF_ABA_MAPA : "Mapa_Formulas");
  if (!mapa || mapa.getLastRow() < 2) {
    dgDizer_(rel, "", "sem Mapa_Formulas — nada e firmado, a pagina fica em formula.");
    return;
  }
  const dados = mapa.getRange(2, 1, mapa.getLastRow() - 1, Math.min(mapa.getLastColumn(), 13)).getValues();
  const marcadas = [];
  dados.forEach(function (l) {
    const m = String(l[2] || "").trim().toUpperCase();
    if (m === "SIM" || m === "S") marcadas.push(String(l[0] || "").trim());
  });
  dgDizer_(rel, "", "marcadas para firmar: " + (marcadas.join(", ") || "nenhuma"));

  // O erro que ja derrubou a pagina uma vez.
  const teto = typeof FM_ULTIMA_COLUNA_DERRAME !== "undefined" ? FM_ULTIMA_COLUNA_DERRAME : 24;
  const proibidas = marcadas.filter(function (letra) {
    const col = typeof fmColunaDeLetra_ === "function" ? fmColunaDeLetra_(letra) : 0;
    return col > 0 && col <= teto;
  });
  if (proibidas.length > 0) {
    dgDizer_(rel, "FALHA", "Coluna(s) " + proibidas.join(", ") + " marcadas SIM estao DENTRO " +
             "do derrame A:X. Firmar ali esvazia B:X e derruba a pagina. Desmarque no Mapa_Formulas.");
  }
}

// 7. As DUAS leituras pesadas do getTransferData, cronometradas em separado.
//    Se a chamada do portal demora, e uma destas — e saber qual muda o que fazer.
function dgLeituraPesada_(rel, sh) {
  dgTitulo_(rel, "7. AS LEITURAS QUE O PORTAL FAZ");
  dgDizer_(rel, "", "o getTransferData chama getValues() E getDisplayValues() sobre o mesmo retangulo.");

  const a = dgEtapa_(rel, "getDataRange().getValues()", function () {
    return sh.getDataRange().getValues().length;
  });
  if (a.ok) dgDizer_(rel, "", "ok   getValues(): " + a.valor + " linhas em " + a.ms + "ms");

  const b = dgEtapa_(rel, "getDataRange().getDisplayValues()", function () {
    return sh.getDataRange().getDisplayValues().length;
  });
  if (b.ok) dgDizer_(rel, "", "ok   getDisplayValues(): " + b.valor + " linhas em " + b.ms + "ms");

  const total = (a.ms || 0) + (b.ms || 0);
  if (total > 30000) {
    dgDizer_(rel, "FALHA", "As duas leituras somam " + (total / 1000).toFixed(1) + "s. " +
             "E tempo suficiente para a chamada do portal desistir — este e o motivo do " +
             "\"Carregando dados...\" que nao sai.");
  }
}

// 8. As planilhas de fora. So o Calculo.gs usa; o portal abre sem elas.
function dgPlanilhasExternas_(rel) {
  dgTitulo_(rel, "8. PLANILHAS EXTERNAS (usadas pelo Calculo.gs)");
  [["Pré-Agendamento", typeof PT_PEDIDOS_DB_ID !== "undefined" ? PT_PEDIDOS_DB_ID : null],
   ["Plano_Transporte", typeof PT_TRANSPORTES_DB_ID !== "undefined" ? PT_TRANSPORTES_DB_ID : null]
  ].forEach(function (par) {
    if (!par[1]) { dgDizer_(rel, "AVISO", par[0] + ": id nao definido (Calculo.gs ausente?)"); return; }
    // AVISO, nao FALHA: o portal nao le estas planilhas para abrir — quem le e o
    // Calculo.gs, na sincronizacao.
    const r = dgEtapa_(rel, par[0], function () {
      return SpreadsheetApp.openById(par[1]).getName();
    }, "AVISO");
    if (r.ok) dgDizer_(rel, "", "ok   " + par[0] + " -> \"" + r.valor + "\" (" + r.ms + "ms)");
    else dgDizer_(rel, "", "     ^ as colunas AO/AP/AQ ficam vazias, mas o portal ABRE do mesmo jeito.");
  });
}

// 9. A prova final: exatamente o que o portal chama ao abrir.
function dgCargaDoPortal_(rel) {
  dgTitulo_(rel, "9. A CHAMADA QUE O PORTAL FAZ AO ABRIR");
  if (typeof getTransferData !== "function") {
    dgDizer_(rel, "FALHA", "getTransferData nao existe — STO-Backend.gs ausente do projeto.");
    return;
  }

  const r = dgEtapa_(rel, "getTransferData()", function () { return getTransferData(); });
  if (!r.ok) {
    dgDizer_(rel, "FALHA", "ESTA E A CAUSA: o portal chama getTransferData() ao abrir, " +
             "e ela lanca. A mensagem e o stack estao logo acima.");
    return;
  }

  const linhas = r.valor.length;
  dgDizer_(rel, "", "ok   devolveu " + linhas + " linha(s) em " + (r.ms / 1000).toFixed(1) + "s");

  if (linhas === 0) {
    dgDizer_(rel, "FALHA", "Devolveu ZERO linhas: o portal abre, mas com a tabela vazia. " +
             "Quase sempre e o derrame A:X morto (etapa 4) ou o cabecalho fora do lugar.");
    return;
  }

  const tam = dgTamanho_(r.valor);
  dgDizer_(rel, "", "payload para o navegador: " + tam);
  const kb = parseFloat(tam);
  if (isFinite(kb) && kb > 8000) {
    dgDizer_(rel, "AVISO", "Payload acima de 8 MB — a ponte google.script.run costuma falhar nesse tamanho.");
  }

  if (r.ms > 30000) {
    dgDizer_(rel, "FALHA", "Levou " + (r.ms / 1000).toFixed(1) + "s. O navegador desiste antes " +
             "disso, e o portal fica no \"Carregando dados...\" para sempre.");
  } else if (r.ms > 15000) {
    dgDizer_(rel, "AVISO", "Levou " + (r.ms / 1000).toFixed(1) + "s — perto do limite do aceitavel.");
  }

  // Uma linha inteira, para conferir se os campos chegam preenchidos.
  const o = r.valor[0];
  dgDizer_(rel, "", "primeira linha: doc=" + o.doc + " item=" + o.item + " planta=" + o.planta +
           " material=" + o.matCod + " qtd=" + o.qtySap + " data=" + o.dataSap);
  dgDizer_(rel, "", "                alerta=\"" + o.alerta + "\" estoque=\"" + o.statusEstqUnificado + "\"");

  const vazios = ["doc", "planta", "matCod"].filter(function (k) {
    return o[k] === "" || o[k] === null || o[k] === undefined;
  });
  if (vazios.length) {
    dgDizer_(rel, "AVISO", "campos vazios na primeira linha: " + vazios.join(", "));
  }
}

// ---------------------------------------------------------------- resumo -----

function dgResumo_(rel) {
  const seg = ((Date.now() - rel.inicio) / 1000).toFixed(1);
  dgTitulo_(rel, "RESUMO");
  dgDizer_(rel, "", "diagnostico levou " + seg + "s");

  if (rel.falhas.length === 0 && rel.avisos.length === 0) {
    dgDizer_(rel, "", "NADA QUEBRADO. O servidor esta inteiro: planilha abre, colunas no lugar " +
             "e getTransferData responde. Se mesmo assim o portal nao abre, a falha esta NO " +
             "NAVEGADOR, nao aqui — abra o F12, aba Console e aba Network, e veja se algum " +
             "arquivo de cdn.jsdelivr.net foi bloqueado pela rede da empresa.");
  } else {
    if (rel.falhas.length) {
      dgDizer_(rel, "", rel.falhas.length + " FALHA(S) — em ordem de importancia:");
      rel.falhas.forEach(function (f, i) { dgDizer_(rel, "", "  " + (i + 1) + ". " + f); });
    }
    if (rel.avisos.length) {
      dgDizer_(rel, "", rel.avisos.length + " aviso(s):");
      rel.avisos.forEach(function (a, i) { dgDizer_(rel, "", "  " + (i + 1) + ". " + a); });
    }
  }

  const texto = rel.linhas.join("\n");
  console.log("\n(o relatorio inteiro tambem foi devolvido pela funcao — " +
              "de para copiar do painel de execucao)");
  return texto;
}

// ====================================================================
// EXTRA: so a chamada do portal, cronometrada, sem o resto da varredura.
// Para reconferir depois de mexer em alguma coisa, sem pagar o diagnostico todo.
// ====================================================================
function diagnosticarCargaRapida() {
  const t0 = Date.now();
  try {
    const dados = getTransferData();
    const ms = Date.now() - t0;
    const msg = "getTransferData(): " + dados.length + " linha(s), " +
                (JSON.stringify(dados).length / 1024).toFixed(1) + " KB, " +
                (ms / 1000).toFixed(1) + "s" +
                (ms > 30000 ? "  <<< LENTO DEMAIS: o navegador desiste antes disso" : "");
    console.log(msg);
    return msg;
  } catch (e) {
    const msg = "getTransferData() LANCOU apos " + ((Date.now() - t0) / 1000).toFixed(1) +
                "s: " + e.message + "\n" + (e.stack || "");
    console.error(msg);
    return msg;
  }
}
