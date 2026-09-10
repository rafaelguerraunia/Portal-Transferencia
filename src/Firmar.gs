// ====================================================================
// FIRMAR AS COLUNAS CALCULADAS DA "Pagina Transferencia"
// ====================================================================
//
// PROBLEMA (o mesmo que a "Resumo P.O" tinha do lado do Portal de Pedidos): a
// pagina e montada por formula de ponta a ponta. Cada recalculo da planilha
// reavalia VLOOKUP na ME2W, cruzamento com RESB / ME2N / Stock e o que mais
// estiver ali — e o recalculo dispara sozinho, sem ninguem pedir: o sync
// reescreve as bases, TODAY() vira a meia-noite, um IMPORTRANGE se atualiza por
// conta propria. Enquanto isso corre, o getTransferData() do portal fica
// esperando: e a planilha em recalculo permanente que faz a tela demorar a abrir.
//
// SOLUCAO: uma vez por sincronizacao, deixar a formula calcular, esperar ela
// parar de mudar e gravar VALOR no lugar dela. Entre duas sincronizacoes a
// pagina e estatica — o portal le e vai embora, sem esperar conta nenhuma.
//
// O CICLO, e por que ele tem essa ordem:
//
//     restaurarFormulasPaginaTransferencia()   formula volta ao lugar
//                  |                            (a coluna firmada e valor,
//                  v                             valor nao recalcula)
//            SpreadsheetApp.flush()            a planilha recalcula
//                  |
//                  v
//              fmAguardar_()                   espera parar de mudar
//                  |
//                  v
//     firmarPaginaTransferencia()              le o resultado e grava valor
//
// Quem faz os tres passos de uma vez e refirmarPaginaTransferencia(), chamado
// pelo Sync no fim de cada sincronizacao. As funcoes soltas existem para rodar
// a mao pelo editor quando se quer inspecionar o meio do caminho.
//
// COMO SE ESCOLHE O QUE FIRMAR: pela coluna "Firmar?" da aba Mapa_Formulas,
// gerada pelo Formulas.gs. Nada e firmado por conta propria — o default de toda
// coluna nova e "NÃO". Trocar formula por valor estatico numa planilha em
// producao e decisao de quem conhece a coluna.
//
// COMO SE DESFAZ: restaurarFormulasPaginaTransferencia() devolve a formula
// arquivada as colunas marcadas. O arquivo da formula vive na propria aba
// Mapa_Formulas (colunas "Fórmula (exemplo)" e "Fórmula R1C1"), entao a volta
// nao depende de backup nenhum.
//
// COM UMA EXCECAO, e ela e o assunto do bloco fmExcluirEfetivo_ mais abaixo: as
// colunas que o Calculo.gs refaz em JS (Y..AQ) NAO voltam a ser formula por
// padrao. Repor AB e AJ e o que deixa a planilha em recalculo permanente e o
// portal sem abrir. Para desfazer essas duas tambem, e assumindo a espera:
// restaurarFormulasPaginaTransferencia({ incluirCalculadas: true }).
//
// EFEITO COLATERAL ACEITO, o mesmo da Resumo: entre duas sincronizacoes, uma
// linha que apareca na pagina fica com as colunas firmadas vazias ate a proxima
// passada. A janela e o intervalo do gatilho.

// Sinaliza que a pagina esta com colunas firmadas. Lido pelo Sync para nao
// tentar firmar o que nunca foi configurado, e pelo diagnostico.
const FM_PROP_FIRMADA = "PT_COLUNAS_FIRMADAS";

// Tempo maximo esperando o recalculo estabilizar, e de quanto em quanto se olha.
// 4 min e o que sobra do gatilho de 6 depois do sync; passar disso e melhor
// deixar a formula no lugar do que firmar dado pela metade.
const FM_ESPERA_MAX_MS = 4 * 60 * 1000;
const FM_ESPERA_INTERVALO_MS = 5000;

// Teto de celulas por setValues(), para nao estourar o payload da chamada.
const FM_LOTE_CELULAS = 20000;

// ATE ONDE VAI O DERRAME DA A2 — e o que NUNCA pode ser firmado.
//
// A A2 nao e "mais uma coluna de formula": e a fonte da pagina inteira. A
// ARRAYFORMULA dela derrama nos DOIS sentidos — o VSTACK empilha ME5A e ME2W na
// vertical, e o HSTACK monta 24 colunas na horizontal. Tudo de A ate X e o
// resultado de UMA formula que mora numa celula so.
//
// Firmar essa coluna e o unico jeito de quebrar a pagina inteira de uma vez. O
// fmGravarValores_ monta corridas so das colunas MARCADAS: com a A marcada
// sozinha a corrida tem largura 1, ele le A2:A e regrava como valor — e gravar
// valor na ancora mata a formula, o que leva junto o derrame de B ate X, que
// ninguem tinha pedido para firmar. Sem N (Material), I (Planta) e G/H/U
// (Documento/Item/Schedule Line), as colunas Y..AQ passam a ler celula vazia e
// nao ha recalculo que se sustente — o portal abre uma pagina em branco.
//
// E nao ha o que se ganhar do outro lado: A:X e dado cru copiado da ME5A e da
// ME2W, sem conta cara para congelar. O custo do recalculo mora em Y..AQ, e e
// so ali que firmar paga.
//
// Quem ja firmou a A antes desta protecao existir volta pelo
// restaurarDerramePaginaTransferencia().
const FM_ULTIMA_COLUNA_DERRAME = 24;   // X

function fmEhErro_(v) {
  return typeof v === "string" && v.charAt(0) === "#";
}

// ====================================================================
// O QUE ESTE ARQUIVO NUNCA RESTAURA SOZINHO — as colunas do Calculo.gs
// ====================================================================
//
// A COLUNA "Firmar?" TEM DOIS LEITORES, E ELES A LEEM AO CONTRARIO UM DO OUTRO:
//
//   Calculo.gs  "SIM" = refaz a conta em JS e grava VALOR. A formula sai.
//   Firmar.gs   "SIM" = REPOE a formula, espera o recalculo e congela o
//               resultado. A formula volta antes de sair.
//
// Para AB ("Dias Disponiveis em Estoque") e AJ ("Dias Disponiveis na BR14") o
// segundo caminho nao e mais caro: e inviavel. As duas sao MAP dentro de MAP
// sobre SEQUENCE(365) — 365 SUMIFS por linha da pagina, cada um varrendo RESB
// inteira. Repor essa formula invalida a coluna toda de uma vez, e o
// SpreadsheetApp.flush() do restaurar fica esperando o recalculo que nao termina:
// a execucao morre no limite de 6 min COM A FORMULA JA REPOSTA. A planilha fica
// em recalculo permanente, o getDataRange() do getTransferData passa a esperar
// por ele, e o portal para de abrir — o sintoma de "firmei e o HTML parou".
//
// O `excluir` sempre existiu para evitar isso, mas era passado por UM chamador
// so (o firmarAposSync_ do Sync.gs). Rodar refirmarPaginaTransferencia() a mao
// pelo editor — que e exatamente o que se faz para "tirar as formulas" — nao
// passava nada, e o caminho generico repunha as 19 colunas, AB e AJ inclusive.
//
// Agora o default e o inverso: sem dizer nada, as colunas do Calculo.gs ficam de
// fora, venha a chamada de onde vier. Quem quiser mesmo a formula de volta pede
// por escrito — restaurarFormulasPaginaTransferencia({incluirCalculadas: true}) —
// e assume a espera.
function fmColunasDoCalculo_() {
  return (typeof ptLetrasCalculadas_ === "function") ? ptLetrasCalculadas_() : [];
}

function fmExcluirEfetivo_(opts) {
  const o = opts || {};
  if (o.excluir) return o.excluir;              // lista explicita manda
  if (o.incluirCalculadas === true) return [];  // opt-in consciente
  return fmColunasDoCalculo_();
}

// ====================================================================
// ATE ONDE A PAGINA REALMENTE VAI
// ====================================================================
//
// getLastRow() responde pela ultima celula com conteudo da ABA — e isso nao e o
// fim da pagina. Abaixo do derrame sobra o que ficou de uma pagina maior: uma
// formula arrastada ate uma linha que hoje nao existe mais, ou o valor de uma
// firmacao feita quando havia mais STOs. No export atual da planilha o derrame
// termina na linha 262 e o getLastRow() diz 695 — 433 linhas com A:X vazias e
// AO/AP/AQ escritas ("Nao", "Nao", "-").
//
// Dimensionar por getLastRow() faz todo o resto trabalhar sobre linha que nao
// existe: o Calculo.gs calcula 694 linhas em vez de 261, o restaurar escreve 694
// formulas por coluna, o fmAssinatura_ soma hash de 694 linhas A CADA sondagem.
// Com AB/AJ na conta, esse 2,7x e a diferenca entre o recalculo caber no gatilho
// e nao caber.
//
// A verdade sobre o tamanho esta no derrame, e dentro dele na coluna Documento:
// toda linha que o portal aproveita tem documento (ver o
// `if (documento === "") continue` do getTransferData).
const FM_COLUNA_CHAVE_DERRAME = 7;   // G  Documento

function fmUltimaLinhaDerrame_(aba) {
  const bruta = aba.getLastRow();
  if (bruta < 2) return bruta;

  const col = aba.getRange(2, FM_COLUNA_CHAVE_DERRAME, bruta - 1, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== "") return i + 2;
  }
  // Nenhum documento em linha nenhuma: o derrame esta morto. Devolver 1 e o que
  // faz os chamadores tratarem como "sem linhas" e NAO gravarem por cima —
  // firmar uma pagina vazia congelaria o vazio.
  return 1;
}

// ====================================================================
// CATALOGO — a aba Mapa_Formulas lida como configuracao
// ====================================================================

// So volta o que esta marcado para firmar E tem formula arquivada. Uma linha
// sem R1C1 nao pode ser restaurada depois, e firmar sem poder voltar atras e
// exatamente o que este arquivo nao faz.
//
// `excluir` e a lista de letras que OUTRO caminho ja resolve — na pratica, as
// colunas que o Calculo.gs refaz em JS. Sem ela o Sync chamava este catalogo
// duas vezes por ciclo e a segunda passada repunha a formula das mesmas colunas
// que a primeira tinha acabado de calcular: AB e AJ voltavam a ser os 365
// SUMIFS por linha que o Calculo.gs existe para nao pagar, o fmAguardar_
// estourava o tempo esperando, e a pagina terminava o sync em formula.
function fmLerCatalogo_(ss, excluir) {
  const aba = ss.getSheetByName(MF_ABA_MAPA);
  if (!aba) {
    return { erro: "aba '" + MF_ABA_MAPA + "' não existe — rode mapearFormulasPaginaTransferencia() primeiro.", itens: [] };
  }
  if (aba.getLastRow() < 2) {
    return { erro: "aba '" + MF_ABA_MAPA + "' está vazia — rode mapearFormulasPaginaTransferencia() primeiro.", itens: [] };
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, MF_MAPA_HEADERS.length).getValues();
  const itens = [], semFormula = [], protegidas = [], delegadas = [];

  const fora = {};
  (excluir || []).forEach(function (l) { fora[String(l).toUpperCase()] = true; });

  dados.forEach(function (linha) {
    const marca = mfTxt_(linha[MF_C_FIRMAR]).toUpperCase();
    if (marca !== "SIM" && marca !== "S" && marca !== "SIM.") return;

    const letra = mfTxt_(linha[MF_C_LETRA]);
    if (fora[letra.toUpperCase()]) { delegadas.push(letra); return; }

    const col = fmColunaDeLetra_(letra);

    // O veto e aqui, e nao na hora de gravar, porque este catalogo e lido pelos
    // TRES caminhos — firmar, restaurar e o calculo em JS. Barrar num lugar so
    // deixaria os outros dois passando por cima do derrame.
    if (col > 0 && col <= FM_ULTIMA_COLUNA_DERRAME) { protegidas.push(letra); return; }

    const r1c1 = mfTxt_(linha[MF_C_R1C1]);
    const a1 = mfTxt_(linha[MF_C_FORMULA]);
    if (!r1c1 && !a1) { semFormula.push(letra); return; }

    itens.push({
      letra: letra,
      col: col,
      nome: mfTxt_(linha[MF_C_NOME]),
      // "solta" restaura como "array" — a ancora sozinha. Espalhar a formula pela
      // coluna inteira consertaria a coluna, mas restaurar e devolver ao estado
      // anterior, nao corrigi-lo: quem conserta e o Calculo.gs, calculando todas
      // as linhas.
      tipo: mfTxt_(linha[MF_C_TIPO]) === "preenchida" ? "preenchida" : "array",
      r1c1: r1c1,
      a1: a1,
      ancora: Number(linha[MF_C_ANCORA]) || 2
    });
  });

  if (semFormula.length > 0) {
    console.warn("Coluna(s) marcada(s) para firmar sem fórmula arquivada, ignorada(s): " +
                 semFormula.join(", ") + ". Regere o mapa com a fórmula no lugar.");
  }
  if (protegidas.length > 0) {
    console.warn("Coluna(s) " + protegidas.join(", ") + " marcada(s) para firmar e RECUSADA(S): " +
                 "estão dentro do derrame da A2 (A:X), que monta a página inteira. " +
                 "Firmar ali esvazia B:X e derruba todas as colunas calculadas. " +
                 "Desmarque no " + MF_ABA_MAPA + "; se a página já ficou vazia, rode " +
                 "restaurarDerramePaginaTransferencia().");
  }
  if (delegadas.length > 0) {
    console.log("Coluna(s) " + delegadas.join(", ") + " ficaram com o Calculo.gs " +
                "(conta refeita em JS) — o caminho genérico não as toca.");
  }
  return { erro: null, itens: itens.filter(function (i) { return i.col > 0; }) };
}

function fmColunaDeLetra_(letra) {
  const s = String(letra || "").toUpperCase().replace(/[^A-Z]/g, "");
  if (!s) return 0;
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

// ====================================================================
// PRONTIDAO — a parte que impede firmar dado pela metade
// ====================================================================

// Retrato do estado atual das colunas que serao firmadas: ate onde a pagina vai
// e o que tem dentro delas. Duas leituras iguais em sequencia = o recalculo
// terminou. So o numero de linhas nao bastaria: a pagina pode manter o tamanho e
// trocar o conteudo (mesma quantidade de STOs, STOs diferentes).
function fmAssinatura_(aba, itens) {
  // fmUltimaLinhaDerrame_ e nao getLastRow(): a assinatura tem de descrever a
  // pagina, nao a sobra abaixo dela. Somar hash de linha orfa a cada sondagem
  // custa tempo de gatilho e nao muda de valor nunca.
  const ultimaLinha = fmUltimaLinhaDerrame_(aba);
  if (ultimaLinha < 2) return { linhas: 0, hash: 0, erro: "" };

  // Um getDisplayValues por coluna seria uma ida ao servidor por coluna A CADA
  // sondagem, e sondagem tem varias. Le o retangulo que cobre todas as colunas de
  // uma vez e escolhe as que interessam na memoria — as do meio que nao serao
  // firmadas vem junto e sao descartadas, o que sai muito mais barato.
  let menor = Infinity, maior = 0;
  itens.forEach(function (i) {
    if (i.col < menor) menor = i.col;
    if (i.col > maior) maior = i.col;
  });

  const bloco = aba.getRange(2, menor, ultimaLinha - 1, maior - menor + 1).getDisplayValues();
  let hash = 0;
  let erro = "";

  for (let k = 0; k < itens.length; k++) {
    const desloc = itens[k].col - menor;
    for (let i = 0; i < bloco.length; i++) {
      const s = bloco[i][desloc];
      if (!erro && (fmEhErro_(s) || s === "Loading...")) {
        erro = itens[k].letra + "=" + s;
      }
      for (let c = 0; c < s.length; c++) hash = (hash * 31 + s.charCodeAt(c)) | 0;
    }
  }
  return { linhas: ultimaLinha, hash: hash, erro: erro };
}

// Espera a pagina parar de mudar. Devolve a assinatura estavel, ou null quando
// estourou o tempo ou ha erro nas colunas. Nunca lanca: o pior caso e nao firmar
// nesta passada, e a formula continua no lugar fazendo o que sempre fez.
function fmAguardar_(aba, itens, esperaMaxMs) {
  const limite = Date.now() + (esperaMaxMs || FM_ESPERA_MAX_MS);
  let anterior = null;

  while (true) {
    SpreadsheetApp.flush();
    const atual = fmAssinatura_(aba, itens);

    if (atual.linhas < 2) {
      console.warn("Página Transferência sem linhas de dados. Nada a firmar.");
      return null;
    }
    if (atual.erro) {
      console.warn("Página Transferência com " + atual.erro +
                   " — origem ainda carregando. Fórmulas mantidas.");
      return null;
    }
    if (anterior && anterior.linhas === atual.linhas && anterior.hash === atual.hash) {
      return atual;
    }
    if (Date.now() >= limite) {
      console.warn("Página Transferência ainda recalculando após " +
                   Math.round((esperaMaxMs || FM_ESPERA_MAX_MS) / 1000) +
                   "s. Fórmulas mantidas para não firmar dado incompleto.");
      return null;
    }
    anterior = atual;
    Utilities.sleep(FM_ESPERA_INTERVALO_MS);
  }
}

// ====================================================================
// GRAVACAO
// ====================================================================

// Grava em corridas de colunas ADJACENTES. Escrever um retangulo do minimo ao
// maximo passaria por cima das colunas NAO marcadas que estao no meio — e essas
// continuam sendo formula, ou preenchimento manual de operacao.
function fmGravarValores_(aba, itens, ultimaLinha) {
  const total = ultimaLinha - 1;
  if (total < 1) return 0;

  // Abaixo do derrame nao ha linha — o que estiver ali e sobra de uma pagina
  // maior. Limpar e o que impede a sobra de inflar o getLastRow() e fazer a
  // proxima passada trabalhar sobre linhas que nao existem.
  const maxLinhas = aba.getMaxRows();

  const ordenadas = itens.slice().sort(function (a, b) { return a.col - b.col; });
  const runs = [];
  ordenadas.forEach(function (item) {
    const ultimo = runs.length ? runs[runs.length - 1] : null;
    if (ultimo && item.col === ultimo[ultimo.length - 1].col + 1) ultimo.push(item);
    else runs.push([item]);
  });

  let gravadas = 0;
  runs.forEach(function (run) {
    const largura = run.length;

    // LER A CORRIDA INTEIRA ANTES DE ESCREVER QUALQUER LINHA. Ler e escrever
    // alternadamente, bloco a bloco, destroi a coluna "array": a ancora dela mora
    // no primeiro bloco, e grava-la mata a formula — o derrame das linhas de baixo
    // some junto, e o bloco seguinte seria lido ja vazio. O resultado era a coluna
    // firmada com o primeiro lote certo e o resto em branco.
    const valores = [];
    const passoLeitura = Math.max(1, Math.floor(FM_LOTE_CELULAS / largura));
    for (let ini = 0; ini < total; ini += passoLeitura) {
      const n = Math.min(passoLeitura, total - ini);
      const parte = aba.getRange(2 + ini, run[0].col, n, largura).getValues();
      for (let i = 0; i < parte.length; i++) valores.push(parte[i]);
    }

    for (let ini = 0; ini < total; ini += passoLeitura) {
      const n = Math.min(passoLeitura, total - ini);
      aba.getRange(2 + ini, run[0].col, n, largura)
         .setValues(valores.slice(ini, ini + n));
      gravadas += n * largura;
    }

    const sobra = maxLinhas - ultimaLinha;
    if (sobra > 0) aba.getRange(ultimaLinha + 1, run[0].col, sobra, largura).clearContent();
  });
  return gravadas;
}

// ====================================================================
// PONTOS DE ENTRADA
// ====================================================================

// Congela em valor as colunas marcadas "SIM" no Mapa_Formulas.
//
// NAO restaura formula antes: parte do principio de que a pagina ja esta
// calculada. Chamado logo depois de restaurarFormulas..., e o segundo passo do
// ciclo; chamado sozinho numa pagina ja firmada, nao faz mal nenhum (regravar
// valor sobre valor e no-op) mas tambem nao atualiza nada.
// Para o ciclo completo use refirmarPaginaTransferencia().
function firmarPaginaTransferencia(opcoes) {
  const opts = opcoes || {};
  const lock = LockService.getScriptLock();
  // tryLock(0): se outra execucao ja esta firmando, esta desiste. Duas passadas
  // simultaneas gravariam resultados de instantes diferentes na mesma coluna.
  // E o MESMO lock que o Sync e os saves do portal disputam — firmar no meio de
  // um clique gravaria a pagina sem a confirmacao que acabou de ser salva.
  if (!lock.tryLock(opts.esperaLockMs || 0)) {
    console.log("Página Transferência: outra execução está com o lock. Firmação ignorada.");
    return null;
  }

  try {
    const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName(MF_ABA_ALVO);
    if (!aba) {
      console.error("Aba '" + MF_ABA_ALVO + "' não encontrada. Nada a firmar.");
      return null;
    }

    const cat = fmLerCatalogo_(ss, fmExcluirEfetivo_(opts));
    if (cat.erro) { console.log("Firmação não configurada: " + cat.erro); return null; }
    if (cat.itens.length === 0) {
      console.log("Nenhuma coluna marcada com \"SIM\" em " + MF_ABA_MAPA + ". Nada a firmar.");
      return null;
    }

    const estavel = fmAguardar_(aba, cat.itens, opts.esperaMaxMs);
    if (!estavel) return null;

    const celulas = fmGravarValores_(aba, cat.itens, estavel.linhas);
    fmMarcarFirmadas_(cat.itens.map(function (i) { return i.letra; }), estavel.linhas - 1);

    console.log("Página Transferência firmada: " + cat.itens.length + " coluna(s) (" +
                cat.itens.map(function (i) { return i.letra; }).join(", ") + ") x " +
                (estavel.linhas - 1) + " linha(s) = " + celulas + " célula(s) agora em valor.");

    return { colunas: cat.itens.length, linhas: estavel.linhas - 1, celulas: celulas };
  } finally {
    lock.releaseLock();
  }
}

// Devolve a formula arquivada as colunas marcadas. E o passo 1 do ciclo, e
// tambem o botao de desfazer: rodar isto e parar de chamar o refirmar deixa a
// pagina exatamente como era antes.
//
// DUAS PASSADAS, nesta ordem, porque a pagina pode ter mudado de tamanho:
//   1. as colunas "array" (a ancora sozinha — o resto e derrame dela);
//   2. flush, para o derrame acontecer e o getLastRow() dizer a verdade;
//   3. as colunas "preenchida", ate a ultima linha que passou a existir.
function restaurarFormulasPaginaTransferencia(opcoes) {
  const opts = opcoes || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(opts.esperaLockMs || 30000)) {
    throw new Error("Lock ocupado — restauração adiada para não escrever junto com o sync ou com um save do portal.");
  }

  try {
    const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName(MF_ABA_ALVO);
    if (!aba) {
      console.error("Aba '" + MF_ABA_ALVO + "' não encontrada. Nada a restaurar.");
      return null;
    }

    // As colunas do Calculo.gs ficam de fora por PADRAO (ver fmExcluirEfetivo_):
    // repor a formula de AB/AJ e o que derruba o portal, e o default nao pode
    // depender de o chamador lembrar de passar a lista. Para desfazer de verdade,
    // inclusive essas duas, chame com {incluirCalculadas: true}.
    const cat = fmLerCatalogo_(ss, fmExcluirEfetivo_(opts));
    if (cat.erro) { console.log("Restauração não configurada: " + cat.erro); return null; }
    if (cat.itens.length === 0) {
      console.log("Nenhuma coluna marcada com \"SIM\" em " + MF_ABA_MAPA + ". Nada a restaurar.");
      return null;
    }

    const arrays = cat.itens.filter(function (i) { return i.tipo === "array"; });
    const preenchidas = cat.itens.filter(function (i) { return i.tipo !== "array"; });

    // Medido UMA vez, antes de limpar qualquer coisa: cada clearContent encurta o
    // getLastRow() da aba, e reler dentro do laco faria a segunda coluna limpar
    // menos linhas que a primeira — sobrando derrame velho embaixo da fórmula nova.
    const antesDeLimpar = aba.getLastRow();
    arrays.forEach(function (item) {
      const linha = item.ancora || 2;
      // Limpa o derrame materializado antes de repor a ancora: com valor abaixo
      // dela a formula nao tem para onde derramar e volta #REF!.
      if (antesDeLimpar > linha) {
        aba.getRange(linha + 1, item.col, antesDeLimpar - linha, 1).clearContent();
      }
      aba.getRange(linha, item.col).setFormula(item.a1 || item.r1c1);
    });

    if (arrays.length > 0) SpreadsheetApp.flush();

    // Duas medidas, de proposito. LIMPAR usa o getLastRow() (limpar de mais e
    // seguro); ESCREVER formula usa o fim do derrame — cada formula reposta numa
    // linha orfa e uma linha inteira de conta paga para produzir nada.
    const ultimaLinha = fmUltimaLinhaDerrame_(aba);
    const ultimaBruta = aba.getLastRow();
    if (ultimaLinha >= 2) {
      preenchidas.forEach(function (item) {
        if (ultimaBruta > ultimaLinha) {
          aba.getRange(ultimaLinha + 1, item.col, ultimaBruta - ultimaLinha, 1).clearContent();
        }
        const faixa = aba.getRange(2, item.col, ultimaLinha - 1, 1);
        // R1C1 na faixa inteira reproduz o arrastar: as referencias relativas
        // resolvem por linha. Fazer isso com a formula em A1 apontaria todas as
        // linhas para a mesma celula da linha 2.
        if (item.r1c1) faixa.setFormulaR1C1(item.r1c1);
        else faixa.setFormula(item.a1);
      });
    }

    SpreadsheetApp.flush();
    // Só as que voltaram: numa restauração parcial (o Sync passa `excluir`) as
    // colunas do Calculo.gs continuam firmadas e continuam no registro.
    fmDesmarcarFirmadas_(cat.itens.map(function (i) { return i.letra; }));

    console.log("Página Transferência restaurada: " + cat.itens.length +
                " coluna(s) voltaram a ser fórmula (" +
                cat.itens.map(function (i) { return i.letra; }).join(", ") + ").");

    return { colunas: cat.itens.length, linhas: Math.max(ultimaLinha - 1, 0) };
  } finally {
    lock.releaseLock();
  }
}

// Acha a formula arquivada de UMA coluna, marcada ou nao. O fmLerCatalogo_ so
// devolve o que esta com "SIM", e a coluna do derrame nunca esta — nem pode
// estar. Sem isto, arquivar a formula da A2 no mapa nao serviria de nada
// justamente no caso em que ela precisa voltar.
function fmFormulaArquivada_(ss, letra) {
  const aba = ss.getSheetByName(MF_ABA_MAPA);
  if (!aba || aba.getLastRow() < 2) {
    return { erro: "aba '" + MF_ABA_MAPA + "' não existe ou está vazia — sem ela não há fórmula arquivada para repor." };
  }

  const alvo = String(letra).toUpperCase();
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, MF_MAPA_HEADERS.length).getValues();

  for (let i = 0; i < dados.length; i++) {
    if (mfTxt_(dados[i][MF_C_LETRA]).toUpperCase() !== alvo) continue;

    const a1 = mfTxt_(dados[i][MF_C_FORMULA]);
    const r1c1 = mfTxt_(dados[i][MF_C_R1C1]);
    if (!a1 && !r1c1) {
      return { erro: "a coluna " + alvo + " está no mapa mas sem fórmula arquivada." };
    }
    return { erro: null, a1: a1, r1c1: r1c1, ancora: Number(dados[i][MF_C_ANCORA]) || 2 };
  }
  return { erro: "a coluna " + alvo + " não está no " + MF_ABA_MAPA + ". Rode mapearFormulasPaginaTransferencia() com a fórmula no lugar." };
}

// SOCORRO: devolve a A2 ao lugar depois de uma firmação que congelou o derrame.
//
// Para quem firmou a coluna A antes de o FM_ULTIMA_COLUNA_DERRAME existir e
// ficou com a página de B até X vazia. Não entra em nenhum ciclo automático: é
// uma correção de uma vez só, rodada à mão pelo editor.
//
// O QUE ELA FAZ DE DIFERENTE de uma tentativa manual: limpa o RETÂNGULO A:X
// inteiro antes de repor a fórmula, não só a coluna A. Com valor parado em
// qualquer célula de B..X a ARRAYFORMULA não tem para onde derramar e volta
// #REF! — e aí o que era uma página vazia vira uma página com erro.
//
// Lê a fórmula ANTES de limpar qualquer coisa: sem ela arquivada no mapa, esta
// função não toca na página.
function restaurarDerramePaginaTransferencia(opcoes) {
  const opts = opcoes || {};
  const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(MF_ABA_ALVO);
  if (!aba) throw new Error("Aba '" + MF_ABA_ALVO + "' não encontrada.");

  const letra = opts.letra || "A";
  const arq = fmFormulaArquivada_(ss, letra);
  if (arq.erro) throw new Error("Derrame não restaurado: " + arq.erro);

  const col = fmColunaDeLetra_(letra);
  if (col < 1 || col > FM_ULTIMA_COLUNA_DERRAME) {
    throw new Error("A coluna " + letra + " não é a âncora do derrame (A:X). " +
                    "Para as colunas calculadas use restaurarFormulasPaginaTransferencia().");
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(opts.esperaLockMs || 30000)) {
    throw new Error("Lock ocupado — restauração adiada para não escrever junto com o sync ou com um save do portal.");
  }

  try {
    const ancora = arq.ancora;
    const ultima = aba.getLastRow();

    // O retângulo inteiro, incluindo a própria âncora: ela está com o valor
    // congelado que matou a fórmula.
    if (ultima >= ancora) {
      aba.getRange(ancora, 1, ultima - ancora + 1, FM_ULTIMA_COLUNA_DERRAME).clearContent();
    }
    aba.getRange(ancora, 1).setFormula(arq.a1 || arq.r1c1);
    SpreadsheetApp.flush();

    const derramou = Math.max(aba.getLastRow() - ancora + 1, 0);
    console.log("Derrame restaurado: fórmula reposta em " + letra + ancora +
                " e " + derramou + " linha(s) preenchida(s) de A até X.");

    if (derramou === 0) {
      console.warn("A fórmula voltou mas não derramou nenhuma linha — confira se as abas " +
                   "ME2W e ME5A estão sincronizadas.");
    }

    // As colunas calculadas continuam com o valor da última firmação, que foi
    // tirado de uma página em outro estado. Elas se acertam sozinhas na próxima
    // passada; dizer isso aqui evita a caça a "dado velho" que não é defeito.
    console.log("As colunas Y..AQ ainda estão com o valor da firmação anterior — " +
                "a próxima sincronização (ou firmarColunasCalculadasTransferencia({todas:true})) as reescreve.");

    return { letra: letra, ancora: ancora, linhas: derramou };
  } finally {
    lock.releaseLock();
  }
}

// O ciclo completo: repoe a formula, espera calcular, grava valor. E esta que o
// Sync chama no fim de cada sincronizacao — o equivalente ao
// firmarColunasCalculadasResumo() do Portal de Pedidos, so que aqui a conta e
// feita pela propria planilha em vez de reimplementada em JS.
//
// Nao lanca: firmar e o passo opcional do sync. Falhar aqui nao pode desfazer
// nem mascarar uma sincronizacao que deu certo, e o pior caso — pagina com
// formula — e o comportamento de sempre.
function refirmarPaginaTransferencia(opcoes) {
  const opts = opcoes || {};
  try {
    const ss = opts.ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const excluir = fmExcluirEfetivo_(opts);
    const cat = fmLerCatalogo_(ss, excluir);
    if (cat.erro || cat.itens.length === 0) {
      // Silencioso de proposito: enquanto ninguem marcou coluna nenhuma, o sync
      // nao tem por que reclamar a cada 15 min.
      //
      // E tambem o caso NORMAL depois do Calculo.gs: se toda coluna marcada tem
      // conta em JS, nao sobra nada para o caminho generico e a passada inteira
      // e pulada — que e justamente o que mantem a pagina firmada em vez de
      // devolve-la para formula.
      return null;
    }

    restaurarFormulasPaginaTransferencia({ ss: ss, excluir: excluir,
                                           esperaLockMs: opts.esperaLockMs || 30000 });
    return firmarPaginaTransferencia({ ss: ss, excluir: excluir,
                                       esperaMaxMs: opts.esperaMaxMs,
                                       esperaLockMs: opts.esperaLockMs || 30000 });
  } catch (e) {
    console.error("Página Transferência não foi firmada nesta passada: " + e.message);
    return null;
  }
}

// ====================================================================
// O REGISTRO DO QUE ESTA FIRMADO
// ====================================================================
//
// A property e escrita pelos DOIS caminhos — o Calculo.gs e o
// firmarPaginaTransferencia daqui — e lida pelo fmPrecisaRefirmar_ e pelo
// diagnostico. Por isso ela e UNIAO, nunca substituicao: cada caminho
// acrescenta as suas colunas em vez de apagar as do outro.
//
// Substituir dava dois defeitos no mesmo ciclo do Sync. O diagnostico passava a
// reportar so as colunas do ultimo caminho a rodar; e, pior, o restaurar apagava
// a property inteira antes de o generico gravar a dele — entao um estouro de
// tempo na espera do recalculo deixava a pagina COM as colunas do Calculo.gs
// firmadas e SEM registro nenhum. O fmPrecisaRefirmar_ lia "nunca firmou" e
// mandava refirmar tudo de novo, a cada gatilho, para sempre.
function fmMarcarFirmadas_(letras, linhas) {
  const props = PropertiesService.getScriptProperties();
  const hoje = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");

  let anteriores = [];
  try {
    const bruto = props.getProperty(FM_PROP_FIRMADA);
    if (bruto) {
      const estado = JSON.parse(bruto);
      // So aproveita registro do proprio dia: o de ontem nao vale mais, porque
      // TODAY() virou debaixo dele.
      if (Utilities.formatDate(new Date(estado.em), TIMEZONE, "yyyy-MM-dd") === hoje) {
        anteriores = estado.colunas || [];
      }
    }
  } catch (e) { /* property ilegivel: recomeca do zero */ }

  const uniao = {};
  anteriores.concat(letras || []).forEach(function (l) { if (l) uniao[l] = true; });

  props.setProperty(FM_PROP_FIRMADA, JSON.stringify({
    em: new Date().toISOString(),
    colunas: Object.keys(uniao).sort(),
    linhas: linhas
  }));
}

// Tira do registro APENAS as colunas que voltaram a ser formula. Some com a
// property so quando nao sobra nenhuma — e o unico caso em que "nada esta
// firmado" e verdade.
function fmDesmarcarFirmadas_(letras) {
  const props = PropertiesService.getScriptProperties();
  const bruto = props.getProperty(FM_PROP_FIRMADA);
  if (!bruto) return;

  let estado;
  try { estado = JSON.parse(bruto); } catch (e) { props.deleteProperty(FM_PROP_FIRMADA); return; }

  const saiu = {};
  (letras || []).forEach(function (l) { if (l) saiu[l] = true; });
  const restam = (estado.colunas || []).filter(function (l) { return !saiu[l]; });

  if (restam.length === 0) props.deleteProperty(FM_PROP_FIRMADA);
  else props.setProperty(FM_PROP_FIRMADA, JSON.stringify({
    em: estado.em, colunas: restam, linhas: estado.linhas
  }));
}

// Vale a pena refirmar nesta passada? Duas coisas envelhecem a pagina, e so
// essas duas:
//
//   1. o sync escreveu base nova (ME2W, RESB, ME2N, Stock mudaram debaixo dela);
//   2. o dia virou — TODAY() e as contas de janela (+7D) passaram a responder
//      outra coisa, mesmo sem nenhum XLSX ter chegado.
//
// Fora desses dois casos a pagina firmada ja esta certa, e refirmar so pagaria um
// recalculo inteiro para chegar no mesmo resultado. E a diferenca entre firmar
// uma ou duas vezes por dia e firmar a cada 15 min sem motivo.
function fmPrecisaRefirmar_(escreveuBase) {
  if (escreveuBase) return true;

  const bruto = PropertiesService.getScriptProperties().getProperty(FM_PROP_FIRMADA);
  if (!bruto) return true;   // nunca firmou

  try {
    const estado = JSON.parse(bruto);
    const hoje = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    const entao = Utilities.formatDate(new Date(estado.em), TIMEZONE, "yyyy-MM-dd");
    return hoje !== entao;
  } catch (e) {
    return true;   // property ilegivel: refirma e ela se conserta
  }
}

// Diagnostico de uma linha: o que esta firmado, desde quando, e o que o mapa
// diz que deveria estar. Util antes de sair investigando dado velho na tela.
function statusFirmacaoPaginaTransferencia() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cat = fmLerCatalogo_(ss);
  const bruto = PropertiesService.getScriptProperties().getProperty(FM_PROP_FIRMADA);
  const estado = bruto ? JSON.parse(bruto) : null;

  const marcadas = cat.itens.map(function (i) { return i.letra; });
  console.log("Marcadas em " + MF_ABA_MAPA + ": " + (marcadas.join(", ") || "nenhuma") +
              (cat.erro ? " (" + cat.erro + ")" : ""));
  console.log(estado
    ? "Firmadas em " + estado.em + ": " + estado.colunas.join(", ") + " x " + estado.linhas + " linha(s)."
    : "Nenhuma firmação registrada — a página está inteira em fórmula.");

  return { marcadas: marcadas, ultimaFirmacao: estado };
}