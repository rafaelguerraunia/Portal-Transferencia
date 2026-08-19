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
// arquivada a todas as colunas marcadas. O arquivo da formula vive na propria
// aba Mapa_Formulas (colunas "Fórmula (exemplo)" e "Fórmula R1C1"), entao a
// volta nao depende de backup nenhum.
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

function fmEhErro_(v) {
  return typeof v === "string" && v.charAt(0) === "#";
}

// ====================================================================
// CATALOGO — a aba Mapa_Formulas lida como configuracao
// ====================================================================

// So volta o que esta marcado para firmar E tem formula arquivada. Uma linha
// sem R1C1 nao pode ser restaurada depois, e firmar sem poder voltar atras e
// exatamente o que este arquivo nao faz.
function fmLerCatalogo_(ss) {
  const aba = ss.getSheetByName(MF_ABA_MAPA);
  if (!aba) {
    return { erro: "aba '" + MF_ABA_MAPA + "' não existe — rode mapearFormulasPaginaTransferencia() primeiro.", itens: [] };
  }
  if (aba.getLastRow() < 2) {
    return { erro: "aba '" + MF_ABA_MAPA + "' está vazia — rode mapearFormulasPaginaTransferencia() primeiro.", itens: [] };
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, MF_MAPA_HEADERS.length).getValues();
  const itens = [], semFormula = [];

  dados.forEach(function (linha) {
    const marca = mfTxt_(linha[MF_C_FIRMAR]).toUpperCase();
    if (marca !== "SIM" && marca !== "S" && marca !== "SIM.") return;

    const letra = mfTxt_(linha[MF_C_LETRA]);
    const r1c1 = mfTxt_(linha[MF_C_R1C1]);
    const a1 = mfTxt_(linha[MF_C_FORMULA]);
    if (!r1c1 && !a1) { semFormula.push(letra); return; }

    itens.push({
      letra: letra,
      col: fmColunaDeLetra_(letra),
      nome: mfTxt_(linha[MF_C_NOME]),
      tipo: mfTxt_(linha[MF_C_TIPO]) === "array" ? "array" : "preenchida",
      r1c1: r1c1,
      a1: a1,
      ancora: Number(linha[MF_C_ANCORA]) || 2
    });
  });

  if (semFormula.length > 0) {
    console.warn("Coluna(s) marcada(s) para firmar sem fórmula arquivada, ignorada(s): " +
                 semFormula.join(", ") + ". Regere o mapa com a fórmula no lugar.");
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
  const ultimaLinha = aba.getLastRow();
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

    const cat = fmLerCatalogo_(ss);
    if (cat.erro) { console.log("Firmação não configurada: " + cat.erro); return null; }
    if (cat.itens.length === 0) {
      console.log("Nenhuma coluna marcada com \"SIM\" em " + MF_ABA_MAPA + ". Nada a firmar.");
      return null;
    }

    const estavel = fmAguardar_(aba, cat.itens, opts.esperaMaxMs);
    if (!estavel) return null;

    const celulas = fmGravarValores_(aba, cat.itens, estavel.linhas);
    PropertiesService.getScriptProperties().setProperty(
      FM_PROP_FIRMADA,
      JSON.stringify({ em: new Date().toISOString(),
                       colunas: cat.itens.map(function (i) { return i.letra; }),
                       linhas: estavel.linhas - 1 })
    );

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

    const cat = fmLerCatalogo_(ss);
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

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha >= 2) {
      preenchidas.forEach(function (item) {
        const faixa = aba.getRange(2, item.col, ultimaLinha - 1, 1);
        // R1C1 na faixa inteira reproduz o arrastar: as referencias relativas
        // resolvem por linha. Fazer isso com a formula em A1 apontaria todas as
        // linhas para a mesma celula da linha 2.
        if (item.r1c1) faixa.setFormulaR1C1(item.r1c1);
        else faixa.setFormula(item.a1);
      });
    }

    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().deleteProperty(FM_PROP_FIRMADA);

    console.log("Página Transferência restaurada: " + cat.itens.length +
                " coluna(s) voltaram a ser fórmula (" +
                cat.itens.map(function (i) { return i.letra; }).join(", ") + ").");

    return { colunas: cat.itens.length, linhas: Math.max(ultimaLinha - 1, 0) };
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
    const cat = fmLerCatalogo_(ss);
    if (cat.erro || cat.itens.length === 0) {
      // Silencioso de proposito: enquanto ninguem marcou coluna nenhuma, o sync
      // nao tem por que reclamar a cada 15 min.
      return null;
    }

    restaurarFormulasPaginaTransferencia({ ss: ss, esperaLockMs: opts.esperaLockMs || 30000 });
    return firmarPaginaTransferencia({ ss: ss, esperaMaxMs: opts.esperaMaxMs,
                                       esperaLockMs: opts.esperaLockMs || 30000 });
  } catch (e) {
    console.error("Página Transferência não foi firmada nesta passada: " + e.message);
    return null;
  }
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
