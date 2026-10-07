// ====================================================================
// AVISO DE TRANSFERÊNCIAS EMERGENCIAIS — e-mail ao armazém
// ====================================================================
//
// STO firmada DEPOIS do corte — 10:00 do dia anterior à entrega, em dias corridos —
// furou a programação do armazém. A cada 15 min (instalarAvisoEmergencial) as que
// ainda não foram avisadas vão numa lista só, dividida pela Plant (destino) da ME2W.
// Sem nada novo, nada é enviado.
//
// A REGRA, por linha da Pagina Transferência:
//   - Status Planejamento "Firme" — o mesmo do portal: confirmação de pé e igual ao
//     SAP em data e quantidade. Completa ou cancelada no SAP não é Firme, e por isso a
//     STO já atendida não entra;
//   - Firmado em (Confirmacoes_Store) a partir do corte. Inclui a firmada no próprio
//     dia da entrega a qualquer hora (D0 antes das 10h também);
//   - a entrega ainda não tinha passado no ato do firme: firmar STO vencida é
//     arrumação de carteira, não emergência;
//   - firme das últimas AVISO_EMERG_JANELA_HORAS e posterior ao marco gravado na
//     instalação — ligar o aviso não despeja no armazém os firmes de antes dele.
//
// UMA VEZ POR FIRME. O Log_Email_Emergencial guarda, por linha enviada, a chave e o
// Firmado em. A STO firmada de novo volta como "Atualizada" só se a entrega ou a
// quantidade mudaram desde o último aviso dela; o mesmo firme repetido não volta.
// A STO avisada que depois é limpa ou reprogramada para fora da regra NÃO gera aviso.
//
// OS NÚMEROS DO TOPO contam só as linhas do próprio e-mail: cada um se refaz contando
// a tabela logo abaixo dele. Como nenhum firme sai em dois avisos, a soma dos avisos
// do dia é o total do dia — e é o "acumulado" que sai do log (que só cresce). Cada
// aviso leva o seu número no dia: um buraco na sequência é um e-mail que não chegou.

const AVISO_EMERG_PARA = "armazembr14@ftatransportes.com.br,bp_fernanda_frazao@colpal.com";
const AVISO_EMERG_REMETENTE = "Smart Hub · Portal de Transferência";
const AVISO_EMERG_HORA_CORTE = 10;
const AVISO_EMERG_JANELA_HORAS = 24;
const AVISO_EMERG_TESTE_DIAS = 7;
const AVISO_EMERG_COL_PLANT = "Plant";
const AVISO_EMERG_COL_PALETES = "Pallet Order";
const PROP_AVISO_EMERG_INICIO = "AVISO_EMERG_INICIO";
const ABA_LOG_EMERG = "Log_Email_Emergencial";

// Uma linha por STO enviada, na ordem do envio. As 7 primeiras são as que a
// deduplicação lê, numa leitura só (LG_*): ficam juntas e no começo. Coluna nova, no FIM.
const LOG_EMERG_HEADERS = [
  "Enviado em", "Aviso nº", "Chave", "Firmado em", "Delivery Date", "Order Quantity", "Plant",
  "Tipo", "Purchasing Document", "Item", "Schedule Line", "Material", "Short Text",
  "Order Unit", "Paletes", "Corte", "Destinatários"
];
const LG_ENVIADO = 0, LG_AVISO = 1, LG_CHAVE = 2, LG_FIRMADO = 3, LG_DATA = 4, LG_QTD = 5, LG_PLANT = 6;
const LG_LARGURA_LIDA = 7;

// GATILHO de 15 min. Rodar à mão também é seguro: só sai o que ainda não foi avisado.
function enviarAvisoEmergencial() {
  return aeExecutar_({});
}

// PARA RODAR PELO EDITOR antes de ligar: monta o aviso com os firmes emergenciais dos
// últimos AVISO_EMERG_TESTE_DIAS dias e manda SÓ para o EMAIL_ALERTA, com [TESTE] no
// assunto. Ignora o log e o marco e não registra o envio — rodar de novo manda o mesmo
// e-mail. (Só cria a aba Log_Email_Emergencial, vazia, se ela ainda não existir.)
function testarAvisoEmergencial() {
  return aeExecutar_({ teste: true });
}

// Liga o gatilho de 15 min e grava o marco: só entram firmes feitos daqui para a frente.
// Rodar de novo recria o gatilho e mantém o marco da primeira instalação.
function instalarAvisoEmergencial() {
  aeRemoverGatilhos_();
  ScriptApp.newTrigger("enviarAvisoEmergencial").timeBased().everyMinutes(15).create();

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_AVISO_EMERG_INICIO)) {
    props.setProperty(PROP_AVISO_EMERG_INICIO, String(Date.now()));
  }
  aeObterAbaLog_(SpreadsheetApp.openById(SPREADSHEET_ID));
  const inicio = new Date(Number(props.getProperty(PROP_AVISO_EMERG_INICIO)));
  console.log("Aviso emergencial ligado: a cada 15 min, para " + AVISO_EMERG_PARA +
              ". Entram firmes a partir de " + aeDiaMesAno_(inicio) + " " + aeHora_(inicio) + ".");
}

function desligarAvisoEmergencial() {
  const n = aeRemoverGatilhos_();
  console.log("Aviso emergencial desligado (" + n + " gatilho(s) removido(s)). O log e o marco ficam.");
}

function aeRemoverGatilhos_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "enviarAvisoEmergencial") { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

// `opc.agora` existe para o teste fora do Apps Script; o gatilho usa a hora real.
function aeExecutar_(opc) {
  const teste = !!opc.teste;

  // Trava de USUÁRIO, e não a do script: a do script é a do Salvar do portal e a do
  // sync, e o aviso não tem por que segurar um clique. Esta só impede dois avisos ao
  // mesmo tempo (o gatilho e uma execução à mão), que mandariam o mesmo firme duas vezes.
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10000)) {
    console.log("Aviso emergencial: outro envio em andamento — fica para o próximo gatilho.");
    return null;
  }
  try {
    const agora = opc.agora || new Date();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. O store diz quem foi firmado na janela e depois do corte. É o passo barato —
    //    e, na maior parte dos gatilhos, o último.
    const firmes = aeFirmesRecentes_(ss, aeDesde_(agora, teste));
    if (firmes.size === 0) {
      console.log("Aviso emergencial: nenhum firme depois do corte" +
                  (teste ? " nos últimos " + AVISO_EMERG_TESTE_DIAS + " dias." : " na janela."));
      return null;
    }

    // 2. O log tira o que já foi avisado: o mesmo firme, ou um firme novo que não mudou
    //    entrega nem quantidade. O teste mostra tudo da janela.
    const log = aeLerLog_(ss, agora);
    if (!teste) {
      firmes.forEach((f, chave) => {
        const ult = log.ultimo.get(chave);
        const jaEnviado = log.enviados.has(aeChaveFirme_(chave, f.firmadoEm));
        const semMudanca = ult && f.confData instanceof Date &&
                           aeIsoDe_(ult.data) === aeIso_(f.confData) && qtdSap_(ult.qtd) === qtdSap_(f.confQtd);
        if (jaEnviado || semMudanca) firmes.delete(chave);
      });
      if (firmes.size === 0) {
        console.log("Aviso emergencial: os firmes depois do corte já foram avisados.");
        return null;
      }
    }

    // 3. A página confirma que ainda é Firme e traz o resto da linha; a ME2W, a Plant.
    const linhas = aeMontarLinhas_(ss, firmes, log);
    if (linhas.length === 0) {
      console.log("Aviso emergencial: " + firmes.size + " firme(s) na fila, nenhum Firme na página agora.");
      return null;
    }

    // 4. Envia e SÓ ENTÃO registra. Falha no envio deixa tudo para o próximo gatilho;
    //    registrar antes perderia o aviso de vez.
    const ctx = { agora: agora, aviso: log.avisosHoje + 1, acumulado: log.stosHoje + linhas.length, teste: teste };
    const email = montarEmailEmergencial_(linhas, ctx);
    const para = teste ? EMAIL_ALERTA : AVISO_EMERG_PARA;
    if (MailApp.getRemainingDailyQuota() < para.split(",").length) {
      throw new Error("Cota diária de e-mail esgotada — " + linhas.length +
                      " STO(s) emergencial(is) aguardam a cota voltar.");
    }
    MailApp.sendEmail({
      to: para,
      subject: (teste ? "[TESTE] " : "") + email.assunto,
      htmlBody: email.html,
      body: email.texto,
      name: AVISO_EMERG_REMETENTE
    });
    if (!teste) aeGravarLog_(log.aba, linhas, ctx, para);

    console.log("Aviso emergencial nº " + ctx.aviso + (teste ? " (TESTE, só para " + para + ")" : "") +
                ": " + linhas.length + " STO(s) — " + linhas.map(l => l.chave + " " + l.tipo).join(" | "));
    return { aviso: ctx.aviso, linhas: linhas.length, assunto: email.assunto };
  } finally {
    lock.releaseLock();
  }
}

// O início da janela: o mais recente entre o marco da instalação e AVISO_EMERG_JANELA_HORAS
// atrás. A janela é folga para gatilho que falhou, não para despejar firme antigo.
function aeDesde_(agora, teste) {
  if (teste) return new Date(agora.getTime() - AVISO_EMERG_TESTE_DIAS * 86400000);
  const props = PropertiesService.getScriptProperties();
  let inicio = Number(props.getProperty(PROP_AVISO_EMERG_INICIO));
  // Gatilho criado à mão, sem o instalarAvisoEmergencial: o marco nasce agora.
  if (!(inicio > 0)) {
    inicio = agora.getTime();
    props.setProperty(PROP_AVISO_EMERG_INICIO, String(inicio));
  }
  return new Date(Math.max(inicio, agora.getTime() - AVISO_EMERG_JANELA_HORAS * 3600000));
}

// ====================================================================
// A REGRA
// ====================================================================

function aeDia_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

// 10:00 do dia anterior à entrega, em dias corridos (a segunda tem corte no domingo).
function aeCorte_(entrega) {
  return new Date(entrega.getFullYear(), entrega.getMonth(), entrega.getDate() - 1, AVISO_EMERG_HORA_CORTE, 0, 0);
}

// Firmado a partir do corte, com a entrega ainda de pé no dia do firme.
function aeEmergencial_(entrega, firmadoEm) {
  return firmadoEm >= aeCorte_(entrega) && aeDia_(firmadoEm) <= aeDia_(entrega);
}

// ====================================================================
// LEITURAS
// ====================================================================

// Do store, só o que decide: a chave e as manuais (A:G) e o Firmado em. Devolve
// Map chave -> { firmadoEm, confData, confQtd } das STOs com confirmação de pé,
// firmadas na janela e depois do corte pela data confirmada. A página confere de novo.
function aeFirmesRecentes_(ss, desde) {
  const aba = obterAbaStore(ss);
  const saida = new Map();
  const n = aba.getLastRow() - 1;
  if (n < 1) return saida;

  const base = aba.getRange(2, 1, n, ST_MANUAIS + 3).getValues();
  const firmados = aba.getRange(2, ST_FIRMADO_EM + 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    const firmadoEm = firmados[i][0];
    if (!(firmadoEm instanceof Date) || isNaN(firmadoEm.getTime()) || firmadoEm < desde) continue;
    const v = base[i];
    const chave = String(v[ST_CHAVE]).trim();
    if (!chave || !temConfirmacao_(v)) continue;
    const confData = v[ST_MANUAIS + 1];
    if (confData instanceof Date && !aeEmergencial_(confData, firmadoEm)) continue;
    saida.set(chave, { firmadoEm: firmadoEm, confData: confData, confQtd: v[ST_MANUAIS + 2] });
  }
  return saida;
}

function aeObterAbaLog_(ss) {
  let aba = ss.getSheetByName(ABA_LOG_EMERG);
  if (!aba) {
    aba = ss.insertSheet(ABA_LOG_EMERG);
    aba.getRange(1, 1, 1, LOG_EMERG_HEADERS.length).setValues([LOG_EMERG_HEADERS]);
    aba.setFrozenRows(1);
  }
  return aba;
}

// O log só recebe linhas no fim, então a última linha de cada chave é o último aviso
// dela. `avisosHoje` é o maior nº de aviso de hoje (o próximo é +1) e `stosHoje` as
// linhas enviadas hoje — a soma de todos os avisos do dia.
function aeLerLog_(ss, agora) {
  const aba = aeObterAbaLog_(ss);
  const out = { aba: aba, enviados: new Set(), ultimo: new Map(), avisosHoje: 0, stosHoje: 0 };
  const n = aba.getLastRow() - 1;
  if (n < 1) return out;

  const hoje = aeIso_(agora);
  aba.getRange(2, 1, n, LG_LARGURA_LIDA).getValues().forEach(l => {
    const chave = String(l[LG_CHAVE]).trim();
    if (!chave) return;
    out.enviados.add(aeChaveFirme_(chave, l[LG_FIRMADO]));
    out.ultimo.set(chave, { data: l[LG_DATA], qtd: l[LG_QTD], plant: String(l[LG_PLANT]).trim() });
    if (l[LG_ENVIADO] instanceof Date && aeIso_(l[LG_ENVIADO]) === hoje) {
      out.stosHoje++;
      out.avisosHoje = Math.max(out.avisosHoje, Number(l[LG_AVISO]) || 0);
    }
  });
  return out;
}

// Chave + Firmado em ao segundo. Os dois lados saem de leitura da planilha, mas o
// milissegundo não precisa sobreviver a ela para a comparação valer.
function aeChaveFirme_(chave, firmadoEm) {
  const t = firmadoEm instanceof Date ? Math.round(firmadoEm.getTime() / 1000) : String(firmadoEm);
  return chave + "@" + t;
}

// Plant da ME2W (o destino da STO), por chave. Null sem as colunas: aí vale a
// Planta_ Destino da página, com aviso no registro de execução.
function aeLerPlantMe2w_(ss) {
  const aba = ss.getSheetByName(ABA_ME2W);
  if (!aba || aba.getLastRow() < 2) return null;
  const n = aba.getLastRow() - 1;
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const idx = CHAVE_ME2W.concat([AVISO_EMERG_COL_PLANT]).map(c => headers.indexOf(c));
  if (idx.some(i => i === -1)) {
    console.warn("Aviso emergencial: ME2W sem a coluna '" + AVISO_EMERG_COL_PLANT +
                 "' ou sem as colunas-chave — destino pela Planta_ Destino da página.");
    return null;
  }

  // Pelo texto exibido, como o Salvar do portal: Date no lugar do número não casa a chave.
  const codigo = i => {
    const r = aba.getRange(2, i + 1, n, 1);
    const v = r.getValues(), e = r.getDisplayValues();
    return v.map((x, k) => normalizarNumeroDoc(x[0], e[k][0]));
  };
  const docs = codigo(idx[0]), itens = codigo(idx[1]), scheds = codigo(idx[2]);
  const plants = aba.getRange(2, idx[3] + 1, n, 1).getDisplayValues();
  const mapa = new Map();
  for (let k = 0; k < n; k++) {
    if (docs[k] !== "") mapa.set(montarChave(docs[k], itens[k], scheds[k]), String(plants[k][0]).trim());
  }
  return mapa;
}

// Pallet Order da página por número de linha (o rowIndex do getTransferData). Só número:
// o "Sem Estoque para Convesão" fica sem palete, e o e-mail diz quantas ficaram assim.
function aeLerPaletes_(ss) {
  const out = new Map();
  const aba = ss.getSheetByName(SHEET_NAME);
  if (!aba || aba.getLastRow() < 2) return out;
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf(AVISO_EMERG_COL_PALETES);
  if (col === -1) return out;
  aba.getRange(2, col + 1, aba.getLastRow() - 1, 1).getValues().forEach((x, k) => {
    if (typeof x[0] === "number" && isFinite(x[0])) out.set(k + 2, x[0]);
  });
  return out;
}

// As linhas do e-mail. Firme é o do portal (getTransferData), para "firmada" querer
// dizer aqui exatamente o que a tela mostra.
function aeMontarLinhas_(ss, firmes, log) {
  const pagina = getTransferData();
  const plants = aeLerPlantMe2w_(ss);
  const paletes = aeLerPaletes_(ss);
  const saida = [];

  pagina.forEach(r => {
    const chave = montarChave(r.doc, r.item, r.sched);
    const f = firmes.get(chave);
    if (!f || r.alerta !== "Firme") return;

    // O store e a página têm de contar a MESMA confirmação. Divergem por um instante
    // quando a leitura cai no meio de um Salvar ou de um sync — aí espera o próximo
    // gatilho, em vez de avisar com a linha velha.
    if ((f.confData instanceof Date && aeIso_(f.confData) !== r.confDateIso) ||
        qtdSap_(f.confQtd) !== qtdSap_(r.confQty)) {
      console.log("Aviso emergencial: " + chave + " com store e página divergentes — fica para o próximo gatilho.");
      return;
    }

    const entrega = aeDataDoIso_(r.dataSapIso);
    if (!entrega || !aeEmergencial_(entrega, f.firmadoEm)) return;

    const qtd = qtdSap_(r.qtySap);
    // O próprio firme já no log só chega aqui pelo teste, que não filtra o log: vai
    // como saiu da primeira vez, e não como atualização de si mesmo.
    const ult = log.enviados.has(aeChaveFirme_(chave, f.firmadoEm)) ? null : log.ultimo.get(chave);
    let antes = null;
    if (ult) {
      antes = {};
      if (aeIsoDe_(ult.data) !== r.dataSapIso && ult.data instanceof Date) antes.entrega = ult.data;
      if (qtdSap_(ult.qtd) !== qtd && qtdSap_(ult.qtd) !== null) antes.qtd = qtdSap_(ult.qtd);
    }

    saida.push({
      chave: chave,
      tipo: ult ? "Atualizada" : "Nova",
      antes: antes,
      entrega: entrega,
      corte: aeCorte_(entrega),
      plant: (plants && plants.get(chave)) || String(r.planta == null ? "" : r.planta).trim() || "Sem destino",
      doc: r.doc, item: r.item, sched: r.sched,
      material: String(r.matCod == null ? "" : r.matCod).trim(),
      descricao: String(r.matDesc == null ? "" : r.matDesc).trim(),
      qtd: qtd === null ? r.qtySap : qtd,
      unidade: String(r.unitSap == null ? "" : r.unitSap).trim(),
      paletes: paletes.has(r.rowIndex) ? paletes.get(r.rowIndex) : null,
      firmadoEm: f.firmadoEm,
      prioridade: r.prioridade,
      fluxo: r.statusFluxo
    });
  });
  return saida;
}

function aeGravarLog_(aba, linhas, ctx, para) {
  const dados = linhas.map(l => [
    ctx.agora, ctx.aviso, l.chave, l.firmadoEm, l.entrega, l.qtd, l.plant,
    l.tipo, l.doc, l.item, l.sched, l.material, l.descricao,
    l.unidade, l.paletes === null ? "" : l.paletes, l.corte, para
  ]);
  const primeira = aba.getLastRow() + 1;
  garantirGradeAba(aba, primeira + dados.length - 1, LOG_EMERG_HEADERS.length);
  aba.getRange(primeira, 1, dados.length, LOG_EMERG_HEADERS.length).setValues(dados);
}

// ====================================================================
// O E-MAIL — JS puro: recebe as linhas prontas e devolve assunto, HTML e texto
// ====================================================================

const AE_MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const AE_DIAS = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
const AE_FONTE = "font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;";

// Mesmas cores das faixas de prioridade do portal (Tema_SmartHub).
const AE_PRIORIDADE = {
  "0-Crítico": { txt: "Crítico", cor: "#b91c1c", bg: "#fef2f2", borda: "#fecaca" },
  "1-Urgente": { txt: "Urgente", cor: "#c2410c", bg: "#fff7ed", borda: "#fed7aa" },
  "2-Alto":    { txt: "Alto",    cor: "#b45309", bg: "#fffbeb", borda: "#fde68a" },
  "3-Normal":  { txt: "Normal",  cor: "#1d4ed8", bg: "#eff6ff", borda: "#bfdbfe" },
  "4-Baixo":   { txt: "Baixo",   cor: "#475569", bg: "#f1f5f9", borda: "#e2e8f0" }
};

function aeP2_(n) { return String(n).padStart(2, "0"); }
function aeIso_(d) { return d.getFullYear() + "-" + aeP2_(d.getMonth() + 1) + "-" + aeP2_(d.getDate()); }
function aeIsoDe_(v) { return (v instanceof Date && !isNaN(v.getTime())) ? aeIso_(v) : ""; }
function aeDataDoIso_(s) {
  const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}
function aeDiaMes_(d) { return aeP2_(d.getDate()) + "/" + AE_MESES[d.getMonth()]; }
function aeDiaMesAno_(d) { return aeDiaMes_(d) + "/" + d.getFullYear(); }
function aeHora_(d) { return aeP2_(d.getHours()) + ":" + aeP2_(d.getMinutes()); }
function aeEsc_(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
                                   .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// 1440 -> "1.440"; 5.43 -> "5,43". Sem Intl: o toLocaleString do Apps Script não é confiável.
function aeNum_(n) {
  const v = Number(n);
  if (n === null || n === "" || !isFinite(v)) return String(n == null ? "" : n);
  const partes = Math.abs(Math.round(v * 100) / 100).toFixed(2).split(".");
  const dec = partes[1].replace(/0+$/, "");
  return (v < 0 ? "-" : "") + partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".") + (dec ? "," + dec : "");
}

function aeAposCorte_(firmadoEm, corte) {
  const min = Math.max(0, Math.round((firmadoEm - corte) / 60000));
  const h = Math.floor(min / 60), m = min % 60;
  return (h > 0 ? h + "h" + aeP2_(m) : m + " min") + " após o corte";
}

function aePill_(txt, cor, bg, borda) {
  return '<span style="display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;' +
         'font-weight:700;line-height:16px;white-space:nowrap;color:' + cor + ';background:' + bg +
         ';border:1px solid ' + borda + ';">' + aeEsc_(txt) + '</span>';
}

function aePlural_(n, um, varios) { return n + " " + (n === 1 ? um : varios); }

// linhas: as de aeMontarLinhas_. ctx: { agora, aviso, acumulado, teste }.
function montarEmailEmergencial_(linhas, ctx) {
  const agora = ctx.agora;
  const hoje = aeDia_(agora);
  const amanha = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + 1);
  const ontem = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 1);
  const ehDia = (d, ref) => aeIso_(d) === aeIso_(ref);

  // Um bloco por destino; dentro, a entrega mais próxima primeiro.
  const grupos = new Map();
  linhas.slice()
    .sort((a, b) => String(a.plant).localeCompare(String(b.plant)) || a.entrega - b.entrega ||
                    String(a.prioridade).localeCompare(String(b.prioridade)) || a.firmadoEm - b.firmadoEm)
    .forEach(l => {
      if (!grupos.has(l.plant)) grupos.set(l.plant, []);
      grupos.get(l.plant).push(l);
    });

  // OS NÚMEROS DO TOPO: todos contados SÓ nas linhas deste e-mail.
  const total = linhas.length;
  const atualizadas = linhas.filter(l => l.tipo === "Atualizada").length;
  const novas = total - atualizadas;
  const nHoje = linhas.filter(l => ehDia(l.entrega, hoje)).length;
  const nAmanha = linhas.filter(l => ehDia(l.entrega, amanha)).length;
  const comPalete = linhas.filter(l => typeof l.paletes === "number");
  const paletes = comPalete.reduce((s, l) => s + l.paletes, 0);
  const semPalete = total - comPalete.length;

  const assunto = "⚠️ Transferências entre Plantas Emergenciais · Aviso " + ctx.aviso + " de " +
                  aeDiaMes_(agora) + " · " + aePlural_(total, "STO", "STOs");

  const kpi = (rot, val, sub, cor, ultimo) =>
    '<td width="25%" valign="top" style="padding:14px 8px 12px;text-align:center;' +
    (ultimo ? '' : 'border-right:1px solid #e3e8ef;') + '">' +
      '<div style="' + AE_FONTE + 'font-size:24px;font-weight:800;line-height:28px;color:' + cor + ';">' + val + '</div>' +
      '<div style="' + AE_FONTE + 'font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#64748b;margin-top:4px;">' + rot + '</div>' +
      '<div style="' + AE_FONTE + 'font-size:11px;color:#8b97a8;margin-top:2px;">' + sub + '</div>' +
    '</td>';

  // Larguras fixas: cada destino é uma tabela própria, e sem elas as colunas de um bloco
  // não alinham com as do outro. São de conteúdo (o padding de 20px soma por fora) e
  // cabem o texto que não quebra linha; o Material fica com o que sobra.
  const th = (txt, largura, alinhar) =>
    '<th' + (largura ? ' width="' + largura + '"' : '') + ' style="' + AE_FONTE + 'padding:9px 10px;font-size:10px;' +
    'font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#dbe3ee;text-align:' + (alinhar || 'left') +
    ';background:#1e293b;' + (largura ? 'width:' + largura + 'px;' : '') + '">' + txt + '</th>';

  const td = (conteudo, extra) =>
    '<td valign="top" style="' + AE_FONTE + 'padding:10px;font-size:13px;color:#0f172a;border-top:1px solid #e3e8ef;' +
    (extra || '') + '">' + conteudo + '</td>';

  const seloEntrega = d => ehDia(d, hoje) ? aePill_("Hoje", "#ffffff", "#b91c1c", "#b91c1c")
                         : ehDia(d, amanha) ? aePill_("Amanhã", "#ffffff", "#c2410c", "#c2410c")
                         : ehDia(d, ontem) ? aePill_("Ontem", "#ffffff", "#475569", "#475569")
                         : aePill_(AE_DIAS[d.getDay()], "#0f172a", "#f1f5f9", "#cdd5e0");

  let blocos = "";
  grupos.forEach((itens, plant) => {
    const palPlant = itens.reduce((s, l) => s + (typeof l.paletes === "number" ? l.paletes : 0), 0);
    let corpo = "";
    itens.forEach((l, k) => {
      const pri = AE_PRIORIDADE[l.prioridade] || AE_PRIORIDADE["3-Normal"];
      const mudou = [];
      if (l.antes && l.antes.entrega) mudou.push("entrega " + aeDiaMes_(l.antes.entrega));
      if (l.antes && l.antes.qtd !== undefined) mudou.push(aeNum_(l.antes.qtd) + " " + aeEsc_(l.unidade));
      const fluxo = String(l.fluxo == null ? "" : l.fluxo).trim();

      corpo += '<tr style="background:' + (k % 2 ? '#f7f9fb' : '#ffffff') + ';">' +
        td(seloEntrega(l.entrega) +
           '<div style="font-size:11px;color:#4b5769;margin-top:4px;white-space:nowrap;">' +
           AE_DIAS[l.entrega.getDay()] + ' ' + aeDiaMes_(l.entrega) + '</div>') +
        td('<div style="font-weight:700;white-space:nowrap;">' + aeEsc_(l.doc) + '</div>' +
           '<div style="font-size:11px;color:#8b97a8;white-space:nowrap;">Item ' + aeEsc_(l.item) +
           ' · SL ' + aeEsc_(l.sched) + '</div>' +
           (l.tipo === "Atualizada"
             ? '<div style="margin-top:4px;">' + aePill_("Atualizada", "#6d28d9", "#f5f3ff", "#ddd6fe") + '</div>' : '')) +
        td('<div style="font-weight:700;">' + aeEsc_(l.material) + '</div>' +
           '<div style="font-size:11px;color:#4b5769;line-height:15px;">' + aeEsc_(l.descricao) + '</div>' +
           (mudou.length ? '<div style="font-size:11px;color:#6d28d9;margin-top:3px;">Antes: ' + mudou.join(' · ') + '</div>' : '')) +
        td('<div style="font-weight:700;white-space:nowrap;">' + aeNum_(l.qtd) + ' ' + aeEsc_(l.unidade) + '</div>' +
           (typeof l.paletes === "number"
             ? '<div style="font-size:11px;color:#8b97a8;white-space:nowrap;">≈ ' + aeNum_(l.paletes) + ' pal.</div>' : ''),
           'text-align:right;') +
        td('<div style="font-weight:700;white-space:nowrap;">' + aeDiaMes_(l.firmadoEm) + ' ' + aeHora_(l.firmadoEm) + '</div>' +
           '<div style="font-size:11px;color:#b91c1c;white-space:nowrap;">' + aeAposCorte_(l.firmadoEm, l.corte) + '</div>') +
        td(aePill_(pri.txt, pri.cor, pri.bg, pri.borda) +
           (fluxo && fluxo !== "-"
             ? '<div style="font-size:11px;color:#8b97a8;margin-top:4px;">' + aeEsc_(fluxo) + '</div>' : '')) +
      '</tr>';
    });

    blocos +=
      '<tr><td style="padding:22px 24px 8px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
          '<td style="' + AE_FONTE + 'font-size:15px;font-weight:800;color:#0f172a;">' +
            '<span style="font-size:10px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#64748b;">Destino</span>&nbsp; ' +
            aePill_(plant, "#1d4ed8", "#eff6ff", "#bfdbfe") + '</td>' +
          '<td style="' + AE_FONTE + 'font-size:12px;color:#64748b;text-align:right;white-space:nowrap;">' +
            aePlural_(itens.length, "STO", "STOs") + (palPlant ? ' · ≈ ' + aeNum_(palPlant) + ' pal.' : '') + '</td>' +
        '</tr></table>' +
      '</td></tr>' +
      '<tr><td style="padding:0 24px;">' +
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid #e3e8ef;">' +
          '<tr>' + th('Entrega', 64) + th('STO', 92) + th('Material') + th('Quantidade', 72, 'right') +
                   th('Firmado em', 104) + th('Prioridade', 92) + '</tr>' +
          corpo +
        '</table>' +
      '</td></tr>';
  });

  const html =
'<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">' +
'<meta name="viewport" content="width=device-width,initial-scale=1"><title>' + aeEsc_(assunto) + '</title></head>' +
'<body style="margin:0;padding:0;background:#eef2f7;">' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#eef2f7"><tr><td align="center" style="padding:24px 12px;">' +
'<table role="presentation" width="760" cellpadding="0" cellspacing="0" border="0" style="width:760px;max-width:100%;background:#ffffff;border:1px solid #e3e8ef;">' +

  (ctx.teste
    ? '<tr><td style="' + AE_FONTE + 'background:#fffbeb;border-bottom:1px solid #fde68a;padding:10px 24px;font-size:12px;color:#92400e;">' +
      '<b>TESTE</b> — enviado só para você. Nada foi registrado no log e o armazém não recebeu.</td></tr>'
    : '') +

  '<tr><td height="5" style="background:#dc2626;font-size:0;line-height:0;">&nbsp;</td></tr>' +
  '<tr><td bgcolor="#0f172a" style="background:#0f172a;padding:22px 24px 20px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td style="' + AE_FONTE + 'font-size:10px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#94a3b8;">Smart Hub · Portal de Transferência de Material</td>' +
      '<td style="' + AE_FONTE + 'font-size:11px;font-weight:800;color:#fca5a5;text-align:right;white-space:nowrap;">AVISO Nº ' + ctx.aviso + ' DE HOJE</td>' +
    '</tr></table>' +
    '<div style="' + AE_FONTE + 'font-size:22px;font-weight:800;line-height:28px;color:#ffffff;margin-top:6px;">Transferências entre Plantas Emergenciais</div>' +
    '<div style="' + AE_FONTE + 'font-size:13px;color:#cbd5e1;margin-top:4px;">STOs firmadas depois das 10:00 do dia anterior à entrega · ' +
      AE_DIAS[agora.getDay()] + ' ' + aeDiaMesAno_(agora) + ' às ' + aeHora_(agora) + '</div>' +
  '</td></tr>' +

  '<tr><td style="padding:20px 24px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td width="4" style="background:#b91c1c;font-size:0;">&nbsp;</td>' +
      '<td style="' + AE_FONTE + 'background:#fef2f2;padding:12px 16px;font-size:14px;line-height:21px;color:#7f1d1d;">' +
        '<b>' + aePlural_(total, "STO", "STOs") + '</b> ' +
        (total === 1 ? 'foi firmada' : 'foram firmadas') + ' pelo Planejamento <b>depois do corte</b> e ' +
        (total === 1 ? 'ficou' : 'ficaram') + ' fora da programação normal do armazém. ' +
        'Favor priorizar a separação e a expedição.' +
      '</td>' +
    '</tr></table>' +
  '</td></tr>' +

  '<tr><td style="padding:16px 24px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e3e8ef;background:#f7f9fb;"><tr>' +
      kpi('STOs neste aviso', total,
          atualizadas ? aePlural_(novas, "nova", "novas") + ' · ' + aePlural_(atualizadas, "atualizada", "atualizadas")
                      : aePlural_(novas, "nova", "novas"), '#0f172a') +
      kpi('Ainda p/ hoje', nHoje, 'em aberto no envio', '#b91c1c') +
      kpi('Para amanhã', nAmanha, 'entrega ' + aeDiaMes_(amanha), '#c2410c') +
      kpi('Paletes (aprox.)', aeNum_(paletes), semPalete ? semPalete + ' sem conversão' : 'neste aviso', '#0f172a', true) +
    '</tr></table>' +
    '<div style="' + AE_FONTE + 'font-size:12px;color:#64748b;margin-top:8px;">' +
      'Acumulado de hoje: <b style="color:#334155;">' + aePlural_(ctx.acumulado, "STO", "STOs") + ' em ' +
      aePlural_(ctx.aviso, "aviso", "avisos") + '</b> (com este).</div>' +
  '</td></tr>' +

  blocos +

  '<tr><td style="padding:24px 24px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f7f9fb;border:1px solid #e3e8ef;"><tr>' +
      '<td style="' + AE_FONTE + 'padding:12px 16px;font-size:12px;line-height:18px;color:#4b5769;">' +
        '<b style="color:#0f172a;">Como esta lista é montada:</b> entra a STO firmada no portal a partir das ' +
        '<b>10:00 do dia anterior à entrega</b> (dias corridos), inclusive no próprio dia da entrega. ' +
        'Cada firme é avisado uma vez só; se a STO for firmada de novo com outra data ou quantidade, volta como ' +
        '<b>Atualizada</b>.<br><br>' +
        '<b style="color:#0f172a;">Os números do topo</b> contam só as STOs deste e-mail. ' +
        '<b>Ainda p/ hoje</b> são as de entrega hoje que estavam firmes e em aberto no envio — as já atendidas ' +
        'e as avisadas antes não entram, então não é o total de entregas do dia. ' +
        'O acumulado soma todos os avisos de hoje; os avisos são numerados, e um número pulado é um e-mail que não chegou.' +
      '</td>' +
    '</tr></table>' +
  '</td></tr>' +
  '<tr><td style="' + AE_FONTE + 'padding:16px 24px 22px;font-size:11px;line-height:16px;color:#8b97a8;">' +
    'E-mail automático do Portal de Transferência de Material — Smart Hub. Dúvidas: responda este e-mail.' +
  '</td></tr>' +

'</table>' +
'</td></tr></table>' +
'</body></html>';

  // Texto puro, para o cliente que não mostra HTML.
  const texto = [
    "TRANSFERÊNCIAS ENTRE PLANTAS EMERGENCIAIS — AVISO Nº " + ctx.aviso + " DE HOJE",
    aePlural_(total, "STO firmada", "STOs firmadas") + " depois das 10:00 do dia anterior à entrega.",
    "Ainda p/ hoje: " + nHoje + " | Para amanhã: " + nAmanha + " | Paletes (aprox.): " + aeNum_(paletes) +
      " | Acumulado de hoje: " + aePlural_(ctx.acumulado, "STO", "STOs") + " em " + aePlural_(ctx.aviso, "aviso", "avisos"),
    ""
  ].concat(Array.from(grupos.entries()).map(([plant, itens]) =>
    "Destino " + plant + "\n" + itens.map(l =>
      "  - Entrega " + AE_DIAS[l.entrega.getDay()] + " " + aeDiaMes_(l.entrega) +
      " | STO " + l.doc + "/" + l.item + "/" + l.sched + (l.tipo === "Atualizada" ? " (Atualizada)" : "") +
      " | " + l.material + " " + l.descricao + " | " + aeNum_(l.qtd) + " " + l.unidade +
      " | firmado " + aeDiaMes_(l.firmadoEm) + " " + aeHora_(l.firmadoEm)).join("\n")))
  .join("\n");

  return { assunto: assunto, html: html, texto: texto };
}
