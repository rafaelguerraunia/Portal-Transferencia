// ====================================================================
// E-MAIL DAS STOs FIRMADAS — informativo aos planejadores das linhas
// ====================================================================
//
// Lista as STOs com Status Planejamento "Firme" (o mesmo do portal), com a data e a
// hora da última atualização da database no topo. Uma tabela por destino (Plant da
// ME2W): STO/Item/Schedule Line, quem firmou, Material, Qtd, Entrega, Entrou no
// sistema em e Firmado em. Fast Track vai com badge e linha rosada.
//
// O horário (16h) é um acionador criado no editor: Acionadores > Adicionar >
// enviarEmailFirmadas > Baseado no tempo > Diário > 16h às 17h.
//
// Reaproveita do EmailEmergencial.gs os helpers ae* (datas, números, escape) e o
// aeLerPlantMe2w_; as imagens da marca estão em EmailFirmadas_Imagens.gs.

const EMAIL_FIRMADAS_PARA = "bp_rafael_guerra@colpal.com";
const EMAIL_FIRMADAS_REMETENTE = "Smart Hub · Portal de Transferência";

function enviarEmailFirmadas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const linhas = efMontarLinhas_(ss);
  if (linhas.length === 0) {
    console.log("E-mail das firmadas: nenhuma STO Firme agora — nada enviado.");
    return null;
  }

  const base = staBaseDados();
  const email = montarEmailFirmadas_(linhas, {
    agora: new Date(),
    base: base ? new Date(base.iso) : null,
    logo: "cid:logo", canto: "cid:canto"
  });
  MailApp.sendEmail({
    to: EMAIL_FIRMADAS_PARA,
    subject: email.assunto,
    htmlBody: email.html,
    body: email.texto,
    name: EMAIL_FIRMADAS_REMETENTE,
    inlineImages: {
      logo: Utilities.newBlob(Utilities.base64Decode(EF_IMG_LOGO), "image/png", "logo.png"),
      canto: Utilities.newBlob(Utilities.base64Decode(EF_IMG_CANTO), "image/png", "canto.png")
    }
  });
  console.log("E-mail das firmadas: " + linhas.length + " STO(s) para " + EMAIL_FIRMADAS_PARA + ".");
  return { linhas: linhas.length, assunto: email.assunto };
}

// Do store, por chave: Entrou no sistema em, Firmado em e Firmado por (P:R).
function efLerRastreio_(ss) {
  const aba = obterAbaStore(ss);
  const mapa = new Map();
  const n = aba.getLastRow() - 1;
  if (n < 1) return mapa;
  const chaves = aba.getRange(2, ST_CHAVE + 1, n, 1).getValues();
  const rastreio = aba.getRange(2, ST_ENTROU + 1, n, ST_FIRMADO_POR - ST_ENTROU + 1).getValues();
  const data = v => (v instanceof Date && !isNaN(v.getTime())) ? v : null;
  for (let i = 0; i < n; i++) {
    const chave = String(chaves[i][0]).trim();
    if (!chave) continue;
    const por = String(rastreio[i][ST_FIRMADO_POR - ST_ENTROU] || "").trim();
    mapa.set(chave, {
      entrou: data(rastreio[i][0]),
      firmadoEm: data(rastreio[i][ST_FIRMADO_EM - ST_ENTROU]),
      firmadoPor: por.indexOf("@") > 0 ? por : ""   // "portal" = sem e-mail de quem clicou
    });
  }
  return mapa;
}

// Fast Track: a causa do portal (automática ou digitada) ou, se a causa foi
// preenchida à mão com outra coisa, a mesma regra do portal — entrega até 1 dia
// da criação da STO.
function efFastTrack_(r) {
  if (/fast\s*track/i.test(String(r.causaDesvio || ""))) return true;
  const doc = aeDataDoIso_(r.docDateIso), ent = aeDataDoIso_(r.dataSapIso);
  return !!(doc && ent && Math.abs(ent - doc) <= 86400000);
}

function efMontarLinhas_(ss) {
  const rastreio = efLerRastreio_(ss);
  const plants = aeLerPlantMe2w_(ss);
  const saida = [];
  getTransferData().forEach(r => {
    if (r.alerta !== "Firme") return;
    const entrega = aeDataDoIso_(r.dataSapIso);
    if (!entrega) return;
    const chave = montarChave(r.doc, r.item, r.sched);
    const t = rastreio.get(chave) || {};
    const qtd = qtdSap_(r.qtySap);
    saida.push({
      plant: (plants && plants.get(chave)) || String(r.planta == null ? "" : r.planta).trim() || "Sem destino",
      doc: r.doc, item: r.item, sched: r.sched,
      material: String(r.matCod == null ? "" : r.matCod).trim(),
      descricao: String(r.matDesc == null ? "" : r.matDesc).trim(),
      qtd: qtd === null ? r.qtySap : qtd,
      unidade: String(r.unitSap == null ? "" : r.unitSap).trim(),
      entrega: entrega,
      apareceu: t.entrou || null,
      firmadoEm: t.firmadoEm || null,
      firmadoPor: t.firmadoPor || "",
      fastTrack: efFastTrack_(r)
    });
  });
  return saida;
}

// ====================================================================
// O E-MAIL — JS puro: recebe as linhas prontas e devolve assunto, HTML e texto
// ====================================================================

const FM = { navy: "#0F2B41", azul: "#1378AC", agua: "#52BDB7", fundo: "#f2f6f9", linha: "#d6e2ea",
             texto: "#0F2B41", texto2: "#34516a", apagado: "#5b7489", destaqueFundo: "#e2f0f7",
             ft: "#be185d", ftFundo: "#fdf0f6" }; // Fast Track: cor de status, fora da paleta da marca
const FM_FONTE = "font-family:Roboto,'Segoe UI',Arial,sans-serif;";

// ctx: { agora, base (Date|null), logo, canto } — logo/canto: src das imagens (cid: no envio).
function montarEmailFirmadas_(linhas, ctx) {
  const agora = ctx.agora;
  const hoje = aeIso_(agora);
  const amanha = aeIso_(new Date(agora.getFullYear(), agora.getMonth(), agora.getDate() + 1));
  const dh = d => d ? aeDiaMes_(d) + " " + aeHora_(d) : "—";

  const grupos = new Map();
  linhas.slice().sort((a, b) => String(a.plant).localeCompare(String(b.plant)) || a.entrega - b.entrega)
    .forEach(l => { if (!grupos.has(l.plant)) grupos.set(l.plant, []); grupos.get(l.plant).push(l); });

  const total = linhas.length;
  const nFast = linhas.filter(l => l.fastTrack).length;
  const proxima = linhas.reduce((m, l) => (!m || l.entrega < m) ? l.entrega : m, null);
  const base = ctx.base ? AE_DIAS[ctx.base.getDay()] + ", " + aeDiaMesAno_(ctx.base) + " às " + aeHora_(ctx.base) : "indisponível";
  const assunto = "Smart Hub · BR14 · " + aePlural_(total, "STO firmada", "STOs firmadas") +
                  (nFast ? " · " + nFast + " Fast Track" : "") +
                  " · base de " + (ctx.base ? aeDiaMes_(ctx.base) + " " + aeHora_(ctx.base) : "—");

  const pill = (txt, cor, bg) => '<span style="display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;' +
    'font-weight:700;line-height:16px;white-space:nowrap;color:' + cor + ';background:' + bg + ';">' + aeEsc_(txt) + '</span>';
  const selo = d => aeIso_(d) === hoje ? pill("Hoje", "#fff", FM.navy)
                  : aeIso_(d) === amanha ? pill("Amanhã", FM.navy, FM.agua)
                  : pill(AE_DIAS[d.getDay()], FM.azul, FM.destaqueFundo);

  const th = (t, w, al) => '<th' + (w ? ' width="' + w + '"' : '') + ' style="' + FM_FONTE + 'padding:9px 10px;font-size:10px;' +
    'font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:' + FM.agua + ';background:' + FM.navy +
    ';text-align:' + (al || 'left') + ';">' + t + '</th>';
  const td = (c, extra) => '<td valign="top" style="' + FM_FONTE + 'padding:10px;font-size:13px;color:' + FM.texto +
    ';border-top:1px solid ' + FM.linha + ';' + (extra || '') + '">' + c + '</td>';
  const sub = t => '<div style="font-size:11px;color:' + FM.apagado + ';white-space:nowrap;margin-top:2px;">' + t + '</div>';

  let blocos = "";
  grupos.forEach((itens, plant) => {
    const corpo = itens.map((l, k) =>
      '<tr style="background:' + (l.fastTrack ? FM.ftFundo : k % 2 ? '#f7fafc' : '#ffffff') + ';">' +
        td('<div style="font-weight:700;white-space:nowrap;">' + aeEsc_(l.doc) + '</div>' +
           sub('Item ' + aeEsc_(l.item) + ' · SL ' + aeEsc_(l.sched)) +
           (l.fastTrack ? '<div style="margin-top:5px;">' + pill('⚡ Fast Track', '#fff', FM.ft) + '</div>' : '') +
           (l.firmadoPor ? '<div style="margin-top:5px;font-size:11px;line-height:14px;white-space:nowrap;">' +
             '<a href="mailto:' + aeEsc_(l.firmadoPor) + '" style="color:' + FM.azul + ';text-decoration:none;">&#9993; ' +
             aeEsc_(l.firmadoPor) + '</a></div>' : ''),
           l.fastTrack ? 'border-left:4px solid ' + FM.ft + ';' : '') +
        td('<div style="font-weight:700;">' + aeEsc_(l.material) + '</div>' +
           '<div style="font-size:11px;line-height:15px;color:' + FM.texto2 + ';">' + aeEsc_(l.descricao) + '</div>') +
        td('<div style="font-weight:700;white-space:nowrap;">' + aeNum_(l.qtd) + '</div>' + sub(aeEsc_(l.unidade)), 'text-align:right;') +
        td(selo(l.entrega) + sub(aeDiaMes_(l.entrega))) +
        td('<div style="white-space:nowrap;color:' + FM.texto2 + ';">' + dh(l.apareceu) + '</div>') +
        td('<div style="font-weight:700;white-space:nowrap;color:' + (l.fastTrack ? FM.ft : FM.azul) + ';">' + dh(l.firmadoEm) + '</div>') +
      '</tr>').join("");
    const fastPlant = itens.filter(l => l.fastTrack).length;

    blocos +=
      '<tr><td class="sh-pad" style="padding:24px 28px 8px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
          '<td style="' + FM_FONTE + 'font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:' + FM.apagado + ';">' +
            'BR14 &#10140; <span style="display:inline-block;padding:3px 10px;margin-left:4px;border-radius:4px;font-size:13px;letter-spacing:.04em;' +
            'color:#fff;background:' + FM.azul + ';">' + aeEsc_(plant) + '</span></td>' +
          '<td style="' + FM_FONTE + 'font-size:12px;color:' + FM.apagado + ';text-align:right;white-space:nowrap;">' +
            aePlural_(itens.length, "STO", "STOs") +
            (fastPlant ? ' · <b style="color:' + FM.ft + ';">' + fastPlant + ' Fast Track</b>' : '') + '</td>' +
        '</tr></table></td></tr>' +
      '<tr><td class="sh-pad" style="padding:0 28px;">' +
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid ' + FM.linha + ';">' +
          '<tr>' + th('STO', 180) + th('Material') + th('Qtd', 64, 'right') + th('Entrega', 64) +
                   th('Entrou no sistema', 86) + th('Firmada em', 86) + '</tr>' + corpo +
        '</table></td></tr>';
  });

  const kpi = (val, rot, ultimo) =>
    '<td width="25%" valign="top" style="padding:14px 8px;text-align:center;' + (ultimo ? '' : 'border-right:1px solid ' + FM.linha + ';') + '">' +
      '<div style="' + FM_FONTE + 'font-size:26px;font-weight:700;line-height:30px;color:' + FM.navy + ';">' + val + '</div>' +
      '<div style="' + FM_FONTE + 'font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:' + FM.apagado + ';margin-top:4px;">' + rot + '</div></td>';

  // Animação só onde o cliente aceita (Apple Mail, Outlook web/novo); o Gmail ignora e fica o estático.
  const estilo = '<style>' +
    '@keyframes shPulso{0%{box-shadow:0 0 0 0 rgba(82,189,183,.7)}70%{box-shadow:0 0 0 9px rgba(82,189,183,0)}100%{box-shadow:0 0 0 0 rgba(82,189,183,0)}}' +
    '@keyframes shVarre{0%{background-position:-200px 0}100%{background-position:800px 0}}' +
    '@keyframes shSobe{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}' +
    '.sh-pulso{animation:shPulso 2s infinite}' +
    '.sh-varre{background-image:linear-gradient(90deg,' + FM.azul + ' 0,' + FM.agua + ' 50%,' + FM.azul + ' 100%)!important;background-size:200px 100%;background-repeat:no-repeat;animation:shVarre 3s linear infinite}' +
    '.sh-sobe{animation:shSobe .7s cubic-bezier(.2,.8,.2,1) both}' +
    '@media (max-width:620px){.sh-arte{display:none!important}.sh-pad{padding-left:16px!important;padding-right:16px!important}}' +
    '</style>';

  const html =
'<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
'<title>' + aeEsc_(assunto) + '</title>' + estilo + '</head>' +
'<body style="margin:0;padding:0;background:' + FM.fundo + ';">' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + FM.fundo + '"><tr><td align="center" style="padding:24px 12px;">' +
'<table role="presentation" width="800" cellpadding="0" cellspacing="0" border="0" style="width:800px;max-width:100%;background:#fff;border-radius:10px;overflow:hidden;border:1px solid ' + FM.linha + ';">' +

  // Faixa do topo: logo, sobretítulo verde-água, título branco e grafismo no canto direito.
  '<tr><td bgcolor="' + FM.navy + '" style="background:' + FM.navy + ';">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td class="sh-pad" valign="middle" style="padding:26px 28px 24px;">' +
        '<img src="' + ctx.logo + '" width="96" alt="Smart Supply Hub" style="display:block;width:96px;height:auto;border:0;">' +
        '<div style="' + FM_FONTE + 'font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:' + FM.agua + ';margin-top:18px;">Transferências · Armazém BR14</div>' +
        '<div style="' + FM_FONTE + 'font-size:26px;font-weight:700;line-height:32px;color:#fff;margin-top:6px;">STOs firmadas</div>' +
        '<div style="' + FM_FONTE + 'font-size:13px;color:#b3c7d4;margin-top:6px;">O que o BR14 vai enviar para a sua linha, com a data confirmada.</div>' +
      '</td>' +
      '<td class="sh-arte" width="170" valign="top" align="right" style="width:170px;">' +
        '<img src="' + ctx.canto + '" width="170" alt="" style="display:block;width:170px;height:auto;border:0;"></td>' +
    '</tr></table></td></tr>' +
  '<tr><td class="sh-varre" height="4" style="background:' + FM.azul + ';font-size:0;line-height:0;">&nbsp;</td></tr>' +

  // A data da base: a primeira coisa depois da faixa.
  '<tr><td class="sh-pad" style="padding:22px 28px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="sh-sobe" style="background:' + FM.destaqueFundo + ';border-radius:8px;"><tr>' +
      '<td width="4" style="background:' + FM.agua + ';font-size:0;border-radius:8px 0 0 8px;">&nbsp;</td>' +
      '<td style="padding:14px 18px;">' +
        '<div style="' + FM_FONTE + 'font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:' + FM.azul + ';">' +
          '<span class="sh-pulso" style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + FM.agua + ';vertical-align:middle;margin-right:8px;"></span>' +
          'Database atualizada em</div>' +
        '<div style="' + FM_FONTE + 'font-size:20px;font-weight:700;color:' + FM.navy + ';margin-top:4px;">' + base + '</div>' +
      '</td>' +
      '<td align="right" style="' + FM_FONTE + 'padding:14px 18px;font-size:11px;color:' + FM.apagado + ';white-space:nowrap;">Enviado ' + dh(agora) + '</td>' +
    '</tr></table></td></tr>' +

  '<tr><td class="sh-pad" style="padding:16px 28px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid ' + FM.linha + ';border-radius:8px;"><tr>' +
      kpi(total, 'STOs firmadas') + kpi(grupos.size, grupos.size === 1 ? 'Destino' : 'Destinos') +
      kpi('<span style="color:' + (nFast ? FM.ft : FM.navy) + ';">' + nFast + '</span>', 'Fast Track') +
      kpi(proxima ? aeDiaMes_(proxima) : '—', 'Entrega mais próxima', true) +
    '</tr></table></td></tr>' +

  blocos +

  '<tr><td class="sh-pad" style="' + FM_FONTE + 'padding:24px 28px 0;font-size:12px;line-height:18px;color:' + FM.texto2 + ';">' +
    '<b style="color:' + FM.navy + ';">Entrou no sistema</b> é quando a STO apareceu na base do SAP; ' +
    '<b style="color:' + FM.navy + ';">Firmada em</b> é quando data e quantidade foram confirmadas no portal, por quem aparece no &#9993; abaixo da STO. ' +
    '<b style="color:' + FM.ft + ';">Fast Track</b>: entrega até 1 dia depois da criação da STO.</td></tr>' +
  '<tr><td style="padding:22px 0 0;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td height="6" width="60%" style="background:' + FM.navy + ';font-size:0;">&nbsp;</td>' +
    '<td height="6" width="25%" style="background:' + FM.azul + ';font-size:0;">&nbsp;</td>' +
    '<td height="6" width="15%" style="background:' + FM.agua + ';font-size:0;">&nbsp;</td></tr></table></td></tr>' +
  '<tr><td bgcolor="' + FM.navy + '" class="sh-pad" style="' + FM_FONTE + 'background:' + FM.navy + ';padding:16px 28px;font-size:11px;line-height:16px;color:#87a1b3;">' +
    'E-mail automático do Portal de Transferência de Material · Colgate Smart Supply Hub. Dúvidas: responda este e-mail.</td></tr>' +
'</table></td></tr></table></body></html>';

  const texto = ["STOs FIRMADAS — ARMAZÉM BR14", "Database atualizada em " + base, ""]
    .concat(Array.from(grupos.entries()).map(([plant, itens]) => "Destino " + plant + "\n" + itens.map(l =>
      "  - STO " + l.doc + "/" + l.item + "/" + l.sched + (l.fastTrack ? " [FAST TRACK]" : "") + " | " +
      l.material + " " + l.descricao + " | " + aeNum_(l.qtd) + " " + l.unidade + " | entrega " + aeDiaMes_(l.entrega) +
      " | entrou " + dh(l.apareceu) + " | firmada " + dh(l.firmadoEm) + (l.firmadoPor ? " por " + l.firmadoPor : "")).join("\n")))
    .join("\n");

  return { assunto: assunto, html: html, texto: texto };
}
