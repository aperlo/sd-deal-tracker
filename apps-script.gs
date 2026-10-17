/**
 * ============================================================
 *  RICEVITORE — Daily Deal Tracker · Sales Design
 *  Conferma immediata + digest supervisione 4x/giorno
 *  v5 — v4 INTATTO (vendite) + doGet API + tab Recruiting
 * ------------------------------------------------------------
 *  >>> La sezione "VENDITE (v4)" qui sotto NON va modificata:
 *      doPost, colonne Master, CODICI, email, digest, dashboard
 *      STATS restano identici al sistema di vendita in produzione.
 *  >>> La sezione "API + RECRUITING (v5)" in fondo e' ADDITIVA:
 *      aggiunge solo lettura dati (doGet) e i tab recruiting.
 * ============================================================
 */

/* ████████████████████████████████████████████████████████████
   ██  SEZIONE VENDITE (v4) — NON MODIFICARE                 ██
   ████████████████████████████████████████████████████████████ */

const SHEET_ID     = "11_RBtNlb9CeD1eWy5Q_meH571AnJ-TqtpzxYycM6zeM";  // foglio VENDITE (account dealtracker.hub)
const MITTENTE     = "Sales Design Deal Tracker";

const MAIL_DESK    = "dealtracker.hub@gmail.com";        // Sales Tracker Desk (mittente script)
const MAIL_DIRCOMM = "direzionecommerciale@sales-design.it";
const MAIL_TEST    = "dealtracker.hub@gmail.com";
const SUPERVISORI  = [MAIL_DESK, MAIL_DIRCOMM];          // digest: desk + direzione commerciale

// === MODO TEST EMAIL =========================================
// true  = TUTTE le email (conferme + digest) vanno SOLO a MAIL_TEST_DEST.
//         Niente invii a canali/agenzie/aree/supervisori reali. (fase di test)
// false = invii reali ai destinatari. (PRODUZIONE)
const MODO_TEST      = true;
const MAIL_TEST_DEST = "dealtracker.hub@gmail.com";     // dove arrivano le email in test

const COLORI = {
  dircomm:     "#FF941A",   // arancio  — Direzione Commerciale
  salesdesk:   "#6366F1",   // indaco   — Sales Tracker Desk
  agent:       "#2EC4A0",   // verde    — Agente / Area
  compilatore: "#60A5FA",   // azzurro  — Compilatore
  test:        "#4A3A1A"    // marrone  — Test banner
};

const CODICI = {
  "SD-DIRECT":              { canale: "Direct Sales",    tipo: "Direct",  gruppo: "master",  saluto: "nome" },
  "SD-AG-WECOMM":           { canale: "WeComm",          tipo: "Agenzia", gruppo: "agenzia", saluto: "team", mail: "weservice@sales-design.it" },
  "SD-AG-FIDITECH":         { canale: "Fiditech",        tipo: "Agenzia", gruppo: "agenzia", saluto: "team", mail: "fiditech@sales-design.it" },
  "SD-AG-PUNTOSEMPLICE":    { canale: "Punto Semplice",  tipo: "Agenzia", gruppo: "agenzia", saluto: "team", mail: "puntosemplice@sales-design.it" },
  "SD-AG-GRANDASOLUTIONS":  { canale: "Granda Solutions",tipo: "Agenzia", gruppo: "agenzia", saluto: "team", mail: "r.dalsasso02@sales-design.it" },
  "SD-AG-TEAMTUTINO":       { canale: "Team Tutino",     tipo: "Agenzia", gruppo: "agenzia", saluto: "team", mail: "s.tutino10@sales-design.it" },
  "SD_NORD":                { canale: "Area Nord",       tipo: "Area",    gruppo: "area",    saluto: "team", mail: "m.mantovelli06@sales-design.it" },
  "SD_CENTRO":              { canale: "Area Centro",     tipo: "Area",    gruppo: "area",    saluto: "team", mail: "marco.cima02@attitudeltd.com" },
  "SD_SUD":                 { canale: "Area Sud",        tipo: "Area",    gruppo: "area",    saluto: "team", mail: "c.massafra21@sales-design.it" },
};

const NOME_TAB = "Master";
const INTESTAZIONI = [
  "Timestamp","Data","Canale","Tipo","Agente","Ragione Sociale",
  "Partita IVA","Fatturato mese","Dispositivo","N dispositivi",
  "Pricing","Tariffa Pay","Tariffa Domestic","Provincia","Note","Email agente",
  "Area Manager"   // (P) 17a colonna ADDITIVA — competenza AM, non tocca le colonne 1-16
];

/* ============================================================
   (P) ATTRIBUZIONE AREA MANAGER — ADDITIVO
   I sales/agenzie SD confluiscono nella RETE DIRETTA SD (sotto il
   suo AM); i sales delle tre Aree restano di competenza dei tre
   AM ATD. I nomi devono combaciare con il roster del foglio
   Recruiting per far salire correttamente l'override (modello C).
   ============================================================ */
// Rete diretta SD (Direct Sales + Agenzie + AM): tutti sotto Andrea Perlo, AM Italia.
const SD_AM_DIRETTA = "Andrea Perlo";
// Aree ATD: ciascuna sotto il suo Area Manager.
const AREA_AM = {
  "Area Nord":   "Massimo Mantovelli",
  "Area Centro": "Marco Cima",
  "Area Sud":    "Claudia Massafra"
};
function areaManagerDi_(canale) {
  return canale.tipo === "Area" ? (AREA_AM[canale.canale] || "") : SD_AM_DIRETTA;
}

/* ============================================================
   ENTRY POINT
   ============================================================ */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    return elabora(d, false);
  } catch(err) {
    return rispondi({ ok: false, errore: "Dati non leggibili" });
  }
}

function elabora(d, modalitaTest) {
  const canale = CODICI[String(d.codice || "").trim()];
  if (!canale) return rispondi({ ok: false, errore: "Codice canale non valido" });

  if (!/^\d{11}$/.test(String(d.partitaIva || "")))
    return rispondi({ ok: false, errore: "Partita IVA non valida" });

  const obbligatori = ["data","agente","ragioneSociale","fatturatoMese",
    "tipoDispositivo","numeroDispositivi","pricing","tariffaPay","tariffaDomestic","provincia"];
  for (const c of obbligatori)
    if (!String(d[c] || "").trim()) return rispondi({ ok: false, errore: "Campo mancante: " + c });

  const adesso = new Date();
  const dataIt = String(d.data).split("-").reverse().join("/");
  const ts     = Utilities.formatDate(adesso, "Europe/Rome", "dd/MM/yyyy HH:mm:ss");

  apriTab().appendRow([
    ts, dataIt, canale.canale, canale.tipo,
    String(d.agente), String(d.ragioneSociale),
    "'" + String(d.partitaIva),
    Number(d.fatturatoMese),
    String(d.tipoDispositivo),
    Number(d.numeroDispositivi),
    String(d.pricing),
    String(d.tariffaPay),
    String(d.tariffaDomestic),
    String(d.provincia),
    String(d.note || ""),
    String(d.emailAgente || ""),
    areaManagerDi_(canale)        // (P) colonna Area Manager
  ]);

  // ---- ROUTING MAIL ----
  const emailComp = String(d.emailAgente || "").trim();
  const bccDesk   = [MAIL_DESK];  // Sales Desk sempre in BCC silenzioso

  if (modalitaTest) {
    inviaConferma(d, canale, dataIt, [MAIL_TEST], [], [], COLORI.test, "[TEST] ", true);

  } else if (canale.gruppo === "master") {
    if (emailComp) {
      inviaConferma(d, canale, dataIt, [emailComp],  [], bccDesk, COLORI.compilatore, "", false);
    } else {
      inviaConferma(d, canale, dataIt, [MAIL_DESK],  [], [],      COLORI.salesdesk,   "", false);
    }

  } else if (canale.gruppo === "agenzia") {
    if (emailComp) {
      inviaConferma(d, canale, dataIt, [emailComp],    [], bccDesk, COLORI.compilatore, "", false);
    }
    inviaConferma(d, canale, dataIt, [canale.mail], [], bccDesk, COLORI.agent,       "", false);

  } else {
    if (emailComp) {
      inviaConferma(d, canale, dataIt, [emailComp],    [], bccDesk, COLORI.compilatore, "", false);
    }
    inviaConferma(d, canale, dataIt, [canale.mail], [], bccDesk, COLORI.agent,       "", false);
  }

  return rispondi({ ok: true });
}

/* ============================================================
   SHEET
   ============================================================ */
function apriTab() {
  const doc = SpreadsheetApp.openById(SHEET_ID);
  let tab = doc.getSheetByName(NOME_TAB);
  if (!tab) { tab = doc.insertSheet(NOME_TAB); tab.appendRow(INTESTAZIONI); tab.setFrozenRows(1); }
  else if (tab.getLastRow() === 0) { tab.appendRow(INTESTAZIONI); tab.setFrozenRows(1); }
  return tab;
}

/* ============================================================
   MAIL CONFERMA
   ============================================================ */
function inviaConferma(d, canale, dataIt, destinatari, ccList, bccList, colore, tag, isTest) {
  const oggetto = tag + 'Registrazione merchant "' + d.ragioneSociale + '" — ' + dataIt;
  const euro    = "€ " + Number(d.fatturatoMese).toLocaleString("it-IT");
  const disp    = d.tipoDispositivo + " × " + d.numeroDispositivi;
  const saluto  = canale.saluto === "team" ? "Complimenti Team " + canale.canale : "Complimenti " + d.agente;
  const listaDeal = elencoDeal(canale, d.emailAgente, d.ragioneSociale, dataIt);
  const rigaCopia = [dataIt, d.ragioneSociale, d.partitaIva, euro, disp,
    d.pricing, "RevPay " + d.tariffaPay, "Dom " + d.tariffaDomestic, "Prov. " + d.provincia].join(" · ");

  const bandaStyle = 'bgcolor="' + colore + '" style="background-color:' + colore + ';padding:8px 28px;color:#FFFFFF;font-size:11px;font-weight:700;text-align:center;letter-spacing:.08em;text-transform:uppercase;"';
  const bandaLabel = isTest ? "&#9888; EMAIL DI TEST" :
    (colore === COLORI.dircomm     ? "Direzione Commerciale"   :
     colore === COLORI.salesdesk   ? "Sales Tracker Desk"      :
     colore === COLORI.agent       ? "Notifica agente / area"  :
                                     "Conferma compilatore");

  const html =
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#0F1716;padding:20px 0;"><tr><td align="center">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" bgcolor="#0F1716" style="max-width:600px;background-color:#0F1716;border-radius:14px;overflow:hidden;border:1px solid #2E3A38;font-family:Poppins,Arial,sans-serif;">' +
    '<tr><td ' + bandaStyle + '>' + bandaLabel + '</td></tr>' +
    '<tr><td bgcolor="#1C2624" style="background-color:#1C2624;border-bottom:3px solid #FF941A;padding:22px 28px 20px;">' +
      '<table width="100%"><tr><td style="color:#FF941A;font-weight:700;font-size:19px;">Sales Design</td>' +
      '<td align="right" style="color:#B8BFBB;font-size:10px;font-weight:600;letter-spacing:2px;text-transform:uppercase;">Daily Deal Tracker</td></tr></table></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:26px 28px 22px;">' +
      '<span style="display:inline-block;background-color:#16302A;color:#3FD9B3;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;padding:6px 12px;border-radius:999px;">&#10003; Registrazione confermata</span>' +
      '<div style="margin-top:16px;color:#FFFFFF;font-size:22px;font-weight:700;line-height:1.3;">Registrazione merchant</div>' +
      '<div style="margin-top:4px;color:#FF941A;font-size:22px;font-weight:700;line-height:1.3;">' + d.ragioneSociale + '</div>' +
      '<div style="margin-top:10px;color:#C8CFCB;font-size:13px;">' + dataIt + ' &nbsp;&middot;&nbsp; canale <strong style="color:#FFFFFF;">' + canale.canale + '</strong> &nbsp;&middot;&nbsp; prov. <strong style="color:#FFFFFF;">' + d.provincia + '</strong></div></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:4px 28px 8px;">' +
      '<div style="color:#FFFFFF;font-size:15px;margin:0 0 6px;font-weight:600;">' + saluto + '</div>' +
      '<div style="color:#C8CFCB;font-size:13px;line-height:1.6;">un nuovo merchant &egrave; entrato nella rete Sales Design. Il deal &egrave; stato registrato e tracciato correttamente.</div></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:18px 28px 4px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#151F1D" style="background-color:#151F1D;border:1px solid #2E3A38;border-radius:12px;"><tr><td style="padding:18px 20px;">' +
        '<div style="color:#FF941A;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:14px;">Dettaglio deal</div>' +
        '<table width="100%" style="border-collapse:collapse;font-size:13.5px;">' +
          rigaTab("Merchant", d.ragioneSociale) + rigaTab("Partita IVA", d.partitaIva) +
          rigaTab("Provincia", d.provincia) +
          rigaTab("Fatturato mese", euro) + rigaTab("Dispositivo", disp) + rigaTab("Pricing", d.pricing) +
          rigaTabVerde("Tariffa RevPay", d.tariffaPay) + rigaTabVerde("Tariffa Domestic", d.tariffaDomestic) +
        '</table></td></tr></table></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:16px 28px 4px;">' +
      '<div style="color:#A8B0AC;font-size:11px;font-weight:600;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px;">Riga riepilogo</div>' +
      '<div style="color:#8A918D;font-size:11.5px;margin-bottom:8px;line-height:1.5;">Se lo desideri, copia questa riga in un tuo file personale per tenere traccia dei tuoi inserimenti.</div>' +
      '<div style="background-color:#1C2624;border:1px dashed #4A5754;border-radius:8px;padding:12px 14px;font-family:monospace;font-size:12px;color:#D4DBD7;word-break:break-all;">' + rigaCopia + '</div></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:20px 28px 8px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#151F1D" style="background-color:#151F1D;border:1px solid #2E3A38;border-radius:12px;"><tr><td style="padding:18px 20px;">' +
        '<div style="color:#FFFFFF;font-size:14px;font-weight:700;margin-bottom:4px;">Ultimi inserimenti</div>' +
        '<div style="color:#A8B0AC;font-size:12px;margin-bottom:14px;">' + (canale.gruppo === "master" ? "I deal che hai registrato tu" : "Deal registrati su " + canale.canale) + '</div>' +
        '<table width="100%" style="border-collapse:collapse;">' + listaDeal + '</table></td></tr></table></td></tr>' +
    '<tr><td bgcolor="#0A0F0E" style="background-color:#0A0F0E;border-top:1px solid #2E3A38;padding:18px 28px 22px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#1E1810" style="background-color:#1E1810;border:1px solid #4A3A1A;border-radius:8px;margin-bottom:16px;"><tr><td style="padding:12px 14px;">' +
        '<div style="color:#FF941A;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:5px;">Assistenza</div>' +
        '<div style="color:#C8CFCB;font-size:12px;line-height:1.6;">Per problematiche scrivete a <a href="mailto:direzionecommerciale@sales-design.it" style="color:#FF941A;text-decoration:none;">direzionecommerciale@sales-design.it</a></div></td></tr></table>' +
      '<div style="color:#8A918D;font-size:11px;line-height:1.7;text-align:center;"><strong style="color:#C8CFCB;">Sales Design S.r.l.</strong><br>Galleria Passarella 1, 20122 &mdash; Milano</div></td></tr>' +
  '</table></td></tr></table>';

  const testo = tag + oggetto + "\n\n" + saluto +
    "\nCanale: " + canale.canale + "\nAgente: " + d.agente +
    "\nMerchant: " + d.ragioneSociale + "\nP.IVA: " + d.partitaIva +
    "\nProvincia: " + d.provincia +
    "\nDispositivo: " + disp + "\nPricing: " + d.pricing +
    "\nRevPay: " + d.tariffaPay + "\nDomestic: " + d.tariffaDomestic +
    "\nRiga: " + rigaCopia;

  const opzioni = { htmlBody: html, name: MITTENTE };
  const dest = MODO_TEST ? [MAIL_TEST_DEST] : destinatari;   // in test: solo a te
  if (!MODO_TEST && !isTest) {
    if (ccList.length)  opzioni.cc  = ccList.join(",");
    if (bccList.length) opzioni.bcc = bccList.join(",");
  }
  MailApp.sendEmail(dest.join(","), (MODO_TEST ? "[TEST] " : "") + oggetto, testo, opzioni);
}

/* ============================================================
   DIGEST SUPERVISIONE — 4x al giorno (11/15/18/21)
   ============================================================ */
function inviaDigestSupervisione() {
  const adesso = new Date();
  const ora    = adesso.getHours();

  let oreFinestra;
  if      (ora >= 21) oreFinestra = 3;
  else if (ora >= 18) oreFinestra = 3;
  else if (ora >= 15) oreFinestra = 3;
  else if (ora >= 11) oreFinestra = 14;
  else                oreFinestra = 24;
  const dal = new Date(adesso.getTime() - oreFinestra * 60 * 60 * 1000);

  const dati = apriTab().getDataRange().getValues();
  const deal = [];

  for (let i = 1; i < dati.length; i++) {
    const tsRaw = dati[i][0];
    if (!tsRaw || String(tsRaw).trim() === "") continue;
    const dataRiga = parseTimestamp(tsRaw);
    if (!dataRiga) continue;
    if (dataRiga < dal || dataRiga > adesso) continue;
    deal.push({
      ts:       Utilities.formatDate(dataRiga, "Europe/Rome", "HH:mm"),
      canale:   String(dati[i][2]),
      agente:   String(dati[i][4]),
      merchant: String(dati[i][5]),
      provincia:String(dati[i][13]),
      disp:     String(dati[i][8]) + " × " + String(dati[i][9])
    });
  }

  if (!deal.length) {
    Logger.log("Digest ore " + ora + ": nessun deal nella finestra " + oreFinestra + "h.");
    return;
  }

  const perCanale = {};
  deal.forEach(function(r) { perCanale[r.canale] = (perCanale[r.canale] || 0) + 1; });

  const oraFmt  = Utilities.formatDate(adesso, "Europe/Rome", "HH:mm");
  const dataFmt = Utilities.formatDate(adesso, "Europe/Rome", "dd/MM/yyyy");
  const oggetto = "Riepilogo rete — " + deal.length + " contratt" + (deal.length === 1 ? "o" : "i") +
    " · " + dataFmt + " ore " + oraFmt;

  const righeHtml = deal.map(function(r, idx) {
    const bg = idx % 2 === 0 ? "#151F1D" : "#0F1716";
    return '<tr>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#6E7773;font-size:11px;">'  + r.ts      + '</td>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#FF941A;font-size:11px;font-weight:600;">' + r.canale  + '</td>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#C8CFCB;font-size:11px;">' + r.agente  + '</td>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#FFFFFF;font-size:11px;font-weight:500;">' + r.merchant + '</td>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#A8B0AC;font-size:11px;">' + r.provincia + '</td>' +
      '<td bgcolor="'+bg+'" style="background-color:'+bg+';padding:9px 12px;color:#C8CFCB;font-size:11px;text-align:right;">' + r.disp + '</td>' +
    '</tr>';
  }).join("");

  const totaliHtml = Object.keys(perCanale).sort().map(function(c) {
    return '<tr>' +
      '<td style="padding:5px 12px;color:#A8B0AC;font-size:12px;">' + c + '</td>' +
      '<td style="padding:5px 12px;color:#FF941A;font-size:16px;font-weight:700;text-align:right;">' + perCanale[c] + '</td>' +
    '</tr>';
  }).join("");

  const html =
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#0F1716;padding:20px 0;"><tr><td align="center">' +
  '<table role="presentation" width="620" cellpadding="0" cellspacing="0" bgcolor="#0F1716" style="max-width:620px;background-color:#0F1716;border-radius:14px;overflow:hidden;border:1px solid #2E3A38;font-family:Poppins,Arial,sans-serif;">' +
    '<tr><td bgcolor="#1C2624" style="background-color:#1C2624;border-bottom:3px solid #FF941A;padding:20px 28px;">' +
      '<table width="100%"><tr>' +
        '<td style="color:#FF941A;font-weight:700;font-size:17px;">Sales Design</td>' +
        '<td align="right" style="color:#B8BFBB;font-size:10px;font-weight:600;letter-spacing:2px;text-transform:uppercase;">Supervisione rete</td>' +
      '</tr></table></td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:22px 28px 16px;">' +
      '<div style="color:#FFFFFF;font-size:20px;font-weight:700;">Riepilogo contratti</div>' +
      '<div style="color:#C8CFCB;font-size:13px;margin-top:4px;">' + dataFmt + ' &nbsp;&middot;&nbsp; aggiornamento ore ' + oraFmt + '</div>' +
    '</td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:0 28px 20px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#151F1D" style="background-color:#151F1D;border:1px solid #2E3A38;border-radius:12px;"><tr><td style="padding:16px 20px;">' +
        '<table width="100%"><tr>' +
          '<td><div style="color:#6E7773;font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px;">CONTRATTI PERIODO</div>' +
              '<div style="color:#FF941A;font-size:36px;font-weight:800;line-height:1;">' + deal.length + '</div></td>' +
          '<td valign="top"><table style="border-collapse:collapse;">' + totaliHtml + '</table></td>' +
        '</tr></table>' +
      '</td></tr></table>' +
    '</td></tr>' +
    '<tr><td bgcolor="#0F1716" style="background-color:#0F1716;padding:0 28px 20px;">' +
      '<div style="color:#FFFFFF;font-size:13px;font-weight:600;margin-bottom:10px;">Dettaglio inserimenti</div>' +
      '<table width="100%" style="border-collapse:collapse;">' +
        '<tr bgcolor="#0A0F0E">' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:left;padding:7px 12px;letter-spacing:1px;">ORA</th>' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:left;padding:7px 12px;letter-spacing:1px;">CANALE</th>' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:left;padding:7px 12px;letter-spacing:1px;">AGENTE</th>' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:left;padding:7px 12px;letter-spacing:1px;">MERCHANT</th>' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:left;padding:7px 12px;letter-spacing:1px;">PROV.</th>' +
          '<th style="color:#6E7773;font-size:10px;font-weight:600;text-align:right;padding:7px 12px;letter-spacing:1px;">DISPOS.</th>' +
        '</tr>' +
        righeHtml +
      '</table>' +
    '</td></tr>' +
    '<tr><td bgcolor="#0A0F0E" style="background-color:#0A0F0E;border-top:1px solid #2E3A38;padding:16px 28px;text-align:center;">' +
      '<div style="color:#8A918D;font-size:11px;line-height:1.7;"><strong style="color:#C8CFCB;">Sales Design S.r.l.</strong><br>Galleria Passarella 1, 20122 &mdash; Milano</div>' +
    '</td></tr>' +
  '</table></td></tr></table>';

  const testo = oggetto + "\n\nContratti: " + deal.length + "\n\n" +
    Object.keys(perCanale).map(function(c){ return c + ": " + perCanale[c]; }).join("\n") +
    "\n\n" + deal.map(function(r){ return r.ts + " · " + r.canale + " · " + r.agente + " · " + r.merchant + " (" + r.provincia + ")"; }).join("\n");

  const destDigest = MODO_TEST ? MAIL_TEST_DEST : SUPERVISORI.join(",");   // in test: solo a te
  MailApp.sendEmail(destDigest, (MODO_TEST ? "[TEST] " : "") + oggetto, testo, { htmlBody: html, name: MITTENTE });
  Logger.log("Digest inviato: " + deal.length + " deal a " + destDigest);
}

/* ============================================================
   PARSER TIMESTAMP
   ============================================================ */
function parseTimestamp(tsRaw) {
  try {
    if (tsRaw instanceof Date) return new Date(tsRaw.getTime());
    if (typeof tsRaw === "number") return new Date((tsRaw - 25569) * 86400000);
    const str = String(tsRaw).trim();
    const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})[,\s]+(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m) {
      return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]),
                      parseInt(m[4]), parseInt(m[5]), parseInt(m[6]));
    }
    return null;
  } catch(e) { return null; }
}

/* ============================================================
   TRIGGER — eseguire UNA SOLA VOLTA dalla tendina
   ============================================================ */
function impostaTriggerDigest() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "inviaDigestSupervisione") ScriptApp.deleteTrigger(t);
  });
  [11, 15, 18, 21].forEach(function(ora) {
    ScriptApp.newTrigger("inviaDigestSupervisione").timeBased().everyDays(1).atHour(ora).create();
  });
  Logger.log("Trigger impostati: 11:00, 15:00, 18:00, 21:00");
}

/* ============================================================
   HELPER TABELLE EMAIL
   ============================================================ */
function rigaTab(k, v) {
  return '<tr><td style="color:#A8B0AC;padding:6px 0;">' + k + '</td><td style="color:#FFFFFF;font-weight:600;text-align:right;">' + v + '</td></tr>';
}
function rigaTabVerde(k, v) {
  return '<tr><td style="color:#A8B0AC;padding:6px 0;">' + k + '</td><td style="color:#3FD9B3;font-weight:700;text-align:right;">' + v + '</td></tr>';
}

/* ============================================================
   ULTIMI DEAL (sezione email)
   ============================================================ */
function elencoDeal(canale, emailCompilatore, merchantCorrente, dataCorrente) {
  try {
    const dati  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOME_TAB).getDataRange().getValues();
    const email = String(emailCompilatore || "").trim().toLowerCase();
    const righe = [];
    for (let i = dati.length - 1; i >= 1 && righe.length < 6; i--) {
      let match;
      if (canale.gruppo === "master") {
        match = email ? (String(dati[i][15]||"").trim().toLowerCase() === email) : (righe.length === 0);
      } else {
        match = (String(dati[i][2]) === canale.canale);
      }
      if (match) {
        const nome = String(dati[i][5]).trim();
        const data = String(dati[i][1]).trim();
        if (nome) righe.push({ nome: nome, data: data, primo: righe.length === 0 });
      }
    }
    if (!righe.length) return '<tr><td style="padding:9px 0;color:#FFFFFF;font-size:13px;">' + merchantCorrente + '</td></tr>';
    return righe.map(function(r, idx) {
      const bordo  = idx < righe.length - 1 ? "border-bottom:1px solid #232E2C;" : "";
      const colore = r.primo ? "#FFFFFF" : "#D4DBD7";
      const peso   = r.primo ? "500" : "400";
      return '<tr><td style="padding:9px 0;' + bordo + 'color:' + colore + ';font-size:13px;font-weight:' + peso + ';">' + r.nome +
        '</td><td align="right" style="padding:9px 0;' + bordo + 'color:#6E7773;font-size:12px;">' + String(r.data).substring(0,5) + '</td></tr>';
    }).join("");
  } catch(e) {
    return '<tr><td style="padding:9px 0;color:#FFFFFF;font-size:13px;">' + merchantCorrente + '</td></tr>';
  }
}

function rispondi(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   TEST
   ============================================================ */
function testInvio() {
  Logger.log(elabora({
    codice: "SD-DIRECT", data: "2026-10-16", agente: "Mario Rossi",
    ragioneSociale: "Azienda Test S.r.l.", partitaIva: "12345678901",
    fatturatoMese: "30000", tipoDispositivo: "Standard", numeroDispositivi: "2",
    pricing: "189", tariffaPay: "0,50%", tariffaDomestic: "0,80%",
    provincia: "MI", note: "Test", emailAgente: "dealtracker.hub@gmail.com"
  }, true).getContent());
}

function testTutteLeLinee() {
  const codici = ["SD-DIRECT","SD-AG-WECOMM","SD-AG-FIDITECH","SD-AG-PUNTOSEMPLICE",
                  "SD-AG-GRANDASOLUTIONS","SD-AG-TEAMTUTINO","SD_NORD","SD_CENTRO","SD_SUD"];
  codici.forEach(function(cod, i) {
    elabora({
      codice: cod, data: "2026-10-16", agente: "Test " + cod,
      ragioneSociale: "Test " + cod + " Srl", partitaIva: "1234567890" + (i < 9 ? i+1 : "1"),
      fatturatoMese: "30000", tipoDispositivo: "Standard", numeroDispositivi: "2",
      pricing: "189", tariffaPay: "0,50%", tariffaDomestic: "0,80%",
      provincia: "TO", note: "Test linea", emailAgente: "dealtracker.hub@gmail.com"
    }, true);
    Utilities.sleep(500);
  });
  Logger.log("OK: 9 linee testate (incl. Granda + Team Tutino)");
}

function testDigest() {
  inviaDigestSupervisione();
}

/* ============================================================
   DASHBOARD — SD (Direct+Agenzia) vs AT (Aree)
   ============================================================ */
function generaDashboard() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vecchia = ss.getSheetByName("STATS");
  if (vecchia) ss.deleteSheet(vecchia);
  const s = ss.insertSheet("STATS", 1);
  s.setHiddenGridlines(true);
  s.setColumnWidth(1, 24);
  [2,3,4,5,6,7,8].forEach(function(c){ s.setColumnWidth(c, 130); });

  s.getRange("B2").setValue("DAILY DEAL TRACKER — Contratti chiusi").setFontSize(16).setFontWeight("bold").setFontColor("#0F1716");
  s.getRange("B3").setValue("Metrica: numero contratti · aggiornata in tempo reale").setFontColor("#6E7773").setFontSize(9);

  s.getRange("B5").setValue("OGGI").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("C5").setValue("MESE CORRENTE").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("D5").setValue("vs MESE PREC.").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("E5").setValue("TRIM. CORRENTE").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("F5").setValue("vs TRIM. PREC.").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("G5").setValue("ANNO").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("B6").setFormula('=COUNTIFS(Master!B2:B;TEXT(TODAY();"dd/mm/yyyy"))').setFontSize(20).setFontWeight("bold").setFontColor("#FF941A");
  s.getRange("C6").setFormula('=SUMPRODUCT((Master!B2:B<>"")*(IFERROR(MONTH(DATEVALUE(Master!B2:B));0)=MONTH(TODAY()))*(IFERROR(YEAR(DATEVALUE(Master!B2:B));0)=YEAR(TODAY())))').setFontSize(20).setFontWeight("bold").setFontColor("#FF941A");
  s.getRange("D6").setFormula('=SUMPRODUCT((Master!B2:B<>"")*(IFERROR(MONTH(DATEVALUE(Master!B2:B));0)=MONTH(EDATE(TODAY();-1)))*(IFERROR(YEAR(DATEVALUE(Master!B2:B));0)=YEAR(EDATE(TODAY();-1))))').setFontSize(20).setFontWeight("bold").setFontColor("#8A918D");
  s.getRange("E6").setFormula('=SUMPRODUCT((Master!B2:B<>"")*(IFERROR(ROUNDUP(MONTH(DATEVALUE(Master!B2:B))/3;0);0)=ROUNDUP(MONTH(TODAY())/3;0))*(IFERROR(YEAR(DATEVALUE(Master!B2:B));0)=YEAR(TODAY())))').setFontSize(20).setFontWeight("bold").setFontColor("#FF941A");
  s.getRange("F6").setFormula('=SUMPRODUCT((Master!B2:B<>"")*(IFERROR(ROUNDUP(MONTH(DATEVALUE(Master!B2:B))/3;0);0)=ROUNDUP(MONTH(EDATE(TODAY();-3))/3;0))*(IFERROR(YEAR(DATEVALUE(Master!B2:B));0)=YEAR(EDATE(TODAY();-3))))').setFontSize(20).setFontWeight("bold").setFontColor("#8A918D");
  s.getRange("G6").setFormula('=SUMPRODUCT((Master!B2:B<>"")*(IFERROR(YEAR(DATEVALUE(Master!B2:B));0)=YEAR(TODAY())))').setFontSize(20).setFontWeight("bold").setFontColor("#FF941A");

  s.getRange("B8").setValue("Δ Mese vs prec:").setFontColor("#6E7773").setFontSize(10);
  s.getRange("C8").setFormula('=IF(D6=0;"—";TEXT((C6-D6)/D6;"+0%;-0%"))').setFontWeight("bold").setFontSize(11);
  s.getRange("E8").setValue("Δ Trim. vs prec:").setFontColor("#6E7773").setFontSize(10);
  s.getRange("F8").setFormula('=IF(F6=0;"—";TEXT((E6-F6)/F6;"+0%;-0%"))').setFontWeight("bold").setFontSize(11);

  s.getRange("B10").setValue("DIREZIONI COMMERCIALI").setFontWeight("bold").setFontColor("#0F1716").setFontSize(11);
  s.getRange("B11").setValue("TOTALE RETE DIRETTA (SD)").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("B12").setFormula('=COUNTIF(Master!D2:D;"Direct")+COUNTIF(Master!D2:D;"Agenzia")').setFontSize(26).setFontWeight("bold").setFontColor("#FF941A");
  s.getRange("D11").setValue("TOTALE AREE (AT)").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("D12").setFormula('=COUNTIF(Master!D2:D;"Area")').setFontSize(26).setFontWeight("bold").setFontColor("#2EC4A0");
  s.getRange("F11").setValue("QUOTA SD").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("F12").setFormula('=IFERROR(TEXT(B12/(B12+D12);"0%");"—")').setFontSize(16).setFontWeight("bold").setFontColor("#FF941A");
  s.getRange("G11").setValue("QUOTA AT").setFontWeight("bold").setFontSize(9).setFontColor("#6E7773");
  s.getRange("G12").setFormula('=IFERROR(TEXT(D12/(B12+D12);"0%");"—")').setFontSize(16).setFontWeight("bold").setFontColor("#2EC4A0");

  s.getRange("B15").setValue("Direzione").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("C15").setValue("Contratti").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("B16").setValue("Rete diretta (SD)");
  s.getRange("C16").setFormula('=B12');
  s.getRange("B17").setValue("Aree (AT)");
  s.getRange("C17").setFormula('=D12');

  s.getRange("E15").setValue("Rete diretta").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("F15").setValue("Contratti").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("E16").setValue("Direct Sales");
  s.getRange("F16").setFormula('=COUNTIF(Master!C2:C;"Direct Sales")');
  s.getRange("E17").setValue("WeComm");
  s.getRange("F17").setFormula('=COUNTIF(Master!C2:C;"WeComm")');
  s.getRange("E18").setValue("Fiditech");
  s.getRange("F18").setFormula('=COUNTIF(Master!C2:C;"Fiditech")');
  s.getRange("E19").setValue("Punto Semplice");
  s.getRange("F19").setFormula('=COUNTIF(Master!C2:C;"Punto Semplice")');
  s.getRange("E20").setValue("Granda Solutions");
  s.getRange("F20").setFormula('=COUNTIF(Master!C2:C;"Granda Solutions")');
  s.getRange("E21").setValue("Team Tutino");
  s.getRange("F21").setFormula('=COUNTIF(Master!C2:C;"Team Tutino")');

  s.getRange("B20").setValue("Aree").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("C20").setValue("Contratti").setFontWeight("bold").setBackground("#F1EFE8");
  s.getRange("B21").setValue("Area Nord");
  s.getRange("C21").setFormula('=COUNTIF(Master!C2:C;"Area Nord")');
  s.getRange("B22").setValue("Area Centro");
  s.getRange("C22").setFormula('=COUNTIF(Master!C2:C;"Area Centro")');
  s.getRange("B23").setValue("Area Sud");
  s.getRange("C23").setFormula('=COUNTIF(Master!C2:C;"Area Sud")');

  s.getRange("B27").setValue("CLASSIFICA SALES / AGENZIA / AREA").setFontWeight("bold").setFontColor("#0F1716").setFontSize(11);
  s.getRange("B28").setFormula('=QUERY(Master!A2:P;"select E, count(G) where E is not null group by E order by count(G) desc label E \'Chi\', count(G) \'Contratti\'";0)');
  s.getRange("E27").setValue("CONTRATTI PER GIORNO").setFontWeight("bold").setFontColor("#0F1716").setFontSize(11);
  s.getRange("E28").setFormula('=QUERY(Master!A2:P;"select B, count(G) where B is not null group by B order by B desc label B \'Giorno\', count(G) \'Contratti\'";0)');

  const chartDirezioni = s.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(s.getRange("B15:C17"))
    .setPosition(40, 2, 0, 0)
    .setOption("title", "Direzioni commerciali · SD vs AT")
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#FF941A", "#2EC4A0"])
    .setOption("width", 620).setOption("height", 300).build();
  s.insertChart(chartDirezioni);

  const chartRete = s.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(s.getRange("E15:F21"))
    .setPosition(58, 2, 0, 0)
    .setOption("title", "Composizione Rete diretta (SD)")
    .setOption("colors", ["#FF941A", "#FFB055", "#E67E00", "#B36400", "#7A4400", "#4A2800"])
    .setOption("pieHole", 0.4)
    .setOption("width", 400).setOption("height", 300).build();
  s.insertChart(chartRete);

  const chartAree = s.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(s.getRange("B20:C23"))
    .setPosition(58, 5, 0, 0)
    .setOption("title", "Composizione Aree (AT)")
    .setOption("colors", ["#2EC4A0", "#3FD9B3", "#1F9B7C"])
    .setOption("pieHole", 0.4)
    .setOption("width", 400).setOption("height", 300).build();
  s.insertChart(chartAree);

  ss.setActiveSheet(s);
  Logger.log("Dashboard STATS generata correttamente.");
}


/* ████████████████████████████████████████████████████████████
   ██  API VENDITE (sola lettura) — ADDITIVA                   ██
   ██  doGet serve i contratti del tab Master al cruscotto.    ██
   ██  Il RECRUITING ha il SUO foglio e il SUO script          ██
   ██  (recruiting-apps-script.gs): qui non c'e' piu'.         ██
   ████████████████████████████████████████████████████████████ */

/* doGet -> { ok, records:[...] }  (contratti dal tab Master) */
function doGet(e) {
  try {
    return rispondi({
      ok: true,
      records: leggiSales_(),
      generato: Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss")
    });
  } catch (err) {
    return rispondi({ ok: false, errore: String(err && err.message ? err.message : err) });
  }
}

/* ---- Lettura VENDITE dal tab Master (sola lettura) ---- */
function leggiSales_() {
  const doc = SpreadsheetApp.openById(SHEET_ID);
  const tab = doc.getSheetByName(NOME_TAB);
  if (!tab || tab.getLastRow() < 2) return [];
  const dati = tab.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < dati.length; i++) {
    const r = dati[i];
    if (!r[0] && !r[1]) continue;
    const tipo = String(r[3] || "").trim();
    const rete = tipo === "Area" ? "ATD" : "SD";   // Direct+Agenzia = SD · Area = ATD (AT)
    out.push({
      ts:         String(r[0] || ""),
      data:       fmtDataIt_(r[1]),
      canale:     String(r[2] || ""),
      tipo:       tipo,
      rete:       rete,
      agente:     String(r[4] || ""),
      merchant:   String(r[5] || ""),
      partitaIva: String(r[6] || "").replace(/^'/, ""),
      fatturato:  Number(r[7]) || 0,
      dispositivo:String(r[8] || ""),
      nDisp:      Number(r[9]) || 0,
      pricing:    String(r[10] || ""),
      tariffaPay: String(r[11] || ""),
      tariffaDom: String(r[12] || ""),
      provincia:  String(r[13] || "").toUpperCase(),
      am:         String(r[16] || "")      // (P) Area Manager di competenza
    });
  }
  return out;
}

/* Formatta la colonna Data in "dd/MM/yyyy" (gestisce Date, ISO, o testo). */
function fmtDataIt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Rome", "dd/MM/yyyy");
  const s = String(v || "").trim();
  if (!s) return "";
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);           // 2026-07-15
  if (iso) return iso[3] + "/" + iso[2] + "/" + iso[1];
  const it = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);      // 15/7/2026
  if (it) return String(it[1]).padStart(2, "0") + "/" + String(it[2]).padStart(2, "0") + "/" + it[3];
  const d = new Date(s);                                      // es. "Wed Jul 15 2026 ..."
  if (!isNaN(d)) return Utilities.formatDate(d, "Europe/Rome", "dd/MM/yyyy");
  return s;
}

/* Test rapido API vendite (log JSON). */
function testApiVendite() {
  Logger.log(doGet({}).getContent());
}
