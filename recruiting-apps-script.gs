/**
 * ============================================================
 *  RECRUITING — backend del 2° Google Sheet (fonte separata)
 *  Alimenta la sezione "Team Recruiting" del cruscotto.
 *  Indipendente dallo script vendite.
 * ------------------------------------------------------------
 *  SETUP (consigliato, senza ID):
 *   1. Crea un NUOVO Google Sheet dedicato al recruiting.
 *   2. DAL FOGLIO: Estensioni -> Apps Script -> incolla questo file.
 *      (cosi' lo script e' "legato" al foglio: NON serve l'ID)
 *   3. Lascia SHEET_ID_REC com'e'. Salva.
 *   4. Esegui setupRecruiting() una volta (crea i 3 tab). Autorizza.
 *   5. Deploy -> App web -> Esegui come Me -> Accesso: Chiunque.
 *   6. Copia l'URL /exec in CFG.ENDPOINT_RECRUITING del cruscotto.
 *
 *  In alternativa (script standalone): incolla in SHEET_ID_REC l'ID
 *  del foglio (la parte tra /d/ e /edit nell'URL). Lo script deve
 *  girare con lo STESSO account che ha accesso al foglio.
 * ============================================================
 */

// ID del foglio recruiting (lascia "" se lo script e' legato al foglio).
const SHEET_ID_REC = "1BcBewjKcBdc1xSomiGsakxXU_x04DGymCmYWbUh56vo";

// Restituisce il foglio: per ID se impostato, altrimenti il foglio legato.
function ssRec_() {
  if (SHEET_ID_REC && !/INCOLLA|^$/.test(SHEET_ID_REC)) return SpreadsheetApp.openById(SHEET_ID_REC);
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (!a) throw new Error("Nessun foglio legato: crea lo script da 'Estensioni > Apps Script' del foglio, oppure imposta SHEET_ID_REC.");
  return a;
}

const TAB_RECRUITING  = "Recruiting";    // persone della rete
const TAB_ATTIVAZIONI = "Attivazioni";   // attivazioni per persona e mese
const TAB_POTENZIALE  = "Potenziale";    // quota retail per regione

const HEAD_RECRUITING = [
  "ID","Nome","Cognome","Ruolo","Tipo contratto","Tipo ingaggio","Citta","Regione",
  "AM riferimento","Fase","Data fase","Data ingaggio","Diretti","Capacita","Agenti",
  "Canale","On dal","Off dal",
  "Esclusivita","Altri mandati","Note CRM"
];
// Esclusivita: "Esclusivo" (solo con noi) | "Multi-mandato" (anche con altri)
// Tipo contratto (compenso): A = Sales · B = Agency · C = Override Mng (AM)
const HEAD_ATTIVAZIONI = ["AgenteID","Mese","Attivazioni"];
const HEAD_POTENZIALE  = ["Regione","Retail"];

const STAGES_REC = ["CALL","COLLOQUIO","INGAGGIO","MANDATO + MAIL","WHATSAPP","MAPPA ATTIVA"];

/* ------------------------------------------------------------
   doGet — API lettura recruiting (JSON)
   restituisce { ok, agenti:[], vendite:[], potenziale:[] }
   ------------------------------------------------------------ */
function doGet(e) {
  try {
    const r = leggiRecruiting_();
    return rispondi_({
      ok: true,
      agenti: r.agenti, vendite: r.vendite, potenziale: r.potenziale,
      generato: Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss")
    });
  } catch (err) {
    return rispondi_({ ok: false, errore: String(err && err.message ? err.message : err) });
  }
}

function leggiRecruiting_() {
  const doc = ssRec_();

  // -- Persone --
  const agenti = [];
  const ta = doc.getSheetByName(TAB_RECRUITING);
  if (ta && ta.getLastRow() >= 2) {
    const d = ta.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const r = d[i];
      const nome = String(r[1] || "").trim();
      const cognome = String(r[2] || "").trim();
      if (!nome && !cognome) continue;
      const id = String(r[0] || "").trim() || ("row" + (i + 1));
      const ruolo = String(r[3] || "SALES").trim().toUpperCase();
      agenti.push({
        id: id,
        nome: (nome + " " + cognome).trim(),
        nomeProprio: nome,
        cognome: cognome,
        ruolo: ruolo,
        tipoContratto: normTipoContratto_(r[4], ruolo),
        tipoIngaggio: String(r[5] || "").trim(),
        citta: String(r[6] || "").trim(),
        regione: String(r[7] || "").trim(),
        parent: String(r[8] || "").trim(),
        stage: clampStage_(r[9]),
        stageSince: toISO_(r[10]),
        ingaggio: toISO_(r[11]),
        diretti: Number(r[12]) || 0,
        capacita: (r[13] === "" || r[13] == null) ? null : (Number(r[13]) || 0),
        agenti: Number(r[14]) || 0,
        canale: String(r[15] || "").trim(),
        onDal: toISO_(r[16]),
        offDal: toISO_(r[17]),
        attivo: !toISO_(r[17]),
        esclusivita: String(r[18] || "").trim(),
        altriMandati: String(r[19] || "").trim(),
        noteCrm: String(r[20] || "").trim()
      });
    }
  }

  // -- Attivazioni --
  const vendite = [];
  const tv = doc.getSheetByName(TAB_ATTIVAZIONI);
  if (tv && tv.getLastRow() >= 2) {
    const d = tv.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const aid = String(d[i][0] || "").trim();
      if (!aid) continue;
      vendite.push({ agenteId: aid, data: toMese_(d[i][1]), attivazioni: Number(d[i][2]) || 0 });
    }
  }

  // -- Potenziale --
  const potenziale = [];
  const tp = doc.getSheetByName(TAB_POTENZIALE);
  if (tp && tp.getLastRow() >= 2) {
    const d = tp.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const reg = String(d[i][0] || "").trim();
      if (!reg) continue;
      potenziale.push({ regione: reg, retail: Number(d[i][1]) || 0 });
    }
  }

  return { agenti: agenti, vendite: vendite, potenziale: potenziale };
}

/* ------------------------------------------------------------
   setupRecruiting — crea i 3 tab con intestazioni + esempi.
   Eseguire UNA VOLTA dalla tendina delle funzioni.
   ------------------------------------------------------------ */
function setupRecruiting() {
  const ss = ssRec_();

  const tabRec = creaTabConHead_(ss, TAB_RECRUITING, HEAD_RECRUITING);
  if (tabRec.getLastRow() < 2) tabRec.getRange(2, 1, ESEMPI_RECRUITING.length, HEAD_RECRUITING.length).setValues(ESEMPI_RECRUITING);

  const tabAtt = creaTabConHead_(ss, TAB_ATTIVAZIONI, HEAD_ATTIVAZIONI);
  if (tabAtt.getLastRow() < 2) tabAtt.getRange(2, 1, ESEMPI_ATTIVAZIONI.length, HEAD_ATTIVAZIONI.length).setValues(ESEMPI_ATTIVAZIONI);

  const tabPot = creaTabConHead_(ss, TAB_POTENZIALE, HEAD_POTENZIALE);
  if (tabPot.getLastRow() < 2) tabPot.getRange(2, 1, ESEMPI_POTENZIALE.length, HEAD_POTENZIALE.length).setValues(ESEMPI_POTENZIALE);

  Logger.log("Tab recruiting pronti su SHEET_ID_REC.");
}

/* ------------------------------------------------------------ HELPERS ------ */
function creaTabConHead_(ss, nome, head) {
  let t = ss.getSheetByName(nome);
  if (!t) t = ss.insertSheet(nome);
  if (t.getLastRow() === 0) {
    t.appendRow(head); t.setFrozenRows(1);
    t.getRange(1, 1, 1, head.length).setFontWeight("bold").setBackground("#111827").setFontColor("#E8ECFF");
  }
  return t;
}
function clampStage_(v) { const n = parseInt(v, 10); if (isNaN(n)) return 0; return Math.max(0, Math.min(5, n)); }
function normTipoContratto_(v, ruolo) {
  const c = String(v || "").trim().toUpperCase().charAt(0);
  if (c === "A" || c === "B" || c === "C") return c;
  if (ruolo === "AM") return "C";
  if (ruolo === "AGENZIA") return "B";
  return "A";
}
function toISO_(v) {
  if (!v && v !== 0) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Rome", "yyyy-MM-dd");
  const s = String(v).trim(); if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return m[3] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[1]).padStart(2, "0");
  return s;
}
function toMese_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Rome", "yyyy-MM");
  const s = String(v || "").trim();
  if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
  const m = s.match(/^(\d{1,2})[\/\-](\d{4})/);
  if (m) return m[2] + "-" + String(m[1]).padStart(2, "0");
  return s;
}
function rispondi_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------ ESEMPI ------ */
const ESEMPI_POTENZIALE = [
  ["Piemonte", 7.3], ["Valle d'Aosta", 0.2], ["Lombardia", 16.8], ["Trentino-Alto Adige", 1.9],
  ["Veneto", 8.1], ["Friuli-Venezia Giulia", 2.0], ["Liguria", 2.6], ["Emilia-Romagna", 7.6],
  ["Toscana", 6.3], ["Umbria", 1.4], ["Marche", 2.5], ["Lazio", 9.6], ["Abruzzo", 2.1],
  ["Molise", 0.5], ["Campania", 9.2], ["Puglia", 6.4], ["Basilicata", 0.9], ["Calabria", 3.0],
  ["Sicilia", 8.0], ["Sardegna", 2.6]
];
// Colonne = HEAD_RECRUITING · Tipo contratto: A Sales · B Agency · C Override Mng
const ESEMPI_RECRUITING = [
  ["AM-001","Marco","Rossi","AM","C","Mandato","Torino","Piemonte","", 5, "2026-09-10", "2026-02-01", 4, 10, 0],
  ["AG-001","Luca","Bianchi","AGENZIA","B","Partita IVA","Milano","Lombardia","Marco Rossi", 5, "2026-08-20", "2026-03-15", 0, 0, 12],
  ["SA-001","Sara","Ferrari","SALES","A","Procacciatore","Torino","Piemonte","Marco Rossi", 4, "2026-10-01", "2026-06-01", 0, 0, 0],
  ["SA-002","Giulia","Russo","SALES","A","Procacciatore","Napoli","Campania","", 2, "2026-10-05", "2026-09-10", 0, 0, 0],
  ["AM-002","Antonio","Esposito","AM","C","Mandato","Roma","Lazio","", 5, "2026-07-01", "2026-01-20", 2, 8, 0]
];
const ESEMPI_ATTIVAZIONI = [
  ["AM-001","2026-10", 6], ["AG-001","2026-10", 9], ["SA-001","2026-10", 3],
  ["AM-002","2026-10", 4], ["SA-002","2026-10", 1],
  ["AM-001","2026-09", 5], ["AG-001","2026-09", 7]
];

/* ------------------------------------------------------------
   popolaRete — scrive la rete REALE nota (AM + Agency + Sales).
   SOVRASCRIVE il tab Recruiting. Eseguire UNA VOLTA.
   Tutti ON da oggi (colonna "Off dal" vuota = attivo).
   Regioni/tipo ingaggio sono stime da rifinire nel foglio.
   ------------------------------------------------------------ */
function popolaRete() {
  const ss = ssRec_();
  let t = ss.getSheetByName(TAB_RECRUITING);
  if (!t) t = ss.insertSheet(TAB_RECRUITING);
  t.clear();
  t.getRange(1, 1, 1, HEAD_RECRUITING.length).setValues([HEAD_RECRUITING])
    .setFontWeight("bold").setBackground("#111827").setFontColor("#E8ECFF");
  t.setFrozenRows(1);
  // completa ogni riga con i campi CRM lite (Esclusivita, Altri mandati, Note CRM)
  const rows = ROSTER_NOTO.map(function(r){ return r.concat(["Esclusivo", "", ""]); });
  t.getRange(2, 1, rows.length, HEAD_RECRUITING.length).setValues(rows);
  Logger.log("Rete popolata: " + rows.length + " persone (AM + Agency + Sales).");
}

// Data di "accensione" (oggi). Modificabile per singola persona nel foglio.
const T_ON = "2026-10-16";
// Colonne = HEAD_RECRUITING (18): ID,Nome,Cognome,Ruolo,TipoContratto,TipoIngaggio,
//   Citta,Regione,AMrif,Fase,DataFase,DataIngaggio,Diretti,Capacita,Agenti,Canale,OnDal,OffDal
const ROSTER_NOTO = [
  // --- AREA MANAGER (C) ---
  ["AM-01","Massimo","Mantovelli","AM","C","Mandato","","Piemonte","",5,T_ON,T_ON,7,10,0,"Area Nord",T_ON,""],
  ["AM-02","Marco","Cima","AM","C","Mandato","","Lazio","",5,T_ON,T_ON,2,8,0,"Area Centro",T_ON,""],
  ["AM-03","","Massafra","AM","C","Mandato","","Campania","",5,T_ON,T_ON,7,10,0,"Area Sud",T_ON,""],
  // --- AGENCY (B) ---
  ["AG-01","WeComm","","AGENZIA","B","","","","",5,T_ON,T_ON,0,0,0,"WeComm",T_ON,""],
  ["AG-02","Fiditech","","AGENZIA","B","","","","",5,T_ON,T_ON,0,0,0,"Fiditech",T_ON,""],
  ["AG-03","Punto Semplice","","AGENZIA","B","","","","",5,T_ON,T_ON,0,0,0,"Punto Semplice",T_ON,""],
  ["AG-04","Granda Solutions","","AGENZIA","B","","","","",5,T_ON,T_ON,0,0,0,"Granda Solutions",T_ON,""],
  ["AG-05","Team Tutino","","AGENZIA","B","","","","",5,T_ON,T_ON,0,0,0,"Team Tutino",T_ON,""],
  // --- SALES (A) · Area Sud → AM Massafra ---
  ["SA-01","Fabio","Pirrelli","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-02","Emanuele","Montagna","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-03","Francesco","Di Vara","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-04","Francesco","De Filippi","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-05","Giovanni","Enzo Rundo","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-06","Domenico","Derasmo","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  ["SA-07","Francesco","Iazzolino","SALES","A","","","Campania","Massafra",5,T_ON,T_ON,0,"",0,"Area Sud",T_ON,""],
  // --- SALES (A) · Area Nord → AM Mantovelli ---
  ["SA-08","Rossano","Brozzu","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-09","Pacino","Gesualdo Giancarlo","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-10","Andrea","Acuti","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-11","Michele","Acquisti","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-12","Daniele","Petrisor Niculae","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-13","Duccio","Morelli","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  ["SA-14","Piero","Salciccia","SALES","A","","","Piemonte","Massimo Mantovelli",5,T_ON,T_ON,0,"",0,"Area Nord",T_ON,""],
  // --- SALES (A) · Area Centro → AM Cima ---
  ["SA-15","Edoardo","Cannella","SALES","A","","","Lazio","Marco Cima",5,T_ON,T_ON,0,"",0,"Area Centro",T_ON,""],
  ["SA-16","Anna","Garzone","SALES","A","","","Lazio","Marco Cima",5,T_ON,T_ON,0,"",0,"Area Centro",T_ON,""],
  // --- SALES (A) · Direct Sales (senza AM) ---
  ["SA-17","Stefano","Arestia","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""],
  ["SA-18","Andrea","Perlo","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""],
  ["SA-19","Rossella","Candura","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""],
  ["SA-20","Angelo","Ambrosino","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""],
  ["SA-21","Cataldo","Vasapolli","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""],
  ["SA-22","Alida","Lerda","SALES","A","","","","",5,T_ON,T_ON,0,"",0,"Direct Sales",T_ON,""]
];

function testApiRecruiting() { Logger.log(doGet({}).getContent()); }
