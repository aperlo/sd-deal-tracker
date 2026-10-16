# SD Deal Tracker — Sales Design / Attitude Ltd

Cruscotto unico, grafica moderna "app", diviso in due sezioni che leggono da **due fonti separate**:

- **Sales Tracking** — contratti chiusi + **proiezione guadagni (Economics)** con i modelli A / B / C.
- **Team Recruiting** — mappa Italia, linee AM, indice performance per zona, pipeline e schede clusterizzate.

Un solo `index.html`. Accesso protetto da **password**. Nessun framework, nessuna build.

---

## 1. Architettura — DUE fonti indipendenti

```
FONTE 1 · VENDITE                         FONTE 2 · RECRUITING
[ Form contratti ]                        [ 2° Google Sheet ]
   (Cloudflare) ─POST→ [ Apps Script       (Recruiting/Attivazioni/Potenziale)
                         VENDITE ]                  │
   [ Google Sheet VENDITE (Master) ]        [ Apps Script RECRUITING ]
                         │ doGet(sales)              │ doGet
                         ▼                           ▼
                 ENDPOINT_SALES              ENDPOINT_RECRUITING
                         └──────────┬────────────────┘
                                    ▼
                        [ index.html ] (Netlify, con password)
```

| Componente | Tecnologia | Deploy |
|---|---|---|
| Cruscotto | HTML + Chart.js 4.4.1 + D3 7.9.0 | **Netlify** |
| Form contratti | HTML statici (`sito-sd/`) | **Cloudflare Pages** |
| Backend vendite | `apps-script.gs` (sul foglio Master) | Google |
| Backend recruiting | `recruiting-apps-script.gs` (sul 2° foglio) | Google |
| Dati | 2 Google Sheets | Google Drive |

> La logica vendite (doPost/colonne/CODICI/email/digest) è **intatta**. L'unica aggiunta è `doGet` (lettura) + `MODO_TEST` per le email.

---

## 2. Modelli di compenso (sezione Economics)

Ogni persona del recruiting ha un **Tipo contratto** (A/B/C) che determina il compenso. La sezione *Economics* (apribile, è una vista a traino) stima il payout su Sales/AM (per agente) e Agency (per canale).

- **A · Sales** — scaglioni sul n° contratti/mese: ≤9 =150 · 10–14 =200 · 15–20 =220 · 21–24 =230 · 25+ =240 (tariffa su tutti i contratti del mese).
- **B · Agency** — 250 € fisso per contratto.
- **C · Override Mng (AM)** — 25 €/contratto team fino a 25, 50 € oltre — **+** compenso Sales (A) personale.

Parametri modificabili nell'oggetto `COMP` in cima allo script del cruscotto.

---

## 3. Deploy — passo per passo

### A) Backend VENDITE (foglio Master esistente)
1. Google Sheet vendite → **Estensioni → Apps Script** → incolla `apps-script.gs`.
2. ⚠️ In cima c'è `const MODO_TEST = true;` → **in test le email vanno SOLO a `MAIL_TEST_DEST`** (niente invii a canali/agenzie/aree/supervisori). Metti `false` solo quando vai in produzione.
3. **Deploy → App web** · Esegui come **Me** · Accesso **Chiunque** → copia l'URL **`/exec`** = **ENDPOINT_SALES**.

### B) Backend RECRUITING (2° foglio, nuovo)
1. Crea un **nuovo Google Sheet** dedicato al recruiting.
2. **Estensioni → Apps Script** → incolla `recruiting-apps-script.gs`.
3. In cima imposta `SHEET_ID_REC` con l'**ID del nuovo foglio** (lo trovi nell'URL: `/d/ID/edit`).
4. Esegui una volta **`setupRecruiting`** (crea i tab `Recruiting`/`Attivazioni`/`Potenziale` con esempi; autorizza).
5. **Deploy → App web** · Esegui come **Me** · Accesso **Chiunque** → copia l'URL **`/exec`** = **ENDPOINT_RECRUITING**.

### C) Cruscotto (`index.html`)
Nell'oggetto `CFG` in cima allo `<script>`:
```js
const CFG = {
  ENDPOINT_SALES:       "https://script.google.com/macros/s/AAA/exec",
  ENDPOINT_RECRUITING:  "https://script.google.com/macros/s/BBB/exec",
  SHEET_URL_SALES:      "https://docs.google.com/.../edit",
  SHEET_URL_RECRUITING: "https://docs.google.com/.../edit",   // 2° foglio
  ACCESS_PASSWORD:      "sd-tracker-2026",   // password accesso ("" = nessuna)
  POLL_MS:              0
};
```

### D) Netlify
GitHub Desktop → commit su `main` → push. Netlify → *Import from Git* → repo `sd-deal-tracker` · Build command vuoto · Publish directory `.`.

---

## 4. Accesso riservato (password)

Il cruscotto mostra un **gate**: senza la `ACCESS_PASSWORD` non si vede nulla. Cambiala in `CFG.ACCESS_PASSWORD`.
> È un blocco lato pagina (soft lock) adatto "per ora": tiene fuori i curiosi. Per sicurezza forte serve la protezione a livello di hosting (Netlify Pro / accesso Identity) — da valutare più avanti.

---

## 5. Verifiche rapide
- `ENDPOINT_SALES` nel browser → deve dare `{"ok":true,"records":[...]}`.
- `ENDPOINT_RECRUITING` nel browser → deve dare `{"ok":true,"agenti":[...],...}`.
- Prova in locale senza doppio clic:
  ```bash
  cd C:\Users\andre\Documents\GitHub\sd-deal-tracker; python -m http.server 8080
  ```

---

## 6. File
```
sd-deal-tracker/
├── index.html                 ← cruscotto (2 tab, 2 fonti, password, Economics)
├── apps-script.gs             ← backend VENDITE (Master) + MODO_TEST
├── recruiting-apps-script.gs  ← backend RECRUITING (2° foglio)
├── data/regioni-min.json      ← GeoJSON Italia
├── netlify.toml · _redirects · .gitignore
└── README.md
```
