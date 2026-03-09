# OFFERTA DI SVILUPPO — Hanna Statistics
**Gibertini Elettronica S.r.l. → Hanna Instruments SRL Nusfalau**
**Data**: Marzo 2026 | **Revisione**: 0.1 Draft

---

## 1. Il Sistema Gibertini presso Hanna Instruments

### 1.1 Panoramica dell'Ecosistema Software

Presso lo stabilimento Hanna Instruments SRL di Nusfalau sono attualmente in uso **tre applicativi**
installati su postazioni Windows dedicate in produzione e laboratorio:

| Software | Versione | Contesto d'uso | Utenti |
|---|---|---|---|
| **Chemical Production** | v1.3.24 | Reparto produzione — reagenti in polvere/tablet | Tecnici di produzione |
| **Chemical QC** | v1.2.31 | Laboratorio QC chimico — letture strumentali | Operatori QC |
| **Chemical MR** | v1.1.40 | Produzione buffer (soluzioni tampone pH) | Tecnici produzione buffer |

Tutti e tre i software sono **installati localmente** su postazioni Windows dedicate,
condividendo il database dei codici prodotto (`dbCode.mdb`).

**Chemical MR** è un applicativo parallelo a Chemical Production, dedicato
alla produzione delle soluzioni buffer. Ha architettura identica: gestione ricette, formulazioni,
preparazioni, proprio ciclo QC interno con correzioni CM, e export Excel stile Chemical QC.
Database proprio: `dbChemiMR.mdb` + `dbCode.mdb` condiviso.

### 1.2 Chemical Production — Funzionalità

Il software gestisce l'intera pipeline di produzione chimica:

```
Ricette (Recipes)
    → Formulazione componenti (materie prime, quantità, set reagenti)
    → Richiesta Materiali (Material Requisition)
    → Preparazione batch (Preparation)
        ↳ Export file PREP .xlsx → salvato localmente sul PC di produzione
        ↳ Stampa etichette Brother (numero lotto, QR code)
    → Stampa QR code "to QC"
        ↳ Etichetta fisica che accompagna i campioni al laboratorio Chemical QC

    ┌─────────────────────────────────────────────────────┐
    │  CICLO ITERATIVO — finché non Passed                │
    │                                                     │
    │  [Chemical QC valuta il campione — vedi sez. 1.3]  │
    │      ↓                                             │
    │  Se PASSED → ciclo terminato                       │
    │  Se FAILED → campione torna in Chemical Production │
    │      ↓                                             │
    │  Preparation QC (modulo in Chemical Production)    │
    │      ↳ Registra: Correction Date, Correction,      │
    │        QC Operator                                  │
    │      ↓                                             │
    │  Nuova preparazione / correzione batch             │
    │      ↓                                             │
    │  Stampa nuovo QR code "to QC" → laboratorio       │
    └─────────────────────────────────────────────────────┘

    → [esito finale Passed] Scan QR code "QC closed" in Chemical Production
        ↳ Acquisizione esito definitivo dal laboratorio
    → Export Hanna Code → aggiornamento dbCode.mdb
        ↳ Database condiviso con tolleranze/codici prodotto (locale, non passa a QC direttamente)
```

### 1.3 Chemical QC — Funzionalità

Il software gestisce il ciclo di qualifica QC di ogni lotto in laboratorio:

```
Ricezione campione dalla produzione (scan QR code "to QC")
    → Information QC: inserimento/verifica metadata lotto
    → Reading QC: inserimento letture strumentali per ogni STD
        ↳ Fino a 4 meter (M1-M4) con MIN/MAX per ogni standard value
        ↳ pH 1/2/3, Turbidità, Peso [mg]
        ↳ Colonna CORRECTION: esclusione/correzione di singole letture outlier
    → Evaluation QC: selezione letture valide (≥70-80%), calcolo medie
        ↳ Esito finale lotto: Passed / Failed
        ↳ STD Mean Value per ogni standard salvato nel DB
    → Graph QC: grafico medie per meter (Gaussiana interna a Chemical QC)
    → Certificate: generazione certificato lotto PDF
    → Export Excel .xlsx → file QC lotto salvato localmente (PC laboratorio)
        ↳ Naming: CP-R{n}_{HannaCode}_{STDval}_LOT{n}_PWW{week}_{year}_QC.xlsx
    → Stampa QR code "QC closed" → etichetta di ritorno alla produzione
```

Il file `.xlsx` esportato da Chemical QC è la **fonte primaria** del modulo Reagenti QC Statistics.

> **Nota**: la colonna CORRECTION in Chemical QC riguarda esclusivamente le letture strumentali
> (outlier, riletture). Le correzioni di produzione buffer (CM, grammi) sono gestite da
> Chemical MR — vedi sezione 1.4.

### 1.4 Chemical MR — Funzionalità

Chemical MR gestisce la produzione delle **soluzioni buffer** (pH 1.68, 4.01, 7.01, 10.01).
L'architettura è parallela a Chemical Production, con lo stesso pattern operativo:

```
Ricette buffer (Recipes)
    → Formulazione componenti (materie prime, quantità, density)
    → Preparazione batch (Preparation)
        ↳ Export file PREP .xlsx → salvato localmente
    → [esito Passed] Export .xlsx per lotto buffer
        ↳ Contiene tabella MR con scostamenti per standard:
```

**Struttura tabella MR nel file export Chemical MR** (da confermare):

| STD Number | STD Value | MR Qty | MR Acquired | Variance | Variance Perc | Note |
|---|---|---|---|---|---|---|
| Numero standard | Valore target | Quantità pianificata | Quantità effettiva | Scostamento [g] | Scostamento [%] | Note |

Questi valori alimentano il **Buffer Statistic Excel** aggregato:
- `corr CM [g]` ← `Variance` (scostamento assoluto in grammi)
- `corr CM [%]` ← `Variance Perc` (scostamento percentuale)
- `Corr Gibertini [%]` ← verosimilmente `Variance Perc` rinominato (calcolato da Chemical MR)

> I file `PREP_HI*.MR*.*` (e simili) esportati da Chemical MR sono la **fonte primaria
> del Modulo 2**, esattamente come i file Chemical QC sono la fonte del Modulo 1.
> La struttura interna della tabella MR è da confermare con Hanna prima dell'implementazione
> del parser.

### 1.5 Flusso di Dati tra i Tre Software

```
┌─────────────────────────┐         ┌─────────────────────────┐
│   CHEMICAL PRODUCTION   │         │      CHEMICAL QC        │
│   VB6 v1.3.24           │         │   VB6 v1.2.31           │
│   reagenti polvere/tab  │         │   laboratorio QC        │
│                         │         │                         │
│  dbChemiProd.mdb        │         │  dbChemicalQC.mdb       │
│  dbFormule.mdb          │         │                         │
│  dbCode.mdb ────────────────────────────────► dbCode.mdb   │
│  (shared)               │         │  (shared — tolleranze)  │
│                         │         │                         │
│  Export PREP .xlsx ─────────── locale ──────────────────►  │
│  QR code "to QC" ──────── etichetta fisica ──────────────► │
│  QR code "QC closed" ──── etichetta fisica ◄────────────── │
│                         │         │                         │
│                         │         │  Export QC .xlsx ─────► │
│                         │         │  Statistic_HI782-0.xlsx │
└─────────────────────────┘         └─────────────────────────┘
                                              │
                                    (manuale/cartella rete)
                                              │
┌─────────────────────────┐                  │
│   CHEMICAL MR           │                  ▼
│   VB6 v1.1.40           │         ┌─────────────────────────┐
│   produzione buffer pH  │         │    HANNA STATISTICS     │
│                         │         │    (import .xlsx)       │
│  dbChemiMR.mdb          │         │    DB centralizzato     │
│  dbCode.mdb (shared)    │         │                         │
│                         │         │  ← QC lots (Chem. QC)   │
│  Ciclo QC buffer:       │         │  ← Buffer Statistic     │
│  Correction + CM ───►   │         │     (Chem. MR)          │
│  Buffer Statistic.xlsx  │         │  ← PREP files           │
│  (Corr Gibertini [%])   │         │     (Chem. Prod + MR)   │
└─────────────────────────┘         └─────────────────────────┘
         │ (manuale/cartella rete)
         └────────────────────────────────────►
```


## 2. Considerazioni Tecnologiche

### 2.1 Stato Attuale: VB6 e MS Access

I tre software sono basati su tecnologia **Visual Basic 6** con database **MS Access (.mdb)**. Questa architettura:

- Funziona correttamente e stabilmente in produzione
- È familiare agli utenti attuali
- **Non è web-based**: non accessibile da browser o dispositivi mobili
- Scalabilità limitata: un'istanza per postazione fisica
- Aggiornamenti richiedono distribuzione del .exe su ogni macchina (SmartUpdate)

### 2.2 Architettura Proposta: Software Desktop con Percorso verso Multi-Utente

Per la **V1**, Hanna Statistics viene sviluppato come **applicativo desktop standalone**,
installato su una singola postazione dedicata (PC del responsabile produzione o laboratorio).
Questa scelta minimizza i rischi, riduce i costi di infrastruttura e permette
un avvio rapido senza dipendenze da server o configurazioni di rete complesse.

**Import dei dati in V1 — due scenari:**

| Scenario | Modalità import |
|---|---|
| **A — Cartella di rete condivisa** | Se esiste una cartella di rete visibile da produzione, QC e PC Statistics: copia automatica o semi-automatica dei file .xlsx esportati |
| **B — Import manuale** | Chiavetta USB o copia diretta: l'operatore copia i file .xlsx sul PC Statistics e avvia l'import |

In entrambi i casi il processo di import (parsing + salvataggio DB) è identico;
cambia solo il modo in cui i file arrivano al PC.

**Percorso evolutivo verso multi-utente (V2):**

L'architettura è progettata per poter essere migrata su server e resa accessibile
via browser da qualsiasi PC della rete interna, con modifiche minime al codice applicativo:

```
V1 — Desktop standalone              V2 — Server + browser (eventuale)
┌────────────────────────┐           ┌──────────────────────────────────┐
│  PC Responsabile       │           │  Server Linux LAN Nusfalau       │
│  Windows               │    →→→    │  (es. 192.168.1.36:83)           │
│  Hanna Statistics      │           │  Hanna Statistics                │
│  MariaDB locale        │           │  MariaDB                         │
│  Import file manuale   │           │  Import da cartella di rete      │
└────────────────────────┘           └──────────────────────────────────┘
                                              ▲ browser
                                     Qualsiasi PC sulla LAN
```

### 2.3 Perché NON Integrare le Statistiche nei Software Esistenti

| | Chemical Production / QC / MR | Hanna Statistics |
|---|---|---|
| **Chi lo usa** | Tecnici produzione, operatori QC, tecnici buffer | Responsabili produzione, responsabili lab |
| **Dove** | Postazione fissa dedicata (una per software) | PC dedicato o qualsiasi PC in rete |
| **Frequenza** | Uso operativo continuativo (ogni lotto) | Uso periodico per analisi e report |
| **Tipo di utilizzo** | Inserimento dati in tempo reale | Consultazione, analisi trend, decisioni |

Aggiungere un modulo statistico complesso ai tre software VB6 esistenti comporterebbe:
- Rischi di stabilità per applicativi critici di produzione che non possono fermarsi
- Complessità inappropriata nell'architettura single-user + MS Access
- Impatto sugli aggiornamenti futuri dei software operativi

### 2.4 Prospettiva Futura: Integrazione con Hanna Core

Una volta maturata l'esperienza con Hanna Statistics V1, il sistema può evolversi
verso l'integrazione con **Hanna Core** (ERP/MES dello stabilimento) come modulo web:

- Database `hanna_statistics` affiancato al DB Hanna Core sullo stesso server
- API REST già presenti → integrazione senza modificare Chemical Production/QC
- Anagrafica prodotti (`hanna_codes`) sincronizzabile automaticamente con Hanna Core
- **Nessuna modifica** necessaria a Chemical Production o Chemical QC:
  il flusso di export dei file QC .xlsx rimane invariato in qualsiasi scenario

---

## 3. Premessa

Il presente documento descrive l'analisi dei requisiti di Hanna Instruments SRL per lo sviluppo
del software **Hanna Statistics**, sistema di analisi statistica e controllo qualità per il
laboratorio QC chimico dello stabilimento di Nusfalau.

Il software sostituisce l'attuale flusso manuale basato su fogli Excel con Power Query, offrendo
un sistema centralizzato con import automatico dei dati, visualizzazione interattiva e
archiviazione strutturata su database.

**V1**: applicativo desktop standalone su singolo PC dedicato (responsabile produzione o laboratorio).
**V2 eventuale**: migrazione su server LAN, accessibile da qualsiasi postazione tramite browser.

---

## 4. Analisi dei Dati Reali

Prima di definire i requisiti funzionali, sono stati analizzati i file Excel attualmente in uso
presso Hanna Instruments per comprendere esattamente la struttura dei dati, i volumi e la
complessità del sistema esistente.

### 4.1 File Analizzato: Statistic_HI782-0_rev0.0_2025.06.03.xlsx

> Statistica QC del prodotto **HI782-0** — Marine Nitrate HR Reagent (CP-R80, Line L57 Powder)

**Struttura del foglio Excel QC statistiche** (5 sheet):

| Sheet | Contenuto |
|---|---|
| **STD All** | Calcoli principali e grafico (133 colonne, foglio molto esteso) |
| **HISTORY LOG** | Log modifiche ai file raw |
| **Lot structure** | Verifica allineamento colonne meter tra lotti |
| **Import Lots Data** | Letture QC raw aggregate da tutti i file lot |
| **Import Hanna Code** | Specifiche prodotto e tolleranze da dbCode? |

**Dati reali HI782-0**:

| Parametro | Valore |
|---|---|
| Prodotto | HI782-0 — Marine Nitrate HR Reagent |
| Ricetta | CP-R80 |
| Linea | L57 Powder |
| STD values | **4 standard**: 0, 15, 35, 60 ppm |
| Sigma per STD | σ(0)=1.00, σ(15)=1.3675, σ(35)=1.8575, σ(60)=2.47 |
| Numero lotti | **25 lotti** (LOT0366 → LOT0972) |
| Periodo | Gennaio 2024 – Marzo 2025 |
| TEST types presenti | ASTM, ASTM D665, EPP:170mg, EPP:171mg, EPP:172mg, EPP:250mg, Old A-D, P/Final, P/Prod, Valid |

**Osservazioni strutturali critiche**:
- Il foglio STD All si estende su **133 colonne** per contenere: selezione filtri (col 1-10),
  letture raw con medie per ogni lot (col 51-58), calcolo running averages e limiti sigma per
  ogni lot per ogni STD (col 60-133)
- I **TEST types** sono più numerosi di quanto indicato nei requisiti formali (non solo P/Final
  e P/Prod, ma anche ASTM, EPP e categorie "Old" per dati storici)
- Le **bande sigma** sono calcolate come valori assoluti per ogni STD: es. per STD=15,
  le bande vanno da 15-3σ (10.33 ppm) a 15+3σ (19.10 ppm)
- Il calcolo include **running averages** cumulativi: la media di ogni lot include tutti i
  lotti precedenti, non solo quello corrente

### 4.2 File Analizzato: Buffer Preparation and Correction Statistic 2024-2025_rev.0.1_2025.09.15.xlsx

> Statistica produzione buffer (soluzioni tampone pH) — aggregazione dati da **Chemical MR** v1.1.40

**Struttura**: 4 sheet pH + 1 sheet lookup materiali

| Sheet | Ricetta | Lotti | Correzioni | Note |
|---|---|---|---|---|
| **pH 1.68** | C158 | 18 | 6 | Volume medio ~350 Kg/lotto |
| **pH 4.01** | C163 | 175 | 54 | Maggiore volume, 31% correzioni |
| **pH 7.01** | C164 | 168 | 33 | 20% correzioni |
| **pH 10.01** | C166 | 88 | 44 | 50% correzioni — più critico |
| **CM info** | — | — | — | Lookup 8 materiali correttivi |
| **TOTALE** | | **449 lotti** | **137 correzioni** | Periodo: Gen 2024 – Set 2025 |

**8 Materiali Correttivi (CM info sheet)**:

| CM Code | Nome | CAS | Usato per |
|---|---|---|---|
| CM011 | Potassium Tetraoxalate Dihydrate | 6100-20-5 | pH 1.68 |
| CM021 | Potassium Hydrogen Phthalate | 877-24-7 | pH 4.01 |
| CM131 | Hiamine | 121-54-0 | pH 4.01 |
| CM480 | Potassium Hydroxide 90% | 1310-58-3 | pH 4.01 |
| CM022 | Potassium Dihydrogen Phosphate | 7778-77-0 | pH 7.01 |
| CM023 | Di-Sodium Hydrogen Phosphate | 7558-79-4 | pH 7.01 |
| CM029 | Sodium Carbonate | 497-19-8 | pH 10.01 |
| CM030 | Sodium Hydrogen Carbonate | 144-55-8 | pH 10.01 |

**Struttura dati per ogni lotto** (colonne del file):
`CODE` (lista prodotti separati da `;`) | `LOT` | `DATE` | `QUANTITY PRODUCED [Kg]` |
`FIRST QC FAILED` (valore pH al primo test fallito) | `CM USED FOR CORRECTION` (descrizione interventi) |
`CM Code` | `Corr Gibertini [%]` | `corr CM [%]` | `corr CM [g]`

### 4.3 Stima Volumi File per Anno

Basandosi sui file di esempio forniti (set parziale della produzione reale):

**Chemical QC — file .xlsx per lotto:**

| Dato | Valore |
|---|---|
| File nel set di esempio | 392 file |
| Codici prodotto distinti | 99 codici (nel set di esempio) |
| Range numeri lotto | LOT0325 → LOT1744 (span: ~1419 lotti) |
| Periodo coperto (da date) | ~16 mesi (Dic 2023 – Apr 2025) |
| **Stima lotti QC per anno** | **~1.000–1.100 file .xlsx/anno** |
| Prodotto più attivo | HI93701-0: 53 file nel set di esempio |

**Chemical Production + Chemical MR — file PREP .xlsx per batch:**

| Dato | Valore |
|---|---|
| File PREP nel set di esempio (CP + MR) | 329 file .xlsx (da ripartire tra i due software) |
| File PREP raw (.prep) 2025 | 223 file in ~17 settimane (gen–apr) |
| **Stima batch PREP per anno (totale)** | **~650–700 file .xlsx/anno** |

> I file `PREP_CP-B*` provengono da Chemical Production; i file `PREP_B*-SOL.*`, `PREP_B*-A/B.*`
> e `PREP_HI*.MR*.*` provengono da Chemical MR. Il parser è identico per entrambi.

**Totale stimato: ~1.700–1.800 file .xlsx/anno** da gestire tra i tre software.

### 4.4 Statistiche Letture QC — Caso HI782-0

Analisi delle letture nel file `Statistic_HI782-0_rev0.0_2025.06.03.xlsx`:

| Parametro | Valore |
|---|---|
| Letture totali (tutti i lotti) | **2.956 readings** |
| Numero lotti | 25 lotti (Gen 2024 – Mar 2025) |
| Media letture per lotto | ~118 letture/lotto |
| Range lotto: min/max | 36 (LOT piccolo) – 295 (LOT0869) letture |
| Standard 0 ppm | 225 letture (7.6%) |
| Standard 15 ppm | 853 letture (28.9%) |
| Standard 35 ppm | 1.013 letture (34.3%) |
| Standard 60 ppm | 865 letture (29.3%) |

**Distribuzione letture nei lotti più grandi** (numero letture totali):

| Lotto | Data | STD0 | STD15 | STD35 | STD60 | Totale |
|---|---|---|---|---|---|---|
| LOT0869 | 2024-12-12 | 14 | 73 | 107 | 76 | **270** |
| LOT0941 | 2025-02-20 | 26 | 85 | 89 | 95 | **295** |
| LOT0970 | 2025-03-19 | 15 | 67 | 84 | 69 | **235** |
| LOT0667 | 2024-08-27 | 2 | 65 | 86 | 70 | **223** |
| LOT0568 | 2024-05-13 | 10 | 52 | 55 | 51 | **168** |

> Questi volumi indicano che prodotti ad alta frequenza di testing possono accumulare
> 200-300 letture per lotto. Con ~1.000 lotti/anno, la stima è **~120.000 nuove letture/anno**
> solo dal modulo Chemical QC (al ritmo attuale del campione esaminato).

---

## 5. Analisi Requisiti Ufficiali

> Fonte: *Procedure Statistic Analysis QC readings — Rev. 0.0 Draft*
> Hanna Instruments SRL Nusfalau

### 5.1 Scopo e Campo di Applicazione

Il sistema gestisce le **statistiche dei reagenti in polvere** del laboratorio QC chimico.

Obiettivi:
- Visualizzare il trend di evoluzione dei valori medi tra lotti successivi
- Verificare la **precisione** e la **ripetibilità** delle preparazioni
- Fornire una valutazione della qualità di ogni lotto rispetto alle tolleranze Hanna

### 5.2 Terminologia e Abbreviazioni

| Sigla | Definizione |
|---|---|
| STD | Reference Standard (valore di riferimento calibrazione) |
| M1, M2, M3, M4 | Meter 1-4: strumenti di misura utilizzati durante il QC |
| **σ** | **50% della Hanna Tolerance** per ogni valore standard (NON deviazione statistica) |
| QC | Quality Control |
| P/Final | Test type: misura finale di produzione |
| P/Prod | Test type: misura intermedia di produzione |

> **DEFINIZIONE CRITICA**: σ è un valore **fisso**, calcolato per ogni STD come 50% della
> tolleranza Hanna del prodotto.

### 5.3 Fonte Dati: File Lot da Software Gibertini

I file raw sono generati dal software Gibertini (.xlsx, un file per lotto):
- **Naming**: `CP-R{recipe}_{HannaCode}-{version}_{STDvalue}_LOT{number}_PW{week}.{n}_QC.xlsx`
- **Contenuto**: metadata prodotto, specifiche reagenti, configurazione meters, tabella letture raw

**Nota critica — Nomi colonne meter**: i nomi variano tra file (`METER 1` vs `M. 1 HI801`).
Il parser deve identificare le 4 colonne meter per **posizione** (colonne 12-15), non per nome.

### 5.4 Flusso Operativo Attuale (da automatizzare)

```
Gibertini software
    → esporta file .xlsx per ogni lotto → cartella LOTS/
    → operatore QC apre foglio Excel statistiche
    → aggiorna Power Query manualmente (3 sheet raw)
    → verifica allineamento meter (Lot structure sheet)
    → seleziona TEST type e STD values (max 6)
    → attiva calcolo automatico
    → visualizza grafico + tabella
```

### 5.5 Requisiti Grafico (Sezione 8 del documento ufficiale)

Il grafico principale mostra **3 serie sovrapposte** per ogni STD selezionato:

| Serie | Tipo chart | Descrizione |
|---|---|---|
| Letture raw | **Scatter plot** | Tutti i punti per lotto, numerati 1 by 1, colonne separate per lotto |
| Medie lot-by-lot | **Scatter + linea retta** | Running average lot per lot (cumulativo) |
| Bande sigma | **Stacked area** | 1σ verde, 2σ blu, 3σ giallo (centrate sul STD value) |

**Calcolo bande sigma**:
- 1σ band = STD Value ± (50% × Hanna Tolerance) → **colore VERDE**
- 2σ band = STD Value ± (100% × Hanna Tolerance) → **colore BLU**
- 3σ band = STD Value ± (150% × Hanna Tolerance) → **colore GIALLO**

Ogni STD ha le proprie bande (la tolleranza varia per STD value).

### 5.6 Requisiti Tabella Statistica (Sezione 8.2)

Per ogni lotto, per ogni STD:

| Colonna | Contenuto |
|---|---|
| Nr. | Numero progressivo lotto |
| Lot | Numero lotto |
| QC Date | Data QC |
| QC Week | Settimana di produzione |
| Grand total tests | Totale test nel lotto (tutti STD) |
| Per ogni STD: Total tests | Test validi per quel STD |
| Per ogni STD: `< ±1σ` | Count e % letture entro 1σ |
| Per ogni STD: `±1σ to ±2σ` | Count e % letture tra 1σ e 2σ |
| Per ogni STD: `±2σ to ±3σ` | Count e % letture tra 2σ e 3σ |
| Per ogni STD: `> ±3σ` | Count e % letture fuori 3σ |

### 5.7 Filtri e Selezioni

- **TEST type**: selezionabile (P/Final, P/Prod, ASTM, EPP:xxx, Old, Valid...)
- **STD values**: fino a 6 standard per prodotto, selezionabili in ordine crescente
- **Prodotto (Hanna Code)**: selezione del codice prodotto

---

## 6. Proposta di Sviluppo — Hanna Statistics V1

### 6.1 Confronto: Sistema Attuale vs Nuovo Software

| Problema attuale (Excel + Power Query) | Soluzione Hanna Statistics |
|---|---|
| Aggiornamento dati manuale (Power Query) | Import automatico file .xlsx |
| Un file Excel per ogni codice prodotto | Database centralizzato, tutti i prodotti |
| Non accessibile in contemporanea da più utenti | Multi-utente, accesso da qualsiasi PC in rete |
| Allineamento meter manuale per ogni lotto | Parser automatico per posizione colonna (col. 12-15) |
| Nessuna storicità aggregata multi-prodotto | Trend multi-anno, comparazione tra prodotti |
| Nessun controllo accessi | Login con ruoli (operatore QC / manager / admin) |
| Foglio con 133 colonne e calcoli complessi | Calcoli automatici su database, UI semplice |

### 6.2 Moduli Previsti

#### Modulo 1 — Reagenti QC Statistics (PRIORITÀ MASSIMA)
*Digitalizzazione diretta dei requisiti ufficiali*

**Funzionalità**:
- Import file QC lot da Gibertini (upload singolo o batch)
- Parsing automatico con gestione nomi meter variabili (identificazione per posizione)
- Control Chart (Shewhart): scatter raw + running averages + bande σ
- Tabella distribuzione sigma per lotto (con count e percentuali per ogni banda)
- Filtri: Hanna Code, TEST type, STD value, range date
- Export tabella in CSV/Excel

**Calcolo automatico**:
- σ = 50% × Hanna Tolerance (per ogni STD, fisso)
- Classificazione ogni lettura: <1σ / 1-2σ / 2-3σ / >3σ
- Running average cumulativo lot-by-lot
- Limiti bande: 3σ low/2σ low/1σ low/1σ high/2σ high/3σ high per ogni STD

#### Modulo 2 — Buffer Production Statistics (ESTENSIONE PROPOSTA)
*Basato sui dati reali analizzati: 449 lotti, 4 valori pH, 8 materiali CM — da confermare*

**Funzionalità**:
- Import file Excel Buffer Statistic (4 sheet pH + CM info)
- Trend quantità prodotta per mese/ricetta/pH
- Analisi correzioni: frequenza per pH, tipologia CM, quantità media per lotto
- KPI: % lotti con correzione, media corr CM [g], tasso correzione per periodo

#### Modulo 3 — Preparation List (ESTENSIONE PROPOSTA)
*File PREP esportati da **Chemical Production** (reagenti) e **Chemical MR** (buffer)*

**Due sorgenti di file PREP — stesso formato, software diverso:**

| Tipo file | Software sorgente | Naming pattern | Contenuto |
|---|---|---|---|
| `PREP_CP-B*.L56.CTK.*` | Chemical Production | Ricetta CP-B + linea | Prodotti finiti (reagenti, polveri) |
| `PREP_B*-SOL.A/B.*`, `PREP_B*-A/B.*` | **Chemical MR** | Ricetta B + soluzione | Soluzioni buffer intermedie |
| `PREP_HI*.MR*.*`, `PREP_DPD*.MR*.*` | **Chemical MR** | HannaCode + ricetta MR | Standard/reagenti via MR |

**Funzionalità**:
- Import file PREP da entrambi i software (parser identico — stesso layout fisso)
- Tracciabilità lotto: componenti, operatore, date preparazione
- Confronto settimana pianificata vs effettiva
- Statistiche per ricetta, software sorgente, mese

### 6.3 Dashboard Riepilogativa

KPI cards:
- Lotti QC importati (ultimi 30 giorni)
- % letture entro 1σ (ultimi 30 giorni)
- Lotti buffer con correzione (mese corrente)
- Alert: lotti con letture fuori 3σ

---

## 7. Architettura Tecnica

| Layer | Tecnologia |
|---|---|
| Interfaccia utente | React 19 + TypeScript + Vite |
| Componenti UI | shadcn/ui + Tailwind CSS 4 |
| Grafici | Apache ECharts |
| Tabelle | TanStack Table v8 |
| Backend / API | Node.js 20 + Express |
| Database | MariaDB 10.11+ (schema dedicato `hanna_statistics`) |
| Accesso rete | Nginx (reverse proxy + HTTPS self-signed) — solo V2 |
| Deploy | Docker Compose |
| Multilingua | EN, IT, RO, HU |

**V1 — Desktop**: MariaDB installato localmente su PC Windows (servizio Windows, nessun server richiesto).
**V2 eventuale**: stessa istanza MariaDB migrabile su server Linux LAN senza modifiche allo schema.

> **Nota MariaDB**: MariaDB è un fork drop-in di MySQL — stessa sintassi SQL, stesso driver JDBC/ODBC,
> compatibilità binaria. Una eventuale migrazione da MySQL esistente a MariaDB locale (V1)
> è trasparente e non richiede adattamenti al codice applicativo.

---

## 8. Schema Database (sintesi)

15 tabelle nel database `hanna_statistics`:

| Tabella | Contenuto |
|---|---|
| `hanna_codes` | Anagrafica prodotti (Hanna Code, descrizione, ricetta) |
| `product_standards` | Configurazione STD e σ per prodotto |
| `production_lots` | Lotti di produzione (numero, date, operatori) |
| `qc_readings` | Letture raw QC (M1-M4, pH, turb, peso) |
| `lot_sigma_distribution` | Cache distribuzione sigma (performance) |
| `lot_running_averages` | Cache running averages per control chart |
| `control_chart_limits` | Limiti σ (3σ/2σ/1σ low+high) per prodotto per STD |
| `buffer_production` | Dati produzione buffer (quantità, correzioni CM) |
| `correction_materials` | Lookup 8 materiali CM (CM011, CM021... CM030) |
| `preparation_batches` | Batch preparazione (da file PREP) |
| `preparation_hanna_codes` | Prodotti per batch (relazione N:M) |
| `users` | Utenti e ruoli |
| `audit_log` | Log operazioni |
| `app_config` | Configurazione applicazione |
| `import_log` | Log import file |

---

## 9. Fasi di Sviluppo

### Fase 1 — Core QC Statistics
Setup infrastruttura + Modulo Reagenti QC completo:
- Import file Chemical QC con parser automatico
- Control Chart (Shewhart) con bande σ e running averages
- Tabella distribuzione sigma per lotto
- Gestione Hanna Codes + product standards

### Fase 2 — Buffer & Preparation
- Modulo Buffer Production Statistics (4 pH, 8 CM, 449 lotti storici)
- Modulo Preparation List Statistics
- Dashboard con KPI aggregati

### Fase 3 — Reporting & Finalizzazione
- Export report PDF/Excel
- Ottimizzazioni performance (cache pre-calcolata, indici DB)
- Multilingua completo (EN/IT/RO/HU)
- Deploy finale con HTTPS

---

## 10. Analisi Aggiuntive: Valore dai Dati Disponibili

Oltre ai requisiti ufficiali, i file già analizzati contengono dati sufficienti per analisi
di maggior valore diagnostico. Di seguito le proposte ordinate per priorità/effort.

### 10.1 Analisi Aggiuntive — Modulo Reagenti QC

#### A. Confronto Inter-Meter (M1 vs M2 vs M3 vs M4)
I file QC contengono letture separate per ogni meter.
I requisiti ufficiali le aggregano; noi possiamo mostrarle in modo disaggregato:

| Analisi | Valore |
|---|---|
| Bias sistematico per meter | Es. M2 legge costantemente +0.8 ppm sopra M1 → revisione strumento |
| Deviazione standard per meter | Quale meter è più preciso? |
| Trend degradazione meter | Il meter X sta diventando sempre più fuori allineamento nel tempo? |
| Heatmap letture per meter × STD | Visione immediata del comportamento di ogni strumento |

> **Dato disponibile**: colonne M1/M2/M3/M4 già presenti in ogni file .xlsx QC.
> Non richiede import aggiuntivi.

---

#### B. Analisi per TEST Type
I file contengono TEST type (P/Final, P/Prod, ASTM, EPP:170mg...) ma i requisiti usano
solo uno o due tipi per l'analisi principale. Comparazione utile:

| Analisi | Valore |
|---|---|
| Distribuzione sigma per TEST type | P/Final è più preciso di P/Prod per lo stesso lotto? |
| Trend separati per TEST type | ASTM vs P/Final: lo stesso lotto visto da due angoli |
| % letture fuori 3σ per TEST type | Quali tipi di test sono più critici? |

---

#### C. Analisi Peso Campione [mg]
Il campo `WEIGHT [mg]` è presente nelle letture ma non utilizzato nei requisiti ufficiali.

| Analisi | Valore |
|---|---|
| Correlazione peso ↔ reading value | La variazione di peso nella pesata influenza il risultato? |
| Scatter peso vs deviazione da STD | Identificare se pesate troppo lontane dalla target producono outlier |
| Range peso per lotto | Operatore X pesa sempre con più variabilità? |

---

#### D. Stabilità Intra-Lotto (Convergenza)
Le letture di un lotto sono raccolte in sessioni diverse (giorni/turni diversi).
Analisi della convergenza della media man mano che si aggiungono letture:

| Analisi | Valore |
|---|---|
| Running mean per lotto (intra-lot) | Dopo quante letture la media si stabilizza? |
| Prima metà vs seconda metà del lotto | Le ultime letture confermano le prime? |
| Variabilità per turno/data di lettura | Se la data è disponibile: letture del mattino vs pomeriggio |

---

#### E. Previsione Trend (Early Warning)
Con dati su 25+ lotti cumulativi, è possibile estrarre:

| Analisi | Valore |
|---|---|
| Regressione lineare sulla running average | La media sta driftando verso ±2σ? → alert preventivo |
| Rate of change lot-by-lot | Il lotto corrente si sta avvicinando al limite più velocemente del solito? |
| Flag "lotto a rischio" | Se trend attuale continua, il lotto N+2 potrebbe essere fuori 2σ |

---

### 10.2 Analisi Aggiuntive — Modulo Buffer Production

#### F. First-Pass Yield (FPY) per Ricetta
KPI industriale standard: % lotti che passano la QC al **primo tentativo senza correzioni**.

| Ricetta | Lotti totali | No correzioni | FPY |
|---|---|---|---|
| C158 (pH 1.68) | 18 | 12 | **66.7%** |
| C163 (pH 4.01) | 175 | 121 | **69.1%** |
| C164 (pH 7.01) | 168 | 135 | **80.4%** |
| C166 (pH 10.01) | 88 | 44 | **50.0%** |

> pH 10.01 è il processo più critico: 1 lotto su 2 richiede correzione.
> Questo KPI è già calcolabile dai dati esistenti.

---

#### G. Analisi Stagionale delle Correzioni
I buffer di pH sono sensibili alla temperatura ambiente. Analisi mensile:

| Analisi | Valore |
|---|---|
| Tasso correzioni per mese | Più correzioni in estate/inverno? (T° laboratorio) |
| Quantità CM media per stagione | Le correzioni estive richiedono più grammi? |
| Confronto anno su anno | Gen 2024 vs Gen 2025: il processo è migliorato? |

---

#### H. Analisi per Materiale Correttivo (CM)
8 CM distinti, ognuno con un ruolo diverso nel processo di correzione:

| Analisi | Valore |
|---|---|
| Frequenza uso per CM | CM021 usato 40 volte, CM480 solo 5 → un CM è quasi mai necessario? |
| Quantità media CM per intervento | Correzioni con CM480 (KOH forte) sono in media più grandi → processo instabile? |
| CM co-occorrenza | Quali CM vengono usati insieme nello stesso lotto? |
| Trend quantità CM nel tempo | Le correzioni sono diventate più grandi nel 2025 rispetto al 2024? |

---

#### I. Analisi Batch Multi-Prodotto
La colonna `CODE` nel Buffer Statistic contiene **lista di prodotti separati da `;`**
(un batch di buffer può alimentare più prodotti finiti). Analisi possibile:

| Analisi | Valore |
|---|---|
| Quali prodotti finiti dipendono da un buffer corretto? | Se pH 4.01 è corretto → impatto su N prodotti |
| Impatto correzione su volume prodotto totale | Correzione = ritardo → quantità totale prodotta per settimana |

---

### 10.3 Analisi Cross-Modulo (QC + Buffer + PREP)

#### J. Tracciabilità Completa del Lotto
Collegando i tre moduli (PREP → Buffer → QC), è possibile ricostruire il ciclo di vita completo:

```
PREP batch (CP-R80.L57.CTK.001.W04.2024)
    ↓
Buffer pH 4.01 — lotto C163-LOT0175 — 2 correzioni (CM021 + CM480)
    ↓
QC lotto HI782-0 LOT0415 — 118 letture — 96.2% entro 1σ
```

| Analisi | Valore |
|---|---|
| Lead time PREP → QC Passed | Quanto tempo passa dalla preparazione al lotto approvato? |
| Correlazione correzioni buffer ↔ qualità QC | I lotti da buffer corretto hanno più letture fuori 2σ? |
| Operatore PREP vs qualità finale | L'operatore A produce lotti con meno correzioni dell'operatore B? |

> **Nota**: questo tipo di analisi richiede che il link PREP–Buffer–QC sia tracciato
> nel database. I file attuali lo permettono tramite numero lotto e codice ricetta.

---

### 10.4 Riepilogo Priorità Proposte

| # | Analisi | Fonte dati | Effort | Valore diagnostico |
|---|---|---|---|---|
| A | Confronto inter-meter M1-M4 | QC .xlsx | Basso | **Alto** — identifica problemi strumentali |
| F | First-Pass Yield per ricetta | Buffer Statistic | Basso | **Alto** — KPI produzione immediato |
| E | Early warning drift trend | QC .xlsx | Medio | **Alto** — prevenzione non conformità |
| G | Stagionalità correzioni buffer | Buffer Statistic | Basso | Medio — ottimizzazione processo |
| H | Analisi per CM materiale | Buffer Statistic | Basso | Medio — gestione materiali correttivi |
| B | Confronto TEST type | QC .xlsx | Basso | Medio — validazione metodi |
| C | Correlazione peso ↔ lettura | QC .xlsx | Medio | Medio — qualità pesatura |
| J | Tracciabilità PREP→Buffer→QC | Tutti | Alto | **Alto** — unico, molto differenziante |
| D | Stabilità intra-lotto | QC .xlsx | Medio | Basso-medio — ottimizzazione sessioni |
| I | Analisi batch multi-prodotto | Buffer Statistic | Medio | Medio — pianificazione produzione |

---

## 11. Domande da Allineare con Hanna Instruments

Prima dell'avvio dello sviluppo:

1. **Ambito V1**: solo Reagenti QC (requisiti ufficiali) o anche Buffer Production?
2. **TEST types**: quali tipi vanno supportati oltre P/Final e P/Prod? (ASTM, EPP, Old, Valid?)
3. **Utenti**: chi userà Hanna Statistics? Solo il responsabile laboratorio QC?
   Solo il responsabile produzione? Entrambi?
4. **Numero postazioni**: quante installazioni di Hanna Statistics sono previste?
   - Se **1 postazione**: standalone semplice, database locale
   - Se **2+ postazioni**: necessario allineamento/sincronizzazione dei dati
     (cartella di rete condivisa, o server centralizzato)
5. **Cartella di rete**: esiste già una cartella di rete accessibile da produzione,
   laboratorio QC e PC Statistics? (condiziona la modalità di import automatico)
6. **Domande/esigenze specifiche**: a quali domande Hanna Instruments vorrebbe che
   le statistiche rispondessero? (es. "qual è il prodotto più fuori tolleranza?",
   "il lotto X è migliorato rispetto al precedente?", "quante correzioni buffer facciamo in media?")
7. **Multilingua**: quali lingue per V1?
8. **Storico iniziale**: i dati pregressi (es. 25 lotti HI782, 449 lotti buffer) vanno importati?
9. **Separazione da Hanna Core**: il database deve rimanere completamente separato?
10. **Retention**: per quanti anni mantenere le letture raw?
11. **Struttura file export Chemical MR per lotto buffer**: confermare che i file `PREP_HI*.MR*.*`
    contengono la tabella `STD Number | STD Value | MR Qty | MR Acquired | Variance | Variance Perc | Note`
    e che questi sono la fonte primaria per il Modulo 2 (come i file Chemical QC lo sono per il Modulo 1).

---

## 12. Note Tecniche per il Parser Import

### File Chemical QC (Reagenti)
- Sheet unico: "Hanna Code"
- Metadata: righe 2-9 (lot, recipe, prep week, exp date, QC type...)
- Standards (fino a 3 per file, fino a 6 nel sistema): righe 36-40
- Header tabella letture: riga 44 (label), riga 45 (colonne), dati da riga 46+
- **Colonne meter: identificate per posizione (col 12-15), non per nome** (nomi variano tra file)
- Colonne chiave: Standard, STD_Value, TEST, QC_DATE, M1-M4, WEIGHT, QC_OPERATOR

### File Buffer Statistic (aggregazione da Chemical MR)
- Fonte: aggregazione dati Chemical MR v1.1.40 (da confermare struttura import)
- Multi-sheet: pH 1.68, 4.01, 7.01, 10.01 (dati) + CM info (lookup materiali)
- Header: riga 1 ("STATISTIC PRODUCTION", col E = ricetta), riga 2 (col E = pH value)
- Header colonne: riga 4; dati da riga 5
- Colonne: CODE (multi-valore separato da `;`), LOT, DATE, QUANTITY_KG, FIRST_QC_FAILED,
  CM_USED, CM_CODE, CORR_GIBERTINI_PCT, CORR_CM_PCT, CORR_CM_G
- `CORR_CM_G` ← `Variance` da tabella MR (scostamento assoluto grammi)
- `CORR_CM_PCT` / `CORR_GIBERTINI_PCT` ← `Variance Perc` da tabella MR (scostamento %)

### File Chemical MR per lotto (fonte primaria Modulo 2 — da confermare)
- File: `PREP_HI*.MR*.*` e simili (stesso formato PREP, software Chemical MR)
- Contiene tabella MR: `STD Number | STD Value | MR Qty | MR Acquired | Variance | Variance Perc | Note`
- Import per Modulo 2 analogo al Modulo 1: un file per lotto, parser automatico

### File PREP (Preparation) — Chemical Production e Chemical MR

**Stesso layout fisso**, software sorgente identificabile dal pattern del nome:

| Pattern | Software | Tipo |
|---|---|---|
| `PREP_CP-B{n}.L56.CTK.*` | Chemical Production | Prodotto finito con Hanna Code Table |
| `PREP_B{n}-SOL.A/B.L56.CTK.*` | **Chemical MR** | Soluzione buffer intermedia |
| `PREP_B{n}-A/B.L56.CTK.*` | **Chemical MR** | Componente buffer A o B |
| `PREP_{HannaCode}.MR{n}.*` | **Chemical MR** | Standard/reagente via Material Requisition |

**Layout comune** (colonna B come origine dati):
- r4: Recipe, Line, Rev, Exp
- r9: Preparation Date, Batch#
- r10: Planned Week / Actual Week
- r11: Note, Operator, Exp Date
- r17+: Component Table
- (solo CP): Hanna Code Table (1+ prodotti per batch)
