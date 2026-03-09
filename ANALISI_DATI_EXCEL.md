# ANALISI COMPLETA DEI FILE EXCEL STATISTICI - Hanna Statistics

> **Documento generato dall'analisi automatizzata dei file Excel attualmente in uso**
> Data analisi: 2026-02-12

---

## SOMMARIO ESECUTIVO

Dall'analisi dei 2 file Excel emergono **DUE TIPOLOGIE COMPLETAMENTE DIVERSE** di prodotti e statistiche:

| Aspetto | File 1: Buffer Preparation | File 2: HI782-0 (Reagenti) |
|---------|---------------------------|----------------------------|
| **Prodotto** | Soluzioni tampone pH | Reagenti chimici (polvere/tablet) |
| **Focus principale** | Correzioni materiale | Controllo statistico Sigma |
| **Metrica chiave** | % materiale correzione usato | Distribuzione nei band sigma |
| **Tipo analisi** | Tracking correzioni batch | Control Chart (Shewhart) |
| **Volume dati** | ~465 righe (4 pH) | ~2957 letture QC raw |
| **Struttura** | 1 file = 1 gruppo ricetta pH | 1 file = 1 Hanna Code SFG |

**ATTENZIONE SCALABILITA**: Questi 2 file rappresentano POCHI prodotti. Con ~1000 SFG codes, il sistema deve gestire milioni di record.

---

## FILE 1: Buffer Preparation and Correction Statistic 2024-2025

### 1.1 Struttura del File

**5 Fogli:**
- `pH 1.68` (22 righe dati)
- `pH 4.01` (179 righe dati)
- `pH 7.01` (172 righe dati)
- `pH 10.01` (92 righe dati)
- `CM info` (tabella lookup materiali di correzione)

### 1.2 Schema Dati per Foglio pH

Ogni foglio pH ha la stessa struttura identica:

| Colonna | Header | Tipo Dato | Descrizione | Esempio |
|---------|--------|-----------|-------------|---------|
| A | CODE | `string` (semicolon-separated) | Codici prodotto SFG | `HI5001-01;HI5001-02;HI5001-11` |
| B | LOT | `string` | Numero lotto | `LOT0015` |
| C | DATE | `date` | Data produzione | `2024-01-17` |
| D | QUANTITY PRODUCED [Kg] | `float` | Quantita prodotta in Kg | `230`, `575.8` |
| E | FIRST QC FAILED (excel record book) | `string/float` | Valore pH primo QC fallito | `1.704`, `pH 4.025` |
| F | CM USED FOR CORRECTION | `string` | Descrizione materiale correzione | `Potassium Tetraoxalate dihydrate` |
| G | CM Code | `string` | Codice materiale correzione | `CM011`, `CM480` |
| H | Corr Gibertini [%] | `float` | Scostamento % calcolato da Chemical MR (`Variance Perc`) | `0.052`, `0.18` |
| I | corr CM [%] | `float` (calcolato) | Percentuale correzione (formula Excel locale) | `0.00052...` |
| J | corr CM [g] | `float` | Grammi materiale correzione (`Variance` da Chemical MR) | `120`, `250` |

**Riga intestazione:** Row 1 = "STATISTIC PRODUCTION", Row 1 Col E = "RECIPE # Cxxx"
**Riga 2:** Col E = "pH Value: x.xx"
**Dati iniziano:** Row 5

### 1.3 Formule Chiave

```
// Colonna H — generata da Chemical MR (da confermare struttura esatta)
Corr Gibertini [%] = Variance Perc  (da tabella MR: (MR_Acquired - MR_Qty) / MR_Qty * 100)

// Colonna I — formula Excel locale nel Buffer Statistic
corr CM [%] = IF(corr_CM_g == 0, "", corr_CM_g / (QUANTITY_KG * 1000))
```

`corr CM [%]` calcola la percentuale del materiale di correzione rispetto al peso totale in grammi.
`Corr Gibertini [%]` è lo scostamento percentuale tra quantità acquisita ed erogata secondo Chemical MR.

### 1.4 Tabella CM Info (Materiali di Correzione)

| Nr | CM Code | CM Name | CAS Number | Used in pH |
|----|---------|---------|------------|------------|
| 1 | CM011 | Potassium Tetraoxalate dihydrate | 6100-20-5 | 1.68 |
| 2 | CM021 | Potassium Hydrogen Phthalate | 877-24-7 | 4.01 |
| 3 | CM131 | Hiamine | 58-96-8 | 4.01 |
| 4 | CM480 | Potassium Hydroxide 90% | 1310-58-3 | 4.01, 7.01 |
| 5 | CM022 | Potassium Dihydrogen Phosphate | 7778-77-0 | 7.01 |
| 6 | CM023 | Di-Sodium Hydrogen Phosphate | 7558-79-4 | 7.01 |
| 7 | CM029 | Sodium Carbonate | 497-19-8 | 10.01 |
| 8 | CM030 | Sodium Hydrogen Carbonate | 144-55-8 | 10.01 |

### 1.5 Peculiarita dei Dati

- **CODE e multi-valore**: I codici prodotto sono separati da `;` (fino a 12+ codici per riga)
- **QC Failed puo essere**: numero (es. `1.704`), testo (es. `pH 4.025`), vuoto
- **Correzione opzionale**: Molte righe non hanno correzione (CM vuoto)
- **Date range**: Gennaio 2024 - Settembre 2025
- **Quantita variabili**: da 10 Kg a 830 Kg per batch

---

## FILE 2: Statistic_HI782-0 (Marine Nitrate HR Reagent)

### 2.1 Struttura del File

**5 Fogli:**
- `STD All` (2460 righe x 133 colonne) - **FOGLIO PRINCIPALE STATISTICHE**
- `HISTORY LOG` (2 righe) - Changelog modifiche
- `Lot structure` (27 righe x 105 col) - Template struttura file QC per lotto
- `Import Lots Data` (2957 righe x 104 col) - **DATI RAW IMPORTATI**
- `Import Hanna Code` (208 righe x 104 col) - Configurazione prodotto

### 2.2 Identita Prodotto (Import Hanna Code)

| Campo | Valore |
|-------|--------|
| Hanna SFG Code | HI782-0 |
| Parameter Formula | NO3 (Nitrati) |
| Description | Marine Nitrate HR Reagent |
| Recipe | CP-R80 |
| Line | L57 Powder |
| QC Department | Laboratory |
| Registration Book | TABLET 3 |
| QC Type | Production |
| Reference Weight | 206 mg (min 196, max 216) |
| Reagent Sets | Configurazione Reagent A/B/C/D/E per Set 1 e Set 2 |

### 2.3 Dati Raw QC (Import Lots Data) - 2957 righe

| Colonna | Header | Tipo Dato | Descrizione | Esempio |
|---------|--------|-----------|-------------|---------|
| A | (Filename) | `string` | Nome file QC sorgente | `R080_HI782-0_0.0_LOT0366_PWW49.1_QC.xlsx` |
| B | Standard | `int` | Numero standard (1-6) | `1`, `2`, `3`, `4` |
| C | STD Value | `float` | Valore standard atteso | `0`, `15`, `35`, `60` |
| D | # | `int` | Numero progressivo test | `1`...`53` |
| E | TEST | `enum string` | Tipo di test | `Valid`, `Old A`, `Old B`, `P/Final`, `P/Prod` |
| F | QC DATE | `datetime` | Data controllo qualita | `2024-01-08` |
| G | QC TIME | `string/time` | Ora controllo | `18:31`, `15:57` |
| H | PROD. DATE | `datetime` | Data produzione | `2024-01-08` |
| I | PROD. TIME | `string/time` | Ora produzione | `14:00`, `08:10` |
| J | PROD. OPERATOR | `string` | Operatore produzione | `ERIKA`, `ELENA` |
| K | HEAD | `int` | Numero testa (linea produzione) | `2` |
| L | M.1 Reading | `float` | Lettura Misuratore 1 | `0.6`, `16.5`, `37.9` |
| M | M.2 Reading | `float` | Lettura Misuratore 2 | `38.5`, `14.5` |
| N | M.3 Reading | `float` | Lettura Misuratore 3 | (spesso vuoto) |
| O | M.4 Reading | `float` | Lettura Misuratore 4 | (spesso vuoto) |
| P | SPECTR. [ABS] | `float` | Spettrometro assorbimento | (spesso vuoto) |
| Q-S | pH 1/2/3 | `float` | Misure pH | `1.716`, `1.760` |
| T | TURB. | `float` | Torbidita | `8.72`, `1.55` |
| U | WEIGHT [mg] | `float` | Peso tablet in mg | `206`, `210`, `180` |
| V | REAGENT SET | `int/string` | Set reagenti usato | `1`, `2` |
| W | QC OPERATOR | `string` | Operatore QC | `Alla`, `DAVID`, `Kinga` |
| X | CORRECTION | `string` | Info correzione | (spesso vuoto) |
| Y | phNumber | `int` | Numero pH (?) | `3` |
| Z | STD | `int` | ID standard | `1`, `2`, `3`, `4` |
| AA | STD_ID | `string` | Identificativo standard | `SW4`, `35404`, `90453` |
| AB | NOTE | `string` | Note | `9.27%` |

**Osservazioni sui dati raw:**
- I modelli misuratore cambiano tra lotti: `HI97115`, `HI97105`, `HI97715`, `Label17`
- Le colonne M2-M4 sono spesso vuote (dipende dal numero di misuratori usati)
- I tipi TEST rappresentano: `Valid` = validazione, `Old A/B/C/D` = vecchi metodi, `P/Final` = produzione finale, `P/Prod` = produzione
- Date range: Dicembre 2023 - Gennaio 2025+
- `phNumber` sembra costante (3) per questo prodotto
- `CORRECTION` in nota (es. `9.27%`) indica correzione applicata

### 2.4 Foglio STD All - Statistiche Aggregate (2460 righe x 133 colonne)

Questo foglio e diviso in **6 ZONE FUNZIONALI** affiancate:

#### ZONA A: Configurazione (Colonne 1-10, Righe 1-15)

| Riga | Contenuto |
|------|-----------|
| 1-9 | Definizione TEST types: `ASTM`, `ASTM D665`, `EPP:170mg`...`EPP:250mg`, `Old A`...`Old D`, `P/Final`, `P/Prod`, `Valid` |
| 1 | Valori STD: Col2=0, Col3=15 (intervalli min-max per standard) |
| 2-5 | STD values: 0, 15, 35, 60 |
| 4 | Sigma σ: 1, 1.3675, 1.8575, 2.47 (per ogni STD) |
| 2 | QC restriction: 100%, Custom QC restriction |
| 3 | Tolerance: Fixed, 2.0 con operatore `&`/`or` e `4.9%` |

#### ZONA B: Distribuzione Sigma per Lotto (Colonne 6-46, da Riga 11)

Headers in riga 11:

| Colonne | Gruppo | Contenuto |
|---------|--------|-----------|
| 6-7 | ID | Nr., Lot, QC Date, QC Week |
| 8-9 | ID | QC Week |
| 10 | Totale | Grand total tests |
| 11-19 | **STD 1 (0)** | Total tests, <1σ, <1σ%, 1σ-2σ, 1σ-2σ%, 2σ-3σ, 2σ-3σ%, >3σ, >3σ% |
| 20-28 | **STD 2 (15)** | Stessa struttura di STD 1 |
| 29-37 | **STD 3 (35)** | Stessa struttura di STD 1 |
| 38-46 | **STD 4 (60)** | Stessa struttura di STD 1 |

**Esempio riga dati (LOT0366):**
- Grand total tests: 19
- STD 1 (0): 4 test totali, 4 sotto 1σ (100%), 0 in 1-2σ, 0 in 2-3σ, 0 oltre 3σ
- STD 2 (15): 5 test, 1 sotto 1σ (20%), 4 in 1-2σ (80%)
- STD 3 (35): 5 test, 1 sotto 1σ (20%), 3 in 1-2σ (60%), 1 in 2-3σ (20%)
- STD 4 (60): 5 test, 0 sotto 1σ (0%), 4 in 1-2σ (80%), 1 in 2-3σ (20%)

#### ZONA C: Dati QC Raw per Grafico (Colonne 49-59)

| Colonna | Header | Descrizione |
|---------|--------|-------------|
| 49-50 | (version?) | Sempre 1.1 |
| 51 | Filename | File sorgente QC |
| 52 | Lot nr | Numero lotto |
| 53 | Lot | ID lotto sequenziale |
| 54 | TEST | Tipo test |
| 55 | STD Value | Valore standard |
| 56-58 | M1/M2/M3 Readings | Letture misuratori |
| 59 | QC Date | Data QC |

#### ZONA D: Medie Mobili per Lotto (Colonne 60-73)

| Colonna | Header | Descrizione |
|---------|--------|-------------|
| 60 | LOT | Numero lotto |
| 61 | LOT nr | Progressivo |
| 62-67 | STD1-STD6 | Valori standard nominali |
| 68-73 | STD1-STD6 AVG | **MEDIE MOBILI CUMULATIVE** per lotto |

**Esempio evoluzione STD2 AVG attraverso i lotti:**
- LOT0366 (lotto 1): 16.74
- LOT0367 (lotto 2): 16.30
- LOT0418 (lotto 3): 16.23
- LOT0428 (lotto 4): 15.36
- LOT0450 (lotto 5): 17.19
...evoluzione della media cumulativa!

#### ZONA E: Limiti Control Chart Sigma (Colonne 74-115)

Per ciascuno dei 6 standard (STD1-STD6), 7 colonne:

| Colonna Pattern | Header | Descrizione |
|-----------------|--------|-------------|
| Base{n} | Base | Valore base per calcolo sigma |
| STDn 3σ low | 3σ low | Limite inferiore 3 sigma |
| STDn 2σ low | 2σ low | Limite inferiore 2 sigma |
| STDn 1σ low | 1σ low | Limite inferiore 1 sigma |
| STDn 1σ high | 1σ high | Limite superiore 1 sigma |
| STDn 2σ high | 2σ high | Limite superiore 2 sigma |
| STDn 3σ high | 3σ high | Limite superiore 3 sigma |

**Valori per STD2 (15):** Base=7.8975, tutti i limiti sigma = 1.3675 (costante per questo set)
**Valori per STD3 (35):** Base=10.325, limiti sigma = 1.8575
**Valori per STD4 (60):** Base=12.0175, limiti sigma = 2.47

#### ZONA F: Medie STD Ripetute (Colonne 116-133) ⚠️ da verificare

Tre blocchi ripetuti di STD1-STD6 AVG (colonne 116-121, 122-127, 128-133). Interpretazione non confermata:
- Blocco 1: Medie cumulative (stessi valori di zona D)
- Blocco 2: Medie parziali (valore 0 = non calcolato)
- Blocco 3: Delta/differenze (valore 0 = non calcolato)

> **Da chiarire con Hanna**: scopo esatto dei tre blocchi ripetuti in ZONA F.

### 2.5 Formule Chiave del File 2

```
// Header dinamico sigma - AND vs OR
=IF(H3="&","Sigma σ (&)","Sigma σ (or)")

// Distribuzione sigma per lotto (percentuali)
=IFERROR(count_in_band / total_tests * 100, "")

// Esempio: % nel band <1σ per STD1
=IFERROR(L12/K12*100, "")  // dove L12=count <1σ, K12=total tests

// Riferimenti a configurazione prodotto
=ʼImport Hanna Codeʼ!I39  // Tolerance
=ʼImport Hanna Codeʼ!J39  // Operator (&/or)
=ʼImport Hanna Codeʼ!K39  // Percent value
=ʼImport Hanna Codeʼ!L39  // QC restriction

// Headers dinamici distribuzione
=CONCATENATE("Std 1 tests distribution( ", C2, " )")
```

---

## FILE 3: PREP Files (Chemical Production + Chemical MR)

### 3.1 Panoramica

I file PREP sono i fogli di preparazione per ogni singolo batch di produzione chimica.
Vengono esportati da **due software distinti** con lo stesso layout fisso:
- `docs/Excel Example Import files/Chemical Production/` → file Chemical Production
- `docs/Excel Example Import files/Chemical MR/` → file Chemical MR

**Tre sottotipi — identificabili dal naming:**
| Sottotipo | Prefisso filename | Software | Descrizione |
|-----------|------------------|----------|-------------|
| **CP** (prodotto finito) | `PREP_CP-B{xxx}.*` | Chemical Production | Preparazione reagente finito (ha Hanna Code Table) |
| **SOL/Buffer** (soluzione intermedia) | `PREP_B{xxx}-SOL.{A\|B}.*`, `PREP_B{xxx}-A/B.*` | **Chemical MR** | Soluzioni buffer pH intermedie |
| **MR** (standard/reagente via MR) | `PREP_{HannaCode}.MR{xxx}.*` | **Chemical MR** | Standard e reagenti gestiti via Material Requisition |

### 3.2 Convenzione Filename

```
PREP_{RecipeCode}.{Line}.CTK.{BatchNum}.{WeekNum}.{Year2d}[.LOT_{LotNum}][.x].xlsx
```

| Parte | Esempio | Descrizione |
|-------|---------|-------------|
| RecipeCode | `CP-B051`, `B036-SOL.A` | Codice ricetta |
| Line | `L56` | Linea produzione |
| CTK | `CTK` | Costante |
| BatchNum | `1`, `2` | Numero batch per questa ricetta/settimana |
| WeekNum | `2`, `10`, `51` | Settimana produzione |
| Year2d | `2025`, `2024` | Anno |
| `.LOT_XXXX` | `.LOT_2266` | (opzionale) Lot number assegnato |
| `.x` | `.x` | (opzionale) Correzione/ricalcolo applicato |

**Esempi:**
- `PREP_CP-B051.L56.CTK.1.2.2025..xlsx` → recipe CP-B051, batch 1, settimana 2/2025, no lot
- `PREP_CP-B036.L56.CTK.1.15.2025.LOT_2420..xlsx` → lot 2420 assegnato
- `PREP_CP-B008.L56.CTK.1.6.2025.LOT_2266.x.xlsx` → lot 2420 + correzione applicata

### 3.3 Struttura File (Layout Fisso)

**Un solo sheet** — nome = filename senza `.xlsx` (troncato a 31 car Excel).
Tutte le celle di dati sono nella colonna B (col 2) e successive. Col A = sempre None.

| Riga | Contenuto | Note |
|------|-----------|------|
| r2 | `"Preparation"` | Header sezione |
| r4 | `Recipe`, `{code}`, `Line`, `L56 CTK`, `Procedure`, (vuoto), `Rev`, `{rev}`, `Exp`, `{anni}` | Header prodotto |
| r5 | `Description`, `{nome}`, `Density`, `{val}`, `MaxQty`, `{val}`, `MinQty`, `{val}`, `Multiple`, `{val}`, `Mix`, `{codici mix}` | Parametri ricetta |
| r9 | `Recipe by`, `{nome}`, `Preparation Date`, `{datetime}`, `# Preparation Week`, `{n}` | Autore e data |
| r10 | `Planned Preparation Week`, `{ww/yyyy}`, `Preparation Week`, `{ww/yyyy}`, `Planning Reference`, `{ref}` | Settimane |
| r11 | `Note`, `{testo}`, `Operator`, `{nome}`, `Exp Date`, `{mm/yyyy}` | Operatore e scadenza |
| r12 (SOL only) | `Preparation Lot (Mix)`, `{numero}` | Lot number interno per soluzioni intermedie |
| r17 | `"Component Table"` | Header sezione |
| r18 | `Code`, `Description`, `Cas`, `%`, `Theoretical weight`, `Real weight`, `Variance`, `Variance Perc`, `Real Perc`, `Note`, `Mix` | Headers componenti |
| r19+ | Righe componenti (variabile) | Una riga per componente |
| dopo componenti | `TotalWeight (Kg)`, `{val}`, `Real Weight (Kg)`, `{val}`, `Variance (Kg)`, `{val}`, `Variance Perc`, `{val}` | Totali peso Kg |
| riga successiva | `TotalWeight (L)`, `{val}`, `Real Weight (L)`, `{val}`, `Variance (L)`, `{val}`, `Variance Perc`, `{val}` | Totali peso in litri |
| variabile | `"Acquisition Table"` | Header sezione acquisti |
| variabile | `Code`, `Description`, `Cas`, `Real Weight (g)`, `Manufacturer`, `Manufacturer Code`, `Manufacturer Lot`, `Delivery Date`, `Qty Delivered`, `Week Delivery`, `Package`, `Note`, `Operator`, `Acquisition Time`, `Recalculation`, `Added Chemical (in recipe)` | Headers acquisti |
| variabile | Righe acquisti (variabile, ≥1 per componente, possibili righe multiple per stesso codice con lotti diversi) | Materiali usati con lotto fornitore |
| variabile (CP only) | `"Hanna Code Table"` | Header sezione output |
| variabile (CP only) | `Code`, `Product Name`, `Line`, `Volume/Weight`, `(um)`, `Q.ty to produce`, `Lot Number` | Headers prodotti finiti |
| variabile (CP only) | 1 o più righe (1 Hanna Code = 1 prodotto, ma può avere più prodotti dallo stesso batch) | Prodotti finiti con lot number |
| ultime righe | `"Preparation Notes"` + `Date`, `Type`, `Description`, `Operator` | Note (solitamente vuote) |

### 3.4 Dati Chiave da Estrarre

**Header (sempre in posizioni fisse):**
| Campo | Riga | Colonna | Tipo |
|-------|------|---------|------|
| Recipe Code | r4 | C | `string` |
| Revision | r4 | I | `float` |
| Expiry years | r4 | K | `string` |
| Description | r5 | C | `string` |
| Density | r5 | E | `float` |
| MaxQty, MinQty, Multiple | r5 | G, I, K | `float` |
| Mix (SOL codes or recipe) | r5 | M | `string` |
| Recipe by | r9 | C | `string` |
| Preparation Date | r9 | E | `datetime` |
| Batch # within week | r9 | G | `int` |
| Planned Prep Week | r10 | C | `string` (es. `"10/2024"`) |
| Actual Prep Week | r10 | E | `string` (es. `"10/2024"`) |
| Planning Reference | r10 | G | `string` |
| Operator | r11 | D | `string` |
| Exp Date | r11 | G | `string` (es. `"09/2027"`) |
| Mix Lot (SOL only) | r12 | C | `int` |

**Hanna Code Table (CP files only, 1+ righe):**
| Campo | Tipo | Esempio |
|-------|------|---------|
| Hanna Code | `string` | `HI3812-0`, `HI772S`, `DEMINERAL 10` |
| Product Name | `string` | `EDTA Solution` |
| Volume/Weight + unit | `float` + `string` | `120 ml`, `30 mL`, `100 g` |
| Q.ty to produce | `int` | `2500`, `4000` |
| Lot Number | `int` | `2220`, `2281` |

### 3.5 Casi Speciali

1. **Multi-prodotto**: Un batch può generare più Hanna Codes (es. `HI755S` e `HI772S` dallo stesso batch CP-B104)
2. **Correzione (`.x`)**: Il file con `.x` nel nome indica ricalcolo — colonna `Recalculation` in Acquisition Table ha valore
3. **Multi-lotto componente**: Lo stesso codice materiale può apparire più volte nell'Acquisition Table con lotti fornitore diversi (batch parziali)
4. **SOL senza Hanna Code Table**: I file SOL sono solo intermedi; il lot number è in r12 (`Preparation Lot (Mix)`)
5. **Qty to produce vuota**: Può essere stringa vuota `''` se la quantità non è stata ancora confermata
6. **Formato settimana**: `ww/yyyy` (es. `51/2024` = settimana 51 anno 2024)

---

## SPECIFICHE PER HANNA STATISTICS: TIPI DI DATI

### 3.1 Entita Database Necessarie

```sql
-- ========================================
-- ENTITA CENTRALI
-- ========================================

-- Prodotti/SFG
CREATE TABLE hanna_codes (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sfg_code VARCHAR(20) NOT NULL UNIQUE,    -- 'HI782-0', 'HI5001-01'
  description VARCHAR(255),                 -- 'Marine Nitrate HR Reagent'
  parameter_formula VARCHAR(50),            -- 'NO3', 'pH'
  recipe VARCHAR(50),                       -- 'CP-R80', 'C237'
  production_line VARCHAR(50),              -- 'L57 Powder'
  qc_department VARCHAR(50),                -- 'Laboratory'
  registration_book VARCHAR(50),            -- 'TABLET 3'
  qc_type VARCHAR(50),                      -- 'Production'
  product_type ENUM('BUFFER','REAGENT','OTHER') NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Configurazione Standard per Prodotto
CREATE TABLE product_standards (
  id INT AUTO_INCREMENT PRIMARY KEY,
  hanna_code_id INT NOT NULL,
  std_number INT NOT NULL,                  -- 1, 2, 3, 4, 5, 6
  std_value DECIMAL(10,4) NOT NULL,         -- 0, 15, 35, 60
  sigma_value DECIMAL(10,6),                -- 1.3675
  tolerance_value DECIMAL(10,4),            -- 2.0
  tolerance_type ENUM('FIXED','PERCENT'),   -- Fixed, Percent
  tolerance_operator ENUM('AND','OR'),      -- &, or
  tolerance_percent DECIMAL(10,4),          -- 4.9
  qc_restriction VARCHAR(50),              -- '100%', 'Custom QC restriction'
  FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id)
);

-- Peso di riferimento (per reagenti tablet/polvere)
CREATE TABLE reference_weights (
  id INT AUTO_INCREMENT PRIMARY KEY,
  hanna_code_id INT NOT NULL,
  ref_weight_mg DECIMAL(10,2),              -- 206
  min_weight_mg DECIMAL(10,2),              -- 196
  max_weight_mg DECIMAL(10,2),              -- 216
  FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id)
);

-- ========================================
-- LOTTI DI PRODUZIONE
-- ========================================

CREATE TABLE production_lots (
  id INT AUTO_INCREMENT PRIMARY KEY,
  hanna_code_id INT NOT NULL,
  lot_number VARCHAR(20) NOT NULL,          -- 'LOT0366'
  lot_sequence INT,                         -- 1, 2, 3...
  preparation_week VARCHAR(20),             -- 'PWW49.1', 'W 12/2'
  qc_date DATE,
  qc_week INT,                              -- settimana anno
  source_filename VARCHAR(255),             -- nome file QC sorgente
  FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id),
  UNIQUE KEY (hanna_code_id, lot_number)
);

-- ========================================
-- TIPO A: BUFFER - Dati Produzione e Correzione
-- ========================================

CREATE TABLE buffer_production (
  id INT AUTO_INCREMENT PRIMARY KEY,
  hanna_code_id INT NOT NULL,
  ph_value DECIMAL(5,2) NOT NULL,           -- 1.68, 4.01, 7.01, 10.01
  product_codes TEXT,                        -- 'HI5001-01;HI5001-02...' (multi-valore)
  lot_number VARCHAR(20) NOT NULL,
  production_date DATE,
  quantity_kg DECIMAL(10,2),                 -- Kg prodotti
  first_qc_failed VARCHAR(100),             -- valore QC fallito (misto testo/numero)
  cm_description VARCHAR(255),              -- 'Potassium Tetraoxalate dihydrate'
  cm_code VARCHAR(20),                       -- 'CM011'
  cm_grams DECIMAL(10,4),                   -- grammi usati
  cm_percentage DECIMAL(10,8),              -- calcolato: cm_grams / (quantity_kg * 1000)
  FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id)
);

-- Lookup materiali di correzione
CREATE TABLE correction_materials (
  id INT AUTO_INCREMENT PRIMARY KEY,
  cm_code VARCHAR(20) NOT NULL UNIQUE,      -- 'CM011'
  cm_name VARCHAR(255),                     -- 'Potassium Tetraoxalate dihydrate'
  cas_number VARCHAR(30),                   -- '6100-20-5'
  used_in_ph TEXT                            -- 'pH 1.68', 'pH 4.01, pH 7.01'
);

-- ========================================
-- TIPO B: REAGENTI - Letture QC Raw
-- ========================================

CREATE TABLE qc_readings (
  id INT AUTO_INCREMENT PRIMARY KEY,
  lot_id INT NOT NULL,
  standard_number INT NOT NULL,             -- 1-6
  std_value DECIMAL(10,4),                  -- 0, 15, 35, 60
  test_number INT,                          -- progressivo test
  test_type ENUM('VALID','OLD_A','OLD_B','OLD_C','OLD_D',
                 'P_FINAL','P_PROD','ASTM','ASTM_D665',
                 'EPP_170','EPP_171','EPP_172','EPP_250','OTHER'),
  qc_date DATE,
  qc_time TIME,
  prod_date DATE,
  prod_time TIME,
  prod_operator VARCHAR(50),                -- 'ERIKA', 'ELENA'
  head_number INT,                          -- testa produzione
  meter1_reading DECIMAL(10,4),             -- lettura M1
  meter2_reading DECIMAL(10,4),             -- lettura M2
  meter3_reading DECIMAL(10,4),             -- lettura M3
  meter4_reading DECIMAL(10,4),             -- lettura M4
  meter1_model VARCHAR(30),                 -- 'HI97115', 'Label17'
  meter2_model VARCHAR(30),
  meter3_model VARCHAR(30),
  meter4_model VARCHAR(30),
  spectr_abs DECIMAL(10,6),                 -- assorbimento spettrometro
  ph1 DECIMAL(6,3),
  ph2 DECIMAL(6,3),
  ph3 DECIMAL(6,3),
  turbidity DECIMAL(10,4),                  -- NTU
  weight_mg DECIMAL(10,2),                  -- peso tablet mg
  reagent_set INT,                          -- 1, 2
  qc_operator VARCHAR(50),                  -- 'Alla', 'DAVID', 'Kinga'
  correction VARCHAR(100),
  note TEXT,
  FOREIGN KEY (lot_id) REFERENCES production_lots(id)
);

-- Set reagenti per lotto
CREATE TABLE lot_reagent_sets (
  id INT AUTO_INCREMENT PRIMARY KEY,
  lot_id INT NOT NULL,
  set_number INT NOT NULL,                  -- 1, 2
  reagent_a_lot VARCHAR(50),
  reagent_b_lot VARCHAR(50),
  reagent_c_lot VARCHAR(50),
  reagent_d_lot VARCHAR(50),
  reagent_e_lot VARCHAR(50),
  FOREIGN KEY (lot_id) REFERENCES production_lots(id)
);

-- ========================================
-- STATISTICHE CALCOLATE (cache per performance)
-- ========================================

-- Distribuzione Sigma per Lotto
CREATE TABLE lot_sigma_distribution (
  id INT AUTO_INCREMENT PRIMARY KEY,
  lot_id INT NOT NULL,
  std_number INT NOT NULL,                  -- 1-6
  total_tests INT,
  count_within_1sigma INT,                  -- <1σ
  pct_within_1sigma DECIMAL(6,2),
  count_1to2_sigma INT,                     -- 1σ-2σ
  pct_1to2_sigma DECIMAL(6,2),
  count_2to3_sigma INT,                     -- 2σ-3σ
  pct_2to3_sigma DECIMAL(6,2),
  count_beyond_3sigma INT,                  -- >3σ
  pct_beyond_3sigma DECIMAL(6,2),
  FOREIGN KEY (lot_id) REFERENCES production_lots(id)
);

-- Medie cumulative per lotto (per control chart)
CREATE TABLE lot_running_averages (
  id INT AUTO_INCREMENT PRIMARY KEY,
  lot_id INT NOT NULL,
  std_number INT NOT NULL,                  -- 1-6
  running_avg DECIMAL(10,6),                -- media cumulativa fino a questo lotto
  FOREIGN KEY (lot_id) REFERENCES production_lots(id)
);

-- Limiti control chart (configurazione per prodotto)
CREATE TABLE control_chart_limits (
  id INT AUTO_INCREMENT PRIMARY KEY,
  hanna_code_id INT NOT NULL,
  std_number INT NOT NULL,
  base_value DECIMAL(10,6),                 -- valore base
  sigma_1_low DECIMAL(10,6),
  sigma_1_high DECIMAL(10,6),
  sigma_2_low DECIMAL(10,6),
  sigma_2_high DECIMAL(10,6),
  sigma_3_low DECIMAL(10,6),
  sigma_3_high DECIMAL(10,6),
  FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id)
);
```

---

## SPECIFICHE: CALCOLI NECESSARI

### 4.1 Calcoli per Buffer (Tipo A)

| # | Calcolo | Formula | Uso |
|---|---------|---------|-----|
| 1 | **CM Percentage** | `cm_grams / (quantity_kg * 1000)` | % materiale correzione su peso totale |
| 2 | **Correction Rate** | `COUNT(cm_used) / COUNT(total_batches) * 100` | % lotti che hanno richiesto correzione |
| 3 | **Avg CM per pH** | `AVG(cm_percentage) WHERE ph = X` | Media correzione per tipo pH |
| 4 | **CM Usage Trend** | Media mobile di cm_percentage nel tempo | Trend uso materiale correzione |
| 5 | **Quantity per Period** | `SUM(quantity_kg) GROUP BY month/week` | Produzione totale per periodo |
| 6 | **First QC Fail Rate** | `COUNT(first_qc_failed NOT NULL) / COUNT(*) * 100` | % primo QC fallito |

### 4.2 Calcoli per Reagenti (Tipo B) - Controllo Statistico

| # | Calcolo | Formula | Uso |
|---|---------|---------|-----|
| 1 | **Media Cumulativa (Running Avg)** | `AVG(reading) FOR lots 1..N` | Media progressiva per control chart |
| 2 | **σ (Sigma)** | `σ = 50% × Hanna Tolerance` (valore **fisso** per ogni STD) | Bande sigma control chart — NON è STDDEV statistico |
| 3 | **Limiti Sigma** | `STD_Value ± (n × σ)` per n=1,2,3 | Bande 1σ/2σ/3σ centrate sul valore nominale dello standard |
| 4 | **Conteggio per Band** | Count readings in cada band sigma | Distribuzione sigma |
| 5 | **% per Band** | `count_in_band / total_tests * 100` | Percentuali distribuzione |
| 6 | **Grand Total Tests** | Somma test per tutti gli standard di un lotto | KPI complessivo |
| 7 | **Weight Deviation** | `(weight - ref_weight) / ref_weight * 100` | Scostamento peso da riferimento |
| 8 | **Cross-Lot Average** | Media letture raggruppate per lotto | Confronto tra lotti |
| 9 | **Weekly Production Rate** | Conteggio test per settimana | Ritmo produzione |
| 10 | **Operator Performance** | Distribuzione sigma per operatore | Analisi per operatore |

### 4.3 Logica AND/OR per QC Restriction

Dal file emerge una logica combinata:
- **AND (`&`)**: TUTTE le condizioni devono essere soddisfatte
- **OR**: ALMENO UNA condizione deve essere soddisfatta
- **Tolerance Fixed**: Valore assoluto (es. +/- 2.0 dal valore standard)
- **Tolerance Percent**: Percentuale (es. 4.9% dal valore standard)
- **QC Restriction 100%**: Tutti i campioni devono passare

```typescript
// Esempio logica validazione
function isWithinSpec(reading: number, stdValue: number, config: ProductConfig): boolean {
  const fixedOk = Math.abs(reading - stdValue) <= config.tolerance;
  const pctOk = Math.abs((reading - stdValue) / stdValue * 100) <= config.tolerancePercent;

  if (config.operator === 'AND') return fixedOk && pctOk;
  if (config.operator === 'OR') return fixedOk || pctOk;
  return fixedOk;
}
```

---

## SPECIFICHE: GRAFICI NECESSARI (Apache ECharts)

### 5.1 Grafici per Buffer (Tipo A)

#### G1: Trend Correzioni nel Tempo
- **Tipo**: Line Chart con scatter points
- **Asse X**: Date (timeline)
- **Asse Y**: cm_percentage (%)
- **Serie**: Una per pH value (4 linee: pH 1.68, 4.01, 7.01, 10.01)
- **Features**: Tooltip con dettagli lotto, zoom temporale, media mobile

#### G2: Distribuzione Uso CM per pH
- **Tipo**: Stacked Bar Chart / Pie Chart
- **Dati**: Frequenza uso di ciascun CM Code per ogni pH
- **Esempio**: pH 4.01 usa CM021 (60%), CM131 (25%), CM480 (15%)

#### G3: Produzione e Correzioni per Mese
- **Tipo**: Bar Chart + Line overlay
- **Barre**: Quantita Kg prodotte per mese
- **Linea**: % lotti corretti per mese
- **Dual Y-axis**: Kg a sinistra, % a destra

#### G4: Box Plot Quantita per pH
- **Tipo**: Box Plot
- **Per ogni pH**: distribuzione quantita prodotte (min, Q1, mediana, Q3, max)

#### G5: Heatmap Correzioni
- **Tipo**: Heatmap (Calendar)
- **Asse X**: Settimane
- **Asse Y**: pH values
- **Colore**: Intensita = % correzione media

### 5.2 Grafici per Reagenti (Tipo B) - I PIU IMPORTANTI

#### G6: Control Chart (Shewhart) - GRAFICO PRINCIPALE
- **Tipo**: Line Chart con zone colorate
- **Asse X**: Lot number (sequenziale) o Date
- **Asse Y**: Valore lettura / Media cumulativa
- **Linee**:
  - Linea centrale (Target/Mean)
  - +/- 1σ (verde)
  - +/- 2σ (giallo)
  - +/- 3σ (rosso)
  - Punti lettura effettivi
  - Media mobile cumulativa
- **Zone colorate**: Verde (<1σ), Blu (1-2σ), Giallo (2-3σ), Rosso (>3σ) — da requisiti ufficiali PDF
- **Uno per ciascun STD** (es. STD1=0, STD2=15, STD3=35, STD4=60)

```typescript
// Esempio ECharts Control Chart
const controlChartOption = {
  title: { text: 'Control Chart - STD 15 ppm' },
  tooltip: { trigger: 'axis' },
  legend: { data: ['Reading', 'Running Avg', '+3σ', '-3σ'] },
  xAxis: { type: 'category', data: lotNumbers },
  yAxis: { type: 'value', name: 'Value' },
  visualMap: {
    pieces: [
      { gt: sigma3High, color: '#FF4444' },              // >3σ rosso
      { gt: sigma2High, lte: sigma3High, color: '#FFFF00' }, // 2-3σ giallo
      { gt: sigma1High, lte: sigma2High, color: '#4488FF' }, // 1-2σ blu
      { gt: sigma1Low, lte: sigma1High, color: '#44BB44' },  // <1σ verde
      { gt: sigma2Low, lte: sigma1Low, color: '#4488FF' },   // 1-2σ blu
      { gt: sigma3Low, lte: sigma2Low, color: '#FFFF00' },   // 2-3σ giallo
      { lte: sigma3Low, color: '#FF4444' },              // >3σ rosso
    ]
  },
  series: [
    { name: 'Reading', type: 'scatter', data: readings },
    { name: 'Running Avg', type: 'line', data: runningAvgs, smooth: true },
    { name: '+3σ', type: 'line', data: sigma3HighLine, lineStyle: { type: 'dashed', color: 'red' } },
    { name: '-3σ', type: 'line', data: sigma3LowLine, lineStyle: { type: 'dashed', color: 'red' } },
    { name: '+2σ', type: 'line', data: sigma2HighLine, lineStyle: { type: 'dashed', color: 'orange' } },
    { name: '-2σ', type: 'line', data: sigma2LowLine, lineStyle: { type: 'dashed', color: 'orange' } },
    { name: '+1σ', type: 'line', data: sigma1HighLine, lineStyle: { type: 'dotted', color: 'green' } },
    { name: '-1σ', type: 'line', data: sigma1LowLine, lineStyle: { type: 'dotted', color: 'green' } },
  ]
};
```

#### G7: Distribuzione Sigma per Lotto (Stacked Bar)
- **Tipo**: 100% Stacked Bar Chart
- **Asse X**: Lot numbers
- **Asse Y**: Percentuale
- **Colori Stack**: <1σ (verde), 1-2σ (giallo), 2-3σ (arancione), >3σ (rosso)
- **Tab/Selector**: Uno per ogni STD value

#### G8: Gauge/KPI per Lotto Corrente
- **Tipo**: Gauge Chart + KPI Cards
- **Metriche**:
  - % test entro 1σ (target > 68%)
  - % test entro 2σ (target > 95%)
  - % test entro 3σ (target > 99.7%)
  - Total tests count

#### G9: Peso Tablet Distribution
- **Tipo**: Histogram + normal curve overlay
- **Asse X**: Weight bins (mg)
- **Asse Y**: Frequency
- **Linee**: Min/Max reference weight, target weight
- **Per**: Confronto tra lotti diversi

#### G10: Confronto Operatori
- **Tipo**: Grouped Bar Chart o Radar Chart
- **Dati**: Media e sigma delle letture per operatore
- **Scopo**: Identificare variabilita tra operatori

#### G11: Timeline Produzione
- **Tipo**: Gantt-like o Calendar Heatmap
- **Dati**: Attivita QC nel tempo, con colore = esito
- **Scopo**: Visualizzare ritmo e pattern produzione

#### G12: Cross-Standard Comparison (Radar)
- **Tipo**: Radar Chart
- **Assi**: Uno per ogni STD (0, 15, 35, 60)
- **Serie**: Diversi lotti sovrapposti
- **Valore**: Scostamento dalla media (normalizzato)

### 5.3 Dashboard Overview (Home Page)

Grafici combinati per vista generale:

| Widget | Tipo | Dati |
|--------|------|------|
| KPI Cards | Numbers + sparkline | Total SFGs, Total Lots, Avg Pass Rate, Active Alerts |
| Production Volume | Area Chart | Quantita prodotte per settimana (tutti i prodotti) |
| Top 10 Problematic | Horizontal Bar | SFG con piu alta % fuori specifica |
| Alert Feed | Table/List | Ultimi lotti con >3σ readings |
| CM Usage Summary | Donut Chart | Distribuzione uso materiali correzione |

---

## CONSIDERAZIONI DI SCALABILITA

### 6.1 Volume Dati Stimato

| Entita | Per SFG | x 1000 SFGs | Per Anno |
|--------|---------|-------------|----------|
| Lotti | ~25/anno | 25,000 | 25,000 |
| Letture QC raw | ~3000/anno | 3,000,000 | 3,000,000 |
| Distribuzione sigma | ~100/anno | 100,000 | 100,000 |
| Buffer batches | ~465/anno | 465,000 | 465,000 |

**Totale stimato: ~3.5M nuovi record/anno**

### 6.2 Strategie Performance

1. **Tabelle statistiche pre-calcolate**: `lot_sigma_distribution` e `lot_running_averages` evitano ricalcoli costosi
2. **Partizionamento per anno**: Tabelle grandi come `qc_readings` dovrebbero essere partizionate
3. **Indici compositi**: `(hanna_code_id, lot_number)`, `(lot_id, std_number)`
4. **Aggregazione lazy**: Calcolare le statistiche alla chiusura del lotto, non ad ogni lettura
5. **Cache lato server**: Redis per dashboard KPIs e grafici frequenti
6. **Paginazione server-side**: TanStack Table con server-side pagination per tabelle > 1000 righe

### 6.3 Import Dati

Il sistema deve supportare:
- **Import massivo**: Upload Excel per import iniziale storico (come i file analizzati)
- **Import incrementale**: Collegamento a file QC individuali (pattern filename: `R080_HI782-0_0.0_LOT0366_PWW49.1_QC.xlsx`)
- **Parsing struttura variabile**: I file QC hanno colonne leggermente diverse per lotto (vedi `Lot structure`)
- **Validazione dati**: Tipo test, range valori, date coerenti

---

## RIEPILOGO REQUISITI FUNZIONALI

### Must Have (V1)
- [ ] Gestione anagrafica prodotti (Hanna Codes) con configurazione standard e sigma
- [ ] Import dati da Excel (batch e singoli file QC da Chemical QC)
- [ ] Control Chart (Shewhart) per ogni STD di ogni prodotto
- [ ] Distribuzione Sigma per lotto (stacked bar)
- [ ] Tabella dati con filtri, ordinamento, export
- [ ] Dashboard overview con KPI cards
- [ ] Gestione lotti e letture QC raw
- [ ] Calcolo automatico medie cumulative e distribuzione sigma
- [ ] Tracking correzioni buffer — Modulo 2 (import file Chemical MR, cm_percentage, trend, FPY) *(se incluso in scope V1)*
- [ ] Preparation List — Modulo 3 (import file PREP da Chemical Production + Chemical MR) *(se incluso in scope V1)*

### Should Have (V2)
- [ ] Confronto tra lotti e tra operatori
- [ ] Alert automatici per letture fuori 3σ
- [ ] Export PDF report con grafici
- [ ] Gauge chart per QC corrente
- [ ] Weight distribution analysis

### Nice to Have (V3)
- [ ] Calendar heatmap produzione
- [ ] Radar chart cross-standard
- [ ] Predictive analytics (trend forecasting)
- [ ] Integrazione diretta con Hanna Core DB (read-only)
- [ ] Notifiche email per alert
- [ ] Role-based access control (Admin, QC Manager, Operator)

---

## MAPPING EXCEL → APP

| Excel Attuale | App Nuova | Note |
|---------------|-----------|------|
| File Excel separato per ogni SFG | Un unico DB con filtro per Hanna Code | Tutto centralizzato |
| Fogli per pH value | Filtro/tab nel frontend | Stessa logica, UI migliore |
| Formule IFERROR in celle | Calcoli server-side automatici | Nessun errore formula |
| Copia-incolla da file QC | Import automatizzato | Risparmio tempo enorme |
| Distribuzione sigma manuale | Calcolata automaticamente | Real-time |
| 133 colonne affiancate | Viste organizzate in tab/sezioni | UX migliore |
| Grafico da rifare in Excel | ECharts interattivi | Zoom, tooltip, export |
| 1 file = 1 prodotto | Confronto cross-prodotto possibile | Nuovo valore |
| Nessun alert | Alert automatici >3σ | Proattivo |
