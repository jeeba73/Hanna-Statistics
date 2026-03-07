# Hanna Core — Specifiche per lo Sviluppo di un Sistema Parallelo

> Documento di analisi e requisiti basato sull'osservazione dell'applicativo **Hanna Core**
> di Hanna Instruments. Obiettivo: riprodurre un sistema gestionale equivalente.

---

## 1. Panoramica del Sistema

### 1.1 Cos'è Hanna Core

Hanna Core è un **ERP/MES (Manufacturing Execution System)** web-based progettato per
gestire l'intero ciclo produttivo di uno stabilimento chimico. Copre:

- Gestione magazzino (materie prime, semilavorati, prodotti finiti)
- Produzione e confezionamento
- Controllo qualità (QC)
- Pianificazione e acquisti
- Manutenzione impianti
- Ticketing e feedback interni

### 1.2 Contesto d'uso

| Aspetto              | Dettaglio                                                  |
| -------------------- | ---------------------------------------------------------- |
| **Settore**          | Produzione chimica / reagenti per strumenti analitici      |
| **Utenti**           | Operatori di linea, QC, magazzinieri, manager, admin       |
| **Dispositivi**      | PC desktop + tablet in produzione (shop floor)             |
| **Scala dati**       | ~6.000 record stock, ~12.000 record packing, in crescita  |
| **Utenti simultanei**| 5-20 stimati                                               |
| **Accesso**          | Rete locale (LAN), nessun accesso internet richiesto       |

---

## 2. Architettura del Sistema

### 2.1 Schema ad alto livello

```
┌─────────────────────────────────────────────────────┐
│                    BROWSER (Client)                  │
│  Bootstrap 5 + jQuery 3.x + DataTables 2.x          │
└──────────────────────┬──────────────────────────────┘
                       │ HTTPS (porta 82)
                       │ Certificato self-signed
┌──────────────────────▼──────────────────────────────┐
│                  WEB SERVER (Nginx)                   │
│              Reverse proxy + static files             │
└──────────────────────┬──────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────┐
│               BACKEND (Application Server)           │
│         PHP 8.x + Laravel  OPPURE  Node.js + Express │
│                                                      │
│  ┌─────────┐ ┌──────────┐ ┌────────┐ ┌───────────┐  │
│  │  Auth   │ │  API     │ │ Export │ │  Business  │  │
│  │ Module  │ │ Routes   │ │ PDF/XLS│ │  Logic     │  │
│  └─────────┘ └──────────┘ └────────┘ └───────────┘  │
└──────────────────────┬──────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────┐
│              DATABASE (MariaDB / MySQL)               │
│                  ~50-75 tabelle                       │
└─────────────────────────────────────────────────────┘
```

### 2.2 Stack tecnologico

| Layer          | Tecnologia                          | Note                              |
| -------------- | ----------------------------------- | --------------------------------- |
| **Frontend**   | HTML5 / CSS3 / JavaScript           | Nessun framework SPA              |
| **UI Library** | Bootstrap 5                         | Layout, navbar, form, bottoni     |
| **Tabelle**    | jQuery 3.7+ / DataTables 2.x       | Paginazione, filtri, export       |
| **Export PDF** | pdfmake (via DataTables Buttons)    | Generazione PDF lato client       |
| **Export XLS** | SheetJS o ExcelJS                   | Generazione Excel                 |
| **Backend**    | PHP 8.2+ con Laravel 11             | Alternativa: Node.js 20 + Express |
| **Database**   | MariaDB 10.11+ o MySQL 8.x         | Engine InnoDB                     |
| **Web Server** | Nginx 1.24+                         | Reverse proxy + HTTPS             |
| **SSL**        | Self-signed o Let's Encrypt         | openssl per generazione           |
| **Hosting**    | NAS Synology / QNAP o server Linux  | On-premise, rete locale           |

---

## 3. Moduli Funzionali

### 3.1 Elenco completo dei moduli

I moduli sono accessibili dalla **barra di navigazione principale** e ciascuno
contiene sotto-sezioni.

#### 3.1.1 Chemical (Chimico)

**Scopo**: Gestione di tutti i prodotti chimici (reagenti, soluzioni, standard).

**Sotto-sezioni osservate**:
- **CH-SFG Stock** (`/stock`) — Inventario semilavorati chimici

**Funzionalita' della tabella SFG Stock**:

| Colonna         | Tipo dato      | Descrizione                                    |
| --------------- | -------------- | ---------------------------------------------- |
| Stock ID        | Integer (PK)   | Identificativo univoco                         |
| Code            | VARCHAR(50)    | Codice prodotto (es. HI97700B)                 |
| Line            | VARCHAR(10)    | Linea produttiva (es. L58)                     |
| Lot             | Integer        | Numero lotto                                   |
| ISO             | VARCHAR(20)    | Riferimento standard ISO                       |
| Exp Date        | DATE           | Data di scadenza                               |
| Quantity        | Integer        | Quantita' disponibile                          |
| Standard        | DECIMAL(10,2)  | Valore standard                                |
| Coverage        | DECIMAL(10,2)  | Percentuale di copertura                       |
| Location (Hall) | VARCHAR(20)    | Ubicazione nel magazzino (es. G STD)           |
| Shelf           | VARCHAR(10)    | Scaffale (es. A2, A3, B2)                      |
| Recipe          | VARCHAR(20)    | Codice ricetta (es. CP-97700)                  |
| History         | Link/Button    | Collegamento allo storico modifiche             |
| Masterfile      | Link/Button    | Collegamento alla scheda tecnica               |

**Azioni disponibili**:
- `+ Insert SFG` — Aggiunta nuovo semilavorato
- `Export Excel` — Esportazione dati in formato .xlsx
- `Export Stock` — Esportazione report stock dedicato
- `Filters` — Filtri avanzati sulle colonne
- `Clear` — Reset filtri

**Regole di visualizzazione**:
- Righe colorate in **rosa/rosso**: prodotti in scadenza o con stock critico
- Righe bianche: prodotti in stato normale
- Paginazione: 10 righe per pagina di default, selezionabile

---

#### 3.1.2 QC — Quality Control

**Scopo**: Tracciamento di tutti i controlli qualita' sui prodotti.

**Sotto-sezioni osservate**:
- **QC Record-book** (`/records`) — Registro controlli qualita'

**Funzionalita' della tabella QC Record-book**:

| Colonna        | Tipo dato      | Descrizione                                    |
| -------------- | -------------- | ---------------------------------------------- |
| Register Nr    | VARCHAR(20) PK | Codice registro (es. QCC12B638)                |
| Date           | DATETIME       | Data e ora registrazione                       |
| Order          | Integer        | Numero ordine                                  |
| Client         | VARCHAR(10)    | Codice cliente (es. L56, L57, L84)             |
| Code           | VARCHAR(30)    | Codice prodotto analizzato                     |
| Lot            | Integer        | Numero lotto                                   |
| Recipe         | VARCHAR(20)    | Codice ricetta (es. CP-B182, CP-B209)          |
| Sampling User  | VARCHAR(50)    | Operatore che ha effettuato il campionamento   |
| Exp Date       | DATE           | Data scadenza (formato MM/YYYY)                |
| Sampling Date  | DATETIME       | Data e ora del campionamento                   |
| QC Type        | ENUM           | Tipo: "First Qc", "Final Qc", "During Production" |

**Azioni disponibili**:
- `+ Add Record` — Nuovo record QC
- `Export PDF` — Esportazione in PDF
- `Export Excel` — Esportazione in Excel
- `Filters` — Filtri avanzati
- `Clear` — Reset filtri

**Tipi di QC**:
1. **First QC** — Primo controllo alla produzione del lotto
2. **During Production** — Controllo in corso di produzione
3. **Final QC** — Controllo finale prima del rilascio

---

#### 3.1.3 Packing (Confezionamento)

**Scopo**: Tracciamento del confezionamento dei prodotti finiti.

**Sotto-sezioni osservate**:
- **Packing Record-book** (`/packing`) — Registro confezionamento

**Funzionalita' della tabella Packing**:

| Colonna          | Tipo dato      | Descrizione                                    |
| ---------------- | -------------- | ---------------------------------------------- |
| LOT              | VARCHAR(20)    | Codice lotto (es. D116/26, R0115/26)           |
| FG Code          | VARCHAR(20)    | Codice prodotto finito (es. HI84532-70U)       |
| Tactile          | VARCHAR(20)    | Riferimento tattile/etichetta                  |
| Status           | STATUS+TIMER   | Stato con timer (verde/giallo/rosso)           |
| Line             | VARCHAR(10)    | Linea produttiva                               |
| SFG 1 - SFG 7   | VARCHAR(50) x7 | Fino a 7 semilavorati che compongono il FG     |
| EXP Date         | DATE           | Data scadenza prodotto finito                  |
| Pcs on Pallet    | Integer        | Pezzi confezionati sul pallet                  |
| Packing Operator | VARCHAR(100)   | Operatore/i di confezionamento                 |
| Claims           | TEXT/Integer   | Reclami o segnalazioni                         |

**Sistema di stato con timer**:
- **Verde** con checkmark + "0d 0h 0m" = Completato nei tempi
- **Giallo** con warning + tempo (es. "0d 0h 33m") = Lieve ritardo
- **Rosso** con X + tempo (es. "0d 1h 6m") = Ritardo significativo

Questo timer misura probabilmente il tempo trascorso dall'inizio del confezionamento
o lo scostamento dal tempo pianificato.

**Azioni disponibili**:
- `+ Add FG` — Aggiunta nuovo prodotto finito
- `Export PDF` / `Export Excel` — Esportazioni
- `Filters` — Filtri avanzati
- `Alias` — Gestione alias/nomi alternativi
- `Clear` — Reset filtri

**Logica BOM (Bill of Materials)**:
Ogni prodotto finito (FG) e' composto da 1 a 7 semilavorati (SFG).
La tabella mostra per ogni SFG: codice, lotto, quantita', data scadenza.
Questo rappresenta la **distinta base** del prodotto finito.

---

#### 3.1.4 Altri moduli (dalla navbar)

| Modulo          | Scopo stimato                                           |
| --------------- | ------------------------------------------------------- |
| **Maintenance** | Gestione manutenzione impianti, attrezzature, scadenze  |
| **Production**  | Ordini di produzione, avanzamento, tempi                 |
| **Finish Good** | Anagrafica e gestione prodotti finiti                    |
| **Activity**    | Log attivita' utenti e operazioni                        |
| **Planning**    | Pianificazione produzione e risorse                      |
| **Purchasing**  | Ordini di acquisto, fornitori, materie prime              |
| **Admin**       | Gestione utenti, ruoli, configurazioni sistema           |
| **Tickets**     | Sistema ticketing per segnalazioni e richieste           |
| **Feedback**    | Raccolta feedback da operatori                           |
| **Calendar**    | Calendario produzione/manutenzione                       |

---

### 3.2 Dashboard principale

La homepage (`/`) mostra **"Hanna Core"** con una griglia di tile/bottoni
di accesso rapido ai sotto-moduli:

- **SFG** — Semi-Finished Goods
- **RM** — Raw Materials (Materie Prime)
- **Equipment** — Attrezzature
- **Prod** — Produzione
- **Finish Good** — Prodotti finiti
- (altri tile parzialmente visibili)

---

## 4. Sistema di Autenticazione

### 4.1 Login

- Pagina di login con logo aziendale su sfondo blu
- Campo **username** con autocomplete/dropdown degli utenti esistenti
- Campo **password**
- La dropdown suggerisce un caricamento AJAX della lista utenti
  (o un datalist HTML5 pre-popolato)

### 4.2 Utenti osservati

| Username       | Ruolo probabile            |
| -------------- | -------------------------- |
| sirca.violeta  | Operatore (nome.cognome)   |
| Boros Csilla   | Operatore (nome completo)  |
| VIOLETA        | Alias o ruolo generico     |
| eugenia        | Operatore                  |
| Tablet13       | Account dedicato a tablet  |
| Csilla         | Alias breve                |

### 4.3 Gestione ruoli (stimata)

Il sistema necessita di almeno questi ruoli:

| Ruolo           | Permessi                                                |
| --------------- | ------------------------------------------------------- |
| **Admin**       | Accesso completo, gestione utenti, configurazioni       |
| **Manager**     | Visualizzazione tutti i moduli, report, planning        |
| **QC Operator** | Accesso QC Record-book, inserimento e modifica record   |
| **Packing Op.** | Accesso Packing, inserimento FG, aggiornamento status   |
| **Warehouse**   | Accesso Stock, inserimento/modifica SFG e RM            |
| **Tablet**      | Accesso limitato, interfaccia semplificata              |
| **Viewer**      | Sola lettura su moduli assegnati                        |

---

## 5. Funzionalita' Trasversali

### 5.1 Tabelle dati (comune a tutti i moduli)

Ogni tabella implementa:

- **Paginazione**: "Showing 1 to 10 of X entries", selettore righe (10/25/50/100)
- **Column visibility**: Toggle per mostrare/nascondere colonne
- **Filtri avanzati**: Popup con filtri per colonna
- **Ordinamento**: Click su header per ordinare ASC/DESC
- **Export PDF**: Generazione PDF della vista corrente
- **Export Excel**: Generazione .xlsx della vista corrente
- **Clear**: Reset di tutti i filtri attivi
- **Colorazione righe**: Logica condizionale (scadenze, stati, alert)

### 5.2 Export

| Formato | Libreria consigliata          | Note                               |
| ------- | ----------------------------- | ---------------------------------- |
| PDF     | pdfmake (client-side)         | Via DataTables Buttons extension   |
| Excel   | SheetJS / ExcelJS             | Via DataTables Buttons extension   |
| CSV     | Nativo DataTables             | Opzione aggiuntiva consigliata     |

### 5.3 Filtri

I filtri operano su ogni colonna e supportano:
- Ricerca testuale (contiene, inizia con, uguale a)
- Range di date (da - a)
- Range numerici (min - max)
- Selezione multipla (per campi ENUM come QC Type)

---

## 6. Schema Database (Stima)

### 6.1 Tabelle principali

```sql
-- ============================================
-- AUTENTICAZIONE E UTENTI
-- ============================================

CREATE TABLE users (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    username        VARCHAR(50) UNIQUE NOT NULL,
    password_hash   VARCHAR(255) NOT NULL,
    full_name       VARCHAR(100),
    role_id         INT NOT NULL,
    is_active       BOOLEAN DEFAULT TRUE,
    last_login      DATETIME,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (role_id) REFERENCES roles(id)
);

CREATE TABLE roles (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    name            VARCHAR(50) UNIQUE NOT NULL,  -- admin, manager, qc_operator, etc.
    permissions     JSON,                          -- oppure tabella separata
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- ============================================
-- CHEMICAL / STOCK
-- ============================================

CREATE TABLE sfg_stock (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(50) NOT NULL,          -- es. HI97700B
    line            VARCHAR(10),                   -- es. L58
    lot             INT NOT NULL,
    iso             VARCHAR(20),
    exp_date        DATE,
    quantity         INT DEFAULT 0,
    standard_value  DECIMAL(10,2),
    coverage        DECIMAL(10,2),
    location_hall   VARCHAR(20),                   -- es. G STD
    shelf           VARCHAR(10),                   -- es. A2, B2
    recipe          VARCHAR(20),                   -- es. CP-97700
    status          ENUM('normal','warning','critical') DEFAULT 'normal',
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE sfg_history (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    sfg_stock_id    INT NOT NULL,
    action          VARCHAR(50),                   -- insert, update, consume
    quantity_change INT,
    user_id         INT,
    notes           TEXT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (sfg_stock_id) REFERENCES sfg_stock(id),
    FOREIGN KEY (user_id) REFERENCES users(id)
);

CREATE TABLE masterfiles (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    product_code    VARCHAR(50) NOT NULL,
    description     TEXT,
    specifications  JSON,
    file_path       VARCHAR(255),                  -- path al documento tecnico
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ============================================
-- QUALITY CONTROL
-- ============================================

CREATE TABLE qc_records (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    register_nr     VARCHAR(20) UNIQUE NOT NULL,   -- es. QCC12B638
    record_date     DATETIME NOT NULL,
    order_number    INT,
    client_code     VARCHAR(10),                   -- es. L56, L57
    product_code    VARCHAR(30) NOT NULL,           -- es. ORGANIC400
    lot             INT NOT NULL,
    recipe          VARCHAR(20),                   -- es. CP-B182
    sampling_user   INT,
    exp_date        DATE,
    sampling_date   DATETIME,
    qc_type         ENUM('First Qc','Final Qc','During Production') NOT NULL,
    result          ENUM('pass','fail','pending') DEFAULT 'pending',
    notes           TEXT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (sampling_user) REFERENCES users(id)
);

CREATE TABLE qc_parameters (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    qc_record_id    INT NOT NULL,
    parameter_name  VARCHAR(50),                   -- es. pH, conductivity
    expected_value  DECIMAL(10,4),
    measured_value  DECIMAL(10,4),
    tolerance       DECIMAL(10,4),
    is_pass         BOOLEAN,
    FOREIGN KEY (qc_record_id) REFERENCES qc_records(id)
);

-- ============================================
-- PACKING / PRODOTTI FINITI
-- ============================================

CREATE TABLE packing_records (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_code        VARCHAR(20) NOT NULL,          -- es. D116/26
    fg_code         VARCHAR(20) NOT NULL,          -- es. HI84532-70U
    tactile         VARCHAR(20),
    status          ENUM('in_progress','completed','delayed','critical') DEFAULT 'in_progress',
    elapsed_time    INT DEFAULT 0,                 -- secondi trascorsi
    line            VARCHAR(10),
    exp_date        DATE,
    pcs_on_pallet   INT DEFAULT 0,
    packing_operator VARCHAR(100),
    claims          TEXT,
    started_at      DATETIME,
    completed_at    DATETIME,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE packing_sfg (
    id                  INT AUTO_INCREMENT PRIMARY KEY,
    packing_record_id   INT NOT NULL,
    sfg_position        TINYINT NOT NULL,          -- 1-7 (SFG 1 a SFG 7)
    sfg_code            VARCHAR(50),
    sfg_lot             INT,
    sfg_quantity        INT,
    sfg_exp_date        DATE,
    FOREIGN KEY (packing_record_id) REFERENCES packing_records(id)
);

-- ============================================
-- PRODUZIONE
-- ============================================

CREATE TABLE production_orders (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    order_number    INT UNIQUE NOT NULL,
    product_code    VARCHAR(50) NOT NULL,
    quantity_planned INT NOT NULL,
    quantity_produced INT DEFAULT 0,
    line            VARCHAR(10),
    status          ENUM('planned','in_progress','completed','cancelled') DEFAULT 'planned',
    planned_date    DATE,
    started_at      DATETIME,
    completed_at    DATETIME,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ============================================
-- RAW MATERIALS
-- ============================================

CREATE TABLE raw_materials (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(50) NOT NULL,
    name            VARCHAR(100),
    supplier_id     INT,
    lot             VARCHAR(30),
    quantity        DECIMAL(10,2) DEFAULT 0,
    unit            VARCHAR(10),                   -- kg, L, pcs
    exp_date        DATE,
    location_hall   VARCHAR(20),
    shelf           VARCHAR(10),
    status          ENUM('available','low','expired','quarantine') DEFAULT 'available',
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ============================================
-- EQUIPMENT
-- ============================================

CREATE TABLE equipment (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(50) UNIQUE NOT NULL,
    name            VARCHAR(100),
    type            VARCHAR(50),
    location        VARCHAR(50),
    status          ENUM('active','maintenance','inactive') DEFAULT 'active',
    last_maintenance DATE,
    next_maintenance DATE,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ============================================
-- MAINTENANCE
-- ============================================

CREATE TABLE maintenance_records (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    equipment_id    INT NOT NULL,
    type            ENUM('preventive','corrective','calibration') NOT NULL,
    description     TEXT,
    performed_by    INT,
    performed_date  DATETIME,
    next_due_date   DATE,
    status          ENUM('scheduled','in_progress','completed') DEFAULT 'scheduled',
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (equipment_id) REFERENCES equipment(id),
    FOREIGN KEY (performed_by) REFERENCES users(id)
);

-- ============================================
-- PURCHASING
-- ============================================

CREATE TABLE suppliers (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(20) UNIQUE NOT NULL,
    name            VARCHAR(100) NOT NULL,
    contact_email   VARCHAR(100),
    phone           VARCHAR(30),
    address         TEXT,
    is_active       BOOLEAN DEFAULT TRUE,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE purchase_orders (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    po_number       VARCHAR(20) UNIQUE NOT NULL,
    supplier_id     INT NOT NULL,
    status          ENUM('draft','sent','confirmed','received','cancelled') DEFAULT 'draft',
    order_date      DATE,
    expected_date   DATE,
    received_date   DATE,
    total_amount    DECIMAL(12,2),
    notes           TEXT,
    created_by      INT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (supplier_id) REFERENCES suppliers(id),
    FOREIGN KEY (created_by) REFERENCES users(id)
);

-- ============================================
-- TICKETS / FEEDBACK
-- ============================================

CREATE TABLE tickets (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    title           VARCHAR(200) NOT NULL,
    description     TEXT,
    priority        ENUM('low','medium','high','critical') DEFAULT 'medium',
    status          ENUM('open','in_progress','resolved','closed') DEFAULT 'open',
    category        VARCHAR(50),
    created_by      INT,
    assigned_to     INT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    resolved_at     DATETIME,
    FOREIGN KEY (created_by) REFERENCES users(id),
    FOREIGN KEY (assigned_to) REFERENCES users(id)
);

-- ============================================
-- CLIENTS
-- ============================================

CREATE TABLE clients (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(10) UNIQUE NOT NULL,   -- es. L56, L57, L84
    name            VARCHAR(100),
    country         VARCHAR(50),
    is_active       BOOLEAN DEFAULT TRUE,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- ============================================
-- RECIPES
-- ============================================

CREATE TABLE recipes (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(20) UNIQUE NOT NULL,   -- es. CP-B182, CP-97700
    product_code    VARCHAR(50),
    description     TEXT,
    version         INT DEFAULT 1,
    is_active       BOOLEAN DEFAULT TRUE,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE recipe_components (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    recipe_id       INT NOT NULL,
    component_code  VARCHAR(50) NOT NULL,          -- codice RM o SFG
    quantity        DECIMAL(10,4) NOT NULL,
    unit            VARCHAR(10),
    sort_order      INT DEFAULT 0,
    FOREIGN KEY (recipe_id) REFERENCES recipes(id)
);

-- ============================================
-- PRODOTTI FINITI (FINISH GOOD)
-- ============================================

CREATE TABLE finished_goods (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    code            VARCHAR(20) UNIQUE NOT NULL,   -- es. HI84532-70U
    name            VARCHAR(100),
    recipe_id       INT,
    category        VARCHAR(50),
    is_active       BOOLEAN DEFAULT TRUE,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (recipe_id) REFERENCES recipes(id)
);
```

### 6.2 Conteggio tabelle stimato

| Area               | Tabelle | Note                          |
| ------------------ | ------- | ----------------------------- |
| Auth/Users         | 2       | users, roles                  |
| Chemical/Stock     | 3       | sfg_stock, sfg_history, masterfiles |
| QC                 | 2       | qc_records, qc_parameters    |
| Packing            | 2       | packing_records, packing_sfg  |
| Production         | 1       | production_orders             |
| Raw Materials      | 1       | raw_materials                 |
| Equipment          | 1       | equipment                     |
| Maintenance        | 1       | maintenance_records           |
| Purchasing         | 3       | suppliers, purchase_orders, po_items |
| Tickets            | 1       | tickets                       |
| Clients            | 1       | clients                       |
| Recipes            | 2       | recipes, recipe_components    |
| Finished Goods     | 1       | finished_goods                |
| **Totale**         | **~21 tabelle core** | Espandibili a 50-75 con dettagli |

---

## 7. API / Routes

### 7.1 Struttura URL osservata

```
GET  /                    → Dashboard (Hanna Core homepage)
GET  /records             → QC Record-book
GET  /stock               → CH-SFG Stock
GET  /packing             → Packing Record-book
```

### 7.2 Routes REST consigliate per ogni modulo

```
# Autenticazione
POST   /api/auth/login
POST   /api/auth/logout
GET    /api/auth/users              (lista utenti per autocomplete)

# SFG Stock
GET    /api/stock                   (lista con paginazione e filtri)
POST   /api/stock                   (inserimento nuovo SFG)
GET    /api/stock/:id               (dettaglio singolo SFG)
PUT    /api/stock/:id               (aggiornamento)
DELETE /api/stock/:id               (eliminazione)
GET    /api/stock/:id/history       (storico modifiche)
GET    /api/stock/export/excel      (export Excel)
GET    /api/stock/export/stock      (export report stock)

# QC Records
GET    /api/records                 (lista con paginazione e filtri)
POST   /api/records                 (nuovo record QC)
GET    /api/records/:id             (dettaglio)
PUT    /api/records/:id             (aggiornamento)
DELETE /api/records/:id             (eliminazione)
GET    /api/records/export/pdf      (export PDF)
GET    /api/records/export/excel    (export Excel)

# Packing
GET    /api/packing                 (lista con paginazione e filtri)
POST   /api/packing                 (nuovo FG)
GET    /api/packing/:id             (dettaglio con SFG associati)
PUT    /api/packing/:id             (aggiornamento stato/timer)
DELETE /api/packing/:id             (eliminazione)
GET    /api/packing/export/pdf      (export PDF)
GET    /api/packing/export/excel    (export Excel)

# Pattern identico per: production, raw-materials, equipment,
# maintenance, purchasing, tickets, clients, recipes, finished-goods
```

---

## 8. Frontend — Dettaglio Implementativo

### 8.1 Layout principale

```html
<!-- Struttura base di ogni pagina -->
<nav class="navbar navbar-expand-lg navbar-dark bg-primary">
    <!-- Logo Hanna Instruments -->
    <!-- Menu: Chemical, Maintenance, Production, QC, Finish Good,
         Activity, Planning, Purchasing, Admin, Tickets, Feedback, Calendar -->
</nav>

<main class="container-fluid mt-3">
    <!-- Titolo sezione (es. "QC Record-book") -->
    <h2>Nome Sezione</h2>

    <!-- Barra azioni -->
    <div class="btn-toolbar mb-3">
        <button class="btn btn-success">+ Add Record</button>
        <button class="btn btn-danger">Export PDF</button>
        <button class="btn btn-info">Export Excel</button>
        <button class="btn btn-warning">Filters</button>
        <button class="btn btn-secondary">Clear</button>
    </div>

    <!-- Tabella DataTables -->
    <table id="dataTable" class="table table-striped table-bordered">
        <!-- ... -->
    </table>
</main>
```

### 8.2 Configurazione DataTables

```javascript
$('#dataTable').DataTable({
    processing: true,
    serverSide: true,                    // paginazione lato server
    ajax: '/api/stock',                  // endpoint API
    pageLength: 10,                      // righe default
    lengthMenu: [10, 25, 50, 100],       // opzioni selettore
    dom: 'Blfrtip',                      // layout con Buttons
    buttons: [
        {
            extend: 'pdfHtml5',
            text: 'Export PDF',
            className: 'btn btn-danger'
        },
        {
            extend: 'excelHtml5',
            text: 'Export Excel',
            className: 'btn btn-info'
        }
    ],
    columns: [
        { data: 'id', title: 'Stock ID' },
        { data: 'code', title: 'Code' },
        { data: 'line', title: 'Line' },
        // ... altre colonne
    ],
    createdRow: function(row, data) {
        // Colorazione righe condizionale
        if (data.status === 'critical') {
            $(row).addClass('table-danger');
        } else if (data.status === 'warning') {
            $(row).addClass('table-warning');
        }
    }
});
```

### 8.3 Sistema di login

```javascript
// Autocomplete utenti nella pagina di login
fetch('/api/auth/users')
    .then(r => r.json())
    .then(users => {
        const datalist = document.getElementById('userList');
        users.forEach(u => {
            const option = document.createElement('option');
            option.value = u.username;
            datalist.appendChild(option);
        });
    });
```

---

## 9. Requisiti Infrastrutturali

### 9.1 Setup MINIMO (MVP / Sviluppo)

| Componente    | Specifica                                               |
| ------------- | ------------------------------------------------------- |
| **Hardware**  | Qualsiasi PC/NAS: 2 core, 2GB RAM, 10GB disco          |
| **OS**        | Ubuntu 22.04, Debian 12, Windows 10/11, o NAS OS        |
| **PHP**       | 8.2+ con estensioni: pdo_mysql, mbstring, json, openssl |
| **Composer**  | 2.x (gestore dipendenze PHP)                            |
| **Node.js**   | 18+ (solo se si sceglie stack Node, oppure per assets)   |
| **MariaDB**   | 10.6+ o MySQL 8.0+                                      |
| **Nginx**     | 1.18+ (oppure Apache 2.4+)                              |
| **SSL**       | Certificato self-signed via openssl                      |
| **Rete**      | LAN con IP statico per il server                         |
| **Browser**   | Chrome/Edge/Firefox moderno                              |
| **Team**      | 1 sviluppatore full-stack                                |
| **Tempo**     | 3-6 mesi per i moduli core                               |

### 9.2 Setup MASSIMO (Produzione Enterprise)

| Componente      | Specifica                                               |
| --------------- | ------------------------------------------------------- |
| **Server App**  | 4-8 core, 8-16GB RAM, SSD 100GB+                       |
| **Server DB**   | Dedicato: 4 core, 8GB RAM, SSD 200GB+                  |
| **OS**          | Ubuntu Server 24.04 LTS o RHEL 9                        |
| **PHP**         | 8.3+ con OPcache, PHP-FPM (pool dedicati)              |
| **Framework**   | Laravel 11 con Horizon (queues), Sanctum (API auth)     |
| **MariaDB**     | 10.11+ con replica read-only per report                 |
| **Redis**       | 7.x per cache, sessioni, timer real-time                |
| **Nginx**       | 1.24+ con rate limiting, gzip, caching statico          |
| **SSL**         | Let's Encrypt (auto-renewal) o certificato aziendale    |
| **Backup**      | Giornaliero automatico DB + file, retention 30 giorni   |
| **Monitoring**  | Grafana + Prometheus o Sentry per error tracking        |
| **CI/CD**       | GitLab CI o GitHub Actions                              |
| **VPN**         | WireGuard o OpenVPN per accesso remoto                  |
| **Rete**        | VLAN dedicata, firewall, accesso controllato             |
| **Team**        | 2-4 sviluppatori + 1 DBA part-time                      |
| **Tempo**       | 6-12 mesi per tutti i moduli                             |

### 9.3 Costi stimati

| Voce                          | Minimo         | Massimo            |
| ----------------------------- | -------------- | ------------------ |
| **Hardware (on-premise)**     | 300-500 EUR    | 3.000-8.000 EUR    |
| **Software/licenze**          | 0 EUR (OSS)    | 0 EUR (OSS)        |
| **Sviluppo (se esterno)**     | 15.000-30.000  | 50.000-120.000 EUR |
| **Manutenzione annua**        | 500 EUR        | 5.000-10.000 EUR   |
| **Hosting cloud (alternativa)** | 20-50 EUR/mese | 200-500 EUR/mese |

---

## 10. Piano di Sviluppo Consigliato

### Fase 1 — Fondamenta (Settimane 1-4)

- [ ] Setup ambiente di sviluppo (Laravel + MariaDB + Nginx)
- [ ] Sistema di autenticazione con ruoli
- [ ] Layout base con navbar e dashboard
- [ ] Componente tabella riutilizzabile (DataTables wrapper)
- [ ] Sistema export PDF/Excel generico

### Fase 2 — Moduli Core (Settimane 5-12)

- [ ] Modulo Chemical / SFG Stock (CRUD + history + colorazione)
- [ ] Modulo QC Record-book (CRUD + tipi QC + filtri)
- [ ] Modulo Packing Record-book (CRUD + BOM SFG + timer stato)
- [ ] Modulo Raw Materials (CRUD + alert scadenza)

### Fase 3 — Moduli Secondari (Settimane 13-18)

- [ ] Modulo Production (ordini, avanzamento)
- [ ] Modulo Finished Goods (anagrafica, collegamento ricette)
- [ ] Modulo Recipes (distinta base, versioning)
- [ ] Modulo Equipment + Maintenance

### Fase 4 — Moduli Supporto (Settimane 19-22)

- [ ] Modulo Planning (calendario produzione)
- [ ] Modulo Purchasing (ordini acquisto, fornitori)
- [ ] Modulo Tickets + Feedback
- [ ] Modulo Admin (gestione utenti, log, configurazioni)

### Fase 5 — Rilascio (Settimane 23-26)

- [ ] Testing completo (funzionale + carico)
- [ ] Ottimizzazione performance
- [ ] Setup produzione (server, SSL, backup)
- [ ] Formazione utenti
- [ ] Go-live e supporto iniziale

---

## 11. Note Tecniche Aggiuntive

### 11.1 Convenzioni codici prodotto Hanna Instruments

Dai dati osservati, i codici seguono questi pattern:

- **HI** + numeri + lettera opzionale = Codice prodotto Hanna (es. HI97700B, HI84532-70U)
- **CP-** + codice = Codice ricetta/procedura (es. CP-B182, CP-97700)
- **QCC** + codice = Registro QC (es. QCC12B638)
- **L** + numero = Codice cliente/linea (es. L56, L57, L58)
- **ORGANIC400** = Prodotto speciale/generico

### 11.2 Gestione timer nel packing

Il timer di stato nel modulo Packing richiede:

```javascript
// Aggiornamento real-time del timer
// Opzione 1: Polling ogni 30 secondi
setInterval(() => {
    fetch('/api/packing/active-timers')
        .then(r => r.json())
        .then(timers => updateTimerDisplay(timers));
}, 30000);

// Opzione 2: WebSocket per aggiornamenti real-time
const ws = new WebSocket('wss://server:82/ws/packing');
ws.onmessage = (event) => {
    const data = JSON.parse(event.data);
    updateTimerDisplay(data);
};
```

Soglie colore osservate:
- **Verde** (checkmark): completato o entro i tempi (0 minuti di ritardo)
- **Giallo** (warning): ritardo lieve (< 1 ora)
- **Rosso** (X): ritardo significativo (> 1 ora)

### 11.3 Sicurezza

Nonostante il sistema originale operi in LAN con certificato self-signed,
per un sistema parallelo si consiglia:

- HTTPS con certificato valido (anche interno con CA aziendale)
- Password hashing con bcrypt/argon2
- CSRF protection su tutti i form
- Rate limiting sulle API
- Sanitizzazione input per prevenire SQL injection e XSS
- Sessioni con timeout automatico (importante per i tablet in produzione)
- Log di audit per tutte le operazioni CRUD

---

> **Documento generato il**: 2026-02-12
> **Basato su**: Analisi visiva di 5 screenshot dell'applicativo Hanna Core
> **Scopo**: Riferimento per lo sviluppo di un sistema gestionale parallelo
