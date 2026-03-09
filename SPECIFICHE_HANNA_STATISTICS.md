# Hanna Statistics — Specifiche Tecniche di Sviluppo

> Applicazione web per statistiche e analisi dati, da installare sulle stesse
> macchine/NAS di Hanna Instruments, accessibile via browser sulla rete locale.
>
> Stack tecnologico derivato dal progetto **MicroDens Logger Pro** (Gibertini),
> evoluto con tecnologie moderne: **shadcn/ui + Tailwind CSS + Apache ECharts**.
> Architettura e pattern mantenuti da MicroDens, UI e grafici completamente rinnovati.

---

## 1. Visione del Progetto

### 1.1 Obiettivo

Sviluppare un'applicazione web **indipendente** da Hanna Core che fornisca
funzionalita' di statistiche, analisi dati e reportistica per lo stabilimento
Hanna Instruments.

### 1.2 Caratteristiche chiave

- **Indipendente**: database proprio, nessuna dipendenza da Hanna Core
- **Coesistente**: gira sullo stesso server/NAS, accessibile via browser
- **Moderna**: shadcn/ui + Tailwind CSS + Apache ECharts (UI premium, grafici spettacolari)
- **Familiare**: stessi pattern architetturali del progetto MicroDens Logger Pro
- **Localizzata**: multilingua (EN, IT, RO, HU) per lo stabilimento rumeno

### 1.3 Accesso

```
URL: https://192.168.1.36:<PORTA>/Statistics
Esempio: https://192.168.1.36:83
Oppure: https://192.168.1.36:82/statistics (se reverse-proxied)
```

---

## 2. Mapping Tecnologico: MicroDens → Hanna Statistics

La tabella seguente mostra come ogni tecnologia usata in MicroDens Logger Pro
viene tradotta nella versione web.

| Aspetto               | MicroDens Logger Pro (Desktop)         | Hanna Statistics (Web)                       |
| --------------------- | -------------------------------------- | -------------------------------------------- |
| **Runtime**           | Electron 39.2.3                        | Node.js 20 LTS + Express.js                 |
| **Frontend**          | React 18.2 + TypeScript 5.9           | React 19 + TypeScript 5.9                   |
| **UI Library**        | MUI v7.3.5 + Emotion                  | **shadcn/ui + Tailwind CSS 4 + Radix UI**   |
| **Build Tool**        | Vite 7.2.4 + vite-plugin-electron     | Vite 7.2.4 (standard, no electron)          |
| **Grafici**           | ReCharts 2.10.3                        | **Apache ECharts 5.5 + echarts-for-react**  |
| **Animazioni**        | *(nessuna)*                            | **Framer Motion 11** (transizioni fluide)    |
| **Tabelle**           | *(basic MUI Table)*                    | **TanStack Table v8** (sorting, filtri, virtualizzazione) |
| **Export PDF**        | pdfkit + jsPDF + jspdf-autotable      | jsPDF + jspdf-autotable                      |
| **Export Excel**      | ExcelJS 4.4.0                          | ExcelJS 4.4.0                               |
| **Auth**              | bcryptjs + ruoli                       | bcryptjs + JWT + ruoli                       |
| **i18n**              | i18next (EN/IT)                        | i18next (EN/IT/RO/HU)                  |
| **Storage**           | JSON su filesystem                     | MariaDB 10.11+ (database relazionale)  |
| **IPC/Comunicazione** | Electron preload bridge                | API REST (Express routes)              |
| **State Management**  | React Context API                      | React Context API (identico)           |
| **Architettura**      | Manager Pattern (Electron main)        | Manager Pattern (Express services)     |
| **Routing Frontend**  | React Router DOM 6.21                  | React Router DOM 6.21 (identico)       |
| **Date**              | date-fns 3.2.0                         | date-fns 3.2.0 (identico)             |
| **Deploy**            | electron-builder (installer)           | Docker / PM2 su NAS                    |

### 2.1 Cosa rimane dal know-how MicroDens

- **Pattern architetturali**: Manager/Service pattern, Context API, hooks custom
- **Logica business**: Calcoli statistici, export PDF/Excel, autenticazione
- **Tooling**: Vite, TypeScript, ESLint, i18next, date-fns, bcryptjs
- **Mentalita'**: React + TypeScript, component-based, type-safe

### 2.2 Cosa si evolve (UI/UX completamente nuova)

| Aspetto         | MicroDens (vecchio)                | Hanna Statistics (nuovo)              |
| --------------- | ---------------------------------- | ------------------------------------- |
| **UI Library**  | MUI v7 (Material Design)           | shadcn/ui + Tailwind (minimal modern) |
| **Grafici**     | ReCharts (basico)                  | Apache ECharts (spettacolare)         |
| **Tabelle**     | MUI Table (limitato)               | TanStack Table (potentissimo)         |
| **Animazioni**  | Nessuna                            | Framer Motion (fluide, premium)       |
| **Styling**     | Emotion CSS-in-JS                  | Tailwind utility-first (piu' veloce) |
| **Look & Feel** | Material Design (standard Google)  | Custom modern dashboard (premium)     |

### 2.3 Cosa cambia nel backend

| Da (Electron)                        | A (Web)                                |
| ------------------------------------ | -------------------------------------- |
| `electron/main.ts`                   | `server/app.ts` (Express server)       |
| `electron/preload.ts` (IPC bridge)   | `server/routes/` (API REST endpoints)  |
| `electron/managers/*.ts`             | `server/services/*.ts` (stessa logica) |
| `window.electronAPI.xxx()`           | `fetch('/api/xxx')` o `axios`          |
| `fs.readFileSync` / `fs.writeFileSync` | Query MariaDB via `mysql2`           |
| `electron-store`                     | Tabella `config` nel DB                |
| `electron-builder`                   | `Dockerfile` + `docker-compose.yml`    |

---

## 3. Architettura del Sistema

### 3.1 Schema ad alto livello

```
┌─────────────────────────────────────────────────────────┐
│                     BROWSER (Client)                     │
│                                                          │
│   React 19 + TypeScript + shadcn/ui + Tailwind + ECharts  │
│   ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌───────────┐  │
│   │  Pages   │ │Components│ │ Contexts │ │  Hooks    │  │
│   └──────────┘ └──────────┘ └──────────┘ └───────────┘  │
│   Build: Vite 7.2.4 → bundle statico (dist/)            │
└────────────────────────┬────────────────────────────────┘
                         │ HTTPS (fetch / axios)
                         │
┌────────────────────────▼────────────────────────────────┐
│              SERVER (Node.js + Express)                   │
│              Porta: 83 (o reverse proxy da :82)          │
│                                                          │
│   ┌─────────────────────────────────────────────────┐    │
│   │              API REST Routes                     │    │
│   │  /api/auth    /api/stats    /api/export          │    │
│   │  /api/data    /api/config   /api/users           │    │
│   └──────────────────────┬──────────────────────────┘    │
│                          │                                │
│   ┌──────────────────────▼──────────────────────────┐    │
│   │           Services (Manager Pattern)             │    │
│   │  AuthService    StatsService    ExportService    │    │
│   │  DataService    ConfigService   UserService      │    │
│   └──────────────────────┬──────────────────────────┘    │
│                          │                                │
│   ┌──────────────────────▼──────────────────────────┐    │
│   │           Database Layer (mysql2)                │    │
│   │           Connection Pool + Query Builder        │    │
│   └──────────────────────┬──────────────────────────┘    │
└────────────────────────────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────┐
│              MariaDB 10.11+ (Database dedicato)          │
│              Database: hanna_statistics                   │
│              Porta: 3306 (locale al NAS)                 │
│              ~25-30 tabelle                               │
└─────────────────────────────────────────────────────────┘
```

### 3.2 Confronto con l'architettura MicroDens

```
MicroDens Logger Pro:                    Hanna Statistics:

┌─────────────────────┐                  ┌─────────────────────┐
│  Renderer Process   │                  │  Browser (Client)   │
│  (React + MUI)      │                  │  (React + shadcn)   │
└────────┬────────────┘                  └────────┬────────────┘
         │ IPC Bridge (preload.ts)                │ HTTP REST (fetch)
┌────────▼────────────┐                  ┌────────▼────────────┐
│  Main Process       │                  │  Express Server     │
│  (Electron + Node)  │                  │  (Node.js)          │
│  ┌────────────────┐ │                  │  ┌────────────────┐ │
│  │   Managers     │ │        ═══>      │  │   Services     │ │
│  │ SessionMgr     │ │                  │  │ StatsService   │ │
│  │ AuthMgr        │ │                  │  │ AuthService    │ │
│  │ ExportMgr      │ │                  │  │ ExportService  │ │
│  │ ConfigMgr      │ │                  │  │ ConfigService  │ │
│  └────────────────┘ │                  │  └────────────────┘ │
│         │           │                  │         │           │
│  ┌──────▼─────────┐ │                  │  ┌──────▼─────────┐ │
│  │  JSON Files    │ │                  │  │  MariaDB       │ │
│  └────────────────┘ │                  │  └────────────────┘ │
└─────────────────────┘                  └─────────────────────┘
```

---

## 3.3 Il Nuovo Stack UI in Dettaglio

### shadcn/ui + Tailwind CSS 4

**Perche' shadcn/ui e non MUI:**

| Aspetto            | MUI v7 (MicroDens)                   | shadcn/ui (Hanna Statistics)           |
| ------------------ | ------------------------------------ | -------------------------------------- |
| **Bundle size**    | ~300KB gzipped                       | ~50KB gzipped (solo quello che usi)    |
| **Personalizzazione** | Override complessi, theme nesting | Copia il componente, modifichi tutto   |
| **Performance**    | CSS-in-JS (runtime overhead)         | Tailwind (zero runtime, solo CSS)      |
| **Look & Feel**    | Material Design (Google standard)    | Minimal, moderno, "premium dashboard"  |
| **Complessita'**   | API enorme, molti props              | Semplice, leggibile, composable        |
| **Trend 2026**     | Maturo ma statico                    | In forte crescita, community attiva    |

**Come funziona shadcn/ui:**
- NON e' una libreria installata da npm
- Si copiano i componenti nel progetto (`src/components/ui/`)
- Ogni componente e' personalizzabile al 100%
- Basato su Radix UI (accessibilita' garantita) + Tailwind CSS
- Si aggiungono componenti con: `npx shadcn@latest add button dialog table`

**Tailwind CSS 4 (2025+):**
- Utility-first: classi come `bg-blue-600 text-white rounded-lg p-4 shadow-md`
- Zero runtime: tutto compilato a CSS statico
- Tailwind v4: configurazione in CSS (non piu' tailwind.config.js)
- Dark mode automatico con `dark:` prefix
- Animazioni con `animate-` utilities

### Apache ECharts 5.5

**Perche' ECharts e non ReCharts:**

| Aspetto           | ReCharts (MicroDens)           | Apache ECharts (Hanna Statistics)       |
| ----------------- | ------------------------------ | --------------------------------------- |
| **Tipi grafici**  | ~15 (bar, line, pie, area...)  | ~30+ (+ heatmap, treemap, radar, 3D)   |
| **Animazioni**    | Basiche                         | Fluide, configurabili, spettacolari     |
| **Interattivita'** | Click, tooltip                 | Zoom, brush, drag, connect, drill-down  |
| **Performance**   | SVG (lento con molti dati)     | Canvas + SVG (veloce con 100K+ punti)  |
| **Temi**          | Manuale                         | Temi pronti (dark, vintage, macarons)   |
| **Responsive**    | Manuale                         | Auto-resize integrato                   |
| **Mappa**         | No                              | Si (mappe geografiche, geoJSON)         |
| **Dimensione**    | ~200KB                          | ~400KB (ma tree-shakeable a ~100KB)     |

**Tipi di grafici ECharts utili per Hanna Statistics:**

```
Dashboard:
├── Gauge Chart         → KPI (QC pass rate, efficienza)
├── Liquid Fill         → Livello stock (visuale effetto acqua)
├── Bar Chart           → Produzione per linea
├── Line Chart          → Trend temporali
└── Pie/Donut           → Distribuzione categorie

Analisi avanzata:
├── Heatmap             → Qualita' per prodotto/periodo (matrice colori)
├── Treemap             → Distribuzione stock per categoria
├── Radar Chart         → Confronto multi-dimensionale prodotti
├── Sankey Diagram      → Flusso produzione (RM → SFG → FG)
├── Boxplot             → Distribuzione parametri QC
├── Scatter             → Correlazioni (tempo vs qualita')
└── Candlestick         → Variazioni stock nel tempo
```

### Framer Motion 11

Aggiunge il **"premium feel"** a tutta l'applicazione:

```
Transizioni pagina:     Fade + slide tra le route
Card KPI:               Contatore animato (0 → 1,523)
Grafici:                Entrata con stagger (uno dopo l'altro)
Tabelle:                Righe che appaiono con spring animation
Sidebar:                Apertura/chiusura fluida
Notifiche:              Slide-in/out con physics-based motion
Loading:                Skeleton shimmer effect
```

### TanStack Table v8

Sostituisce le tabelle MUI e DataTables con una soluzione piu' potente:

```
Funzionalita':
├── Sorting multi-colonna
├── Filtri globali e per colonna
├── Paginazione server-side
├── Column resizing (drag per ridimensionare)
├── Column reordering (drag per riordinare)
├── Row selection (checkbox)
├── Row expansion (dettagli inline)
├── Virtualizzazione (render solo righe visibili → veloce con 100K righe)
├── Export integrato
├── Sticky headers
└── Faceted filtering (filtri con conteggio)
```

---

## 4. Struttura del Progetto

```
hanna-statistics/
│
├── src/                          # Frontend React (evoluto da MicroDens)
│   ├── components/               # Componenti React riutilizzabili
│   │   ├── ui/                  # shadcn/ui components (button, dialog, card, etc.)
│   │   ├── Auth/                 # Login, ruoli
│   │   ├── Dashboard/            # Dashboard statistiche
│   │   ├── Layout/               # MainLayout, Sidebar, Navbar
│   │   ├── Charts/               # Componenti grafici (wrapper ECharts)
│   │   ├── Tables/               # Tabelle dati (wrapper TanStack Table)
│   │   ├── Export/               # UI per export PDF/Excel
│   │   ├── Filters/              # Filtri e selettori
│   │   └── common/               # Componenti condivisi
│   │
│   ├── pages/                    # Pagine (lazy-loaded)
│   │   ├── Dashboard.tsx         # Dashboard principale (KPI, alert, overview)
│   │   ├── ReagentiQC.tsx        # Reagenti QC: Control Chart + Sigma distribution
│   │   ├── BufferProduction.tsx  # Buffer Production: trend correzioni per pH
│   │   ├── PreparationList.tsx   # Preparation List: statistiche PREP files
│   │   ├── TrendAnalysis.tsx     # Analisi trend cross-modulo
│   │   ├── Reports.tsx           # Generazione report PDF/Excel
│   │   ├── DataImport.tsx        # Import dati (Chemical QC + Chemical Production)
│   │   ├── Settings.tsx          # Configurazioni
│   │   └── Users.tsx             # Gestione utenti
│   │
│   ├── contexts/                 # React Context (STESSO PATTERN MicroDens)
│   │   ├── AppContext.tsx        # Config globale, lingua, tema
│   │   ├── AuthContext.tsx       # Autenticazione e ruoli
│   │   ├── DataContext.tsx       # Dati statistici caricati
│   │   ├── FilterContext.tsx     # Stato filtri attivi
│   │   ├── ThemeContext.tsx      # Tema chiaro/scuro
│   │   └── NotificationContext.tsx
│   │
│   ├── hooks/                    # Custom hooks
│   │   ├── useStats.ts          # Hook per calcoli statistici
│   │   ├── useFilters.ts        # Hook per gestione filtri
│   │   ├── useExport.ts         # Hook per export
│   │   └── useApi.ts            # Hook wrapper per fetch API
│   │
│   ├── services/                 # Servizi frontend (chiamate API)
│   │   ├── api.ts               # Client HTTP (fetch/axios)
│   │   ├── authApi.ts           # Chiamate auth
│   │   ├── statsApi.ts          # Chiamate statistiche
│   │   ├── dataApi.ts           # Chiamate dati
│   │   └── exportApi.ts         # Chiamate export
│   │
│   ├── utils/                    # Utility (RIUTILIZZABILI da MicroDens)
│   │   ├── statisticsCalculator.ts  # Calcoli statistici
│   │   ├── chartHelpers.ts      # Helper per grafici
│   │   ├── formatters.ts        # Formattazione numeri/date
│   │   └── validators.ts        # Validazione input
│   │
│   ├── types/                    # TypeScript types
│   │   ├── auth.types.ts
│   │   ├── stats.types.ts
│   │   ├── data.types.ts
│   │   ├── export.types.ts
│   │   └── config.types.ts
│   │
│   ├── config/
│   │   └── i18n.ts              # Setup i18next
│   │
│   ├── locales/                  # Traduzioni
│   │   ├── en.json
│   │   ├── it.json
│   │   ├── ro.json              # Rumeno (per operatori)
│   │   └── hu.json              # Ungherese (per operatori)
│   │
│   ├── lib/
│   │   └── utils.ts              # cn() helper per shadcn/ui + Tailwind
│   │
│   ├── App.tsx                   # Root component
│   ├── routes.tsx                # Route definitions (lazy-loaded)
│   ├── globals.css               # Tailwind base + shadcn CSS variables
│   ├── main.tsx                  # Entry point React
│   └── index.css                 # Global styles
│
├── server/                       # Backend Express (SOSTITUISCE electron/)
│   ├── app.ts                    # Express app setup + middleware
│   ├── server.ts                 # Server entry point (listen)
│   │
│   ├── routes/                   # API REST (SOSTITUISCE preload.ts)
│   │   ├── auth.routes.ts       # POST /api/auth/login, logout, etc.
│   │   ├── stats.routes.ts      # GET  /api/stats/production, quality, etc.
│   │   ├── data.routes.ts       # CRUD /api/data/records, stock, packing
│   │   ├── export.routes.ts     # GET  /api/export/pdf, excel
│   │   ├── config.routes.ts     # GET/PUT /api/config
│   │   ├── users.routes.ts      # CRUD /api/users
│   │   └── import.routes.ts     # POST /api/import (upload dati)
│   │
│   ├── services/                 # Business logic (SOSTITUISCE managers/)
│   │   ├── AuthService.ts       # ← da AuthManager.ts
│   │   ├── StatsService.ts      # ← NUOVO: calcoli statistici server-side
│   │   ├── DataService.ts       # ← da SessionManager.ts
│   │   ├── ExportService.ts     # ← da ExportManager.ts
│   │   ├── ConfigService.ts     # ← da ConfigManager.ts
│   │   ├── UserService.ts       # ← da AuthManager.ts (parte utenti)
│   │   ├── ImportService.ts     # ← NUOVO: importazione dati
│   │   └── index.ts             # Export centralizzato
│   │
│   ├── middleware/               # Express middleware
│   │   ├── auth.middleware.ts   # Verifica JWT token
│   │   ├── role.middleware.ts   # Verifica ruoli utente
│   │   ├── error.middleware.ts  # Gestione errori centralizzata
│   │   └── cors.middleware.ts   # CORS per sviluppo locale
│   │
│   ├── database/                 # Database layer
│   │   ├── connection.ts        # Pool connessioni MariaDB (mysql2)
│   │   ├── migrations/          # Script migrazione schema
│   │   │   ├── 001_create_users.sql
│   │   │   ├── 002_create_hanna_codes.sql
│   │   │   ├── 003_create_production_lots.sql
│   │   │   ├── 004_create_qc_readings.sql
│   │   │   ├── 005_create_sigma_cache.sql
│   │   │   ├── 006_create_buffer_production.sql
│   │   │   ├── 007_create_preparation_batches.sql
│   │   │   └── 008_create_config.sql
│   │   └── seed/                # Dati iniziali
│   │       ├── default_users.sql
│   │       └── default_config.sql
│   │
│   └── types/                    # Types server-side
│       ├── express.d.ts
│       └── database.types.ts
│
├── shared/                       # Codice condiviso (IDENTICO a MicroDens)
│   └── types/
│       ├── stats.types.ts       # Tipi statistici condivisi
│       ├── data.types.ts        # Tipi dati condivisi
│       └── index.ts
│
├── public/                       # Asset statici
│   ├── favicon.ico              # Icona Hanna
│   └── logo.svg                 # Logo Hanna Instruments
│
├── docker/                       # Configurazione Docker per NAS
│   ├── Dockerfile               # Build multi-stage (client + server)
│   ├── docker-compose.yml       # App + MariaDB
│   ├── nginx.conf               # Reverse proxy config
│   └── .env.example             # Variabili d'ambiente template
│
├── scripts/                      # Script di utilita'
│   ├── setup-db.sh              # Inizializzazione database
│   ├── backup-db.sh             # Backup automatico
│   └── migrate.ts               # Runner migrazioni
│
├── components.json               # Configurazione shadcn/ui
├── postcss.config.js             # PostCSS (solo se richiesto da alcuni tool; Tailwind v4 usa Vite plugin)
├── vite.config.ts                # Configurazione Vite
├── tsconfig.json                 # TypeScript config frontend
├── tsconfig.server.json          # TypeScript config server
├── package.json                  # Dipendenze e script
├── .env                          # Variabili d'ambiente (non committato)
├── .env.example                  # Template variabili
└── README.md                     # Documentazione setup
```

---

## 5. Stack Tecnologico Completo

### 5.1 Dipendenze Frontend (da package.json)

```json
{
  "dependencies": {
    "react": "^19.0.0",
    "react-dom": "^19.0.0",
    "react-router-dom": "^7.1.0",

    "@radix-ui/react-dialog": "^1.1.0",
    "@radix-ui/react-dropdown-menu": "^2.1.0",
    "@radix-ui/react-select": "^2.1.0",
    "@radix-ui/react-tabs": "^1.1.0",
    "@radix-ui/react-toast": "^1.2.0",
    "@radix-ui/react-tooltip": "^1.1.0",
    "@radix-ui/react-slot": "^1.1.0",

    "echarts": "^5.5.0",
    "echarts-for-react": "^3.0.2",

    "@tanstack/react-table": "^8.13.0",

    "framer-motion": "^11.15.0",

    "class-variance-authority": "^0.7.0",
    "clsx": "^2.1.0",
    "tailwind-merge": "^2.2.0",
    "lucide-react": "^0.468.0",

    "i18next": "^23.7.16",
    "react-i18next": "^14.0.0",
    "date-fns": "^3.2.0",
    "axios": "^1.6.0",
    "jspdf": "^2.5.1",
    "jspdf-autotable": "^3.8.1",
    "exceljs": "^4.4.0"
  },
  "devDependencies": {
    "typescript": "^5.9.3",
    "vite": "^7.2.4",
    "@vitejs/plugin-react": "^4.2.0",
    "@types/react": "^19.0.0",
    "@types/react-dom": "^19.0.0",
    "tailwindcss": "^4.0.0",
    "@tailwindcss/vite": "^4.0.0",
    "autoprefixer": "^10.4.0",
    "eslint": "^9.39.1",
    "@typescript-eslint/eslint-plugin": "^7.0.0"
  }
}
```

### 5.2 Dipendenze Backend (da package.json)

```json
{
  "dependencies": {
    "express": "^4.18.0",
    "mysql2": "^3.9.0",
    "bcryptjs": "^2.4.3",
    "jsonwebtoken": "^9.0.0",
    "cors": "^2.8.5",
    "helmet": "^7.1.0",
    "morgan": "^1.10.0",
    "multer": "^1.4.5",
    "dotenv": "^16.3.0",
    "compression": "^1.7.4"
  },
  "devDependencies": {
    "tsx": "^4.7.0",
    "nodemon": "^3.0.0",
    "@types/express": "^4.17.0",
    "@types/bcryptjs": "^2.4.0",
    "@types/jsonwebtoken": "^9.0.0",
    "@types/cors": "^2.8.0",
    "@types/morgan": "^1.9.0",
    "@types/multer": "^1.4.0",
    "@types/compression": "^1.7.0"
  }
}
```

### 5.3 Scripts (da package.json)

```json
{
  "scripts": {
    "dev": "concurrently \"npm run dev:client\" \"npm run dev:server\"",
    "dev:client": "vite",
    "dev:server": "nodemon --exec tsx server/server.ts",
    "build": "tsc && vite build && tsc -p tsconfig.server.json",
    "start": "node dist-server/server.js",
    "db:migrate": "tsx scripts/migrate.ts",
    "db:seed": "tsx scripts/seed.ts",
    "lint": "eslint .",
    "typecheck": "tsc --noEmit"
  }
}
```

---

## 6. Traduzione dei Manager Pattern

### 6.1 Da AuthManager (Electron) a AuthService (Express)

**MicroDens (electron/managers/AuthManager.ts)**:
```typescript
// Comunicazione via IPC
ipcMain.handle('auth:login', async (_, username, password) => {
    const user = await this.findUser(username);
    const valid = await bcrypt.compare(password, user.passwordHash);
    return { success: valid, user };
});
```

**Hanna Statistics (server/services/AuthService.ts)**:
```typescript
// Comunicazione via REST API
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import { pool } from '../database/connection';

export class AuthService {
    async login(username: string, password: string) {
        const [rows] = await pool.execute(
            'SELECT * FROM users WHERE username = ? AND is_active = 1',
            [username]
        );
        const user = (rows as any[])[0];
        if (!user) throw new Error('User not found');

        const valid = await bcrypt.compare(password, user.password_hash);
        if (!valid) throw new Error('Invalid password');

        const token = jwt.sign(
            { id: user.id, role: user.role },
            process.env.JWT_SECRET!,
            { expiresIn: '8h' }
        );
        return { token, user: { id: user.id, username: user.username, role: user.role } };
    }
}
```

### 6.2 Da ExportManager a ExportService

**MicroDens (electron/managers/ExportManager.ts)**:
```typescript
// Scrive su filesystem locale
const workbook = new ExcelJS.Workbook();
const ws = workbook.addWorksheet('Data');
ws.addRows(data);
await workbook.xlsx.writeFile(filePath);
```

**Hanna Statistics (server/services/ExportService.ts)**:
```typescript
// Restituisce buffer via HTTP response
import ExcelJS from 'exceljs';
import { Response } from 'express';

export class ExportService {
    async exportExcel(data: any[], res: Response) {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Statistics');
        // ... stessa logica di formattazione MicroDens ...
        ws.addRows(data);

        res.setHeader('Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition',
            'attachment; filename=statistics_report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    }

    async exportPdf(data: any[], config: PdfConfig, res: Response) {
        const doc = new jsPDF();
        // ... stessa logica di formattazione MicroDens ...
        autoTable(doc, { head: [headers], body: rows });

        const buffer = doc.output('arraybuffer');
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition',
            'attachment; filename=statistics_report.pdf');
        res.send(Buffer.from(buffer));
    }
}
```

### 6.3 Da electronAPI (preload) a API Client (fetch)

**MicroDens (frontend — chiama via IPC)**:
```typescript
// src/contexts/AuthContext.tsx nel progetto MicroDens
const login = async (username: string, password: string) => {
    const result = await window.electronAPI.auth.login(username, password);
    setUser(result.user);
};
```

**Hanna Statistics (frontend — chiama via HTTP)**:
```typescript
// src/contexts/AuthContext.tsx nel progetto Hanna Statistics
const login = async (username: string, password: string) => {
    const response = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
    });
    const result = await response.json();
    localStorage.setItem('token', result.token);
    setUser(result.user);
};
```

### 6.4 Mappa completa della traduzione

| MicroDens Manager        | Hanna Statistics Service | Cambiamento principale              |
| ------------------------ | ----------------------- | ----------------------------------- |
| `AuthManager`            | `AuthService`           | + JWT token, DB invece di JSON file |
| `SessionManager`         | `DataService`           | Query SQL invece di fs.read/write   |
| `ExportManager`          | `ExportService`         | Output su HTTP response             |
| `ConfigManager`          | `ConfigService`         | Tabella config invece di JSON file  |
| `BLESerialManager`       | *(non necessario)*      | Nessun hardware da gestire          |
| `BLEManager`             | *(non necessario)*      | Nessun hardware da gestire          |
| `LicenseManager`         | *(opzionale)*           | Puo' non servire su rete interna    |
| `UpdateManager`          | *(non necessario)*      | Docker gestisce gli aggiornamenti   |
| *(nuovo)*                | `StatsService`          | Calcoli statistici server-side      |
| *(nuovo)*                | `ImportService`         | Importazione dati da file/CSV       |

---

## 7. Database Schema

### 7.1 Connessione

```typescript
// server/database/connection.ts
import mysql from 'mysql2/promise';

export const pool = mysql.createPool({
    host: process.env.DB_HOST || '127.0.0.1',
    port: parseInt(process.env.DB_PORT || '3306'),
    user: process.env.DB_USER || 'hanna_stats',
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME || 'hanna_statistics',
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});
```

### 7.2 Tabelle principali

```sql
-- ============================================
-- DATABASE: hanna_statistics
-- Indipendente da Hanna Core
-- ============================================

CREATE DATABASE IF NOT EXISTS hanna_statistics
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE hanna_statistics;

-- ============================================
-- UTENTI E AUTENTICAZIONE
-- ============================================

CREATE TABLE users (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    username        VARCHAR(50) UNIQUE NOT NULL,
    password_hash   VARCHAR(255) NOT NULL,
    full_name       VARCHAR(100),
    role            ENUM('admin','manager','operator','viewer') NOT NULL DEFAULT 'viewer',
    language        VARCHAR(5) DEFAULT 'en',
    is_active       BOOLEAN DEFAULT TRUE,
    last_login      DATETIME,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE audit_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    user_id         INT,
    action          VARCHAR(50) NOT NULL,
    details         JSON,
    ip_address      VARCHAR(45),
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- ============================================
-- ANAGRAFICA PRODOTTI (HANNA CODES)
-- ============================================

CREATE TABLE hanna_codes (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    sfg_code        VARCHAR(30) NOT NULL UNIQUE,    -- 'HI782-0', 'HI3812-0'
    description     VARCHAR(255),                   -- 'Marine Nitrate HR Reagent'
    parameter_formula VARCHAR(50),                  -- 'NO3', 'pH'
    recipe          VARCHAR(50),                    -- 'CP-R80', 'CP-B051'
    production_line VARCHAR(50),                    -- 'L57 Powder', 'L56 CTK'
    qc_department   VARCHAR(50),
    registration_book VARCHAR(50),
    qc_type         VARCHAR(50),
    product_type    ENUM('REAGENT','BUFFER','OTHER') NOT NULL DEFAULT 'REAGENT',
    ref_weight_mg   DECIMAL(10,2),                  -- peso riferimento (solo reagenti tablet)
    ref_weight_min_mg DECIMAL(10,2),
    ref_weight_max_mg DECIMAL(10,2),
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Configurazione standard e sigma per prodotto (fino a 6 STD per prodotto)
CREATE TABLE product_standards (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    std_number      TINYINT NOT NULL,               -- 1-6
    std_value       DECIMAL(10,4) NOT NULL,         -- 0, 15, 35, 60
    -- σ = 50% Hanna Tolerance (NON deviazione standard statistica)
    sigma_value     DECIMAL(10,6),                  -- es. 1.3675
    tolerance_fixed DECIMAL(10,4),                  -- tolleranza fissa assoluta (es. 2.0)
    tolerance_operator ENUM('AND','OR'),            -- operatore logico tra fixed e percent
    tolerance_percent DECIMAL(10,4),                -- tolleranza percentuale (es. 4.9)
    qc_restriction  VARCHAR(50),                    -- '100%', 'Custom QC restriction'
    ph_value        DECIMAL(6,3),
    ph_min          DECIMAL(6,3),
    ph_max          DECIMAL(6,3),
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id) ON DELETE CASCADE,
    UNIQUE KEY uk_hc_std (hanna_code_id, std_number)
);

-- ============================================
-- MODULO REAGENTI QC (Chemical QC files)
-- ============================================

-- Lotti di produzione per reagenti QC
CREATE TABLE production_lots (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    lot_number      VARCHAR(20) NOT NULL,           -- 'LOT0366'
    lot_sequence    INT,                            -- numero progressivo per grafici
    preparation_week VARCHAR(20),                   -- 'PWW49.1'
    first_qc_date   DATE,
    source_filename VARCHAR(255),                   -- nome file QC sorgente
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id),
    UNIQUE KEY uk_lot (hanna_code_id, lot_number)
);

CREATE INDEX idx_lot_hc ON production_lots(hanna_code_id);
CREATE INDEX idx_lot_date ON production_lots(first_qc_date);

-- Letture QC raw (da Chemical QC files — letture meter per standard)
CREATE TABLE qc_readings (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_id          INT NOT NULL,
    std_number      TINYINT NOT NULL,               -- 1-6
    std_value       DECIMAL(10,4),                  -- 0, 15, 35, 60
    test_number     INT,                            -- progressivo test
    test_type       ENUM('VALID','OLD_A','OLD_B','OLD_C','OLD_D','P_FINAL','P_PROD') NOT NULL,
    qc_date         DATE,
    qc_time         TIME,
    prod_date       DATE,
    prod_time       TIME,
    prod_operator   VARCHAR(50),
    head_number     TINYINT,
    -- Colonne meter: identificate per posizione (12-15), non per nome (varia tra file)
    meter1_reading  DECIMAL(10,4),
    meter2_reading  DECIMAL(10,4),
    meter3_reading  DECIMAL(10,4),
    meter4_reading  DECIMAL(10,4),
    spectr_abs      DECIMAL(10,6),
    ph1             DECIMAL(6,3),
    ph2             DECIMAL(6,3),
    ph3             DECIMAL(6,3),
    turbidity       DECIMAL(10,4),
    weight_mg       DECIMAL(10,2),
    reagent_set     TINYINT,
    qc_operator     VARCHAR(50),
    correction      VARCHAR(100),
    note            TEXT,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE
);

CREATE INDEX idx_qc_lot ON qc_readings(lot_id);
CREATE INDEX idx_qc_date ON qc_readings(qc_date);
CREATE INDEX idx_qc_type ON qc_readings(test_type);

-- Cache distribuzione sigma per lotto/STD (pre-calcolata per performance)
CREATE TABLE lot_sigma_distribution (
    id                  INT AUTO_INCREMENT PRIMARY KEY,
    lot_id              INT NOT NULL,
    std_number          TINYINT NOT NULL,
    total_tests         INT DEFAULT 0,
    count_within_1sigma INT DEFAULT 0,
    pct_within_1sigma   DECIMAL(6,2),               -- <1σ → VERDE
    count_1to2_sigma    INT DEFAULT 0,
    pct_1to2_sigma      DECIMAL(6,2),               -- 1-2σ → BLU
    count_2to3_sigma    INT DEFAULT 0,
    pct_2to3_sigma      DECIMAL(6,2),               -- 2-3σ → GIALLO
    count_beyond_3sigma INT DEFAULT 0,
    pct_beyond_3sigma   DECIMAL(6,2),               -- >3σ → fuori spec
    calculated_at       DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE,
    UNIQUE KEY uk_lsd (lot_id, std_number)
);

-- Cache medie cumulative per control chart
CREATE TABLE lot_running_averages (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_id          INT NOT NULL,
    std_number      TINYINT NOT NULL,
    running_avg     DECIMAL(10,6),
    calculated_at   DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE,
    UNIQUE KEY uk_lra (lot_id, std_number)
);

-- Limiti sigma per control chart (configurazione per hanna_code + STD)
CREATE TABLE control_chart_limits (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    std_number      TINYINT NOT NULL,
    base_value      DECIMAL(10,6),
    sigma_1_low     DECIMAL(10,6),
    sigma_1_high    DECIMAL(10,6),
    sigma_2_low     DECIMAL(10,6),
    sigma_2_high    DECIMAL(10,6),
    sigma_3_low     DECIMAL(10,6),
    sigma_3_high    DECIMAL(10,6),
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id) ON DELETE CASCADE,
    UNIQUE KEY uk_ccl (hanna_code_id, std_number)
);

-- ============================================
-- MODULO BUFFER PRODUCTION (Buffer Statistic xlsx)
-- ============================================

CREATE TABLE buffer_production (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    recipe_code     VARCHAR(20),                    -- 'C163', 'C166'
    ph_value        DECIMAL(5,2) NOT NULL,          -- 1.68, 4.01, 7.01, 10.01
    product_codes   TEXT,                           -- 'HI5001-01;HI5001-02...' multi-valore
    lot_number      VARCHAR(20) NOT NULL,
    production_date DATE,
    quantity_kg     DECIMAL(10,2),
    first_qc_failed VARCHAR(100),                   -- valore QC fallito (misto testo/numero)
    cm_description  VARCHAR(255),
    cm_code         VARCHAR(20),
    cm_grams        DECIMAL(10,4),
    cm_percentage   DECIMAL(12,8),                  -- cm_grams / (quantity_kg * 1000)
    source_filename VARCHAR(255),
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_buf_date ON buffer_production(production_date);
CREATE INDEX idx_buf_ph ON buffer_production(ph_value);

-- Lookup materiali di correzione
CREATE TABLE correction_materials (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    cm_code         VARCHAR(20) NOT NULL UNIQUE,
    cm_name         VARCHAR(255),
    cas_number      VARCHAR(30),
    used_in_ph      TEXT                            -- es. 'pH 4.01, pH 7.01'
);

-- ============================================
-- MODULO PREPARATION LIST (PREP files)
-- ============================================

-- Batch di preparazione (da ogni file PREP_*.xlsx)
CREATE TABLE preparation_batches (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    recipe_code     VARCHAR(30) NOT NULL,           -- 'CP-B051', 'B036-SOL.A'
    batch_type      ENUM('CP','SOL') NOT NULL,      -- CP = prodotto finito, SOL = intermedio
    description     VARCHAR(255),
    production_line VARCHAR(20),                    -- 'L56'
    revision        DECIMAL(4,2),
    expiry_years    VARCHAR(10),
    density         DECIMAL(8,4),
    preparation_date DATE,
    batch_number    TINYINT,                        -- 1 o 2 (batch nella stessa settimana)
    planned_week    VARCHAR(10),                    -- 'ww/yyyy' es. '10/2024'
    actual_week     VARCHAR(10),
    planning_reference VARCHAR(50),
    operator        VARCHAR(50),
    exp_date        VARCHAR(10),                    -- 'mm/yyyy'
    mix_lot_number  INT,                            -- solo per SOL: Preparation Lot (Mix)
    source_filename VARCHAR(255),
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_prep_recipe ON preparation_batches(recipe_code);
CREATE INDEX idx_prep_date ON preparation_batches(preparation_date);
CREATE INDEX idx_prep_week ON preparation_batches(actual_week);

-- Prodotti Hanna generati da ogni batch (Hanna Code Table nei PREP files CP)
CREATE TABLE preparation_hanna_codes (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    batch_id        INT NOT NULL,
    hanna_code      VARCHAR(30) NOT NULL,           -- 'HI3812-0', 'HI772S', 'DEMINERAL 10'
    product_name    VARCHAR(255),
    volume_weight   DECIMAL(10,2),
    unit            VARCHAR(10),                    -- 'ml', 'g', 'mL'
    qty_to_produce  INT,
    lot_number      INT,
    FOREIGN KEY (batch_id) REFERENCES preparation_batches(id) ON DELETE CASCADE
);

-- ============================================
-- CONFIGURAZIONE E LOG
-- ============================================

CREATE TABLE app_config (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    config_key      VARCHAR(100) UNIQUE NOT NULL,
    config_value    JSON NOT NULL,
    description     VARCHAR(255),
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE import_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    filename        VARCHAR(255) NOT NULL,
    file_type       ENUM('xlsx','csv','json') NOT NULL,
    module          ENUM('REAGENTI_QC','BUFFER_PRODUCTION','PREPARATION_LIST','OTHER'),
    rows_imported   INT DEFAULT 0,
    rows_skipped    INT DEFAULT 0,
    rows_errors     INT DEFAULT 0,
    status          ENUM('pending','processing','completed','failed') DEFAULT 'pending',
    error_details   JSON,
    user_id         INT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at    DATETIME,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- ============================================
-- DATI INIZIALI
-- ============================================

INSERT INTO users (username, password_hash, full_name, role) VALUES
('admin', '$2a$10$placeholder_hash_change_me', 'Administrator', 'admin');

INSERT INTO app_config (config_key, config_value, description) VALUES
('general.language', '"en"', 'Lingua predefinita'),
('general.dateFormat', '"dd/MM/yyyy"', 'Formato data'),
('general.theme', '"light"', 'Tema UI (light/dark)'),
('stats.defaultTestTypes', '["P_FINAL","P_PROD"]', 'TEST types usati per calcoli sigma'),
('stats.refreshInterval', '300', 'Intervallo refresh automatico (secondi)'),
('export.companyName', '"Hanna Instruments"', 'Nome azienda per report'),
('export.logoPath', '"/public/logo.svg"', 'Path logo per report');
```

### 7.3 Conteggio tabelle

| Area                    | Tabelle | Tabelle                                          |
| ----------------------- | ------- | ------------------------------------------------ |
| Auth / Users            | 2       | users, audit_log                                 |
| Anagrafica              | 2       | hanna_codes, product_standards                   |
| Reagenti QC             | 4       | production_lots, qc_readings, lot_sigma_distribution, lot_running_averages |
| Control Chart           | 1       | control_chart_limits                             |
| Buffer Production       | 2       | buffer_production, correction_materials          |
| Preparation List        | 2       | preparation_batches, preparation_hanna_codes     |
| Config / Log            | 2       | app_config, import_log                           |
| **Totale**              | **15**  | Espandibile con tabelle cache aggiuntive         |

---

## 8. Moduli Funzionali dell'Applicazione

### 8.1 Dashboard (Pagina principale)

**Scopo**: Vista d'insieme con KPI principali e grafici riassuntivi.

**Componenti** (tutti con Apache ECharts):
- KPI Cards: Hanna Codes tracciati, lotti QC nel mese, % letture entro 1σ, alert lotti >3σ
- Control Chart overview: trend ultime medie per prodotti attivi
- Sigma distribution: stacked bar dei lotti piu' recenti
- Tabella alert: ultimi lotti con letture fuori 3σ

**Filtri globali**:
- Periodo temporale (30gg, 90gg, personalizzato)
- Hanna Code / prodotto
- TEST type (P/Final, P/Prod, Valid...)

### 8.2 Reagenti QC Statistics

**Scopo**: Modulo principale — statistiche sigma sui file Chemical QC (da Gibertini software).
Un file per lotto per prodotto: letture da 1-4 meter per ogni Standard.

**Metriche**:
- Distribuzione sigma per lotto per STD: <1σ (verde), 1-2σ (blu), 2-3σ (giallo), >3σ
- Medie cumulative lot-by-lot per ogni STD (running average per control chart)
- Totale test per lotto, per STD, per operatore QC
- Peso tablet: deviazione da ref_weight per lotto
- σ = 50% Hanna Tolerance (valore fisso, NON deviazione standard statistica)

**Grafici**:
- **Control Chart (Shewhart)**: scatter letture + linea medie cumulative + bande sigma colorate. Un grafico per STD (es. 0, 15, 35, 60 ppm). Bande: 1σ verde, 2σ blu, 3σ giallo.
- **Sigma Distribution**: 100% Stacked Bar per lotto. Colori: verde/blu/giallo/rosso.
- **Gauge KPI**: % letture entro 1σ, 2σ, 3σ per lotto selezionato.

**Filtri**:
- Hanna Code (prodotto)
- Standard (STD 1-6)
- TEST type (P/Final, P/Prod, Valid, Old A/B/C/D)
- Range lotti / range date

### 8.3 Buffer Production Statistics

**Scopo**: Statistiche produzione soluzioni tampone pH (da Buffer Statistic xlsx).

**Metriche**:
- Quantita' prodotta per pH value (1.68, 4.01, 7.01, 10.01) nel tempo
- % lotti con correzione CM per pH
- Distribuzione CM usati per pH (quali materiali, quanti grammi, quale %)
- Trend correzione CM nel tempo (media mobile)

**Grafici**:
- LineChart: Trend cm_percentage per pH nel tempo
- BarChart: Quantita' Kg prodotti per mese per pH
- PieChart/Donut: Distribuzione CM codes per pH
- AreaChart: Pianificato vs prodotto (se disponibile)

**Filtri**:
- pH value (1.68 / 4.01 / 7.01 / 10.01)
- CM Code
- Range date

### 8.4 Preparation List Statistics

**Scopo**: Statistiche dai file PREP (Chemical Production) — compliance, operatori, ricette.

**Metriche**:
- Compliance settimane: planned week vs actual week per ricetta
- Numero batch per ricetta per mese
- Varianza pesi: (real_weight - theoretical_weight) / theoretical_weight
- Distribuzione operatori per ricetta
- Tracking lot number assegnati per prodotto

**Grafici**:
- BarChart: Batch per settimana per linea
- Scatter: Compliance planned vs actual week (delta settimane)
- BarChart: Varianza peso per ricetta (media ± dev)
- Timeline: Gantt-like visualizzazione batch nel tempo

**Filtri**:
- Recipe Code (CP-B051, ecc.)
- Operatore
- Range settimane / date

### 8.5 Trend Analysis

**Scopo**: Analisi trend su periodi personalizzati.

**Funzionalita'**:
- Selezione multipla metriche da confrontare
- Periodi personalizzabili
- Moving average (media mobile)
- Confronto periodo attuale vs precedente
- Esportazione serie temporali

### 8.7 Reports

**Scopo**: Generazione report formali.

**Template disponibili**:
- Report Reagenti QC: Control Chart + Sigma Distribution per prodotto/lotto
- Report Buffer Production: trend correzioni CM per periodo e pH
- Report Preparation List: compliance settimane, operatori, ricette
- Report personalizzato (selezione metriche cross-modulo)

**Formati di output**:
- PDF (jsPDF + jspdf-autotable, come MicroDens)
- Excel (ExcelJS, come MicroDens)
- CSV

### 8.8 Data Import

**Scopo**: Importazione dati da fonti esterne.

**Formati supportati**: Excel `.xlsx` (tutti i file sorgente sono Excel)

**Moduli di import**:
- **Chemical QC** (`CP-R14_HI93705B-0_LOT0676_*.xlsx`): letture QC raw per lotto, Standards, meter readings. Parser per posizione colonne (non per nome — nomi meter variano tra file).
- **Buffer Statistic** (`Buffer Preparation and Correction Statistic*.xlsx`): fogli pH 1.68/4.01/7.01/10.01, dati correzioni CM.
- **Preparation PREP** (`PREP_CP-B*.xlsx`, `PREP_B*-SOL.*.xlsx`): parsing layout fisso (posizioni righe note), estrazione Hanna Code Table.

**Funzionalita'**:
- Upload singolo file o batch (cartella)
- Preview struttura rilevata + dati da importare
- Validazione (test_type validi, range valori, date coerenti)
- Log importazioni con rollback
- Deduplicazione automatica (stesso file gia' importato)

### 8.9 Settings

**Scopo**: Configurazione applicazione.

**Sezioni** (come in MicroDens):
- Generale: lingua, tema, formato data
- Statistiche: periodo predefinito, intervallo refresh
- Export: nome azienda, logo, template
- Database: connessione, backup manuale
- Notifiche: soglie alert, destinatari

### 8.10 Users

**Scopo**: Gestione utenti e ruoli.

**Funzionalita'** (come AuthManager MicroDens):
- CRUD utenti
- Assegnazione ruoli (admin, manager, operator, viewer)
- Reset password
- Attivazione/disattivazione utenti
- Log attivita' utente (audit)

---

## 9. Tema e Branding (shadcn/ui + Tailwind CSS)

### 9.1 CSS Variables (globals.css)

shadcn/ui usa CSS custom properties per il tema, non un oggetto JavaScript
come MUI. Questo rende il dark mode automatico e le personalizzazioni immediate.

```css
/* src/globals.css — Tailwind v4: si usa @import, non @tailwind directives */
@import "tailwindcss";

@layer base {
  :root {
    /* Hanna Instruments Brand Colors */
    --background: 0 0% 98%;           /* #FAFAFA grigio chiaro */
    --foreground: 222 47% 11%;        /* Testo scuro */

    --primary: 224 100% 30%;          /* #003399 Hanna blue */
    --primary-foreground: 0 0% 100%;  /* Bianco su primary */

    --secondary: 24 100% 50%;         /* #FF6600 Hanna orange accent */
    --secondary-foreground: 0 0% 100%;

    --accent: 210 100% 60%;           /* #0066CC Hanna light blue */
    --accent-foreground: 0 0% 100%;

    --destructive: 4 90% 58%;         /* Rosso per errori/QC fail */
    --success: 122 39% 49%;           /* Verde per QC pass */
    --warning: 36 100% 50%;           /* Giallo per alert */

    --card: 0 0% 100%;
    --card-foreground: 222 47% 11%;
    --border: 214 32% 91%;
    --input: 214 32% 91%;
    --ring: 224 100% 30%;             /* Focus ring = Hanna blue */

    --radius: 0.5rem;
    --sidebar-width: 16rem;

    /* ECharts custom palette */
    --chart-1: 224 100% 30%;          /* Hanna blue */
    --chart-2: 24 100% 50%;           /* Hanna orange */
    --chart-3: 122 39% 49%;           /* Success green */
    --chart-4: 210 100% 60%;          /* Light blue */
    --chart-5: 4 90% 58%;             /* Red */
    --chart-6: 262 83% 58%;           /* Purple */
  }

  .dark {
    --background: 222 47% 6%;
    --foreground: 210 40% 98%;

    --primary: 217 91% 60%;           /* Lighter blue per dark mode */
    --primary-foreground: 222 47% 6%;

    --card: 222 47% 8%;
    --card-foreground: 210 40% 98%;
    --border: 217 33% 17%;
    --input: 217 33% 17%;
    --ring: 217 91% 60%;
  }
}
```

### 9.2 Tailwind Config

```css
/* Tailwind v4: configurazione direttamente in CSS */
@import "tailwindcss";

@theme {
  --font-sans: "Inter", "Roboto", "Helvetica", "Arial", sans-serif;
  --font-mono: "JetBrains Mono", "Fira Code", monospace;
}
```

### 9.3 ECharts Theme (matching Hanna brand)

```typescript
// src/config/echartsTheme.ts
export const hannaEChartsTheme = {
    color: [
        '#003399',  // Hanna blue
        '#FF6600',  // Hanna orange
        '#4CAF50',  // Success green
        '#0066CC',  // Light blue
        '#F44336',  // Red
        '#9C27B0',  // Purple
        '#00BCD4',  // Cyan
        '#FF9800',  // Amber
    ],
    backgroundColor: 'transparent',
    title: {
        textStyle: { color: '#1a1a2e', fontSize: 16, fontWeight: 600 },
    },
    tooltip: {
        backgroundColor: 'rgba(255,255,255,0.95)',
        borderColor: '#e5e7eb',
        textStyle: { color: '#1a1a2e' },
    },
    legend: {
        textStyle: { color: '#64748b' },
    },
};
```

### 9.4 Esempio Componente Dashboard con shadcn + ECharts

```tsx
// src/pages/Dashboard.tsx
import { motion } from 'framer-motion';
import ReactECharts from 'echarts-for-react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { TrendingUp, Package, CheckCircle, AlertTriangle } from 'lucide-react';

// Animazione container con stagger (figli appaiono uno dopo l'altro)
const containerVariants = {
    hidden: { opacity: 0 },
    visible: {
        opacity: 1,
        transition: { staggerChildren: 0.1 }
    }
};

const itemVariants = {
    hidden: { opacity: 0, y: 20 },
    visible: { opacity: 1, y: 0, transition: { duration: 0.5 } }
};

export function Dashboard() {
    return (
        <motion.div
            variants={containerVariants}
            initial="hidden"
            animate="visible"
            className="space-y-6 p-6"
        >
            {/* KPI Cards Row */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <motion.div variants={itemVariants}>
                    <Card className="border-l-4 border-l-primary">
                        <CardHeader className="flex flex-row items-center justify-between pb-2">
                            <CardTitle className="text-sm font-medium text-muted-foreground">
                                Lotti QC (30gg)
                            </CardTitle>
                            <Package className="h-4 w-4 text-primary" />
                        </CardHeader>
                        <CardContent>
                            <div className="text-2xl font-bold">47</div>
                            <p className="text-xs text-muted-foreground">
                                <span className="text-green-500">+3</span> vs mese precedente
                            </p>
                        </CardContent>
                    </Card>
                </motion.div>

                <motion.div variants={itemVariants}>
                    <Card className="border-l-4 border-l-green-500">
                        <CardHeader className="flex flex-row items-center justify-between pb-2">
                            <CardTitle className="text-sm font-medium text-muted-foreground">
                                Entro 1σ (30gg)
                            </CardTitle>
                            <CheckCircle className="h-4 w-4 text-green-500" />
                        </CardHeader>
                        <CardContent>
                            <div className="text-2xl font-bold">84.3%</div>
                            <Badge variant="secondary" className="mt-1">Target &gt;68%</Badge>
                        </CardContent>
                    </Card>
                </motion.div>

                {/* ... altre KPI cards ... */}
            </div>

            {/* Charts Row */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <motion.div variants={itemVariants}>
                    <Card>
                        <CardHeader>
                            <CardTitle>Sigma Distribution — Ultimi 10 Lotti</CardTitle>
                        </CardHeader>
                        <CardContent>
                            <ReactECharts
                                option={{
                                    tooltip: { trigger: 'axis' },
                                    xAxis: {
                                        type: 'category',
                                        data: ['Lun', 'Mar', 'Mer', 'Gio', 'Ven']
                                    },
                                    yAxis: { type: 'value' },
                                    series: [{
                                        type: 'bar',
                                        data: [1200, 1350, 1100, 1500, 1420],
                                        itemStyle: {
                                            borderRadius: [6, 6, 0, 0],
                                            color: {
                                                type: 'linear',
                                                x: 0, y: 0, x2: 0, y2: 1,
                                                colorStops: [
                                                    { offset: 0, color: '#0066CC' },
                                                    { offset: 1, color: '#003399' }
                                                ]
                                            }
                                        },
                                        animationDelay: (idx) => idx * 100
                                    }],
                                    animationEasing: 'elasticOut'
                                }}
                                style={{ height: 350 }}
                            />
                        </CardContent>
                    </Card>
                </motion.div>

                <motion.div variants={itemVariants}>
                    <Card>
                        <CardHeader>
                            <CardTitle>Running Average — STD 15 ppm</CardTitle>
                        </CardHeader>
                        <CardContent>
                            <ReactECharts
                                option={{
                                    tooltip: { trigger: 'axis' },
                                    xAxis: { type: 'category', data: ['Gen','Feb','Mar','Apr','Mag'] },
                                    yAxis: { type: 'value', min: 90, max: 100 },
                                    series: [{
                                        type: 'line',
                                        data: [96.5, 97.2, 97.8, 98.1, 98.2],
                                        smooth: true,
                                        areaStyle: {
                                            color: {
                                                type: 'linear',
                                                x: 0, y: 0, x2: 0, y2: 1,
                                                colorStops: [
                                                    { offset: 0, color: 'rgba(76,175,80,0.3)' },
                                                    { offset: 1, color: 'rgba(76,175,80,0.05)' }
                                                ]
                                            }
                                        },
                                        lineStyle: { color: '#4CAF50', width: 3 },
                                        symbol: 'circle',
                                        symbolSize: 8,
                                        animationDuration: 2000
                                    }]
                                }}
                                style={{ height: 350 }}
                            />
                        </CardContent>
                    </Card>
                </motion.div>
            </div>
        </motion.div>
    );
}
```

### 9.5 Esempio TanStack Table con shadcn/ui

```tsx
// src/components/Tables/DataTable.tsx
import {
    useReactTable, getCoreRowModel, getSortedRowModel,
    getFilteredRowModel, getPaginationRowModel, flexRender,
    type ColumnDef
} from '@tanstack/react-table';
import {
    Table, TableBody, TableCell, TableHead,
    TableHeader, TableRow
} from '@/components/ui/table';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { ArrowUpDown, Download } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

interface DataTableProps<T> {
    columns: ColumnDef<T, any>[];
    data: T[];
}

export function DataTable<T>({ columns, data }: DataTableProps<T>) {
    const table = useReactTable({
        data,
        columns,
        getCoreRowModel: getCoreRowModel(),
        getSortedRowModel: getSortedRowModel(),
        getFilteredRowModel: getFilteredRowModel(),
        getPaginationRowModel: getPaginationRowModel(),
    });

    return (
        <div className="space-y-4">
            {/* Toolbar */}
            <div className="flex items-center justify-between">
                <Input
                    placeholder="Cerca..."
                    className="max-w-sm"
                    onChange={(e) => table.setGlobalFilter(e.target.value)}
                />
                <div className="flex gap-2">
                    <Button variant="outline" size="sm">
                        <Download className="h-4 w-4 mr-2" /> Export Excel
                    </Button>
                    <Button variant="outline" size="sm">
                        <Download className="h-4 w-4 mr-2" /> Export PDF
                    </Button>
                </div>
            </div>

            {/* Table */}
            <div className="rounded-md border">
                <Table>
                    <TableHeader>
                        {table.getHeaderGroups().map(hg => (
                            <TableRow key={hg.id}>
                                {hg.headers.map(h => (
                                    <TableHead
                                        key={h.id}
                                        className="cursor-pointer select-none"
                                        onClick={h.column.getToggleSortingHandler()}
                                    >
                                        <div className="flex items-center gap-1">
                                            {flexRender(h.column.columnDef.header, h.getContext())}
                                            <ArrowUpDown className="h-3 w-3" />
                                        </div>
                                    </TableHead>
                                ))}
                            </TableRow>
                        ))}
                    </TableHeader>
                    <TableBody>
                        <AnimatePresence>
                            {table.getRowModel().rows.map((row, i) => (
                                <motion.tr
                                    key={row.id}
                                    initial={{ opacity: 0, x: -20 }}
                                    animate={{ opacity: 1, x: 0 }}
                                    transition={{ delay: i * 0.03 }}
                                    className="border-b hover:bg-muted/50"
                                >
                                    {row.getVisibleCells().map(cell => (
                                        <TableCell key={cell.id}>
                                            {flexRender(cell.column.columnDef.cell, cell.getContext())}
                                        </TableCell>
                                    ))}
                                </motion.tr>
                            ))}
                        </AnimatePresence>
                    </TableBody>
                </Table>
            </div>

            {/* Pagination */}
            <div className="flex items-center justify-between">
                <p className="text-sm text-muted-foreground">
                    Showing {table.getRowModel().rows.length} of {data.length} entries
                </p>
                <div className="flex gap-2">
                    <Button
                        variant="outline" size="sm"
                        onClick={() => table.previousPage()}
                        disabled={!table.getCanPreviousPage()}
                    >
                        Precedente
                    </Button>
                    <Button
                        variant="outline" size="sm"
                        onClick={() => table.nextPage()}
                        disabled={!table.getCanNextPage()}
                    >
                        Successivo
                    </Button>
                </div>
            </div>
        </div>
    );
}
```

---

## 10. Deploy su NAS / Server Hanna

### 10.1 Opzione A: Docker (consigliata)

```yaml
# docker/docker-compose.yml
version: '3.8'

services:
  app:
    build:
      context: ..
      dockerfile: docker/Dockerfile
    ports:
      - "83:3000"                     # Accessibile su porta 83
    environment:
      - NODE_ENV=production
      - DB_HOST=db
      - DB_PORT=3306
      - DB_USER=hanna_stats
      - DB_PASSWORD=${DB_PASSWORD}
      - DB_NAME=hanna_statistics
      - JWT_SECRET=${JWT_SECRET}
    depends_on:
      - db
    restart: unless-stopped

  db:
    image: mariadb:10.11
    volumes:
      - db_data:/var/lib/mysql
      - ./init:/docker-entrypoint-initdb.d   # Script inizializzazione
    environment:
      - MYSQL_ROOT_PASSWORD=${MYSQL_ROOT_PASSWORD}
      - MYSQL_DATABASE=hanna_statistics
      - MYSQL_USER=hanna_stats
      - MYSQL_PASSWORD=${DB_PASSWORD}
    ports:
      - "3307:3306"                   # Porta diversa da eventuale altro MariaDB
    restart: unless-stopped

volumes:
  db_data:
```

```dockerfile
# docker/Dockerfile
# Stage 1: Build frontend
FROM node:20-alpine AS frontend-build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build

# Stage 2: Build server
FROM node:20-alpine AS server-build
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY --from=frontend-build /app/dist ./dist
COPY --from=frontend-build /app/dist-server ./dist-server

# Stage 3: Production
FROM node:20-alpine
WORKDIR /app
COPY --from=server-build /app .
EXPOSE 3000
CMD ["node", "dist-server/server.js"]
```

### 10.2 Opzione B: PM2 (senza Docker)

```bash
# Installazione su NAS con Node.js
npm install -g pm2
npm run build
pm2 start dist-server/server.js --name hanna-statistics
pm2 save
pm2 startup
```

### 10.3 Reverse Proxy Nginx (opzionale)

```nginx
# Per servire su https://192.168.1.36:82/statistics
# aggiungendo un location block al Nginx esistente

location /statistics {
    proxy_pass http://127.0.0.1:3000;
    proxy_http_version 1.1;
    proxy_set_header Upgrade $http_upgrade;
    proxy_set_header Connection 'upgrade';
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_cache_bypass $http_upgrade;
}
```

---

## 11. Variabili d'Ambiente

```env
# .env.example

# Server
NODE_ENV=production
PORT=3000
HOST=0.0.0.0

# Database
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=hanna_stats
DB_PASSWORD=changeme_strong_password
DB_NAME=hanna_statistics

# Auth
JWT_SECRET=changeme_random_64_char_string
JWT_EXPIRATION=8h
BCRYPT_ROUNDS=10

# App
DEFAULT_LANGUAGE=en
COMPANY_NAME=Hanna Instruments
```

---

## 12. Requisiti Infrastrutturali

### 12.1 Requisiti MINIMI (stesse macchine di Hanna Core)

| Componente       | Specifica                                          |
| ---------------- | -------------------------------------------------- |
| **Hardware**     | Gia' presente: NAS Synology/QNAP con Docker        |
| **RAM aggiuntiva** | +512MB (Node.js) + 256MB (MariaDB) = ~768MB      |
| **Disco**        | +2GB per app + 5-10GB per DB = ~12GB               |
| **Porta rete**   | 1 porta aggiuntiva (es. 83) o reverse proxy        |
| **Node.js**      | 20 LTS (in Docker o installato su NAS)             |
| **MariaDB**      | 10.11+ (in Docker o pacchetto NAS)                 |
| **Browser**      | Chrome/Edge/Firefox moderno (gia' presente)        |
| **Team**         | 1 sviluppatore (tu, con esperienza MicroDens)      |

### 12.2 Requisiti MASSIMI (produzione enterprise)

| Componente       | Specifica                                          |
| ---------------- | -------------------------------------------------- |
| **Server**       | Dedicato: 4 core, 8GB RAM, SSD 100GB               |
| **Redis**        | Per cache statistiche precalcolate                  |
| **Backup**       | Automatico giornaliero DB (cron + mysqldump)        |
| **Monitoring**   | PM2 monitoring o Grafana                            |
| **SSL**          | Certificato valido (CA interna o Let's Encrypt)     |
| **Team**         | 1-2 sviluppatori                                    |

### 12.3 Confronto risorse con MicroDens

| Aspetto              | MicroDens Logger Pro        | Hanna Statistics            |
| -------------------- | --------------------------- | --------------------------- |
| **Runtime**          | Electron (200-300MB RAM)    | Node.js (100-200MB RAM)    |
| **Storage**          | JSON files (leggero)        | MariaDB (piu' strutturato) |
| **CPU**              | Per rendering Chromium       | Per query e calcoli stats  |
| **Disco**            | ~100MB installazione         | ~50MB app + DB variabile   |
| **Aggiornamenti**    | electron-updater             | docker pull / git pull     |
| **Complessita' ops** | Installazione desktop       | Docker compose up          |

---

## 13. Piano di Sviluppo

### Fase 1 — Fondamenta (Settimane 1-3)

- [ ] Inizializzazione progetto (Vite + React + TS + shadcn/ui + Tailwind CSS 4)
- [ ] Setup server Express con TypeScript
- [ ] Connessione MariaDB (mysql2 + pool)
- [ ] AuthService + JWT (tradotto da AuthManager MicroDens)
- [ ] Layout base con shadcn/ui (navbar Hanna blue, sidebar)
- [ ] ThemeContext e i18n (copiati da MicroDens)
- [ ] Docker Compose (app + MariaDB)

### Fase 2 — Importazione Dati (Settimane 4-5)

- [ ] ImportService (CSV, Excel, JSON)
- [ ] UI importazione con preview e mapping colonne
- [ ] Migrazioni DB (tutte le tabelle)
- [ ] Log importazioni

### Fase 3 — Moduli Statistici (Settimane 6-10)

- [ ] Dashboard con KPI cards e grafici Apache ECharts
- [ ] Reagenti QC Statistics (Control Chart, distribuzione sigma, letture per lotto)
- [ ] Buffer Production Statistics (trend correzioni, quantita' prodotte per pH)
- [ ] Preparation List (statistiche prep per ricetta, compliance settimane)
- [ ] Trend Analysis cross-modulo

### Fase 4 — Export e Report (Settimane 11-12)

- [ ] ExportService (tradotto da ExportManager MicroDens)
- [ ] Report PDF con template Hanna
- [ ] Report Excel formattati
- [ ] Export CSV

### Fase 5 — Rilascio (Settimane 13-14)

- [ ] Deploy Docker su NAS Hanna
- [ ] Test con dati reali
- [ ] Configurazione rete (porta, SSL, reverse proxy)
- [ ] Formazione utenti
- [ ] Go-live

### Tempo totale stimato: 14 settimane (3.5 mesi)

Grazie al riutilizzo dell'80% del codice frontend MicroDens, i tempi si riducono
significativamente rispetto allo sviluppo da zero (~6 mesi).

---

## 14. Riepilogo: Cosa si Riutilizza da MicroDens e Cosa Evolve

### 14.1 Riutilizzo diretto (know-how e pattern)

| Elemento                       | Riutilizzo | Note                                      |
| ------------------------------ | ---------- | ----------------------------------------- |
| React + TypeScript setup       | 100%       | Identico, upgrade a React 19              |
| Context pattern (Auth, Theme)  | 90%        | Auth aggiunge JWT, ThemeContext invariato  |
| ExportService (Excel, PDF)     | 85%        | Output su HTTP anziche' file              |
| statisticsCalculator.ts        | 100%       | Riutilizzabile direttamente               |
| i18n setup                     | 90%        | Aggiunta lingue RO, HU                   |
| Types e interfaces             | 70%        | Nuovi types per i dati Hanna              |
| Manager → Service pattern      | 80%        | Stessa architettura, diverso I/O          |
| Build system (Vite)            | 90%        | Rimosso plugin Electron                   |

### 14.2 Evoluzione (nuove tecnologie)

| Da (MicroDens)         | A (Hanna Statistics)        | Vantaggio                                    |
| ---------------------- | --------------------------- | -------------------------------------------- |
| MUI v7 + Emotion       | shadcn/ui + Tailwind CSS 4  | -250KB bundle, UI premium, zero runtime CSS  |
| ReCharts               | Apache ECharts 5.5          | +20 tipi grafico, animazioni, Canvas perf.   |
| MUI Table (basic)      | TanStack Table v8           | Virtualizzazione, resize, faceted filters    |
| Nessuna animazione     | Framer Motion 11            | Transizioni pagina, stagger, spring physics  |
| Material Icons         | Lucide React                | 1000+ icone, tree-shakeable, ~3KB vs ~40KB  |
| Electron (desktop)     | Node.js + Express (web)     | Accessibile da qualsiasi browser/dispositivo |
| JSON files             | MariaDB                     | Query SQL, relazioni, scalabilita'           |

### 14.3 Stima complessiva

```
Know-how riutilizzato:      ~80% (React, TS, Vite, pattern, logica business)
Codice riutilizzato:        ~40% (services, utils, types, contexts)
Codice nuovo:               ~60% (UI shadcn, grafici ECharts, backend Express, DB)
Tempo risparmiato:          ~35-40% rispetto a sviluppo da zero
```

---

## 15. Stack Finale a Colpo d'Occhio

```
┌────────────────────────────────────────────────────────────┐
│                    HANNA STATISTICS                          │
│                                                              │
│  Frontend:                                                   │
│    React 19 + TypeScript 5.9 + Vite 7.2                     │
│    shadcn/ui + Tailwind CSS 4 + Radix UI                    │
│    Apache ECharts 5.5 (grafici interattivi)                  │
│    TanStack Table v8 (tabelle performanti)                   │
│    Framer Motion 11 (animazioni fluide)                      │
│    Lucide React (icone moderne)                              │
│    i18next (EN/IT/RO/HU)                                     │
│                                                              │
│  Backend:                                                    │
│    Node.js 20 LTS + Express.js                               │
│    mysql2 (connection pool)                                   │
│    bcryptjs + JWT (autenticazione)                            │
│    ExcelJS + jsPDF (export)                                   │
│                                                              │
│  Database:                                                   │
│    MariaDB 10.11+ (indipendente)                             │
│                                                              │
│  Deploy:                                                     │
│    Docker Compose (app + DB)                                  │
│    Su NAS/server Hanna Instruments                           │
│    Accessibile: https://192.168.1.36:83                      │
│                                                              │
│  Pattern: Manager/Service (da MicroDens Logger Pro)          │
└────────────────────────────────────────────────────────────┘
```

---

> **Documento generato il**: 2026-02-12
> **Basato su**: Analisi progetto MicroDens Logger Pro + screenshot Hanna Core
> **Tecnologia scelta**: shadcn/ui + Tailwind CSS + Apache ECharts + Framer Motion
> **Architettura di riferimento**: Manager Pattern (da MicroDens Logger Pro)
> **Destinazione**: NAS/server Hanna Instruments, rete locale
