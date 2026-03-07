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
│   │   ├── Dashboard.tsx         # Dashboard principale
│   │   ├── ProductionStats.tsx   # Statistiche produzione
│   │   ├── QualityStats.tsx      # Statistiche QC
│   │   ├── PackingStats.tsx      # Statistiche confezionamento
│   │   ├── StockAnalysis.tsx     # Analisi stock/magazzino
│   │   ├── TrendAnalysis.tsx     # Analisi trend temporali
│   │   ├── Reports.tsx           # Generazione report
│   │   ├── DataImport.tsx        # Importazione dati
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
│   │   │   ├── 002_create_production_data.sql
│   │   │   ├── 003_create_quality_data.sql
│   │   │   ├── 004_create_packing_data.sql
│   │   │   ├── 005_create_stock_data.sql
│   │   │   ├── 006_create_statistics_cache.sql
│   │   │   └── 007_create_config.sql
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
├── tailwind.config.ts            # Configurazione Tailwind CSS (v4: in globals.css)
├── postcss.config.js             # PostCSS per Tailwind
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
-- (Stessa struttura di AuthManager MicroDens)
-- ============================================

CREATE TABLE users (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    username        VARCHAR(50) UNIQUE NOT NULL,
    password_hash   VARCHAR(255) NOT NULL,          -- bcryptjs (come MicroDens)
    full_name       VARCHAR(100),
    role            ENUM('admin','manager','operator','viewer') NOT NULL DEFAULT 'viewer',
    language        VARCHAR(5) DEFAULT 'en',        -- en, it, ro, hu
    is_active       BOOLEAN DEFAULT TRUE,
    last_login      DATETIME,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE audit_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    user_id         INT,
    action          VARCHAR(50) NOT NULL,           -- login, export, import, etc.
    details         JSON,
    ip_address      VARCHAR(45),
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- ============================================
-- DATI DI PRODUZIONE
-- (Importati o inseriti manualmente)
-- ============================================

CREATE TABLE production_data (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    order_number    VARCHAR(30),
    product_code    VARCHAR(50) NOT NULL,
    product_name    VARCHAR(100),
    line            VARCHAR(10),
    lot             VARCHAR(30),
    quantity_planned INT,
    quantity_produced INT,
    quantity_rejected INT DEFAULT 0,
    start_date      DATETIME,
    end_date        DATETIME,
    operator        VARCHAR(100),
    shift           ENUM('morning','afternoon','night'),
    notes           TEXT,
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    source          VARCHAR(50) DEFAULT 'manual'     -- manual, import, api
);

CREATE INDEX idx_prod_date ON production_data(start_date);
CREATE INDEX idx_prod_product ON production_data(product_code);
CREATE INDEX idx_prod_line ON production_data(line);

-- ============================================
-- DATI CONTROLLO QUALITA'
-- ============================================

CREATE TABLE quality_data (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    register_nr     VARCHAR(30),
    record_date     DATETIME NOT NULL,
    product_code    VARCHAR(50) NOT NULL,
    lot             VARCHAR(30),
    client_code     VARCHAR(20),
    recipe          VARCHAR(30),
    qc_type         ENUM('first','during','final') NOT NULL,
    result          ENUM('pass','fail','pending') DEFAULT 'pending',
    sampling_user   VARCHAR(100),
    sampling_date   DATETIME,
    exp_date        DATE,
    notes           TEXT,
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    source          VARCHAR(50) DEFAULT 'manual'
);

CREATE INDEX idx_qc_date ON quality_data(record_date);
CREATE INDEX idx_qc_product ON quality_data(product_code);
CREATE INDEX idx_qc_result ON quality_data(result);

CREATE TABLE quality_parameters (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    quality_data_id INT NOT NULL,
    parameter_name  VARCHAR(50) NOT NULL,
    expected_value  DECIMAL(12,4),
    measured_value  DECIMAL(12,4),
    tolerance_min   DECIMAL(12,4),
    tolerance_max   DECIMAL(12,4),
    unit            VARCHAR(20),
    is_pass         BOOLEAN,
    FOREIGN KEY (quality_data_id) REFERENCES quality_data(id) ON DELETE CASCADE
);

-- ============================================
-- DATI CONFEZIONAMENTO (PACKING)
-- ============================================

CREATE TABLE packing_data (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_code        VARCHAR(30) NOT NULL,
    fg_code         VARCHAR(30) NOT NULL,
    tactile         VARCHAR(30),
    line            VARCHAR(10),
    status          ENUM('in_progress','completed','delayed','critical') DEFAULT 'completed',
    elapsed_seconds INT DEFAULT 0,
    pcs_on_pallet   INT DEFAULT 0,
    packing_operator VARCHAR(150),
    exp_date        DATE,
    started_at      DATETIME,
    completed_at    DATETIME,
    claims          TEXT,
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    source          VARCHAR(50) DEFAULT 'manual'
);

CREATE TABLE packing_sfg (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    packing_data_id INT NOT NULL,
    sfg_position    TINYINT NOT NULL,               -- 1-7
    sfg_code        VARCHAR(50),
    sfg_lot         VARCHAR(30),
    sfg_quantity    INT,
    sfg_exp_date    DATE,
    FOREIGN KEY (packing_data_id) REFERENCES packing_data(id) ON DELETE CASCADE
);

CREATE INDEX idx_pack_date ON packing_data(completed_at);
CREATE INDEX idx_pack_fg ON packing_data(fg_code);

-- ============================================
-- DATI STOCK / MAGAZZINO
-- ============================================

CREATE TABLE stock_data (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    stock_code      VARCHAR(50) NOT NULL,
    line            VARCHAR(10),
    lot             VARCHAR(30),
    iso             VARCHAR(30),
    exp_date        DATE,
    quantity        INT DEFAULT 0,
    standard_value  DECIMAL(10,2),
    coverage        DECIMAL(10,2),
    location_hall   VARCHAR(30),
    shelf           VARCHAR(20),
    recipe          VARCHAR(30),
    snapshot_date   DATE NOT NULL,                  -- data dello snapshot
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    source          VARCHAR(50) DEFAULT 'manual'
);

CREATE INDEX idx_stock_date ON stock_data(snapshot_date);
CREATE INDEX idx_stock_code ON stock_data(stock_code);

-- ============================================
-- STATISTICHE PRECALCOLATE (CACHE)
-- ============================================

CREATE TABLE statistics_cache (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    stat_type       VARCHAR(50) NOT NULL,           -- production_daily, qc_pass_rate, etc.
    period_start    DATE NOT NULL,
    period_end      DATE NOT NULL,
    dimensions      JSON,                           -- {"line": "L58", "product": "HI97700B"}
    metrics         JSON,                           -- {"total": 1500, "avg": 75.5, "stddev": 2.3}
    calculated_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uk_stat (stat_type, period_start, period_end, dimensions(200))
);

-- ============================================
-- CONFIGURAZIONE APPLICAZIONE
-- (Sostituisce electron-store)
-- ============================================

CREATE TABLE app_config (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    config_key      VARCHAR(100) UNIQUE NOT NULL,
    config_value    JSON NOT NULL,
    description     VARCHAR(255),
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ============================================
-- IMPORTAZIONI
-- ============================================

CREATE TABLE import_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    filename        VARCHAR(255) NOT NULL,
    file_type       ENUM('csv','xlsx','json') NOT NULL,
    target_table    VARCHAR(50) NOT NULL,
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

-- Admin di default (password: admin — da cambiare al primo accesso)
INSERT INTO users (username, password_hash, full_name, role) VALUES
('admin', '$2a$10$placeholder_hash_change_me', 'Administrator', 'admin');

-- Configurazione di default
INSERT INTO app_config (config_key, config_value, description) VALUES
('general.language', '"en"', 'Lingua predefinita'),
('general.dateFormat', '"dd/MM/yyyy"', 'Formato data'),
('general.theme', '"light"', 'Tema UI (light/dark)'),
('stats.defaultPeriod', '"30d"', 'Periodo predefinito per le statistiche'),
('stats.refreshInterval', '300', 'Intervallo refresh automatico (secondi)'),
('export.companyName', '"Hanna Instruments"', 'Nome azienda per report'),
('export.logoPath', '"/public/logo.svg"', 'Path logo per report');
```

### 7.3 Conteggio tabelle

| Area                  | Tabelle | Note                                |
| --------------------- | ------- | ----------------------------------- |
| Auth / Users          | 2       | users, audit_log                    |
| Production            | 1       | production_data                     |
| Quality Control       | 2       | quality_data, quality_parameters    |
| Packing               | 2       | packing_data, packing_sfg           |
| Stock                 | 1       | stock_data                          |
| Statistics Cache      | 1       | statistics_cache                    |
| Config                | 1       | app_config                          |
| Import                | 1       | import_log                          |
| **Totale**            | **11**  | Espandibile a 20-25                 |

---

## 8. Moduli Funzionali dell'Applicazione

### 8.1 Dashboard (Pagina principale)

**Scopo**: Vista d'insieme con KPI principali e grafici riassuntivi.

**Componenti** (tutti con ReCharts, come in MicroDens):
- KPI Cards: Produzione totale, QC Pass Rate, Stock critico, Packing efficienza
- Grafico produzione giornaliera (BarChart)
- Grafico andamento QC (LineChart)
- Grafico distribuzione stock (PieChart)
- Tabella ultime attivita' / alert

**Filtri globali**:
- Periodo temporale (oggi, 7gg, 30gg, personalizzato)
- Linea produttiva
- Prodotto

### 8.2 Production Stats

**Scopo**: Analisi dettagliata della produzione.

**Metriche**:
- Produzione per linea / per prodotto / per turno
- Quantita' pianificata vs prodotta (rendimento)
- Trend produttivo nel tempo
- Scarti e quantita' rifiutate
- Tempi di produzione medi

**Grafici**:
- BarChart: Produzione per linea
- LineChart: Trend giornaliero/settimanale/mensile
- AreaChart: Pianificato vs Prodotto
- Scatter: Correlazione quantita'/tempo

### 8.3 Quality Stats

**Scopo**: Analisi controllo qualita'.

**Metriche**:
- QC Pass Rate (% superamento) globale e per prodotto
- Distribuzione per tipo QC (First, During, Final)
- Parametri fuori specifica (frequenza e entita')
- Tempo medio tra campionamento e risultato
- Trend qualita' nel tempo

**Grafici**:
- PieChart: Pass/Fail ratio
- LineChart: Andamento pass rate nel tempo
- BarChart: Parametri piu' critici
- Heatmap: Qualita' per prodotto/periodo

### 8.4 Packing Stats

**Scopo**: Analisi efficienza confezionamento.

**Metriche**:
- Tempo medio di confezionamento per prodotto
- Distribuzione stati (verde/giallo/rosso)
- Pezzi per pallet medi
- Efficienza per operatore
- Claims per prodotto/periodo

**Grafici**:
- BarChart: Tempi per operatore
- LineChart: Trend efficienza nel tempo
- PieChart: Distribuzione stati
- BarChart: Claims per categoria

### 8.5 Stock Analysis

**Scopo**: Analisi magazzino e scorte.

**Metriche**:
- Livelli stock attuali per prodotto
- Prodotti in scadenza (alert)
- Copertura stock (giorni rimanenti)
- Rotazione magazzino
- Trend giacenze nel tempo

**Grafici**:
- BarChart: Stock per location/shelf
- LineChart: Andamento giacenze
- Gauge: Copertura stock
- Treemap: Distribuzione per categoria

### 8.6 Trend Analysis

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
- Report produzione giornaliero/settimanale/mensile
- Report QC con dettaglio parametri
- Report stock con alert scadenze
- Report packing con efficienza operatori
- Report personalizzato (selezione metriche)

**Formati di output**:
- PDF (jsPDF + jspdf-autotable, come MicroDens)
- Excel (ExcelJS, come MicroDens)
- CSV

### 8.8 Data Import

**Scopo**: Importazione dati da fonti esterne.

**Formati supportati**:
- CSV (con selezione separatore e mapping colonne)
- Excel (.xlsx)
- JSON

**Funzionalita'**:
- Preview dati prima dell'importazione
- Mapping colonne sorgente → colonne DB
- Validazione dati con report errori
- Log importazioni con possibilita' di rollback
- Import schedulato (opzionale)

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
/* src/globals.css */
@tailwind base;
@tailwind components;
@tailwind utilities;

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
                                Produzione Oggi
                            </CardTitle>
                            <Package className="h-4 w-4 text-primary" />
                        </CardHeader>
                        <CardContent>
                            <div className="text-2xl font-bold">1,523</div>
                            <p className="text-xs text-muted-foreground">
                                <span className="text-green-500">+12%</span> vs ieri
                            </p>
                        </CardContent>
                    </Card>
                </motion.div>

                <motion.div variants={itemVariants}>
                    <Card className="border-l-4 border-l-green-500">
                        <CardHeader className="flex flex-row items-center justify-between pb-2">
                            <CardTitle className="text-sm font-medium text-muted-foreground">
                                QC Pass Rate
                            </CardTitle>
                            <CheckCircle className="h-4 w-4 text-green-500" />
                        </CardHeader>
                        <CardContent>
                            <div className="text-2xl font-bold">98.2%</div>
                            <Badge variant="secondary" className="mt-1">Eccellente</Badge>
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
                            <CardTitle>Produzione Settimanale</CardTitle>
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
                            <CardTitle>QC Pass Rate Trend</CardTitle>
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

- [ ] Inizializzazione progetto (Vite + React + TS + MUI)
- [ ] Setup server Express con TypeScript
- [ ] Connessione MariaDB (mysql2 + pool)
- [ ] AuthService + JWT (tradotto da AuthManager MicroDens)
- [ ] Layout base con MUI (navbar Hanna blue, sidebar)
- [ ] ThemeContext e i18n (copiati da MicroDens)
- [ ] Docker Compose (app + MariaDB)

### Fase 2 — Importazione Dati (Settimane 4-5)

- [ ] ImportService (CSV, Excel, JSON)
- [ ] UI importazione con preview e mapping colonne
- [ ] Migrazioni DB (tutte le tabelle)
- [ ] Log importazioni

### Fase 3 — Moduli Statistici (Settimane 6-10)

- [ ] Dashboard con KPI cards e grafici ReCharts
- [ ] Production Stats (metriche + grafici)
- [ ] Quality Stats (pass rate, parametri, trend)
- [ ] Packing Stats (efficienza, tempi, claims)
- [ ] Stock Analysis (livelli, scadenze, copertura)

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
