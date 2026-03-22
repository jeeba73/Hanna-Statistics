# Hanna Statistics — Guida Sviluppo

> Setup completo: dal PC Windows a Milano al server Hetzner in Germania.
> Tutto quello che serve per sviluppare con VSCode + Claude e deployare.

---

## 1. Installazioni sul tuo PC (una volta sola)

### 1.1 Node.js 22 LTS

Scarica e installa da: **https://nodejs.org** (bottone LTS, quello verde)

Verifica dopo l'installazione (apri nuovo PowerShell):
```powershell
node --version    # deve dire v22.x.x
npm --version     # deve dire 11.x.x
```

### 1.2 Docker Desktop

Scarica e installa da: **https://www.docker.com/products/docker-desktop/**

- Durante l'installazione lascia tutto di default
- Potrebbe chiederti di abilitare WSL2 — dì di sì
- Potrebbe richiedere un riavvio del PC
- Dopo il riavvio, apri Docker Desktop e aspetta che dica "running"

Verifica:
```powershell
docker --version          # deve dire Docker version 2x.x.x
docker compose version    # deve dire Docker Compose version v2.x.x
```

### 1.3 Estensione VSCode: Remote - SSH

In VSCode:
1. Apri Extensions (Ctrl+Shift+X)
2. Cerca **"Remote - SSH"** (di Microsoft)
3. Installa

Questa estensione ti permette di aprire VSCode direttamente sul server Hetzner,
come se fosse una cartella locale. Molto utile per debug sul server.

---

## 2. Struttura del Progetto

```
hanna-statistics/
├── client/                  # Frontend React (quello che vede l'utente)
│   ├── src/
│   │   ├── components/      # Componenti React (grafici, tabelle, filtri)
│   │   ├── pages/           # Pagine (Dashboard, QC Statistics, Buffer, ecc.)
│   │   ├── services/        # Chiamate API al backend (fetch)
│   │   ├── i18n/            # Traduzioni EN/RO
│   │   ├── lib/             # Utility (cn() per classnames)
│   │   └── App.tsx
│   ├── index.html
│   ├── package.json
│   ├── vite.config.ts
│   └── components.json      # Configurazione shadcn/ui
│
├── server/                  # Backend Node.js (API + logica)
│   ├── src/
│   │   ├── routes/          # Endpoint API (GET /api/health, ecc.)
│   │   ├── services/        # Logica business (import Excel, calcoli sigma)
│   │   ├── database/        # Connessione DB + migrations + seed
│   │   │   ├── connection.ts
│   │   │   ├── migrations/  # 7 file SQL (001-007, schema 15 tabelle)
│   │   │   └── seed/        # Dati iniziali
│   │   ├── app.ts           # Express app (middleware, routes, static)
│   │   └── server.ts        # Entry point (dotenv + avvio)
│   ├── package.json
│   └── tsconfig.json
│
├── shared/                  # Tipi TypeScript condivisi client/server
│   └── types/
│       └── index.ts
│
├── docker/
│   ├── docker-compose.yml       # Per il server (produzione/demo)
│   ├── docker-compose.dev.yml   # Per lo sviluppo locale
│   └── Dockerfile               # Multi-stage build (3 stage)
│
├── .env                     # Variabili ambiente (git-ignored)
├── .env.example             # Template per .env
├── .gitignore
└── package.json             # Root package.json (script condivisi)
```

### Confronto con Electron (quello che conosci)

```
ELECTRON (prima)                    WEB APP (adesso)
─────────────────                   ──────────────────
src/main/         → main process    server/    → Express (API)
src/renderer/     → React UI        client/    → React UI (identico!)
src/preload/      → bridge IPC      (non serve — HTTP è il bridge)
electron-builder  → packaging       docker/    → containerizzazione
```

---

## 3. Flusso di Lavoro Quotidiano

### 3.1 Sviluppo locale (il tuo PC, ogni giorno)

Metodo rapido (un solo comando dalla root del progetto):
```powershell
cd hanna-statistics
npm run dev
# Avvia tutto: MariaDB (Docker) + Backend (tsx) + Frontend (Vite)
```

Oppure manualmente in 3 terminali separati:
```
Terminal 1 — Database (Docker)
───────────────────────────────
cd hanna-statistics
npm run db:up
# MariaDB parte su localhost:3306

Terminal 2 — Backend
───────────────────────────────
cd hanna-statistics/server
npm run dev
# Express parte su localhost:3000, hot reload con tsx watch

Terminal 3 — Frontend
───────────────────────────────
cd hanna-statistics/client
npm run dev
# Vite parte su localhost:5173, hot reload immediato
# Proxy automatico: /api/* → localhost:3000
```

Apri **http://localhost:5173** nel browser.
Modifichi un file → il browser si aggiorna da solo.
**Identico a quando facevi Electron**, solo senza la finestra Electron.

### 3.2 Deploy su Hetzner (quando vuoi aggiornare la demo)

```powershell
# Dal tuo PC, PowerShell
git add . && git commit -m "descrizione modifica" && git push

# Collegati al server
ssh root@178.104.16.85

# Sul server
cd /opt/Hanna-Statistics
git pull
cd hanna-statistics
docker compose -f docker/docker-compose.yml up -d --build

# Fine. L'app è aggiornata su http://demo.bilsoft.it
```

Tempo totale: **2 minuti**.

---

## 4. Come Parlano Frontend e Backend (la differenza chiave)

### In Electron (prima):

```typescript
// Frontend (renderer) chiama il backend (main) via IPC
const lots = await window.electronAPI.getLots({ hannaCode: 'HI782-0' });
```

```typescript
// Backend (main) risponde via IPC
ipcMain.handle('getLots', async (event, { hannaCode }) => {
  const db = getDatabase();
  return db.query('SELECT * FROM lots WHERE hanna_code = ?', [hannaCode]);
});
```

### In Web App (adesso):

```typescript
// Frontend (React) chiama il backend (Express) via HTTP
const response = await fetch('/api/lots?hannaCode=HI782-0');
const lots = await response.json();
```

```typescript
// Backend (Express) risponde via HTTP
app.get('/api/lots', async (req, res) => {
  const { hannaCode } = req.query;
  const lots = await db.query('SELECT * FROM lots WHERE hanna_code = ?', [hannaCode]);
  res.json(lots);
});
```

**La logica è identica. Cambia solo il "trasporto": IPC → HTTP.**

---

## 5. Database: da SQLite/JSON a MariaDB

### Prima (Electron):
```typescript
// SQLite o file JSON
const db = new Database('data.sqlite');
const rows = db.prepare('SELECT * FROM lots').all();
```

### Adesso (MariaDB):
```typescript
// MariaDB via mysql2
import mysql from 'mysql2/promise';

const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  port: 3306,
  user: 'hanna_stats',
  password: 'HannaStats2026!',
  database: 'hanna_statistics'
});

const [rows] = await pool.query('SELECT * FROM lots WHERE hanna_code = ?', ['HI782-0']);
```

**Stessa sintassi SQL.** MariaDB è MySQL-compatibile, la query language è identica.
La differenza è che MariaDB è un server separato (gira in un container Docker),
quindi usi un "connection pool" per connetterti via rete invece di aprire un file.

---

## 6. Comandi di Riferimento

### Sviluppo locale
```powershell
# Primo setup (una volta)
git clone https://github.com/jeeba73/Hanna-Statistics.git
cd Hanna-Statistics/hanna-statistics
npm run install:all    # installa deps root + client + server

# Ogni giorno (metodo rapido)
npm run dev            # avvia tutto (Docker DB + backend + frontend)

# Oppure manualmente
npm run db:up                                               # avvia MariaDB
cd server && npm run dev                                     # avvia backend
cd client && npm run dev                                     # avvia frontend
```

### Server Hetzner
```powershell
# Collegamento
ssh root@178.104.16.85

# Deploy / aggiornamento
cd /opt/Hanna-Statistics && git pull
cd hanna-statistics
docker compose -f docker/docker-compose.yml up -d --build

# Comandi utili sul server
docker logs hanna-stats-app                # vedi log dell'app
docker logs hanna-stats-db                 # vedi log del database
docker compose -f docker/docker-compose.yml down    # ferma tutto
docker compose -f docker/docker-compose.yml restart app  # riavvia solo l'app
docker ps                                  # vedi container attivi
```

### Git (già lo sai, ma per completezza)
```powershell
git add .
git commit -m "messaggio"
git push origin main
```

---

## 7. Mappa Mentale: Electron → Web App

| Concetto | Electron | Web App |
|----------|----------|---------|
| **L'utente apre** | .exe installato | Browser → URL |
| **Frontend gira su** | Chromium embeddato | Browser dell'utente |
| **Backend gira su** | Stesso PC (main process) | Server remoto (Express) |
| **Comunicazione** | IPC (Inter-Process Communication) | HTTP (fetch / axios) |
| **Database** | SQLite locale / JSON file | MariaDB (server, via rete) |
| **Aggiornamento** | electron-updater → nuovo .exe | git pull + docker rebuild |
| **Multi-utente** | Impossibile (1 PC = 1 utente) | Naturale (N browser → 1 server) |
| **Installazione** | Installer su ogni PC | Nulla — basta il browser |
| **DevTools** | Ctrl+Shift+I nella finestra Electron | F12 nel browser (identico) |
| **Hot reload** | Vite HMR (identico) | Vite HMR (identico) |
| **React** | Identico | Identico |
| **TypeScript** | Identico | Identico |
| **npm** | Identico | Identico |

---

## 8. File Chiave da Conoscere

| File | Cosa fa | Equivalente Electron |
|------|---------|---------------------|
| `server/src/server.ts` | Entry point backend, avvia Express | `src/main/index.ts` |
| `server/src/routes/*.ts` | Endpoint API (GET, POST, ecc.) | `ipcMain.handle(...)` |
| `client/src/App.tsx` | Entry point React | Identico |
| `client/vite.config.ts` | Config Vite + proxy API | Simile, senza electron plugin |
| `docker-compose.yml` | Definisce i container (app + db) | Non esisteva |
| `Dockerfile` | Come costruire l'immagine dell'app | `electron-builder.yml` |
| `.env` | Variabili d'ambiente (password DB, ecc.) | `.env` (identico) |

---

## 9. Indirizzi Utili

| Cosa | URL / Indirizzo |
|------|----------------|
| **Demo pubblica** | http://demo.bilsoft.it |
| **Demo via IP** | http://178.104.16.85 |
| **Server SSH** | `ssh root@178.104.16.85` |
| **Hetzner Console** | https://console.hetzner.cloud |
| **Pannello Aruba DNS** | https://admin.aruba.it |
| **Dev locale frontend** | http://localhost:5173 |
| **Dev locale backend** | http://localhost:3000 |
| **Dev locale DB** | localhost:3306 |

---

> **Regola d'oro:** se una cosa funziona in locale su localhost:5173,
> funzionerà identica su demo.bilsoft.it dopo il deploy.
> L'unica differenza è dove gira il server. Il codice è lo stesso.

---

## 10. Stack Tecnologico (versioni attuali)

| Componente | Versione | Note |
|---|---|---|
| **Node.js** | 22 LTS | Runtime + Docker base image |
| **React** | 19 | Frontend UI |
| **TypeScript** | ~5.7 | Ovunque (client + server + shared) |
| **Vite** | 6 | Build tool + dev server client |
| **Tailwind CSS** | v4 | Styling (plugin `@tailwindcss/vite`, niente config file) |
| **shadcn/ui** | new-york style | Componenti UI (Radix UI sotto) |
| **Express** | 4 | Backend API |
| **MariaDB** | 10.11 | Database (Docker) |
| **ECharts** | 5.5 | Grafici statistici |
| **i18next** | — | Internazionalizzazione EN/RO |
| **tsx** | — | Dev runner per TypeScript server (al posto di nodemon) |

### Note importanti sullo stack

- **ESM**: Il server usa `"type": "module"`. `__dirname` non esiste in ESM.
  Usare sempre: `fileURLToPath(import.meta.url)` + `path.dirname()`
- **Tailwind v4**: Niente `tailwind.config.ts`. La config è nel CSS via `@theme` e il plugin Vite `@tailwindcss/vite`.
- **shadcn/ui**: Configurato in `components.json`. I componenti si installano con `npx shadcn@latest add <nome>`.
