# Setup Sviluppo Remoto (Italia) e Deploy (Romania)

> Guida operativa per lo sviluppo di Hanna Statistics dall'Italia
> con deploy sullo stabilimento Hanna Instruments in Romania.

---

## 1. Cosa chiedere al team IT rumeno

### 1.1 Accesso VPN (PRIORITA' MASSIMA)

```
RICHIESTA AL TEAM IT:

"Ho bisogno di una connessione VPN alla vostra rete locale per poter
sviluppare e deployare l'applicazione Hanna Statistics.
Mi servono le seguenti informazioni:"

[ ] Tipo di VPN disponibile (WireGuard, OpenVPN, IPSec, Synology VPN?)
[ ] File di configurazione VPN (.ovpn, .conf, o profilo WireGuard)
[ ] Credenziali VPN (username + password o chiave privata)
[ ] Indirizzo IP del gateway VPN (IP pubblico del router/firewall rumeno)
[ ] Porta VPN (es. 1194 per OpenVPN, 51820 per WireGuard)
[ ] Subnet della rete interna (es. 192.168.1.0/24)
[ ] Politiche di accesso: orari, durata sessione, MFA richiesto?
```

> **NOTA**: Senza VPN non puoi accedere a nulla. Questa e' la prima cosa
> da ottenere. Se non hanno una VPN gia' configurata, la soluzione piu'
> veloce e' **WireGuard** (setup in 15 minuti) o il **VPN Server del NAS
> Synology** (gia' integrato).

### 1.2 Accesso al Server/NAS

```
RICHIESTA AL TEAM IT:

"Una volta connesso in VPN, ho bisogno di accesso al server dove
gira Hanna Core per deployare la nuova applicazione:"

[ ] IP del server/NAS (es. 192.168.1.36 — gia' noto dalle immagini)
[ ] Accesso SSH abilitato? (porta 22 o custom)
[ ] Credenziali SSH (username + password o chiave pubblica da autorizzare)
[ ] Sistema operativo del NAS (Synology DSM? QNAP QTS? Linux puro?)
[ ] Versione del sistema operativo
[ ] Docker e' installato? Docker Compose disponibile?
[ ] Quanta RAM e disco libero ci sono?
[ ] Ho permessi di creare container Docker?
[ ] Ho permessi di creare nuovi database MariaDB?
```

### 1.3 Database

```
RICHIESTA AL TEAM IT:

"Ho bisogno di creare un database MariaDB SEPARATO per la nuova
applicazione (non tocchero' il database di Hanna Core):"

[ ] MariaDB/MySQL e' gia' installato sul server? Quale versione?
[ ] Posso creare un nuovo database chiamato 'hanna_statistics'?
[ ] Credenziali di accesso root o un utente con privilegi CREATE DATABASE
[ ] Porta MariaDB (default 3306 o custom?)
[ ] Il DB e' accessibile solo da localhost o anche dalla LAN?
```

### 1.4 Rete e Porte

```
RICHIESTA AL TEAM IT:

"L'applicazione avra' bisogno di una porta di rete per essere
accessibile dai browser sulla vostra LAN:"

[ ] Quale porta posso usare? (suggerisco 83 o 8083)
[ ] C'e' un firewall interno? Devo chiedere l'apertura della porta?
[ ] C'e' un reverse proxy Nginx/Apache che posso configurare?
[ ] L'URL finale sara': https://192.168.1.36:<PORTA>
    oppure preferite: https://192.168.1.36:82/statistics (sottopercorso)?
[ ] Volete HTTPS? Se si, avete un certificato o ne genero uno self-signed?
```

### 1.5 Accesso a Hanna Core (opzionale ma utile)

```
RICHIESTA AL TEAM IT:

"Per capire meglio i dati e le strutture, sarebbe utile avere:"

[ ] Un account di sola lettura su Hanna Core (per vedere i dati reali)
[ ] Documentazione delle tabelle/API di Hanna Core (se esiste)
[ ] Export di dati di esempio (CSV/Excel) da usare per lo sviluppo
[ ] Schema del database di Hanna Core (se disponibile)
```

### 1.6 Backup e Recovery

```
RICHIESTA AL TEAM IT:

"Per la sicurezza dei dati:"

[ ] Esiste un sistema di backup automatico sul NAS?
[ ] Posso schedulare un backup del mio database?
[ ] C'e' una policy di retention dei backup?
[ ] In caso di problemi, chi contatto? (contatto IT di emergenza)
```

---

## 2. Email template da inviare al team IT

Ecco un'email pronta da inviare (in inglese, per il team rumeno):

```
Subject: Access Request - Hanna Statistics Application Development

Dear IT Team,

I am developing a new web application called "Hanna Statistics" that will
run alongside Hanna Core on your server. The application will have its own
independent database and will NOT modify or interact with Hanna Core's data.

To develop and deploy this application remotely from Italy, I need the
following access:

1. VPN ACCESS
   - VPN configuration file and credentials to connect to your LAN
   - Preferred: WireGuard or OpenVPN
   - If not available, I can help set up WireGuard (quick setup)

2. SERVER ACCESS (192.168.1.36 or the appropriate server)
   - SSH access (username + key or password)
   - Docker permissions (to create and manage containers)
   - Server OS version and available resources (RAM, disk)

3. DATABASE
   - Permission to create a new MariaDB database: "hanna_statistics"
   - Database credentials (or root access to create a dedicated user)
   - MariaDB version currently installed

4. NETWORK
   - An available port for the web application (suggested: 83 or 8083)
   - Firewall rules if needed
   - SSL certificate (or permission to use self-signed)

5. SAMPLE DATA (optional but helpful)
   - Export of sample data from Hanna Core (CSV/Excel format)
   - This helps me build and test the statistics features

The application will be accessible at:
https://192.168.1.36:<PORT>/

Please let me know the best way to proceed and if you need any
additional information from my side.

Thank you,
[Il tuo nome]
```

---

## 3. Come testare in locale (dall'Italia)

### 3.1 Setup sviluppo locale (SENZA VPN)

Puoi sviluppare e testare **tutto in locale** sul tuo PC senza bisogno
della VPN. Usi la VPN solo per il deploy finale.

#### Prerequisiti sul tuo PC

```
Installa:
[ ] Node.js 20 LTS          → https://nodejs.org
[ ] Docker Desktop           → https://docker.com/products/docker-desktop
[ ] Git                      → (probabilmente gia' installato)
[ ] VS Code                  → (probabilmente gia' installato)
[ ] Un client MariaDB        → HeidiSQL (Windows) o DBeaver (cross-platform)
```

### 3.2 Docker Compose locale (replica l'ambiente rumeno)

Crea un ambiente locale identico a quello di produzione:

```yaml
# docker/docker-compose.dev.yml
# Questo file gira sul TUO PC in Italia

version: '3.8'

services:
  # ==========================================
  # MariaDB locale (replica del DB rumeno)
  # ==========================================
  db:
    image: mariadb:10.11
    container_name: hanna-stats-db
    ports:
      - "3306:3306"          # Accessibile da localhost:3306
    environment:
      MYSQL_ROOT_PASSWORD: root_dev_password
      MYSQL_DATABASE: hanna_statistics
      MYSQL_USER: hanna_stats
      MYSQL_PASSWORD: dev_password
    volumes:
      - db_data:/var/lib/mysql
      - ../server/database/migrations:/docker-entrypoint-initdb.d
    healthcheck:
      test: ["CMD", "healthcheck.sh", "--connect", "--innodb_initialized"]
      interval: 10s
      timeout: 5s
      retries: 3

  # ==========================================
  # phpMyAdmin (opzionale, per vedere il DB)
  # ==========================================
  phpmyadmin:
    image: phpmyadmin/phpmyadmin
    container_name: hanna-stats-phpmyadmin
    ports:
      - "8080:80"            # Accessibile da localhost:8080
    environment:
      PMA_HOST: db
      PMA_USER: root
      PMA_PASSWORD: root_dev_password
    depends_on:
      - db

volumes:
  db_data:
```

### 3.3 Workflow di sviluppo quotidiano

```
PASSO 1: Avvia il database locale (una volta sola)
──────────────────────────────────────────────────
> cd hanna-statistics
> docker compose -f docker/docker-compose.dev.yml up -d

Risultato:
  - MariaDB su localhost:3306
  - phpMyAdmin su http://localhost:8080 (opzionale)


PASSO 2: Avvia il dev server (ogni giorno)
──────────────────────────────────────────────────
> npm run dev

Risultato:
  - Frontend React su http://localhost:5173 (Vite HMR, hot reload)
  - Backend Express su http://localhost:3000 (nodemon, auto-restart)
  - Proxy Vite: le chiamate /api/* vanno automaticamente a :3000


PASSO 3: Sviluppa normalmente
──────────────────────────────────────────────────
  - Modifica un file React → il browser si aggiorna in <100ms
  - Modifica un file server → Express si riavvia in ~1s
  - Modifica una query SQL → testa su phpMyAdmin o dal codice
  - Tutto funziona identico a come funzionera' in Romania


PASSO 4: Quando hai finito, spegni il DB
──────────────────────────────────────────────────
> docker compose -f docker/docker-compose.dev.yml down

(I dati del DB persistono nel volume Docker. Al prossimo `up` li ritrovi.)
```

### 3.4 File .env per sviluppo locale

```env
# .env.development (usato in locale)

NODE_ENV=development
PORT=3000
HOST=localhost

# Database locale (Docker)
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=hanna_stats
DB_PASSWORD=dev_password
DB_NAME=hanna_statistics

# Auth
JWT_SECRET=dev_secret_change_in_production_abc123
JWT_EXPIRATION=24h
BCRYPT_ROUNDS=10

# App
DEFAULT_LANGUAGE=it
COMPANY_NAME=Hanna Instruments (DEV)
```

### 3.5 File .env per produzione (Romania)

```env
# .env.production (usato sul server rumeno)

NODE_ENV=production
PORT=3000
HOST=0.0.0.0

# Database produzione
DB_HOST=127.0.0.1            # o IP del container MariaDB
DB_PORT=3306
DB_USER=hanna_stats
DB_PASSWORD=<PASSWORD_FORTE_GENERATA>
DB_NAME=hanna_statistics

# Auth
JWT_SECRET=<STRINGA_RANDOM_64_CARATTERI>
JWT_EXPIRATION=8h
BCRYPT_ROUNDS=12

# App
DEFAULT_LANGUAGE=en
COMPANY_NAME=Hanna Instruments
```

### 3.6 Vite config con proxy (per sviluppo locale)

```typescript
// vite.config.ts
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

export default defineConfig({
    plugins: [react(), tailwindcss()],
    resolve: {
        alias: {
            '@': '/src',
            '@shared': '/shared',
        },
    },
    server: {
        port: 5173,
        // Proxy: le chiamate /api/* vanno al backend Express
        proxy: {
            '/api': {
                target: 'http://localhost:3000',
                changeOrigin: true,
            },
        },
    },
});
```

### 3.7 Dati di test (seed locale)

Per testare le statistiche servono dati realistici.
Crea uno script di seed che popola il DB locale:

```typescript
// scripts/seed-dev.ts
// Genera dati finti ma realistici per lo sviluppo

import { pool } from '../server/database/connection';

async function seed() {
    console.log('Seeding development database...');

    // Utente admin di default
    await pool.execute(`
        INSERT IGNORE INTO users (username, password_hash, full_name, role)
        VALUES ('admin', '$2a$10$...hash...', 'Admin Dev', 'admin')
    `);

    // Genera 1000 record di produzione (ultimi 6 mesi)
    for (let i = 0; i < 1000; i++) {
        const date = randomDate(6); // ultimi 6 mesi
        const products = ['HI97700B','HI97749B','HI97716B','HI84532-70U','HI701-25'];
        const lines = ['L56','L57','L58','L84'];
        const operators = ['Marinela','Florina','Krisztina','Erika','Roland'];

        await pool.execute(`
            INSERT INTO production_data
            (product_code, line, lot, quantity_planned, quantity_produced,
             quantity_rejected, start_date, operator, shift)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `, [
            randomPick(products),
            randomPick(lines),
            1000 + i,
            randomInt(100, 500),
            randomInt(90, 500),
            randomInt(0, 10),
            date,
            randomPick(operators),
            randomPick(['morning','afternoon','night'])
        ]);
    }

    // Genera 500 record QC
    for (let i = 0; i < 500; i++) {
        await pool.execute(`
            INSERT INTO quality_data
            (register_nr, record_date, product_code, lot, client_code,
             qc_type, result, sampling_user, sampling_date)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `, [
            `QCC${randomString(6)}`,
            randomDate(6),
            randomPick(['HI97700B','ORGANIC400','HI784A-0','HI7006L']),
            1000 + i,
            randomPick(['L56','L57','L84']),
            randomPick(['first','during','final']),
            Math.random() > 0.05 ? 'pass' : 'fail', // 95% pass rate
            randomPick(['Ana','Krisztina','Erika','Roland']),
            randomDate(6)
        ]);
    }

    // Genera 800 record packing
    // ... logica simile ...

    console.log('Seed completed!');
    process.exit(0);
}

// Helper functions
function randomDate(monthsBack: number): string { /* ... */ }
function randomPick<T>(arr: T[]): T { return arr[Math.floor(Math.random() * arr.length)]; }
function randomInt(min: number, max: number): number { return Math.floor(Math.random() * (max - min + 1)) + min; }
function randomString(len: number): string { return Math.random().toString(36).substring(2, 2 + len).toUpperCase(); }

seed();
```

Eseguilo con:
```bash
npm run db:seed
# oppure: npx tsx scripts/seed-dev.ts
```

---

## 4. Workflow di Deploy (Italia → Romania)

### 4.1 Flusso completo

```
┌──────────────────────────────────────────────────────────────┐
│                        TUO PC (Italia)                        │
│                                                                │
│  1. Sviluppi e testi in locale                                 │
│     └── localhost:5173 (frontend)                              │
│     └── localhost:3000 (backend)                               │
│     └── localhost:3306 (DB locale Docker)                      │
│                                                                │
│  2. Commit e push su Git                                       │
│     └── git push origin main                                   │
│                                                                │
│  3. Connetti VPN → 192.168.1.x                                │
│     └── Ora sei "dentro" la rete rumena                        │
│                                                                │
│  4. SSH al server                                              │
│     └── ssh user@192.168.1.36                                  │
│                                                                │
│  5. Deploy                                                     │
│     └── git pull && docker compose up -d --build               │
│                                                                │
│  6. Verifica                                                   │
│     └── Apri https://192.168.1.36:83 nel browser              │
│         (funziona perche' sei in VPN)                          │
│                                                                │
│  7. Disconnetti VPN                                            │
│     └── L'app continua a girare in Romania                     │
└──────────────────────────────────────────────────────────────┘
```

### 4.2 Primo deploy (setup iniziale, una volta sola)

```bash
# =========================================
# Dal tuo PC, connesso in VPN
# =========================================

# 1. Connetti via SSH al server rumeno
ssh user@192.168.1.36

# 2. Clona il repository
git clone <URL_REPO> /opt/hanna-statistics
cd /opt/hanna-statistics

# 3. Crea il file .env di produzione
cp .env.example .env
nano .env   # inserisci credenziali di produzione

# 4. Avvia tutto con Docker Compose
docker compose -f docker/docker-compose.yml up -d --build

# 5. Esegui le migrazioni del database
docker compose exec app npm run db:migrate

# 6. Crea l'utente admin iniziale
docker compose exec app npm run db:seed

# 7. Verifica che funzioni
curl -k https://localhost:83
# oppure apri nel browser: https://192.168.1.36:83

# 8. Verifica i log
docker compose logs -f app
```

### 4.3 Deploy successivi (aggiornamenti)

```bash
# =========================================
# Dal tuo PC, connesso in VPN
# =========================================

ssh user@192.168.1.36
cd /opt/hanna-statistics

# Aggiorna il codice
git pull origin main

# Rebuild e restart (zero-downtime con --build)
docker compose up -d --build

# Se ci sono nuove migrazioni DB
docker compose exec app npm run db:migrate

# Verifica
docker compose logs --tail=50 app
```

### 4.4 Script di deploy automatico (opzionale)

```bash
#!/bin/bash
# scripts/deploy.sh
# Esegui dal tuo PC: ./scripts/deploy.sh

SERVER="user@192.168.1.36"
APP_DIR="/opt/hanna-statistics"

echo "=== Deploying Hanna Statistics ==="

# Verifica connessione VPN
ping -c 1 192.168.1.36 > /dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "ERRORE: Non sei connesso alla VPN!"
    echo "Connettiti prima alla VPN rumena."
    exit 1
fi

echo "1/4 Pulling latest code..."
ssh $SERVER "cd $APP_DIR && git pull origin main"

echo "2/4 Building Docker image..."
ssh $SERVER "cd $APP_DIR && docker compose up -d --build"

echo "3/4 Running migrations..."
ssh $SERVER "cd $APP_DIR && docker compose exec -T app npm run db:migrate"

echo "4/4 Checking health..."
sleep 5
ssh $SERVER "curl -sk https://localhost:83/api/health"

echo ""
echo "=== Deploy completato! ==="
echo "URL: https://192.168.1.36:83"
```

---

## 5. Scenari e soluzioni per problemi comuni

### 5.1 "Non riesco a connettermi in VPN"

```
Possibili cause:
├── Porta VPN bloccata dal tuo ISP → Prova porta 443 (HTTPS)
├── Firewall aziendale in Romania → Chiedi al team IT di aprire la porta
├── Credenziali errate → Verifica username/password
├── IP pubblico rumeno cambiato → Chiedi il nuovo IP (o usate Dynamic DNS)
└── VPN non configurata → Suggerisci WireGuard sul NAS Synology
```

### 5.2 "La VPN e' lenta, il dev e' impossibile"

```
Soluzione: NON SVILUPPARE IN VPN!
├── Sviluppa TUTTO in locale (Docker + Node.js sul tuo PC)
├── Usa la VPN SOLO per il deploy (5 minuti)
├── Testa in locale, deploya in remoto
└── Il 99% del tempo lavori senza VPN
```

### 5.3 "Non hanno Docker sul NAS"

```
Alternativa senza Docker:
├── Installa Node.js 20 direttamente sul NAS
│   ├── Synology: Pacchetto "Node.js" dal Package Center
│   └── QNAP: App Center → Node.js
├── Installa PM2 globalmente: npm install -g pm2
├── Copia i file compilati via SCP
│   └── scp -r dist/ dist-server/ user@192.168.1.36:/opt/hanna-statistics/
├── Avvia con PM2: pm2 start dist-server/server.js --name hanna-stats
└── MariaDB: Pacchetto "MariaDB 10" dal Package Center del NAS
```

### 5.4 "Non ho accesso SSH"

```
Alternative:
├── Synology DSM → File Station per upload file + Task Scheduler per comandi
├── SFTP → Upload file via FileZilla/WinSCP
├── Synology Web Station → Deploy come PHP/Node app da interfaccia web
├── Portainer → Se installato, gestisci Docker via web UI
└── TeamViewer/AnyDesk → Accesso remoto al desktop del server (ultimo resort)
```

### 5.5 "Come faccio a vedere i log in tempo reale?"

```bash
# Via SSH (connesso in VPN)
ssh user@192.168.1.36

# Log del container Docker
docker compose -f /opt/hanna-statistics/docker-compose.yml logs -f app

# Oppure con PM2
pm2 logs hanna-stats

# Oppure salva log su file e scaricali
scp user@192.168.1.36:/opt/hanna-statistics/logs/app.log ./logs/
```

---

## 6. Sicurezza: Checklist

```
[ ] VPN con crittografia forte (WireGuard: ChaCha20, OpenVPN: AES-256)
[ ] Credenziali SSH con chiave pubblica (non solo password)
[ ] Password DB di produzione forte (generata, 32+ caratteri)
[ ] JWT_SECRET di produzione forte (generato, 64+ caratteri)
[ ] File .env MAI committato su Git (verificare .gitignore)
[ ] HTTPS abilitato (anche self-signed, meglio di HTTP)
[ ] Porta del DB (3306) NON esposta su internet (solo localhost o LAN)
[ ] Backup DB automatico configurato
[ ] Account VPN personale (non condiviso con altri)
[ ] Accesso SSH limitato al tuo utente (non root se possibile)
```

---

## 7. Riepilogo: cosa serve prima di iniziare a codare

### Dal team IT rumeno (BLOCCANTE):

```
PRIORITA' 1 (senza queste non puoi deployare):
  [1] Accesso VPN (configurazione + credenziali)
  [2] Accesso SSH al server (IP + porta + credenziali)
  [3] Docker installato (o permesso di installarlo)
  [4] Permesso di creare un database MariaDB

PRIORITA' 2 (utili ma non bloccanti per iniziare):
  [5] Porta rete assegnata (es. 83)
  [6] Dati di esempio da Hanna Core (CSV export)
  [7] Account di sola lettura su Hanna Core

PRIORITA' 3 (nice to have):
  [8] Schema DB di Hanna Core
  [9] Documentazione API Hanna Core
  [10] Contatto diretto con un tecnico IT per emergenze
```

### Da parte tua (puoi fare SUBITO, senza aspettare la Romania):

```
PUOI INIZIARE OGGI:
  [x] Installa Docker Desktop sul tuo PC
  [x] Installa Node.js 20 LTS
  [x] Crea il progetto (npm create vite, shadcn init, ecc.)
  [x] Avvia MariaDB locale con Docker
  [x] Sviluppa frontend (Dashboard, grafici, tabelle)
  [x] Sviluppa backend (Express, API, services)
  [x] Testa tutto in locale con dati finti (seed)
  [x] Il giorno che arriva la VPN → deploy in 30 minuti
```

---

> **Messaggio chiave**: Puoi sviluppare il 100% dell'applicazione in locale
> dall'Italia senza mai toccare il server rumeno. La VPN ti serve SOLO per
> il deploy finale e per i test con dati reali. Non aspettare la VPN per
> iniziare a codare!
