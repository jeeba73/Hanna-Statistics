import dotenv from 'dotenv';
import { resolve } from 'path';

// Load .env from project root BEFORE any other imports
dotenv.config({ path: resolve(process.cwd(), '../.env') });
dotenv.config(); // fallback for .env in CWD

async function main() {
  const { default: app } = await import('./app.js');

  const PORT = parseInt(process.env.SERVER_PORT || '3000', 10);

  app.listen(PORT, () => {
    console.log(`[server] Hanna Statistics API running on http://localhost:${PORT}`);
    console.log(`[server] Environment: ${process.env.NODE_ENV || 'development'}`);
  });
}

main();
