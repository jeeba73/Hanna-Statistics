import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import morgan from 'morgan';
import compression from 'compression';
import path from 'path';
import { healthRouter } from './routes/health.routes.js';

const app = express();

// Middleware
app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(compression());
app.use(morgan('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// API Routes
app.use('/api', healthRouter);

// Serve frontend static files in production
const publicDir = path.resolve(__dirname, '../public');
app.use(express.static(publicDir));

// SPA fallback: serve index.html for all non-API routes
app.get('*', (_req, res) => {
  res.sendFile(path.join(publicDir, 'index.html'));
});

export default app;
