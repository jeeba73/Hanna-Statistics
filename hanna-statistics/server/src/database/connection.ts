import mysql from 'mysql2/promise';

export const pool = mysql.createPool({
  host: process.env.DB_HOST || '127.0.0.1',
  port: parseInt(process.env.DB_PORT || '3306', 10),
  user: process.env.DB_USER || 'hanna_stats',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'hanna_statistics',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});
