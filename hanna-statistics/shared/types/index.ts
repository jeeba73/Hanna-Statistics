// Shared types between client and server

export interface User {
  id: number;
  username: string;
  fullName?: string;
  role: 'admin' | 'manager' | 'operator' | 'viewer';
  language: string;
}

export interface HannaCode {
  id: number;
  sfgCode: string;
  description?: string;
  parameterFormula?: string;
  recipe?: string;
  productionLine?: string;
  productType: 'REAGENT' | 'BUFFER' | 'OTHER';
}

export interface HealthResponse {
  status: string;
  timestamp: string;
  database: 'connected' | 'disconnected';
}
