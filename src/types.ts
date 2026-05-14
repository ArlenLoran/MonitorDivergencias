export type Status = 'ok' | 'error';

export interface MetricHistory {
  timestamp: string;
  value: number;
}

export interface Metric {
  id: string;
  title: string;
  value: number | string;
  status: Status;
  lastUpdate: string;
  objective?: string;
  rules?: string[];
  query?: string;
  history?: MetricHistory[];
  refreshInterval?: number;
  details?: {
    id: string;
    posicao: string;
    item: string;
    validade: string;
    lote: string;
    quantidade: number;
    motivo: string;
  }[];
}

export type AccessLevel = 'admin' | 'viewer' | 'editor';

export interface User {
  id: string;
  name: string;
  email: string;
  role: AccessLevel;
  lastLogin: string;
}

export interface Section {
  id: string;
  title: string;
  metrics: Metric[];
}
