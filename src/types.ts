export type Status = 'ok' | 'critical' | 'info';

export interface MetricHistory {
  date: string;
  value: number;
}

export interface Metric {
  id: string;
  spId?: number; // SharePoint Item ID
  title: string;
  value: string | number;
  status: Status;
  lastUpdate: string;
  objective?: string;
  rules?: string[];
  query?: string;
  history?: MetricHistory[];
  refreshInterval?: number;
  resultTable?: string; // Name of the SharePoint list storing results
  details?: any[];
}

export type AccessLevel = 'admin' | 'viewer' | 'editor';

export interface User {
  id: string;
  spId?: number;
  name: string;
  email: string;
  role: AccessLevel;
  status: string;
  lastActive: string;
}

export interface Section {
  id: string;
  spId?: number; // SharePoint Item ID
  title: string;
  metrics: Metric[];
}

export interface ApiResponseItem {
  id: string;
  status: Status;
  value: string | number;
  [key: string]: any;
}
