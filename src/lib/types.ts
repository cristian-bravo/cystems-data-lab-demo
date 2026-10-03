export interface SaleRecord {
  id: string;
  fecha: string;
  anio: number;
  mes: string;
  documento: string;
  cliente: string;
  segmento: string;
  producto: string;
  categoria: string;
  region: string;
  canal: string;
  cantidad: number;
  precio_unitario: number;
  descuento_pct: number;
  venta_neta: number;
  costo_total: number;
  margen: number;
  margen_pct: number;
  estado_cobro: string;
  dias_cobro: number;
}

export interface DataResult {
  data: SaleRecord[];
  source: 'google-sheets' | 'demo';
  updatedAt: string;
  warning?: string;
}
