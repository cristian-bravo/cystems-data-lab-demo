import { parseCsv, asNumber } from './csv';
import { generateDemoSales } from './demoData';
import type { DataResult, SaleRecord } from './types';

let cache: DataResult | null = null;
let cacheExpiresAt = 0;

const text = (row: Record<string, string>, ...keys: string[]) => {
  for (const key of keys) {
    const value = row[key];
    if (value !== undefined && value !== '') return value.trim();
  }
  return '';
};

function normalizeDate(value: string) {
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  const match = trimmed.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (match) return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
  const date = new Date(trimmed);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString().slice(0, 10);
}

function normalizeRow(row: Record<string, string>, index: number): SaleRecord | null {
  const fecha = normalizeDate(text(row, 'fecha', 'date'));
  const ventaNeta = asNumber(row.venta_neta ?? row.ventas ?? row.importe);
  const costoTotal = asNumber(row.costo_total ?? row.costo ?? row.costos);
  if (!fecha || !ventaNeta) return null;
  const date = new Date(`${fecha}T00:00:00Z`);
  const margin = asNumber(row.margen) || ventaNeta - costoTotal;
  const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

  return {
    id: text(row, 'id', 'id_venta') || `G${String(index + 1).padStart(4, '0')}`,
    fecha,
    anio: asNumber(row.anio) || date.getUTCFullYear(),
    mes: text(row, 'mes') || monthNames[date.getUTCMonth()],
    documento: text(row, 'documento', 'factura', 'comprobante'),
    cliente: text(row, 'cliente', 'nombre_cliente'),
    segmento: text(row, 'segmento'),
    producto: text(row, 'producto', 'servicio'),
    categoria: text(row, 'categoria'),
    region: text(row, 'region', 'ciudad'),
    canal: text(row, 'canal'),
    cantidad: asNumber(row.cantidad) || 1,
    precio_unitario: asNumber(row.precio_unitario ?? row.precio),
    descuento_pct: asNumber(row.descuento_pct ?? row.descuento),
    venta_neta: ventaNeta,
    costo_total: costoTotal,
    margen: margin,
    margen_pct: asNumber(row.margen_pct) || (ventaNeta ? margin / ventaNeta : 0),
    estado_cobro: text(row, 'estado_cobro', 'estado') || 'Pendiente',
    dias_cobro: asNumber(row.dias_cobro),
  };
}

async function fetchSheet(url: string): Promise<SaleRecord[]> {
  const response = await fetch(url, {
    headers: { accept: 'text/csv,text/plain;q=0.9,*/*;q=0.5' },
    signal: AbortSignal.timeout(8000),
  });
  if (!response.ok) throw new Error(`Google Sheets respondió ${response.status}`);
  const rows = parseCsv(await response.text());
  const normalized = rows.map(normalizeRow).filter((row): row is SaleRecord => Boolean(row));
  if (!normalized.length) throw new Error('La hoja no contiene filas válidas');
  return normalized;
}

export async function getSalesData(force = false): Promise<DataResult> {
  const now = Date.now();
  if (!force && cache && now < cacheExpiresAt) return cache;

  const url = import.meta.env.GOOGLE_SHEET_CSV_URL || process.env.GOOGLE_SHEET_CSV_URL;
  const ttl = Number(import.meta.env.DATA_CACHE_SECONDS || process.env.DATA_CACHE_SECONDS || 15);

  if (url) {
    try {
      cache = { data: await fetchSheet(url), source: 'google-sheets', updatedAt: new Date().toISOString() };
      cacheExpiresAt = now + Math.max(5, ttl) * 1000;
      return cache;
    } catch (error) {
      cache = {
        data: generateDemoSales(),
        source: 'demo',
        updatedAt: new Date().toISOString(),
        warning: error instanceof Error ? error.message : 'No se pudo leer Google Sheets',
      };
      cacheExpiresAt = now + 10_000;
      return cache;
    }
  }

  cache = {
    data: generateDemoSales(),
    source: 'demo',
    updatedAt: new Date().toISOString(),
    warning: 'GOOGLE_SHEET_CSV_URL no está configurada',
  };
  cacheExpiresAt = now + 30_000;
  return cache;
}

export function filterSales(data: SaleRecord[], params: URLSearchParams) {
  const desde = params.get('desde') || '';
  const hasta = params.get('hasta') || '';
  const region = (params.get('region') || '').toLowerCase();
  const categoria = (params.get('categoria') || '').toLowerCase();
  const canal = (params.get('canal') || '').toLowerCase();
  const cliente = (params.get('cliente') || '').toLowerCase();
  const anio = Number(params.get('anio') || 0);

  return data.filter((row) => {
    if (desde && row.fecha < desde) return false;
    if (hasta && row.fecha > hasta) return false;
    if (anio && row.anio !== anio) return false;
    if (region && row.region.toLowerCase() !== region) return false;
    if (categoria && row.categoria.toLowerCase() !== categoria) return false;
    if (canal && row.canal.toLowerCase() !== canal) return false;
    if (cliente && !row.cliente.toLowerCase().includes(cliente)) return false;
    return true;
  });
}

export function summarize(data: SaleRecord[]) {
  const sales = data.reduce((sum, row) => sum + row.venta_neta, 0);
  const costs = data.reduce((sum, row) => sum + row.costo_total, 0);
  const margin = sales - costs;
  const byMonth = new Map<string, number>();
  const byProduct = new Map<string, number>();
  const byRegion = new Map<string, number>();

  data.forEach((row) => {
    const month = row.fecha.slice(0, 7);
    byMonth.set(month, (byMonth.get(month) || 0) + row.venta_neta);
    byProduct.set(row.producto, (byProduct.get(row.producto) || 0) + row.venta_neta);
    byRegion.set(row.region, (byRegion.get(row.region) || 0) + row.venta_neta);
  });

  return {
    ventas: Math.round(sales * 100) / 100,
    costos: Math.round(costs * 100) / 100,
    margen: Math.round(margin * 100) / 100,
    margen_pct: sales ? Math.round((margin / sales) * 10000) / 10000 : 0,
    transacciones: data.length,
    clientes: new Set(data.map((row) => row.cliente)).size,
    meses: [...byMonth.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([mes, ventas]) => ({ mes, etiqueta: mes.slice(5), ventas: Math.round(ventas * 100) / 100 })),
    productos: [...byProduct.entries()].sort((a, b) => b[1] - a[1]).map(([nombre, ventas]) => ({ nombre, ventas: Math.round(ventas * 100) / 100 })),
    regiones: [...byRegion.entries()].sort((a, b) => b[1] - a[1]).map(([nombre, ventas]) => ({ nombre, ventas: Math.round(ventas * 100) / 100 })),
  };
}
