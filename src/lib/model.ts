import type { ProductRecord, SaleRecord } from './types';

const round = (value: number, digits = 2) => {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
};

export function productIdFromName(name: string) {
  const slug = name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
  return `PROD-${slug || 'SIN-NOMBRE'}`;
}

export function buildProducts(sales: SaleRecord[]): ProductRecord[] {
  const groups = new Map<string, { producto: string; categoria: string; precio: number; costo: number; count: number }>();

  sales.forEach((sale) => {
    const current = groups.get(sale.producto_id) || {
      producto: sale.producto,
      categoria: sale.categoria,
      precio: 0,
      costo: 0,
      count: 0,
    };
    current.precio += sale.precio_unitario;
    current.costo += sale.cantidad ? sale.costo_total / sale.cantidad : 0;
    current.count += 1;
    groups.set(sale.producto_id, current);
  });

  return [...groups.entries()]
    .map(([producto_id, value]) => ({
      producto_id,
      producto: value.producto,
      categoria: value.categoria,
      precio_referencia: round(value.precio / value.count),
      costo_unitario_referencia: round(value.costo / value.count),
    }))
    .sort((a, b) => a.producto.localeCompare(b.producto, 'es'));
}
