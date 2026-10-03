import type { APIRoute } from 'astro';
import { getSalesData } from '../../../lib/data';
import { buildProducts } from '../../../lib/model';
import { json, options } from '../../../lib/http';

export const prerender = false;
export const OPTIONS: APIRoute = () => options();

export const GET: APIRoute = async ({ url }) => {
  const result = await getSalesData(url.searchParams.get('refresh') === '1');
  const products = buildProducts(result.data);

  return json({
    ok: true,
    meta: {
      source: result.source,
      generated_at: result.updatedAt,
      warning: result.warning || null,
      primary_key: 'producto_id',
      related_table: 'ventas',
      relationship: 'productos.producto_id (1) -> ventas.producto_id (*)',
    },
    total: products.length,
    data: products,
  });
};
