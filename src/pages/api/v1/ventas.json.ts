import type { APIRoute } from 'astro';
import { filterSales, getSalesData } from '../../../lib/data';
import { json, options } from '../../../lib/http';

export const prerender = false;
export const OPTIONS: APIRoute = () => options();

export const GET: APIRoute = async ({ url }) => {
  const result = await getSalesData(url.searchParams.get('refresh') === '1');
  const filtered = filterSales(result.data, url.searchParams);
  const requestedLimit = Number(url.searchParams.get('limit') || filtered.length);
  const limit = Math.min(Math.max(requestedLimit, 1), 5000);
  const offset = Math.max(Number(url.searchParams.get('offset') || 0), 0);
  const rows = filtered.slice(offset, offset + limit);

  return json({
    ok: true,
    meta: {
      source: result.source,
      generated_at: result.updatedAt,
      warning: result.warning || null,
    },
    total: filtered.length,
    count: rows.length,
    offset,
    data: rows,
  });
};
