import type { APIRoute } from 'astro';
import { filterSales, getSalesData, summarize } from '../../../lib/data';
import { json, options } from '../../../lib/http';

export const prerender = false;
export const OPTIONS: APIRoute = () => options();

export const GET: APIRoute = async ({ url }) => {
  const result = await getSalesData(url.searchParams.get('refresh') === '1');
  const filtered = filterSales(result.data, url.searchParams);
  return json({
    ok: true,
    meta: {
      source: result.source,
      generated_at: result.updatedAt,
      warning: result.warning || null,
    },
    filters: Object.fromEntries(url.searchParams.entries()),
    data: summarize(filtered),
  });
};
