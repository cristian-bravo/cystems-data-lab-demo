import type { APIRoute } from 'astro';
import { getSalesData } from '../../../lib/data';
import { json, options } from '../../../lib/http';

export const prerender = false;
export const OPTIONS: APIRoute = () => options();

export const GET: APIRoute = async () => {
  const result = await getSalesData();
  return json({
    ok: true,
    service: 'CYSTEMS Data Lab API',
    version: '1.0.0',
    source: result.source,
    records: result.data.length,
    updated_at: result.updatedAt,
    warning: result.warning || null,
  });
};
