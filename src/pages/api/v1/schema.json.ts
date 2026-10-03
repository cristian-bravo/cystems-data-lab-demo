import type { APIRoute } from 'astro';
import { json, options } from '../../../lib/http';

export const prerender = false;
export const OPTIONS: APIRoute = () => options();

export const GET: APIRoute = () => json({
  ok: true,
  resource: 'ventas',
  primary_key: 'id',
  foreign_keys: [
    { field: 'producto_id', references: 'productos.producto_id', relationship: 'muchos a uno' },
  ],
  fields: [
    ['id', 'texto', 'Identificador único de la venta'],
    ['producto_id', 'texto', 'Clave foránea que identifica el producto'],
    ['fecha', 'fecha ISO', 'Fecha de emisión'],
    ['anio', 'entero', 'Año calendario'],
    ['mes', 'texto', 'Nombre del mes'],
    ['documento', 'texto', 'Número de factura'],
    ['cliente', 'texto', 'Razón social o cliente'],
    ['segmento', 'texto', 'Segmento comercial'],
    ['producto', 'texto', 'Producto o servicio'],
    ['categoria', 'texto', 'Familia del producto'],
    ['region', 'texto', 'Ciudad o región'],
    ['canal', 'texto', 'Canal comercial'],
    ['cantidad', 'decimal', 'Unidades vendidas'],
    ['precio_unitario', 'moneda', 'Precio antes de descuento'],
    ['descuento_pct', 'decimal', 'Descuento expresado entre 0 y 1'],
    ['venta_neta', 'moneda', 'Ingreso después del descuento'],
    ['costo_total', 'moneda', 'Costo asociado'],
    ['margen', 'moneda', 'Venta neta menos costo total'],
    ['margen_pct', 'decimal', 'Margen dividido para venta neta'],
    ['estado_cobro', 'texto', 'Cobrado, Pendiente o Vencido'],
    ['dias_cobro', 'entero', 'Días acordados o transcurridos'],
  ].map(([name, type, description]) => ({ name, type, description })),
});
