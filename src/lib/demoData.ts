import type { SaleRecord } from './types';

const clients = [
  ['Andes Market', 'Corporativo'],
  ['Nova Salud', 'Corporativo'],
  ['Tierra Verde', 'Pyme'],
  ['Pacífico Retail', 'Corporativo'],
  ['Mundo Office', 'Pyme'],
  ['Soluciones Delta', 'Pyme'],
  ['Grupo Horizonte', 'Corporativo'],
  ['Comercial Yasuní', 'Emprendedor'],
] as const;

const products = [
  ['ERP Contable', 'Software', 1480, 720],
  ['Dashboard BI', 'Analítica', 970, 390],
  ['API Empresarial', 'Integración', 820, 330],
  ['Soporte Pro', 'Servicios', 420, 165],
  ['Automatización', 'Servicios', 760, 280],
  ['Portal Web', 'Software', 1150, 510],
  ['Capacitación', 'Formación', 360, 105],
] as const;

const regions = ['Quito', 'Guayaquil', 'Cuenca', 'Manta'];
const channels = ['Venta directa', 'Referido', 'Web', 'Alianza'];
const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

function round(value: number, digits = 2) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

export function generateDemoSales(): SaleRecord[] {
  const rows: SaleRecord[] = [];
  let sequence = 1;

  for (const year of [2025, 2026]) {
    for (let month = 1; month <= 12; month += 1) {
      for (let order = 0; order < 15; order += 1) {
        const client = clients[(month * 3 + order + year) % clients.length];
        const product = products[(month + order * 2 + year) % products.length];
        const region = regions[(month + order) % regions.length];
        const channel = channels[(order + year) % channels.length];
        const quantity = 1 + ((month + order + year) % 4);
        const season = 1 + ((month - 6) * 0.008);
        const growth = year === 2026 ? 1.11 : 1;
        const price = round(product[2] * season * growth, 2);
        const cost = round(product[3] * quantity * (year === 2026 ? 1.045 : 1), 2);
        const discount = [0, 0.03, 0.05, 0.08][(month + order) % 4];
        const gross = round(price * quantity, 2);
        const net = round(gross * (1 - discount), 2);
        const margin = round(net - cost, 2);
        const day = 2 + ((order * 2 + month) % 25);
        const paymentDays = [0, 15, 30, 45][(order + month) % 4];

        rows.push({
          id: `V${String(sequence).padStart(4, '0')}`,
          fecha: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
          anio: year,
          mes: monthNames[month - 1],
          documento: `FAC-${year}-${String(sequence).padStart(5, '0')}`,
          cliente: client[0],
          segmento: client[1],
          producto: product[0],
          categoria: product[1],
          region,
          canal: channel,
          cantidad: quantity,
          precio_unitario: price,
          descuento_pct: discount,
          venta_neta: net,
          costo_total: cost,
          margen: margin,
          margen_pct: net ? round(margin / net, 4) : 0,
          estado_cobro: paymentDays <= 15 ? 'Cobrado' : paymentDays === 30 ? 'Pendiente' : 'Vencido',
          dias_cobro: paymentDays,
        });
        sequence += 1;
      }
    }
  }

  return rows;
}
