import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from 'file:///C:/Users/desktop/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputDir = path.join(root, 'materiales');
const previewDir = path.join(root, 'work', 'workbook-previews');
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const colors = {
  navy: '#0B1124', slate: '#17213A', violet: '#7C3AED', blue: '#3B82F6',
  teal: '#14B8A6', pale: '#F4F6FB', line: '#D8DFEA', white: '#FFFFFF', muted: '#64748B',
  green: '#DCFCE7', amber: '#FEF3C7', red: '#FEE2E2',
};
const font = 'Aptos';
const money = '$#,##0.00';
const pct = '0.0%';

function round(value, digits = 2) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function buildSales() {
  const clients = [
    ['Andes Market', 'Corporativo'], ['Nova Salud', 'Corporativo'], ['Tierra Verde', 'Pyme'],
    ['Pacífico Retail', 'Corporativo'], ['Mundo Office', 'Pyme'], ['Soluciones Delta', 'Pyme'],
    ['Grupo Horizonte', 'Corporativo'], ['Comercial Yasuní', 'Emprendedor'],
  ];
  const products = [
    ['ERP Contable', 'Software', 1480, 720], ['Dashboard BI', 'Analítica', 970, 390],
    ['API Empresarial', 'Integración', 820, 330], ['Soporte Pro', 'Servicios', 420, 165],
    ['Automatización', 'Servicios', 760, 280], ['Portal Web', 'Software', 1150, 510],
    ['Capacitación', 'Formación', 360, 105],
  ];
  const regions = ['Quito', 'Guayaquil', 'Cuenca', 'Manta'];
  const channels = ['Venta directa', 'Referido', 'Web', 'Alianza'];
  const months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const rows = [];
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
        const price = round(product[2] * season * growth);
        const cost = round(product[3] * quantity * (year === 2026 ? 1.045 : 1));
        const discount = [0, 0.03, 0.05, 0.08][(month + order) % 4];
        const net = round(price * quantity * (1 - discount));
        const margin = round(net - cost);
        const day = 2 + ((order * 2 + month) % 25);
        const paymentDays = [0, 15, 30, 45][(order + month) % 4];
        rows.push([
          `V${String(sequence).padStart(4, '0')}`,
          `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
          year, months[month - 1], `FAC-${year}-${String(sequence).padStart(5, '0')}`,
          client[0], client[1], product[0], product[1], region, channel, quantity,
          price, discount, net, cost, margin, net ? round(margin / net, 4) : 0,
          paymentDays <= 15 ? 'Cobrado' : paymentDays === 30 ? 'Pendiente' : 'Vencido', paymentDays,
        ]);
        sequence += 1;
      }
    }
  }
  return rows;
}

function titleBlock(sheet, title, subtitle, columns = 8) {
  const end = String.fromCharCode(64 + columns);
  sheet.mergeCells(`A1:${end}2`);
  sheet.getRange('A1').values = [[title]];
  sheet.getRange(`A1:${end}2`).format = {
    fill: colors.navy,
    font: { name: font, size: 22, bold: true, color: colors.white },
    verticalAlignment: 'center',
  };
  sheet.mergeCells(`A3:${end}3`);
  sheet.getRange('A3').values = [[subtitle]];
  sheet.getRange(`A3:${end}3`).format = {
    fill: colors.slate,
    font: { name: font, size: 10, color: '#CBD5E1' },
    verticalAlignment: 'center', wrapText: true,
  };
  sheet.getRange('A1').format.rowHeight = 30;
  sheet.getRange('A2').format.rowHeight = 30;
  sheet.getRange('A3').format.rowHeight = 28;
}

function sectionHeader(range) {
  range.format = {
    fill: colors.violet,
    font: { name: font, size: 11, bold: true, color: colors.white },
    verticalAlignment: 'center',
  };
}

const sales = buildSales();
const workbook = Workbook.create();

// 1. Inicio
const intro = workbook.worksheets.add('Inicio');
intro.showGridLines = false;
intro.tabColor = colors.violet;
titleBlock(intro, 'CYSTEMS DATA LAB', 'Práctica guiada: API + Google Sheets + Power BI · Proyecto integrador 4.3', 8);
intro.mergeCells('A5:H5');
intro.getRange('A5').values = [['Objetivo de la sesión']];
sectionHeader(intro.getRange('A5:H5'));
intro.mergeCells('A6:H7');
intro.getRange('A6').values = [['Conectar Power BI con una API, transformar datos de ventas, construir indicadores ejecutivos y presentar tres insights basados en evidencia.']];
intro.getRange('A6:H7').format = { fill: colors.white, font: { name: font, size: 11, color: colors.slate }, wrapText: true, verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: colors.line } };
intro.getRange('A9:H9').merge(); intro.getRange('A9').values = [['Flujo de trabajo']]; sectionHeader(intro.getRange('A9:H9'));
const steps = [
  ['01', 'Explorar', 'Revisa Ventas_En_Vivo y comprende cada campo.'],
  ['02', 'Publicar', 'Convierte el archivo en Google Sheets y publica la hoja Ventas_En_Vivo como CSV.'],
  ['03', 'Conectar', 'Usa el endpoint /api/v1/ventas.json desde Power BI.'],
  ['04', 'Analizar', 'Crea medidas, visuales, filtros y conclusiones ejecutivas.'],
];
steps.forEach((step, index) => {
  const row = 11 + index * 2;
  intro.getRange(`A${row}`).values = [[step[0]]];
  intro.getRange(`A${row}:A${row + 1}`).merge();
  intro.getRange(`A${row}:A${row + 1}`).format = { fill: colors.blue, font: { name: font, size: 14, bold: true, color: colors.white }, horizontalAlignment: 'center', verticalAlignment: 'center' };
  intro.getRange(`B${row}:H${row}`).merge(); intro.getRange(`B${row}`).values = [[step[1]]];
  intro.getRange(`B${row}:H${row}`).format = { fill: '#EEF2FF', font: { name: font, size: 11, bold: true, color: colors.violet }, verticalAlignment: 'center' };
  intro.getRange(`B${row + 1}:H${row + 1}`).merge(); intro.getRange(`B${row + 1}`).values = [[step[2]]];
  intro.getRange(`B${row + 1}:H${row + 1}`).format = { fill: colors.white, font: { name: font, size: 10, color: colors.muted }, wrapText: true, verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: colors.line } };
});
intro.getRange('A20:H20').merge(); intro.getRange('A20').values = [['URL preparada para Power BI']]; sectionHeader(intro.getRange('A20:H20'));
intro.getRange('A21:H22').merge(); intro.getRange('A21').values = [['https://api.cystems.ec/api/v1/ventas.json']];
intro.getRange('A21:H22').format = { fill: '#ECFEFF', font: { name: 'Consolas', size: 11, bold: true, color: '#0F766E' }, horizontalAlignment: 'center', verticalAlignment: 'center', borders: { preset: 'all', style: 'thin', color: '#99F6E4' } };
intro.getRange('A24:H25').merge(); intro.getRange('A24').values = [['Importante: edita únicamente las filas de Ventas_En_Vivo. Conserva los encabezados para que la API pueda interpretar la hoja.']];
intro.getRange('A24:H25').format = { fill: colors.amber, font: { name: font, size: 10, bold: true, color: '#92400E' }, wrapText: true, verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: '#FCD34D' } };
intro.getRange('A1:H25').format.font.name = font;
intro.getRange('A1:A25').format.columnWidth = 10;
intro.getRange('B1:H25').format.columnWidth = 15;
intro.freezePanes.freezeRows(3);

// 2. Ventas_En_Vivo
const data = workbook.worksheets.add('Ventas_En_Vivo');
data.showGridLines = false;
data.tabColor = colors.teal;
const headers = ['id','fecha','anio','mes','documento','cliente','segmento','producto','categoria','region','canal','cantidad','precio_unitario','descuento_pct','venta_neta','costo_total','margen','margen_pct','estado_cobro','dias_cobro'];
data.getRangeByIndexes(0, 0, sales.length + 1, headers.length).values = [headers, ...sales];
data.getRange('A1:T1').format = { fill: colors.navy, font: { name: font, size: 9, bold: true, color: colors.white }, wrapText: true, verticalAlignment: 'center', horizontalAlignment: 'center', borders: { preset: 'all', style: 'thin', color: '#334155' } };
data.getRange('A2:T361').format = { font: { name: font, size: 9, color: colors.slate }, borders: { preset: 'inside', style: 'thin', color: '#E5E7EB' }, verticalAlignment: 'center' };
data.getRange('B2:B361').setNumberFormat('yyyy-mm-dd');
data.getRange('M2:M361').setNumberFormat(money);
data.getRange('N2:N361').setNumberFormat(pct);
data.getRange('O2:Q361').setNumberFormat(money);
data.getRange('R2:R361').setNumberFormat(pct);
data.getRange('N2:N361').format.fill = '#FFF7ED';
data.getRange('O2:R361').format.fill = '#F0FDF4';
data.getRange('R2:R361').conditionalFormats.add('colorScale', { colors: ['#FEE2E2', '#FEF3C7', '#DCFCE7'], thresholds: ['min', '50%', 'max'] });
data.getRange('S2:S361').conditionalFormats.add('containsText', { text: 'Vencido', format: { fill: colors.red, font: { color: '#991B1B', bold: true } } });
data.tables.add('A1:T361', true, 'TablaVentasCystems').style = 'TableStyleMedium2';
data.freezePanes.freezeRows(1);
data.freezePanes.freezeColumns(2);
['A','C','L','T'].forEach(c => data.getRange(`${c}1:${c}361`).format.columnWidth = 10);
['B','D','E','G','I','J','K','S'].forEach(c => data.getRange(`${c}1:${c}361`).format.columnWidth = 14);
['F','H'].forEach(c => data.getRange(`${c}1:${c}361`).format.columnWidth = 20);
['M','N','O','P','Q','R'].forEach(c => data.getRange(`${c}1:${c}361`).format.columnWidth = 15);
data.getRange('1:1').format.rowHeight = 32;

// 3. Control
const control = workbook.worksheets.add('Control');
control.showGridLines = false;
control.tabColor = colors.blue;
titleBlock(control, 'PANEL DE CONTROL', 'Totales de validación para comprobar la conexión y los resultados de Power BI.', 8);
control.getRange('A5:B5').values = [['Indicador', 'Resultado esperado']];
sectionHeader(control.getRange('A5:B5'));
control.getRange('A6:A11').values = [['Ventas netas'],['Costos'],['Margen'],['Margen %'],['Transacciones'],['Clientes únicos']];
control.getRange('B6:B10').formulas = [['=SUM(Ventas_En_Vivo!O2:O361)'],['=SUM(Ventas_En_Vivo!P2:P361)'],['=SUM(Ventas_En_Vivo!Q2:Q361)'],['=IFERROR(B8/B6,0)'],['=COUNTA(Ventas_En_Vivo!A2:A361)']];
control.getRange('B11').values = [[8]];
control.getRange('A6:B11').format = { font: { name: font, size: 10, color: colors.slate }, borders: { preset: 'all', style: 'thin', color: colors.line }, verticalAlignment: 'center' };
control.getRange('A6:A11').format.fill = '#EEF2FF'; control.getRange('A6:A11').format.font.bold = true;
control.getRange('B6:B8').setNumberFormat(money); control.getRange('B9').setNumberFormat(pct); control.getRange('B10:B11').setNumberFormat('0');
control.getRange('D5:E5').values = [['Año', 'Ventas netas']]; sectionHeader(control.getRange('D5:E5'));
control.getRange('D6:D7').values = [[2025],[2026]];
control.getRange('E6:E7').formulas = [['=SUMIF(Ventas_En_Vivo!C2:C361,D6,Ventas_En_Vivo!O2:O361)'],['=SUMIF(Ventas_En_Vivo!C2:C361,D7,Ventas_En_Vivo!O2:O361)']];
control.getRange('D6:E7').format = { font: { name: font, size: 10, color: colors.slate }, borders: { preset: 'all', style: 'thin', color: colors.line } }; control.getRange('E6:E7').setNumberFormat(money);
control.getRange('D10:E10').values = [['Comprobación', 'Estado']]; sectionHeader(control.getRange('D10:E10'));
control.getRange('D11').values = [['Filas esperadas']]; control.getRange('E11').formulas = [['=IF(B10=360,"OK","REVISAR")']];
control.getRange('D11:E11').format = { fill: '#ECFDF5', font: { name: font, size: 10, bold: true, color: '#065F46' }, borders: { preset: 'all', style: 'thin', color: '#A7F3D0' } };
control.getRange('A14:H14').merge(); control.getRange('A14').values = [['Cómo usar esta hoja']]; sectionHeader(control.getRange('A14:H14'));
control.getRange('A15:H18').merge(); control.getRange('A15').values = [['1. Antes de iniciar, registra estos totales.\n2. Modifica una venta en Ventas_En_Vivo o en Google Sheets.\n3. Actualiza Power BI y compara el indicador afectado.\n4. Si el número no cambia, revisa la URL, los tipos de datos y la credencial anónima.']];
control.getRange('A15:H18').format = { fill: colors.white, font: { name: font, size: 10, color: colors.slate }, wrapText: true, verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: colors.line } };
control.getRange('A1:H20').format.columnWidth = 14; control.getRange('A1:A20').format.columnWidth = 22; control.getRange('B1:B20').format.columnWidth = 24;
control.freezePanes.freezeRows(3);

// 4. Metas
const goals = workbook.worksheets.add('Metas');
goals.showGridLines = false;
goals.tabColor = '#F59E0B';
titleBlock(goals, 'METAS COMERCIALES', 'Valores de referencia opcionales para comparar desempeño mensual.', 5);
goals.getRange('A5:E5').values = [['anio','mes_numero','mes','meta_ventas','meta_margen_pct']]; sectionHeader(goals.getRange('A5:E5'));
const monthNames = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const goalRows = [];
for (const year of [2025, 2026]) for (let m = 1; m <= 12; m += 1) goalRows.push([year,m,monthNames[m-1],round(30000*(year===2026?1.12:1)*(1+(m-6)*0.01)),year===2026?0.58:0.55]);
goals.getRange('A6:E29').values = goalRows;
goals.getRange('A6:E29').format = { font: { name: font, size: 10, color: colors.slate }, borders: { preset: 'all', style: 'thin', color: colors.line } };
goals.getRange('D6:D29').setNumberFormat(money); goals.getRange('E6:E29').setNumberFormat(pct);
goals.tables.add('A5:E29', true, 'TablaMetasCystems').style = 'TableStyleMedium4';
goals.getRange('A1:E29').format.columnWidth = 18; goals.getRange('C1:C29').format.columnWidth = 20; goals.freezePanes.freezeRows(5);

// 5. Diccionario
const dictionary = workbook.worksheets.add('Diccionario_API');
dictionary.showGridLines = false;
dictionary.tabColor = '#64748B';
titleBlock(dictionary, 'DICCIONARIO DE DATOS', 'Estructura oficial de /api/v1/ventas.json', 5);
dictionary.getRange('A5:E5').values = [['campo','tipo','descripción','ejemplo','uso en Power BI']]; sectionHeader(dictionary.getRange('A5:E5'));
const dictionaryRows = [
  ['id','Texto','Identificador único','V0001','Clave de la transacción'],['fecha','Fecha','Fecha de emisión','2026-01-03','Eje temporal'],['anio','Entero','Año calendario',2026,'Segmentador'],['mes','Texto','Nombre del mes','Enero','Etiqueta'],['documento','Texto','Número de factura','FAC-2026-00181','Detalle'],['cliente','Texto','Cliente o razón social','Andes Market','Ranking'],['segmento','Texto','Segmento comercial','Corporativo','Segmentador'],['producto','Texto','Producto o servicio','Dashboard BI','Ranking'],['categoria','Texto','Familia del producto','Analítica','Segmentador'],['region','Texto','Ciudad o región','Quito','Mapa o barras'],['canal','Texto','Canal comercial','Web','Segmentador'],['cantidad','Entero','Unidades vendidas',2,'SUM'],['precio_unitario','Moneda','Precio antes de descuento',970,'Promedio'],['descuento_pct','Decimal','Descuento entre 0 y 1',0.05,'Porcentaje'],['venta_neta','Moneda','Ingreso neto',1843,'SUM'],['costo_total','Moneda','Costo asociado',780,'SUM'],['margen','Moneda','Venta menos costo',1063,'SUM'],['margen_pct','Decimal','Margen / venta neta',0.5768,'Porcentaje'],['estado_cobro','Texto','Estado de cartera','Cobrado','Segmentador'],['dias_cobro','Entero','Días de cobro',15,'Promedio'],
];
dictionary.getRange('A6:E25').values = dictionaryRows;
dictionary.getRange('A6:E25').format = { font: { name: font, size: 9, color: colors.slate }, borders: { preset: 'all', style: 'thin', color: colors.line }, wrapText: true, verticalAlignment: 'center' };
dictionary.getRange('A6:A25').format.fill = '#EEF2FF'; dictionary.getRange('A6:A25').format.font = { name: 'Consolas', size: 9, bold: true, color: colors.violet };
dictionary.getRange('A1:A25').format.columnWidth = 20; dictionary.getRange('B1:B25').format.columnWidth = 13; dictionary.getRange('C1:C25').format.columnWidth = 32; dictionary.getRange('D1:D25').format.columnWidth = 22; dictionary.getRange('E1:E25').format.columnWidth = 22;
dictionary.freezePanes.freezeRows(5);

workbook.recalculate();

const checks = await workbook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'final formula error scan' });
await fs.writeFile(path.join(previewDir, 'formula-scan.json'), JSON.stringify(checks, null, 2));

for (const sheetName of ['Inicio','Ventas_En_Vivo','Control','Metas','Diccionario_API']) {
  const preview = await workbook.render({ sheetName, autoCrop: 'all', scale: sheetName === 'Ventas_En_Vivo' ? 0.55 : 1, format: 'png' });
  await fs.writeFile(path.join(previewDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
const outputPath = path.join(outputDir, 'Practica_API_Cystems_Power_BI.xlsx');
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, rows: sales.length, sheets: 5, previewDir }));
