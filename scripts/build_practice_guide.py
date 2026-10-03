from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "materiales" / "Guion_Practica_API_Cystems_90_Minutos.docx"
OUT.parent.mkdir(parents=True, exist_ok=True)

NAVY = "0B1124"
SLATE = "17213A"
VIOLET = "7C3AED"
BLUE = "3B82F6"
TEAL = "14B8A6"
PALE = "F4F6FB"
MUTED = "64748B"
LINE = "D8DFEA"
WHITE = "FFFFFF"
AMBER = "FEF3C7"


def shade(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), color)


def cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_text(cell, text, *, color=SLATE, bold=False, size=9, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = "Aptos"
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell_margins(cell)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Página ")
    run.font.name = "Aptos"
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(MUTED)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    rel_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(color)
    r_pr.append(underline)
    run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.add_run(text)
    return p


def add_body(doc, text, *, bold_lead=None, after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.12
    if bold_lead and text.startswith(bold_lead):
        p.add_run(bold_lead).bold = True
        p.add_run(text[len(bold_lead):])
    else:
        p.add_run(text)
    return p


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(2)
    p.add_run(text)
    return p


def add_number(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(3)
    p.add_run(text)
    return p


def add_callout(doc, title, body, color=VIOLET, fill="F5F3FF"):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(0.13)
    table.columns[1].width = Inches(6.7)
    left, right = table.rows[0].cells
    shade(left, color)
    shade(right, fill)
    set_cell_text(left, "", color=WHITE)
    right.text = ""
    p = right.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(title)
    r.bold = True
    r.font.name = "Aptos"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string(color)
    p2 = right.add_paragraph(body)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.08
    cell_margins(right, 120, 160, 120, 160)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_timeline_table(doc, rows):
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [0.9, 1.85, 2.4, 1.65]
    for i, width in enumerate(widths):
        table.columns[i].width = Inches(width)
    headers = ["Minutos", "Acción docente", "Acción del estudiante", "Evidencia"]
    for i, header in enumerate(headers):
        shade(table.rows[0].cells[i], NAVY)
        set_cell_text(table.rows[0].cells[i], header, color=WHITE, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_repeat_table_header(table.rows[0])
    for index, row in enumerate(rows):
        cells = table.add_row().cells
        fill = WHITE if index % 2 == 0 else PALE
        for i, value in enumerate(row):
            shade(cells[i], fill)
            set_cell_text(cells[i], value, color=SLATE, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)
    return table


def add_step(doc, minute, title, objective, teacher, students, checks, contingency=None):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(1.05)
    table.columns[1].width = Inches(5.8)
    left, right = table.rows[0].cells
    shade(left, VIOLET)
    shade(right, NAVY)
    set_cell_text(left, minute, color=WHITE, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(right, title, color=WHITE, bold=True, size=12)
    add_body(doc, f"Objetivo: {objective}", bold_lead="Objetivo:")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Qué hace y dice el docente")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(VIOLET)
    for item in teacher:
        add_bullet(doc, item)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Qué hacen los estudiantes")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(BLUE)
    for item in students:
        add_bullet(doc, item)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Comprobación antes de avanzar")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(TEAL)
    for item in checks:
        add_bullet(doc, item)
    if contingency:
        add_callout(doc, "Plan B inmediato", contingency, color="B45309", fill=AMBER)


doc = Document()
doc.settings.odd_and_even_pages_header_footer = True
section = doc.sections[0]
section.top_margin = Inches(0.62)
section.bottom_margin = Inches(0.62)
section.left_margin = Inches(0.72)
section.right_margin = Inches(0.72)

styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"].font.size = Pt(9.5)
styles["Normal"].font.color.rgb = RGBColor.from_string(SLATE)
for style_name, size, color in (("Title", 30, NAVY), ("Subtitle", 13, MUTED), ("Heading 1", 19, NAVY), ("Heading 2", 14, VIOLET), ("Heading 3", 11, BLUE)):
    style = styles[style_name]
    style.font.name = "Aptos Display" if style_name in ("Title", "Heading 1") else "Aptos"
    style.font.size = Pt(size)
    style.font.bold = style_name != "Subtitle"
    style.font.color.rgb = RGBColor.from_string(color)
    style.paragraph_format.space_before = Pt(12 if style_name != "Title" else 0)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.keep_with_next = True

header = section.header
header_p = header.paragraphs[0]
header_p.text = "CYSTEMS  ·  DATA LAB"
header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for run in header_p.runs:
    run.font.name = "Aptos"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(VIOLET)
footer = section.footer
footer_p = footer.paragraphs[0]
footer_p.text = "Práctica API + Power BI en navegador  ·  "
add_page_number(footer_p)
even_header_p = section.even_page_header.paragraphs[0]
even_header_p.text = "CYSTEMS  ·  DATA LAB"
even_header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for run in even_header_p.runs:
    run.font.name = "Aptos"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(VIOLET)
even_footer_p = section.even_page_footer.paragraphs[0]
even_footer_p.text = "Práctica API + Power BI en navegador  ·  "
add_page_number(even_footer_p)

# Cover
top = doc.add_table(rows=1, cols=1)
top.alignment = WD_TABLE_ALIGNMENT.CENTER
shade(top.cell(0, 0), NAVY)
set_cell_text(top.cell(0, 0), "CYSTEMS DATA LAB", color=WHITE, bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph().paragraph_format.space_after = Pt(24)
p = doc.add_paragraph(style="Title")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Guion de práctica")
p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p2.add_run("API + GOOGLE SHEETS + POWER BI EN NAVEGADOR")
r.bold = True
r.font.name = "Aptos"
r.font.size = Pt(14)
r.font.color.rgb = RGBColor.from_string(VIOLET)
p3 = doc.add_paragraph(style="Subtitle")
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.add_run("Duración completa: 90 minutos · Ruta comprimida: 60 minutos")
doc.add_paragraph().paragraph_format.space_after = Pt(18)
cover_table = doc.add_table(rows=4, cols=2)
cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
cover_table.autofit = False
cover_table.columns[0].width = Inches(2.1)
cover_table.columns[1].width = Inches(4.3)
cover_data = [
    ("Modalidad", "Práctica guiada, 100 % en navegador"),
    ("Fuente", "api.cystems.ec + hoja editable en Google Sheets"),
    ("Producto de la sesión", "Reporte ejecutivo de una página"),
    ("Nivel", "Intermedio, con acompañamiento paso a paso"),
]
for row, (label, value) in zip(cover_table.rows, cover_data):
    shade(row.cells[0], "EEF2FF")
    shade(row.cells[1], WHITE)
    set_cell_text(row.cells[0], label, color=VIOLET, bold=True, size=9.5)
    set_cell_text(row.cells[1], value, color=SLATE, size=9.5)
doc.add_paragraph().paragraph_format.space_after = Pt(18)
add_callout(doc, "Resultado final", "Cada estudiante deja creado un modelo semántico conectado a la API, siete medidas DAX, seis visuales, dos segmentadores y tres conclusiones gerenciales.", color=TEAL, fill="ECFEFF")
doc.add_page_break()

# Purpose and exact unit titles
add_heading(doc, "Propósito y relación con la unidad", 1)
add_body(doc, "Esta práctica prepara el cierre del curso mediante un flujo completo: fuente editable → API → modelo semántico → reporte → insight. Los títulos curriculares se mantienen exactamente como fueron definidos.")
for title in [
    "4.3 Proyecto integrador",
    "4.3.1 Desarrollo de un sistema completo de reporting contable",
    "4.3.2 Análisis de caso real con datos proporcionados",
    "4.3.3 Presentación de insights estratégicos derivados de los datos",
]:
    add_callout(doc, title, "Aplicación directa durante la práctica.", color=VIOLET, fill="F5F3FF")

add_heading(doc, "Objetivos observables", 2)
for item in [
    "Reconocer la arquitectura de una solución sencilla de datos en la nube.",
    "Conectar Power BI Service con un endpoint JSON usando Power Query Online.",
    "Configurar tipos de datos y conservar una tabla limpia llamada Ventas_En_Vivo.",
    "Crear siete medidas DAX y utilizarlas en visualizaciones ejecutivas.",
    "Comprobar una actualización originada por un cambio en Google Sheets.",
    "Redactar tres insights con evidencia, interpretación y recomendación.",
]:
    add_bullet(doc, item)

add_heading(doc, "Criterio de éxito al terminar", 2)
success = doc.add_table(rows=1, cols=3)
success.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, header_text in enumerate(["Componente", "Mínimo esperado", "Comprobación"]):
    shade(success.rows[0].cells[i], NAVY)
    set_cell_text(success.rows[0].cells[i], header_text, color=WHITE, bold=True, size=9)
for row_data in [
    ("Modelo", "1 tabla con 20 columnas y 360 filas", "Sin errores de tipo"),
    ("DAX", "7 medidas", "Devuelven valores"),
    ("Reporte", "4 tarjetas + 2 gráficos + 2 filtros", "Interacciones activas"),
    ("Actualización", "1 cambio visible", "Antes y después registrados"),
    ("Insights", "3 conclusiones", "Dato + lectura + acción"),
]:
    cells = success.add_row().cells
    for i, value in enumerate(row_data):
        shade(cells[i], WHITE if len(success.rows) % 2 == 0 else PALE)
        set_cell_text(cells[i], value, size=8.5)
doc.add_page_break()

# Prep
add_heading(doc, "Preparación del docente antes de la clase", 1)
add_callout(doc, "Realizar antes de que ingresen los estudiantes", "La práctica depende de una conexión pública de lectura. Verifique la API, la cuenta de Power BI y la hoja de Google al menos 15 minutos antes.", color="B45309", fill=AMBER)
for item in [
    "Abrir https://api.cystems.ec/api/v1/health.json y confirmar ok: true, records: 360 y source: google-sheets.",
    "Abrir https://api.cystems.ec/api/v1/ventas.json?limit=3 y confirmar que aparecen tres registros dentro de data.",
    "Abrir el Excel Practica_API_Cystems_Power_BI.xlsx y revisar las hojas Inicio, Ventas_En_Vivo, Control, Metas y Diccionario_API.",
    "Mantener abierta la hoja Ventas_En_Vivo en Google Sheets para la demostración de actualización.",
    "Comprobar que los estudiantes pueden crear contenido en Mi área de trabajo o en el espacio asignado.",
    "Confirmar que la edición web del modelo semántico está habilitada por el administrador del tenant.",
    "Escribir en la pizarra la URL: https://api.cystems.ec/api/v1/ventas.json?refresh=1",
]:
    add_bullet(doc, item)

add_heading(doc, "Archivos y pestañas que debe tener abiertas", 2)
for item in [
    "Pestaña 1: CYSTEMS Data Lab, sección Estudiar.",
    "Pestaña 2: explorador de la base de datos.",
    "Pestaña 3: Power BI Service en https://app.powerbi.com.",
    "Pestaña 4: Google Sheets con la fuente Ventas_En_Vivo.",
    "Documento de apoyo: este guion, abierto en una segunda pantalla si es posible.",
]:
    add_bullet(doc, item)

add_heading(doc, "Si el conector o la edición web no aparecen", 2)
add_body(doc, "No detenga toda la clase. Use el archivo Excel como ruta alternativa: Nuevo elemento → Modelo semántico → Excel → cargar Practica_API_Cystems_Power_BI.xlsx → seleccionar TablaVentasCystems. La estructura, las medidas y los visuales se mantienen; únicamente se omite la demostración de actualización por API.")
add_body(doc, "La ruta principal usa Power Query Online con el conector Web API. Microsoft documenta que este conector admite contenido JSON, autenticación anónima y uso en servicios de nube sin gateway para Web.Contents.")
p = doc.add_paragraph()
add_hyperlink(p, "Referencia oficial: conector Web de Power Query", "https://learn.microsoft.com/power-query/connectors/web/web")
p.add_run("  ·  ")
add_hyperlink(p, "Referencia oficial: edición de modelos en Power BI Service", "https://learn.microsoft.com/es-es/power-bi/transform-model/service-edit-data-models")
doc.add_page_break()

# Agenda
add_heading(doc, "Mapa completo de la sesión · 90 minutos", 1)
agenda = [
    ("0–5", "Presentar el reto", "Escuchar y abrir el portal", "Pregunta gerencial clara"),
    ("5–12", "Explorar datos y arquitectura", "Revisar tabla y endpoint", "Campos identificados"),
    ("12–27", "Guiar la conexión Web API", "Crear consulta en Power Query Online", "360 filas cargadas"),
    ("27–38", "Validar modelo", "Tipos, nombres y medidas base", "Modelo sin errores"),
    ("38–53", "Construir siete medidas DAX", "Copiar, validar y formatear", "Medidas visibles"),
    ("53–70", "Diseñar reporte ejecutivo", "Crear KPIs, tendencia y rankings", "Página funcional"),
    ("70–79", "Simular actualización", "Registrar antes/después y refrescar", "Cambio comprobado"),
    ("79–87", "Convertir hallazgos en insights", "Redactar tres conclusiones", "Tres decisiones"),
    ("87–90", "Cerrar y guardar", "Nombrar y revisar", "Reporte guardado"),
]
add_timeline_table(doc, agenda)

add_heading(doc, "Ruta comprimida · 60 minutos", 2)
compressed = [
    ("0–5", "Reto y arquitectura", "Abrir portal y API", "Objetivo claro"),
    ("5–17", "Conexión", "Crear consulta", "Tabla cargada"),
    ("17–28", "DAX", "Crear 4 medidas: ventas, costos, margen y margen %", "4 medidas"),
    ("28–45", "Reporte", "4 tarjetas, 1 tendencia, 1 ranking, 2 filtros", "Página lista"),
    ("45–53", "Actualización", "Refrescar y comparar", "Cambio visible"),
    ("53–60", "Insights y cierre", "Redactar 2 conclusiones", "Reporte guardado"),
]
add_timeline_table(doc, compressed)
add_callout(doc, "Qué se recorta en 60 minutos", "No se crea Ticket Promedio, Clientes ni Transacciones; se omite el gráfico regional y se redactan dos insights en lugar de tres.", color=BLUE, fill="EFF6FF")
doc.add_page_break()

# Detailed steps
add_heading(doc, "Guion minuto a minuto", 1)
add_step(
    doc, "0–5", "Abrir con una decisión, no con una herramienta",
    "Dar sentido empresarial a la práctica antes de mostrar menús.",
    [
        "Proyecte el portal CYSTEMS Data Lab y diga: «Hoy no vamos a importar un archivo estático. Vamos a construir un reporte que consulta una API y que cambia cuando cambia la fuente». ",
        "Plantee el caso: «Gerencia necesita conocer ventas, margen, productos líderes y ciudades con mayor aporte. Al final deberán recomendar una acción». ",
        "Pregunte: «¿Qué diferencia existe entre ver datos y tomar una decisión con datos?» Escuche dos respuestas breves.",
        "Muestre el resultado final en /dashboard durante no más de 30 segundos. No explique todavía cada gráfico.",
    ],
    [
        "Ingresan al portal y abren Estudiar.",
        "Escriben en una nota la pregunta: ¿qué está explicando la rentabilidad de CYSTEMS?",
    ],
    [
        "Todos tienen abierto el portal y Power BI Service.",
        "Al menos dos estudiantes pueden repetir el objetivo de la práctica con sus propias palabras.",
    ],
    "Si el portal no abre, comparta la URL directa de la API y continúe desde Power BI Service. El portal es apoyo; la API es la fuente.",
)

add_step(
    doc, "5–12", "Leer la arquitectura y reconocer la base",
    "Comprender qué viaja desde la hoja hasta el reporte.",
    [
        "Dibuje o muestre la cadena: Google Sheets → API CYSTEMS → Power BI → decisión.",
        "Abra Base de datos y señale fecha, cliente, producto, región, canal, venta_neta, costo_total y margen.",
        "Diga: «Una fila es una transacción. Las columnas describen cuándo ocurrió, a quién, qué se vendió y qué resultado financiero produjo». ",
        "Abra /api/v1/ventas.json?limit=3 y ubique las propiedades ok, meta, total y data. Explique que Power BI debe expandir data.",
    ],
    [
        "Abren el explorador de datos y aplican un filtro de año o ciudad.",
        "Localizan una fila y comprueban manualmente: margen = venta_neta − costo_total.",
        "Abren el endpoint con limit=3 y reconocen que data es una lista de registros.",
    ],
    [
        "Pueden identificar al menos dos dimensiones y dos métricas.",
        "Comprenden que modificar el JSON directamente no es parte del ejercicio; la fuente editable es Google Sheets.",
    ],
)

add_step(
    doc, "12–27", "Crear el modelo desde la API en Power BI Service",
    "Importar y transformar el JSON completamente desde el navegador.",
    [
        "Pida abrir app.powerbi.com y entrar al área de trabajo asignada.",
        "Indique: Crear o Nuevo elemento → Modelo semántico → Obtener datos. Si aparece una galería de conectores, buscar Web API.",
        "Pida pegar exactamente https://api.cystems.ec/api/v1/ventas.json?refresh=1 y seleccionar autenticación Anónima.",
        "Cuando aparezca Power Query Online, seleccione el registro data. Si se muestra como Lista, elija Convertir en tabla.",
        "En Column1 seleccione Expandir y marque todos los campos. Desactive «usar el nombre original como prefijo». ",
        "Renombre la consulta como Ventas_En_Vivo.",
        "Configure: fecha = Fecha; anio, cantidad y dias_cobro = Número entero; precio_unitario, venta_neta, costo_total y margen = Número decimal fijo/moneda; descuento_pct y margen_pct = Número decimal.",
        "Seleccione Crear un reporte como destino final si la opción está disponible. Si no, cree solo el modelo y luego use Nuevo reporte.",
    ],
    [
        "Repiten la ruta en su navegador.",
        "Comprueban que la vista previa tenga 20 columnas y que total indique 360 registros.",
        "Guardan el modelo como Modelo_Cystems_Apellido_Nombre.",
    ],
    [
        "La tabla se llama Ventas_En_Vivo.",
        "No existe una columna llamada Column1 ni encabezados con prefijos data.",
        "Los importes se alinean como números y la fecha se reconoce como fecha.",
    ],
    "Si Web API no aparece, use Obtener datos → Web o pegue el código M del anexo. Si el tenant no permite Power Query Online, cargue el Excel y seleccione TablaVentasCystems.",
)

add_step(
    doc, "27–38", "Validar el modelo antes de diseñar",
    "Evitar que un error de tipo o nombre se convierta en un error visual.",
    [
        "Abra el modelo semántico en modo Edición y ubique la tabla Ventas_En_Vivo.",
        "Revise la suma predeterminada: id y documento no se resumen; venta_neta, costo_total y margen sí se suman.",
        "Formatee descuento_pct y margen_pct como porcentaje con un decimal.",
        "Diga: «Primero hacemos confiable el modelo; después hacemos bonito el reporte». ",
        "Compare los totales con la hoja Control del Excel: 360 transacciones y ventas aproximadas de $781.519,16.",
    ],
    [
        "Corrigen tipos y formatos.",
        "Verifican que Ventas_En_Vivo[venta_neta] tenga total $781.519,16.",
        "Mantienen abierta la hoja Control como referencia.",
    ],
    [
        "Las ventas y los costos coinciden con Control.",
        "No existen errores ni columnas numéricas tratadas como texto.",
    ],
)

add_step(
    doc, "38–53", "Crear las siete medidas DAX",
    "Centralizar los indicadores que usarán todos los visuales.",
    [
        "Seleccione Ventas_En_Vivo → Nueva medida. Pegue una medida cada vez y pulse Enter.",
        "Después de cada medida, arrástrela temporalmente a una tarjeta para comprobar que devuelve un valor.",
        "Explique la diferencia: una columna guarda un valor por fila; una medida se calcula según el contexto de filtros.",
        "Formatee Ventas Netas, Costos Totales, Margen Total y Ticket Promedio como moneda; Margen % como porcentaje; Transacciones y Clientes como entero.",
    ],
    [
        "Crean y nombran exactamente las siete medidas del bloque DAX de la página siguiente.",
        "Corrigen cualquier nombre de tabla o columna que IntelliSense marque en rojo.",
        "Ocultan la tarjeta temporal cuando la comprobación termina.",
    ],
    [
        "Ventas Netas devuelve aproximadamente $781.519,16.",
        "Margen % devuelve aproximadamente 58,1 %.",
        "Transacciones devuelve 360 y Clientes devuelve 8.",
    ],
)

add_heading(doc, "Bloque DAX para copiar", 2)
dax_table = doc.add_table(rows=1, cols=2)
dax_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, header_text in enumerate(["Medida", "Fórmula DAX"]):
    shade(dax_table.rows[0].cells[i], NAVY)
    set_cell_text(dax_table.rows[0].cells[i], header_text, color=WHITE, bold=True, size=9)
dax_rows = [
    ("Ventas Netas", "Ventas Netas = SUM(Ventas_En_Vivo[venta_neta])"),
    ("Costos Totales", "Costos Totales = SUM(Ventas_En_Vivo[costo_total])"),
    ("Margen Total", "Margen Total = [Ventas Netas] - [Costos Totales]"),
    ("Margen %", "Margen % = DIVIDE([Margen Total], [Ventas Netas], 0)"),
    ("Transacciones", "Transacciones = COUNTROWS(Ventas_En_Vivo)"),
    ("Clientes", "Clientes = DISTINCTCOUNT(Ventas_En_Vivo[cliente])"),
    ("Ticket Promedio", "Ticket Promedio = DIVIDE([Ventas Netas], [Transacciones], 0)"),
]
for idx, row_data in enumerate(dax_rows):
    cells = dax_table.add_row().cells
    fill = WHITE if idx % 2 == 0 else PALE
    for i, value in enumerate(row_data):
        shade(cells[i], fill)
        set_cell_text(cells[i], value, color=SLATE if i == 0 else VIOLET, bold=i == 0, size=8.5)

add_step(
    doc, "53–70", "Diseñar una página ejecutiva",
    "Convertir las medidas en una lectura visual clara y filtrable.",
    [
        "Cree un reporte nuevo desde el modelo. Nombre la página Resumen Ejecutivo.",
        "Añada cuatro tarjetas en la primera fila: Ventas Netas, Margen Total, Margen % y Transacciones.",
        "Añada un gráfico de líneas: fecha en el eje X con nivel Mes y Ventas Netas en valores.",
        "Añada un gráfico de barras: producto en eje y Ventas Netas en valores; orden descendente y Top 7.",
        "Añada un segundo gráfico de barras: region en eje y Margen Total en valores.",
        "Añada dos segmentadores: anio y canal. Active selección única solo para anio si desea comparar un período a la vez.",
        "Use fondo muy claro, violeta para ventas, azul para tendencia y turquesa para margen. Mantenga títulos breves y unidades consistentes.",
        "Abra Editar interacciones y compruebe que los segmentadores filtren todos los visuales.",
    ],
    [
        "Construyen los ocho elementos mínimos: 4 tarjetas, 2 gráficos y 2 segmentadores.",
        "Alinean tarjetas y aplican el mismo tamaño.",
        "Prueban año 2025 y 2026; observan que todos los valores cambian.",
    ],
    [
        "No hay gráficos que muestren Conteo de venta_neta por error.",
        "Los títulos permiten entender el visual sin explicación oral.",
        "El reporte responde a año y canal.",
    ],
)

add_step(
    doc, "70–79", "Demostrar una actualización de extremo a extremo",
    "Comprobar que el reporte puede reflejar un cambio hecho en la fuente.",
    [
        "Pida que todos seleccionen 2026 y anoten el valor de Ventas Netas.",
        "En Google Sheets localice una fila de 2026 y cambie venta_neta en un monto claramente visible. Actualice margen y margen_pct de esa misma fila para mantener coherencia.",
        "Espere entre 15 y 20 segundos o use el endpoint con refresh=1.",
        "En el espacio de trabajo, abra el menú del modelo semántico → Actualizar ahora. Espere a que el historial indique Completado.",
        "Vuelva al reporte y seleccione Actualizar visuales.",
        "Pregunte: «¿Qué medida cambió y por qué? ¿Qué debería ocurrir con el margen?»",
        "Al terminar, restituya los tres valores originales de la fila para que la base quede limpia.",
    ],
    [
        "Registran el valor antes y después.",
        "Actualizan el modelo y luego los visuales; no confunden ambos pasos.",
        "Explican el recorrido del cambio: hoja → API → actualización del modelo → visual.",
    ],
    [
        "El historial de actualización finaliza sin error.",
        "El valor mostrado cambia en la cantidad esperada.",
        "Después de restituir la fila y actualizar nuevamente, el valor vuelve al total de control.",
    ],
    "Si la actualización falla, abra health.json. Si source muestra demo, la API aún no está leyendo Google Sheets. Continúe la explicación con una captura del antes/después y no haga que cada estudiante repita el error.",
)

add_step(
    doc, "79–87", "Redactar tres insights estratégicos",
    "Pasar de describir gráficos a recomendar decisiones.",
    [
        "Escriba la estructura: Evidencia → Interpretación → Acción.",
        "Modele un ejemplo: «En 2026, Guayaquil concentra la mayor venta; esto indica dependencia comercial de una ciudad; recomendamos comparar su margen y replicar el canal más efectivo en Quito». ",
        "Pida un insight sobre producto, uno sobre territorio y uno sobre rentabilidad.",
        "Solicite que cada conclusión incluya al menos un número visible en el reporte.",
    ],
    [
        "Redactan tres conclusiones en un cuadro de texto del reporte o en una nota de entrega.",
        "Evitan frases vagas como «las ventas son buenas». Usan cifras, contexto y una acción.",
        "Comparten una conclusión con el compañero más cercano para comprobar si se entiende sin explicar el gráfico.",
    ],
    [
        "Cada insight contiene evidencia cuantitativa.",
        "La acción recomendada se relaciona directamente con la evidencia.",
        "No se confunde venta alta con rentabilidad alta.",
    ],
)

add_step(
    doc, "87–90", "Guardar, nombrar y cerrar",
    "Dejar el trabajo identificable y listo para continuar en el proyecto final.",
    [
        "Pida guardar el reporte como Dashboard_API_Cystems_Apellido_Nombre.",
        "Realice una revisión relámpago: filtros, títulos, formatos, totales e insights.",
        "Cierre diciendo: «Hoy construyeron el flujo mínimo de un sistema de reporting contable. En el proyecto final ampliarán este mismo patrón con más páginas y decisiones». ",
    ],
    [
        "Guardan y verifican que el reporte aparezca en el espacio de trabajo.",
        "Toman nota de un ajuste pendiente para el proyecto integrador final.",
    ],
    [
        "El nombre contiene apellido y nombre.",
        "El reporte abre desde el espacio de trabajo sin perder visuales.",
    ],
)

doc.add_page_break()
add_heading(doc, "Preguntas de acompañamiento durante la práctica", 1)
questions = [
    ("Al revisar la fuente", "¿Cuál es la diferencia entre una dimensión y una métrica en esta tabla?"),
    ("Al expandir JSON", "¿Por qué debemos entrar en data y no quedarnos en el registro principal?"),
    ("Al configurar tipos", "¿Qué problema tendría una suma si venta_neta fuera texto?"),
    ("Al crear DAX", "¿Qué cambia en la medida cuando seleccionamos 2025?"),
    ("Al diseñar", "¿Este visual responde a una pregunta o solo ocupa espacio?"),
    ("Al actualizar", "¿Actualizar visuales es igual que volver a consultar la fuente?"),
    ("Al concluir", "¿Qué decisión concreta puede tomar la gerencia con este hallazgo?"),
]
qtable = doc.add_table(rows=1, cols=2)
qtable.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, header_text in enumerate(["Momento", "Pregunta para guiar sin dar la respuesta"]):
    shade(qtable.rows[0].cells[i], NAVY)
    set_cell_text(qtable.rows[0].cells[i], header_text, color=WHITE, bold=True, size=9)
for idx, row_data in enumerate(questions):
    cells = qtable.add_row().cells
    for i, value in enumerate(row_data):
        shade(cells[i], WHITE if idx % 2 == 0 else PALE)
        set_cell_text(cells[i], value, bold=i == 0, size=8.8)

add_heading(doc, "Errores frecuentes y respuesta rápida", 2)
errors = [
    ("Power BI muestra Record", "Entrar en data; si es lista, convertir en tabla y luego expandir Column1."),
    ("No aparece Web API", "Buscar Web. Si tampoco aparece, usar la ruta alternativa con Excel."),
    ("Credenciales", "Elegir Anónima y aplicar el nivel a https://api.cystems.ec/."),
    ("Importes como texto", "Cambiar tipo con configuración regional que reconozca punto decimal."),
    ("Gráfico muestra Conteo", "Sustituir la columna por la medida Ventas Netas."),
    ("Meses desordenados", "Usar fecha con nivel Mes; no usar solo el nombre del mes sin columna de orden."),
    ("Cambió Sheets pero no el reporte", "Actualizar el modelo semántico, esperar Completado y luego actualizar visuales."),
    ("API devuelve demo", "Revisar health.json y la variable GOOGLE_SHEET_CSV_URL del despliegue."),
]
etable = doc.add_table(rows=1, cols=2)
etable.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, header_text in enumerate(["Síntoma", "Corrección"]):
    shade(etable.rows[0].cells[i], VIOLET)
    set_cell_text(etable.rows[0].cells[i], header_text, color=WHITE, bold=True, size=9)
for idx, row_data in enumerate(errors):
    cells = etable.add_row().cells
    for i, value in enumerate(row_data):
        shade(cells[i], WHITE if idx % 2 == 0 else PALE)
        set_cell_text(cells[i], value, bold=i == 0, size=8.6)

doc.add_page_break()
add_heading(doc, "Anexo A · Código M de respaldo", 1)
add_body(doc, "Utilícelo únicamente si Power Query Online permite abrir el Editor avanzado. La consulta extrae la lista data y asigna los tipos principales.")
code = '''let
    Origen = Json.Document(
        Web.Contents("https://api.cystems.ec/api/v1/ventas.json?refresh=1")
    ),
    Registros = Origen[data],
    Tabla = Table.FromRecords(Registros),
    Tipos = Table.TransformColumnTypes(
        Tabla,
        {
            {"fecha", type date},
            {"anio", Int64.Type},
            {"cantidad", Int64.Type},
            {"precio_unitario", Currency.Type},
            {"descuento_pct", Percentage.Type},
            {"venta_neta", Currency.Type},
            {"costo_total", Currency.Type},
            {"margen", Currency.Type},
            {"margen_pct", Percentage.Type},
            {"dias_cobro", Int64.Type}
        }
    )
in
    Tipos'''
ctable = doc.add_table(rows=1, cols=1)
ctable.alignment = WD_TABLE_ALIGNMENT.CENTER
shade(ctable.cell(0, 0), NAVY)
ctable.cell(0, 0).text = ""
cp = ctable.cell(0, 0).paragraphs[0]
cp.paragraph_format.space_after = Pt(0)
cr = cp.add_run(code)
cr.font.name = "Consolas"
cr.font.size = Pt(8)
cr.font.color.rgb = RGBColor.from_string("E2E8F0")
cell_margins(ctable.cell(0, 0), 160, 180, 160, 180)

add_heading(doc, "Anexo B · Lista de comprobación del docente", 1)
for item in [
    "□ Portal visible y rutas funcionando.",
    "□ health.json responde ok: true.",
    "□ API indica source: google-sheets.",
    "□ Hoja Ventas_En_Vivo abierta y valores originales registrados.",
    "□ Estudiantes pueden crear un modelo semántico en navegador.",
    "□ Totales coinciden con la hoja Control.",
    "□ Actualización de prueba realizada y luego revertida.",
    "□ Reportes guardados con nombre identificable.",
]:
    add_bullet(doc, item)

add_callout(doc, "Cierre pedagógico", "No cierre preguntando únicamente si entendieron. Pida que un estudiante explique el flujo completo y que otro justifique una decisión con un indicador del reporte.", color=TEAL, fill="ECFEFF")

doc.core_properties.title = "Guion de práctica API CYSTEMS y Power BI en navegador"
doc.core_properties.subject = "Proyecto integrador 4.3"
doc.core_properties.author = "CYSTEMS"
doc.core_properties.keywords = "Power BI, API, Google Sheets, DAX, reporting contable"
doc.save(OUT)
print(OUT)
