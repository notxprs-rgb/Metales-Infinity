# -*- coding: utf-8 -*-
"""
Genera la propuesta comercial (Word .docx) para Gael - Metales Infinity.
Marca: Aimarktech. Paleta y tipografías tomadas de soyaimarktech.com.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ------------------------------------------------------------------ paleta
BLUE = "0A74DA"; CYAN = "00C2FF"; YELLOW = "FFCE00"; GREEN = "28A745"
INK = "0E1B2C"; INK_SOFT = "46566B"; BG_SOFT = "F4F8FD"; DARK = "071426"
LINE = "E3EBF4"; WHITE = "FFFFFF"; WHATS = "25D366"
HEAD = "Montserrat"; BODY = "Lato"

# ------------------------------------------------------------------ helpers
def _set_font(run, font=BODY, size=10.5, color=INK, bold=False, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rfonts.set(qn(a), font)

def para(doc_or_cell, text="", size=10.5, color=INK, bold=False, italic=False,
         font=BODY, align=None, space_after=6, space_before=0, line=1.15):
    p = doc_or_cell.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after); pf.space_before = Pt(space_before)
    pf.line_spacing = line
    if text:
        r = p.add_run(text); _set_font(r, font, size, color, bold, italic)
    return p

def runs(p, parts):
    """parts: list of (text, dict-kwargs)"""
    for text, kw in parts:
        r = p.add_run(text); _set_font(r, **kw)
    return p

def bottom_border(p, color=BLUE, sz=14):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), str(sz))
    b.set(qn('w:space'), '4'); b.set(qn('w:color'), color)
    pbdr.append(b); pPr.append(pbdr)

def cell_bg(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('bottom', bottom), ('start', left), ('end', right),
                     ('left', left), ('right', right)):
        e = OxmlElement('w:' + tag); e.set(qn('w:w'), str(val)); e.set(qn('w:type'), 'dxa'); m.append(e)
    tcPr.append(m)

def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge); e.set(qn('w:val'), 'none'); borders.append(e)
    tblPr.append(borders)

def grid_borders(table, color=LINE, sz=6):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge); e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), str(sz)); e.set(qn('w:space'), '0'); e.set(qn('w:color'), color)
        borders.append(e)
    tblPr.append(borders)

def top_accent(cell, color=BLUE, sz=24):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    t = OxmlElement('w:top'); t.set(qn('w:val'), 'single'); t.set(qn('w:sz'), str(sz))
    t.set(qn('w:space'), '0'); t.set(qn('w:color'), color); borders.append(t)
    tcPr.append(borders)

def section_title(doc, number, text):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(16); pf.space_after = Pt(8); pf.keep_with_next = True
    r = p.add_run(f"{number}  "); _set_font(r, HEAD, 14, CYAN, bold=True)
    r = p.add_run(text); _set_font(r, HEAD, 14, BLUE, bold=True)
    bottom_border(p, LINE, 10)
    return p

def bullet(doc, label, text):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.left_indent = Cm(0.4); p.paragraph_format.line_spacing = 1.12
    r = p.add_run("● "); _set_font(r, BODY, 10.5, BLUE, bold=True)
    if label:
        r = p.add_run(label + " "); _set_font(r, BODY, 10.5, INK, bold=True)
    r = p.add_run(text); _set_font(r, BODY, 10.5, INK_SOFT)
    return p

# ------------------------------------------------------------------ documento
doc = Document()
st = doc.styles['Normal']; st.font.name = BODY; st.font.size = Pt(10.5)
st.font.color.rgb = RGBColor.from_string(INK)
st.element.rPr.rFonts.set(qn('w:eastAsia'), BODY)

sec = doc.sections[0]
sec.top_margin = Cm(1.9); sec.bottom_margin = Cm(1.9)
sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
CONTENT_W = sec.page_width - sec.left_margin - sec.right_margin

# footer
foot = sec.footer.paragraphs[0]; foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = foot.add_run("Aimarktech  ·  Propuesta para Metales Infinity  ·  Confidencial")
_set_font(r, BODY, 8, INK_SOFT)

# ============================================================== PORTADA
cover = doc.add_table(rows=1, cols=1); cover.alignment = WD_TABLE_ALIGNMENT.CENTER
cover.columns[0].width = CONTENT_W
c = cover.cell(0, 0); c.width = CONTENT_W
cell_bg(c, DARK); cell_margins(c, top=360, bottom=360, left=360, right=360)
c.paragraphs[0].text = ""
p = c.paragraphs[0]
r = p.add_run("AIMARKTECH"); _set_font(r, HEAD, 13, CYAN, bold=True)
r = p.add_run("   ·   Sistemas digitales e IA para PyMEs"); _set_font(r, BODY, 9.5, "B9C6D6")
para(c, "Propuesta de tienda en línea", size=26, color=WHITE, bold=True,
     font=HEAD, space_before=10, space_after=2)
para(c, "para Metales Infinity", size=26, color=YELLOW, bold=True, font=HEAD, space_after=10)
para(c, "Tu propia tienda bilingüe (español / inglés) para vender con orden, "
        "cobrar con anticipo y depender menos del día a día.",
     size=11.5, color="D7E4F0", space_after=4)

meta = doc.add_table(rows=1, cols=1); meta.columns[0].width = CONTENT_W
mc = meta.cell(0, 0); mc.width = CONTENT_W
cell_bg(mc, BG_SOFT); cell_margins(mc, 200, 200, 240, 240); top_accent(mc, BLUE, 24)
mc.paragraphs[0].text = ""
runs(mc.paragraphs[0], [("Preparado para:  ", dict(bold=True, size=10.5)),
                        ("Gael — Metales Infinity · San Juan del Río, Qro.", dict(size=10.5, color=INK_SOFT))])
runs(para(mc, space_after=3), [("Elaborado por:  ", dict(bold=True, size=10.5)),
                        ("José Antonio Aguilar · Aimarktech", dict(size=10.5, color=INK_SOFT))])
runs(para(mc, space_after=0), [("Fecha:  ", dict(bold=True, size=10.5)),
                        ("Septiembre 2026", dict(size=10.5, color=INK_SOFT))])

para(doc, "Diagnóstico + propuesta. Sin promesas mágicas: el aumento de ventas se mide "
          "después de lanzar, con tus números reales.", size=9.5, color=INK_SOFT,
     italic=True, space_before=10, align=WD_ALIGN_PARAGRAPH.CENTER)

# ============================================================== 1. EN CORTO
section_title(doc, "1", "En corto")
para(doc, "Metales Infinity fabrica sellos y figuras vaqueras personalizadas de calidad, "
          "y hoy vende principalmente por Facebook y WhatsApp. Esta propuesta plantea el "
          "siguiente paso: tu propia tienda en línea, bilingüe (español e inglés), donde tus "
          "clientes vean productos y precios, puedan pedir y —cuando lo decidas— pagar en "
          "línea con un anticipo.", color=INK_SOFT)
para(doc, "El objetivo no es prometerte un número mágico de ventas, sino darte orden: "
          "menos tiempo contestando lo mismo, menos riesgo de cancelaciones y una vitrina "
          "profesional que trabaje por ti las 24 horas, también para el mercado latino de "
          "Estados Unidos.", color=INK_SOFT)

# ============================================================== 2. DIAGNÓSTICO
section_title(doc, "2", "Diagnóstico breve (lo esencial)")
para(doc, "Cuatro puntos concretos que hoy frenan el crecimiento y que esta propuesta busca "
          "resolver:", color=INK_SOFT, space_after=8)
bullet(doc, "Todo depende de ti.", "Muchas horas al día contestando mensajes y mandando fotos una por una; si tú no estás, la venta se detiene.")
bullet(doc, "Sin catálogo con precios visibles.", "Cada venta es manual y lenta; el cliente no puede explorar a su ritmo.")
bullet(doc, "Sin anticipo.", "Fabricas primero y cobras al final: si cancelan, tú absorbes el material y el tiempo invertido.")
bullet(doc, "Solo Facebook y WhatsApp.", "No tienes tienda propia; competidores como Sellare ya venden con tienda y precios visibles, y te toman ventaja en lo digital.")

# callout mercado
box = doc.add_table(rows=1, cols=1); box.columns[0].width = CONTENT_W
bc = box.cell(0, 0); bc.width = CONTENT_W
cell_bg(bc, BG_SOFT); cell_margins(bc, 160, 160, 220, 220); top_accent(bc, GREEN, 20)
bc.paragraphs[0].text = ""
runs(bc.paragraphs[0], [("A tu favor.  ", dict(bold=True, color=GREEN)),
    ("En Querétaro casi no tienes competencia directa en tu nicho, y ya vendes a Estados "
     "Unidos. Una tienda bilingüe te posiciona antes que los competidores, que siguen "
     "vendiendo solo por chat.", dict(color=INK_SOFT))])

# nota fiscal
box = doc.add_table(rows=1, cols=1); box.columns[0].width = CONTENT_W
bc = box.cell(0, 0); bc.width = CONTENT_W
cell_bg(bc, BG_SOFT); cell_margins(bc, 160, 160, 220, 220); top_accent(bc, YELLOW, 20)
bc.paragraphs[0].text = ""
runs(bc.paragraphs[0], [("Sobre lo fiscal (opcional y a tu ritmo).  ", dict(bold=True, color=INK)),
    ("Darte de alta en el SAT no es necesario para arrancar tu tienda en México ni para "
     "tener el sitio en inglés. Solo se vuelve útil más adelante si quieres facturar "
     "formalmente o cobrar y retirar en dólares. Lo dejamos como referencia, no como presión.",
     dict(color=INK_SOFT))])

# ============================================================== 3. PROPUESTA
section_title(doc, "3", "La propuesta: tu propia tienda en línea")
para(doc, "Te proponemos una tienda propia (no rentada): el sitio es tuyo, en tu dominio, "
          "con el código respaldado en GitHub y alojado en Cloudflare. Bilingüe desde el "
          "inicio, con precios visibles, catálogo, pedidos personalizados, reseñas de "
          "clientes y botón de WhatsApp. Para arrancar, dos caminos:", color=INK_SOFT, space_after=10)

opts = doc.add_table(rows=1, cols=2)
opts.alignment = WD_TABLE_ALIGNMENT.CENTER
half = int(CONTENT_W / 2)
for i in range(2):
    opts.columns[i].width = half; opts.cell(0, i).width = half
grid_borders(opts, WHITE, 8)  # espacio visual

o1 = opts.cell(0, 0); cell_bg(o1, BG_SOFT); cell_margins(o1, 200, 200, 220, 220); top_accent(o1, INK_SOFT, 20)
o1.paragraphs[0].text = ""
para(o1, "OPCIÓN 1", size=9, color=INK_SOFT, bold=True, font=HEAD, space_after=1)
para(o1, "Catálogo + compra por WhatsApp", size=12, color=INK, bold=True, font=HEAD, space_after=2)
para(o1, "Ligera y rápida", size=9.5, color=INK_SOFT, italic=True, space_after=6)
para(o1, "Tu catálogo con fotos, medidas y precios; el cliente pide con un botón que abre "
         "WhatsApp con el mensaje ya escrito.", size=10, color=INK_SOFT, space_after=6)
runs(para(o1, space_after=3), [("A favor:  ", dict(bold=True, size=10, color=GREEN)),
     ("rápida y económica; conservas tu trato personal; ideal para personalizados.", dict(size=10, color=INK_SOFT))])
runs(para(o1, space_after=0), [("Límite:  ", dict(bold=True, size=10, color="C62F3B")),
     ("no cobra en línea (el pago sigue como hoy); no es autoservicio completo.", dict(size=10, color=INK_SOFT))])

o2 = opts.cell(0, 1); cell_bg(o2, BG_SOFT); cell_margins(o2, 200, 200, 220, 220); top_accent(o2, BLUE, 26)
o2.paragraphs[0].text = ""
runs(o2.paragraphs[0], [("OPCIÓN 2   ", dict(size=9, color=BLUE, bold=True, font=HEAD)),
     ("★ RECOMENDADA", dict(size=8.5, color=YELLOW, bold=True, font=HEAD))])
para(o2, "Tienda con carrito y pago en línea", size=12, color=INK, bold=True, font=HEAD, space_after=2)
para(o2, "Completa", size=9.5, color=INK_SOFT, italic=True, space_after=6)
para(o2, "Además del catálogo: carrito, pago en línea (tarjeta / Mercado Pago), anticipo y un "
         "flujo especial para personalizados (cotización → anticipo → aprobación → "
         "fabricación → liquidación). Incluye un panel para que tú administres productos, "
         "precios y pedidos.", size=10, color=INK_SOFT, space_after=6)
runs(para(o2, space_after=3), [("A favor:  ", dict(bold=True, size=10, color=GREEN)),
     ("vende sola 24/7, cobra anticipo, ordena pedidos y saldos, y queda lista para EE.UU.", dict(size=10, color=INK_SOFT))])
runs(para(o2, space_after=0), [("Necesita:  ", dict(bold=True, size=10, color="C62F3B")),
     ("tener claros tus precios y el medio de pago.", dict(size=10, color=INK_SOFT))])

para(doc, "", space_after=2)
box = doc.add_table(rows=1, cols=1); box.columns[0].width = CONTENT_W
bc = box.cell(0, 0); bc.width = CONTENT_W
cell_bg(bc, "EAF2FC"); cell_margins(bc, 160, 160, 220, 220); top_accent(bc, BLUE, 20)
bc.paragraphs[0].text = ""
runs(bc.paragraphs[0], [("Nuestra recomendación.  ", dict(bold=True, color=BLUE)),
    ("Avanzar hacia la Opción 2 en su versión propia y bilingüe. Podemos lanzar rápido una "
     "primera versión (catálogo + compra) e ir sumando el pago en línea y el panel conforme "
     "lo validemos con tus clientes.", dict(color=INK_SOFT))])

# ============================================================== 4. PLAN DE TRABAJO
section_title(doc, "4", "Plan de trabajo (por etapas, no por fechas prometidas)")
plan = [
    ("Etapa 1 · Definición",
     "Acordamos catálogo inicial, precios, reglas de personalización y anticipo, medio de pago y envíos."),
    ("Etapa 2 · Construcción y piloto",
     "Tienda bilingüe funcionando; tú administras productos y probamos los dos recorridos de compra (catálogo y personalizado)."),
    ("Etapa 3 · Operación y mejora",
     "Medimos pedidos, tiempos de atención, anticipos y entregas; ajustamos lo que haga falta."),
]
t = doc.add_table(rows=len(plan) + 1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
grid_borders(t, LINE, 6)
w1 = int(CONTENT_W * 0.32); w2 = CONTENT_W - w1
hdr = t.rows[0].cells
for j, txt in enumerate(("Etapa", "Qué hacemos")):
    hdr[j].width = (w1, w2)[j]; cell_bg(hdr[j], BLUE); cell_margins(hdr[j])
    hdr[j].paragraphs[0].text = ""
    r = hdr[j].paragraphs[0].add_run(txt); _set_font(r, HEAD, 10, WHITE, bold=True)
for i, (etapa, desc) in enumerate(plan, start=1):
    a, b = t.rows[i].cells; a.width = w1; b.width = w2
    fill = WHITE if i % 2 else BG_SOFT
    for cc in (a, b):
        cell_bg(cc, fill); cell_margins(cc)
    a.paragraphs[0].text = ""; r = a.paragraphs[0].add_run(etapa); _set_font(r, BODY, 10, INK, bold=True)
    b.paragraphs[0].text = ""; r = b.paragraphs[0].add_run(desc); _set_font(r, BODY, 10, INK_SOFT)
para(doc, "", space_after=2)
runs(para(doc), [("Cómo medimos el avance (sin humo):  ", dict(bold=True, color=INK, size=10)),
    ("cotizaciones que terminan en pago, minutos de atención por pedido, entregas a tiempo y "
     "reseñas conseguidas. No prometemos porcentajes de venta antes de tener datos reales.",
     dict(color=INK_SOFT, size=10))])

# ============================================================== 5. QUÉ NECESITAMOS
section_title(doc, "5", "Lo que necesitamos de ti (5 preguntas)")
para(doc, "Para hacerte una propuesta con números reales —y no inventados— necesitamos saber:",
     color=INK_SOFT, space_after=8)
qs = [
    "¿Cuántas ventas tienes al mes, en promedio, en los últimos meses? (no el mejor mes, el promedio real).",
    "¿Qué parte de tus ventas es producto de catálogo (precio fijo) y qué parte es 100% personalizado?",
    "Precios y, de ser posible, el margen de tus principales productos (por ejemplo, juegos de sellos vs. pines).",
    "¿Cómo cobras hoy y cómo te gustaría cobrar (transferencia, tarjeta, Mercado Pago)? ¿Aplicarías anticipo?",
    "¿Quién va a administrar la tienda en el día a día (tú o alguien de tu equipo)?",
]
for i, q in enumerate(qs, 1):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(5); p.paragraph_format.left_indent = Cm(0.2)
    r = p.add_run(f"{i}.  "); _set_font(r, BODY, 10.5, BLUE, bold=True)
    r = p.add_run(q); _set_font(r, BODY, 10.5, INK_SOFT)

# ============================================================== 6. INVERSIÓN
section_title(doc, "6", "Inversión (con transparencia)")
para(doc, "Separamos los costos para que veas claramente en qué se va cada peso. El precio de "
          "Aimarktech cubre nuestro trabajo; el dominio, el alojamiento, las comisiones de pago "
          "y la publicidad se cotizan y pagan aparte.", color=INK_SOFT, space_after=8)
inv = [
    ("Servicios Aimarktech (mensual)",
     "Acompañamiento, implementación y ajustes. Planes desde $2,500 (Esencial); $3,500 Crecimiento (recomendado)."),
    ("Desarrollo inicial de la tienda (una vez)",
     "Se cotiza al terminar la Etapa 1, cuando conozcamos el catálogo y el alcance real."),
    ("Servicios externos",
     "Dominio (~$200–400 MXN/año), alojamiento en Cloudflare (desde $0) y correo."),
    ("Comisiones de pago (si eliges Opción 2)",
     "≈ 3.6% + $3 MXN + IVA por transacción. Crecen con las ventas; las cobra la pasarela, no Aimarktech."),
    ("Pauta publicitaria (opcional)",
     "Ajustable mes con mes; se invierte 100% en la plataforma de anuncios."),
]
t = doc.add_table(rows=len(inv) + 1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
grid_borders(t, LINE, 6)
w1 = int(CONTENT_W * 0.38); w2 = CONTENT_W - w1
hdr = t.rows[0].cells
for j, txt in enumerate(("Concepto", "Detalle")):
    hdr[j].width = (w1, w2)[j]; cell_bg(hdr[j], BLUE); cell_margins(hdr[j])
    hdr[j].paragraphs[0].text = ""
    r = hdr[j].paragraphs[0].add_run(txt); _set_font(r, HEAD, 10, WHITE, bold=True)
for i, (con, det) in enumerate(inv, start=1):
    a, b = t.rows[i].cells; a.width = w1; b.width = w2
    fill = WHITE if i % 2 else BG_SOFT
    for cc in (a, b):
        cell_bg(cc, fill); cell_margins(cc)
    a.paragraphs[0].text = ""; r = a.paragraphs[0].add_run(con); _set_font(r, BODY, 10, INK, bold=True)
    b.paragraphs[0].text = ""; r = b.paragraphs[0].add_run(det); _set_font(r, BODY, 10, INK_SOFT)

para(doc, "", space_after=2)
runs(para(doc), [("¿Por qué tienda propia y no Shopify?  ", dict(bold=True, color=INK, size=10)),
    ("Una plataforma rentada cuesta poco al mes, pero te ata a sus reglas y comisiones. Tu "
     "tienda propia es tuya, se adapta a tu forma de vender (personalizados y anticipos) y "
     "crece contigo. Si prefieres arrancar en una plataforma rentada, también lo evaluamos "
     "contigo.", dict(color=INK_SOFT, size=10))])
runs(para(doc), [("Nota honesta:  ", dict(bold=True, color="C62F3B", size=10)),
    ("no prometemos un retorno garantizado. Cualquier proyección de ventas la haremos "
     "contigo, con tus números reales, después de la Etapa 1.", dict(color=INK_SOFT, size=10))])

# ============================================================== 7. POR QUÉ AIMARKTECH
section_title(doc, "7", "Por qué Aimarktech")
for label, txt in [
    ("Diagnóstico antes de recomendar.", "Revisamos el negocio completo antes de proponerte publicidad, software o automatizaciones."),
    ("Una visión conectada.", "Integramos procesos, tecnología, IA y marketing en un mismo plan ejecutable."),
    ("Acompañamiento real.", "No entregamos un documento y desaparecemos: apoyamos la implementación y la transferencia a tu equipo."),
    ("Resultados medibles, sin humo.", "Hablamos claro de límites, dependencias y costos. +28 años de experiencia tecnológica."),
]:
    bullet(doc, label, txt)

# ============================================================== 8. SIGUIENTES PASOS
section_title(doc, "8", "Siguientes pasos")
for i, txt in enumerate([
    "Elige con cuál opción arrancamos (1 o 2).",
    "Una llamada corta para responder las 5 preguntas.",
    "Iniciamos la Etapa 1 (definición) y te entregamos el plan detallado.",
], 1):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(5); p.paragraph_format.left_indent = Cm(0.2)
    r = p.add_run(f"{i}.  "); _set_font(r, BODY, 10.5, BLUE, bold=True)
    r = p.add_run(txt); _set_font(r, BODY, 10.5, INK_SOFT)

# contacto
contact = doc.add_table(rows=1, cols=1); contact.columns[0].width = CONTENT_W
cc = contact.cell(0, 0); cc.width = CONTENT_W
cell_bg(cc, DARK); cell_margins(cc, 240, 240, 300, 300)
cc.paragraphs[0].text = ""
para(cc, "José Antonio Aguilar · Aimarktech", size=12.5, color=WHITE, bold=True, font=HEAD, space_after=4)
runs(para(cc, space_after=2), [("WhatsApp:  ", dict(bold=True, color=CYAN)), ("+52 56 3963 7740", dict(color="D7E4F0"))])
runs(para(cc, space_after=2), [("Oficina:  ", dict(bold=True, color=CYAN)), ("55 8967 7250", dict(color="D7E4F0"))])
runs(para(cc, space_after=2), [("Correo:  ", dict(bold=True, color=CYAN)), ("jantonioab@gmail.com", dict(color="D7E4F0"))])
runs(para(cc, space_after=2), [("Web:  ", dict(bold=True, color=CYAN)), ("https://soyaimarktech.com", dict(color="D7E4F0"))])
para(cc, "Base en Valle de Chalco, con alcance en todo México.", size=9.5, color="B9C6D6", italic=True, space_before=4)

para(doc, "Empecemos con información, no con presión: primero ordenamos, luego crecemos.",
     size=10.5, color=BLUE, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     space_before=12)

import os
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Propuesta_Metales_Infinity_Aimarktech.docx")
doc.save(out)
print("OK ->", out)
