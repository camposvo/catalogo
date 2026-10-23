# -*- coding: utf-8 -*-
"""
Generador de catálogo PDF - Distribuidora Bellingieri
Uso: python generar_catalogo.py
"""

import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
    Table, TableStyle, Image, PageBreak, NextPageTemplate
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from PIL import Image as PILImage
from pypdf import PdfWriter, PdfReader

# Importamos svglib para leer archivos SVG
from svglib.svglib import svg2rlg

# ============================================================
# CONFIGURACIÓN
# ============================================================
EXCEL_PATH     = "catalogo.xlsx"
IMAGENES_DIR   = "imagenes"
LOGO_PATH      = "logo.png"  
CARRITO_PATH   = "carrito.png"  # <--- Puede ser .svg o .png
PORTADA_PDF    = "portada.pdf" 
OUTPUT_PDF     = "Catalogo_Temporal.pdf" 
OUTPUT_FINAL   = "Catalogo_Distribuidora_Bellingieri_Final.pdf" 

COLOR_PRIMARIO   = colors.HexColor("#1F3864")
COLOR_TEXTO      = colors.HexColor("#222222")
COLOR_PRECIO     = colors.HexColor("#C00000")
COLOR_DORADO     = colors.HexColor("#C5A059") 
COLOR_CELESTE    = colors.HexColor("#00AEEF") 
COLOR_GRIS_TABLA = colors.HexColor("#666666")

EMPRESA   = "Distribuidora Bellingieri"
SUBTITULO = "Catálogo de Productos"

# ============================================================
# ESTILOS
# ============================================================
styles = getSampleStyleSheet()

style_nombre = ParagraphStyle(
    "nombre", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=11.5, leading=13.5, 
    textColor=COLOR_PRIMARIO, spaceAfter=2,
)
style_categoria_producto = ParagraphStyle(
    "categoria_producto", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=8.5, leading=10, 
    textColor=COLOR_DORADO, spaceAfter=3,
)
style_activos = ParagraphStyle(
    "activos", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=8, leading=10, 
    textColor=COLOR_PRIMARIO, spaceAfter=3,
)
style_desc = ParagraphStyle(
    "desc", parent=styles["Normal"],
    fontName="Helvetica", fontSize=8.5, leading=11, 
    textColor=COLOR_TEXTO, spaceAfter=4,
)
style_modo = ParagraphStyle(
    "modo", parent=styles["Normal"],
    fontName="Helvetica", fontSize=7.5, leading=9.5, 
    textColor=colors.HexColor("#444444"), spaceAfter=4,
)

style_marca_titulo = ParagraphStyle(
    "marca_titulo", parent=styles["Heading1"],
    fontName="Helvetica-Bold", fontSize=36, leading=42,
    textColor=COLOR_PRIMARIO, alignment=TA_LEFT, spaceBefore=150,
)

# Estilos para la tabla de precios
style_tabla_icono = ParagraphStyle(
    "tabla_icono", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=11, leading=13, textColor=colors.white,
    alignment=TA_CENTER
)
style_tabla_nombre = ParagraphStyle(
    "tabla_nombre", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=6, leading=7.5, textColor=colors.black,
    alignment=TA_LEFT
)
style_tabla_precio = ParagraphStyle(
    "tabla_precio", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=7, leading=8.5, textColor=COLOR_PRECIO,
    alignment=TA_RIGHT
)

# ============================================================
# CARGA DE DATOS
# ============================================================
def cargar_datos(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No se encontró el archivo Excel: {path}")
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    
    if "Marca" not in df.columns:
        raise ValueError("El Excel debe contener una columna llamada 'Marca'.")
        
    df = df.fillna("")
    
    columnas_a_limpiar = [
        "Nombre imagen", "Modo de uso", "Activos principales", "Marca", "Categoría",
        "Presentacion 1", "Precio 1", "Presentacion 2", "Precio 2", "Presentacion 3", "Precio 3"
    ]
    for col in columnas_a_limpiar:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    df["_orden"] = range(len(df))
    return df

def agrupar_por_marca(df):
    grupos = []
    vistos = []
    for marca in df["Marca"]:
        marca_limpia = str(marca).strip()
        if marca_limpia and marca_limpia not in vistos:
            vistos.append(marca_limpia)
            
    for marca in vistos:
        subset = df[df["Marca"].astype(str).str.strip() == marca].sort_values("_orden")
        grupos.append((marca, subset.to_dict("records")))
    return grupos

# ============================================================
# CELDA DE PRODUCTO
# ============================================================
def _placeholder(w, h):
    t = Table([["Sin imagen"]], colWidths=[w], rowHeights=[h])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#EEEEEE")),
        ("TEXTCOLOR", (0,0), (-1,-1), colors.grey),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica-Oblique"),
    ]))
    return t

def obtener_imagen_ajustada(img_path, max_w, max_h):
    if os.path.isfile(img_path):
        try:
            with PILImage.open(img_path) as im:
                w, h = im.size
            ratio = min(max_w / w, max_h / h)
            img = Image(img_path, width=w*ratio, height=h*ratio)
            img.hAlign = "CENTER"
            return img
        except Exception:
            return _placeholder(max_w, max_h)
    return _placeholder(max_w, max_h)

def celda_producto(row, ancho_celda, alto_celda):
    nombre_imagen = str(row.get("Nombre imagen", "")).strip()
    img_path = os.path.join(IMAGENES_DIR, nombre_imagen) if nombre_imagen and nombre_imagen.lower() != "nan" else ""
    
    ancho_bloque_izq = ancho_celda * 0.33
    ancho_col_derecha = ancho_celda * 0.67
    
    ancho_img_area = ancho_bloque_izq * 0.63
    ancho_tabla_area = ancho_bloque_izq * 0.37
    
    img = obtener_imagen_ajustada(img_path, ancho_img_area - 0.1*cm, alto_celda - 0.6*cm)

    # --- CONSTRUCCIÓN DE LA TABLA DE PRESENTACIONES ---
    filas_cuerpo = []
    for i in range(1, 4):
        col_pres = f"Presentacion {i}"
        col_prec = f"Precio {i}"
        
        pres_txt = str(row.get(col_pres, "")).strip()
        prec_txt = str(row.get(col_prec, "")).strip()
        
        if pres_txt and pres_txt.lower() != "nan":
            try:
                prec_str = f"${float(prec_txt):,.2f}"
            except ValueError:
                prec_str = prec_txt
                
            filas_cuerpo.append([
                Paragraph(pres_txt, style_tabla_nombre),
                Paragraph(prec_str, style_tabla_precio)
            ])
            
    if len(filas_cuerpo) > 0:
        ancho_col_izq = ancho_tabla_area * 0.48
        ancho_col_der = ancho_tabla_area * 0.52

        # 1. EL ADORNO DORADO CON SOPORTE PARA SVG Y PNG
        if os.path.isfile(CARRITO_PATH):
            try:
                if CARRITO_PATH.lower().endswith('.svg'):
                    # Cargar SVG
                    icono = svg2rlg(CARRITO_PATH)
                    if icono:
                        # Escalar el SVG para que mida 14 puntos
                        escala = 14 / max(icono.width, icono.height)
                        icono.scale(escala, escala)
                        icono.hAlign = "CENTER"
                    else:
                        icono = Paragraph("PRECIO", style_tabla_icono)
                else:
                    # Cargar PNG
                    icono = Image(CARRITO_PATH, width=13, height=13)
                    icono.hAlign = "CENTER"
            except Exception:
                icono = Paragraph("PRECIO", style_tabla_icono)
        else:
            icono = Paragraph("PRECIO", style_tabla_icono)

        tabla_adorno = Table([["", icono]], colWidths=[ancho_col_izq, ancho_col_der], rowHeights=[0.6*cm])
        tabla_adorno.setStyle(TableStyle([
            ("BACKGROUND", (1,0), (1,0), COLOR_DORADO),
            ("BOX", (1,0), (1,0), 1, COLOR_DORADO),
            ("ROUNDEDCORNERS", [5, 5, 0, 0]),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("VALIGN", (1,0), (1,0), "MIDDLE"), 
            ("TOPPADDING", (1,0), (1,0), 0),
            ("BOTTOMPADDING", (1,0), (1,0), 0),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ]))

        # 2. LA TABLA DE DATOS
        tabla_cuerpo = Table(filas_cuerpo, colWidths=[ancho_col_izq, ancho_col_der], rowHeights=[0.6*cm] * len(filas_cuerpo))
        tabla_cuerpo.setStyle(TableStyle([
            ("ALIGN", (0,0), (0,-1), "LEFT"),
            ("ALIGN", (1,0), (1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 2),
            ("RIGHTPADDING", (0,0), (-1,-1), 1), 
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            
            ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC")),
            ("LINEBELOW", (0,0), (-1,-2), 0.5, colors.HexColor("#CCCCCC")),
            ("LINEAFTER", (0,0), (0,-1), 0.5, colors.HexColor("#CCCCCC")),
            ("ROUNDEDCORNERS", [0, 0, 5, 5]),
        ]))

        # 3. UNIMOS EL ADORNO Y LA TABLA
        tabla_pres = Table([[tabla_adorno], [tabla_cuerpo]], colWidths=[ancho_tabla_area])
        tabla_pres.setStyle(TableStyle([
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ]))
        
        col_izquierda = Table(
            [[img, tabla_pres]], 
            colWidths=[ancho_img_area, ancho_tabla_area],
            rowHeights=[alto_celda - 0.4*cm]
        )
        col_izquierda.setStyle(TableStyle([
            ("ALIGN", (0,0), (0,0), "CENTER"), 
            ("VALIGN", (0,0), (0,0), "MIDDLE"), 
            ("ALIGN", (1,0), (1,0), "CENTER"), 
            ("VALIGN", (1,0), (1,0), "BOTTOM"), 
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 10), 
        ]))
    else:
        col_izquierda = img

    # --- TEXTOS DE LA DERECHA ---
    nombre = Paragraph(str(row.get("Nombre del producto", "")), style_nombre)
    categoria_txt = str(row.get("Categoría", "")).strip()
    categoria_prod = Paragraph(categoria_txt.upper() if categoria_txt and categoria_txt.lower() != "nan" else "", style_categoria_producto)
    activos_txt = str(row.get("Activos principales", "")).strip()
    activos = Paragraph(f"<b>ACTIVOS:</b> {activos_txt}" if activos_txt and activos_txt.lower() != "nan" else "", style_activos)
    desc = Paragraph(str(row.get("Descripción", "")), style_desc)
    modo_txt = str(row.get("Modo de uso", "")).strip()
    modo = Paragraph(f"<b>MODO DE USO:</b> {modo_txt}" if modo_txt and modo_txt.lower() != "nan" else "", style_modo)

    margen_total = 30 
    ancho_real_texto = ancho_col_derecha - margen_total

    texto_derecha = Table(
        [[nombre], [categoria_prod], [activos], [desc], [modo]],
        colWidths=[ancho_real_texto]
    )
    texto_derecha.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "LEFT"), 
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0), 
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0), 
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))

    contenido = Table(
        [[col_izquierda, texto_derecha]], 
        colWidths=[ancho_bloque_izq, ancho_col_derecha], 
        rowHeights=[alto_celda]
    )
    contenido.setStyle(TableStyle([
        ("ALIGN", (0,0), (0,0), "CENTER"), 
        ("VALIGN", (0,0), (0,0), "MIDDLE"), 
        ("ALIGN", (1,0), (1,0), "LEFT"),   
        ("VALIGN", (1,0), (1,0), "MIDDLE"), 
        ("LEFTPADDING", (0,0), (0,0), 10), 
        ("RIGHTPADDING", (0,0), (0,0), 0),
        ("LEFTPADDING", (1,0), (1,0), 15), 
        ("RIGHTPADDING", (1,0), (1,0), 15), 
        ("TOPPADDING", (0,0), (-1,-1), 10), 
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("BOX", (0,0), (-1,-1), 3.5, COLOR_DORADO), 
        ("ROUNDEDCORNERS", [12, 12, 12, 12]),
    ]))
    return contenido

# ============================================================
# PÁGINA DE PRODUCTOS
# ============================================================
def pagina_productos(productos):
    ancho_util = A4[0] - 2.0*cm 
    alto_util  = A4[1] - 5.5*cm 
    gap = 0.3*cm 
    alto_celda = (alto_util - (gap * 3)) / 4

    celdas = [celda_producto(p, ancho_util, alto_celda) for p in productos]
    while len(celdas) < 4: celdas.append("")

    grid = Table(
        [[celdas[0]], [Spacer(1, gap)], [celdas[1]], [Spacer(1, gap)],
         [celdas[2]], [Spacer(1, gap)], [celdas[3]]],
        colWidths=[ancho_util],
        rowHeights=[alto_celda, gap, alto_celda, gap, alto_celda, gap, alto_celda]
    )
    grid.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), 
        ("LEFTPADDING", (0,0), (-1,-1), 0), 
        ("RIGHTPADDING", (0,0), (-1,-1), 0)
    ]))
    return grid

# ============================================================
# ENCABEZADOS
# ============================================================
def header_limpio(canvas, doc):
    pass

def header_footer_productos(canvas, doc):
    canvas.saveState()
    
    alto_header = 2.5*cm
    canvas.setFillColor(colors.black)
    canvas.rect(0, A4[1] - alto_header, A4[0], alto_header, fill=1, stroke=0)
    
    if os.path.exists(LOGO_PATH):
        try:
            with PILImage.open(LOGO_PATH) as im:
                w, h = im.size
            ratio = min(6*cm / w, 1.8*cm / h)
            logo_w, logo_h = w * ratio, h * ratio
            canvas.drawImage(LOGO_PATH, A4[0] - logo_w - 1.5*cm, A4[1] - (alto_header/2) - (logo_h/2), width=logo_w, height=logo_h, mask='auto')
        except Exception: pass

    marca_actual = getattr(doc, '_marca_actual', '')
    
    if marca_actual:
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 20)
        canvas.drawString(1.5*cm, A4[1] - 1.5*cm, marca_actual.upper())

    canvas.setStrokeColor(COLOR_DORADO)
    canvas.setLineWidth(2)
    canvas.line(0, A4[1] - alto_header, A4[0], A4[1] - alto_header)

    canvas.setStrokeColor(colors.lightgrey)
    canvas.setLineWidth(0.5)
    canvas.line(1.5*cm, 1.3*cm, A4[0]-1.5*cm, 1.3*cm)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.5*cm, 0.9*cm, EMPRESA)
    canvas.drawRightString(A4[0]-1.5*cm, 0.9*cm, f"Página {doc.page}")
    canvas.restoreState()

# ============================================================
# GENERACIÓN DEL PDF Y FUSIÓN
# ============================================================
def generar_y_fusionar():
    df = cargar_datos(EXCEL_PATH)
    grupos = agrupar_por_marca(df)

    doc = BaseDocTemplate(
        OUTPUT_PDF, pagesize=A4,
        leftMargin=1.0*cm, rightMargin=1.0*cm,
        topMargin=3.5*cm, bottomMargin=1.8*cm, 
        title=f"{EMPRESA} - {SUBTITULO}", author=EMPRESA,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    
    doc.addPageTemplates([
        PageTemplate(id="limpia", frames=[frame], onPage=header_limpio)
    ])

    story = []
    total_grupos = len(grupos)
    
    story.append(NextPageTemplate("limpia"))

    for idx, (marca, prods) in enumerate(grupos, start=1):
        categoria_ejemplo = str(prods[0].get("Categoría", "")).strip()
        nombre_plantilla_prods = f"prods_{idx}"
        
        def crear_header_con_datos(m, c):
            def custom_header(canvas, doc):
                doc._marca_actual = m
                doc._categoria_actual = c
                header_footer_productos(canvas, doc)
            return custom_header

        doc.addPageTemplates([
            PageTemplate(id=nombre_plantilla_prods, frames=[frame], onPage=crear_header_con_datos(marca, categoria_ejemplo))
        ])
        
        # Separador
        story.append(Paragraph(marca.upper(), style_marca_titulo))
        story.append(NextPageTemplate(nombre_plantilla_prods))
        story.append(PageBreak())

        # Productos
        for i in range(0, len(prods), 4):
            grid = pagina_productos(prods[i:i+4])
            story.append(grid)
            if i + 4 < len(prods):
                story.append(PageBreak())
            
        # Siguiente Separador
        story.append(NextPageTemplate("limpia"))
        if idx < total_grupos:
            story.append(PageBreak())

    doc.build(story)
    print(f"📄 Cuerpo del catálogo generado temporalmente: {OUTPUT_PDF}")

    if os.path.exists(PORTADA_PDF):
        print(f"🖇️  Añadiendo la portada '{PORTADA_PDF}' al inicio...")
        writer = PdfWriter()
        reader_portada = PdfReader(PORTADA_PDF)
        for page in reader_portada.pages:
            writer.add_page(page)
        reader_catalogo = PdfReader(OUTPUT_PDF)
        for page in reader_catalogo.pages:
            writer.add_page(page)
        with open(OUTPUT_FINAL, "wb") as f:
            writer.write(f)
        print(f"✅ ¡Catálogo Final generado exitosamente!: {OUTPUT_FINAL}")
        os.remove(OUTPUT_PDF)
    else:
        print(f"⚠️ No se encontró '{PORTADA_PDF}'. El catálogo se generó sin portada.")
        os.rename(OUTPUT_PDF, OUTPUT_FINAL)
        print(f"✅ Catálogo Final (sin portada): {OUTPUT_FINAL}")

if __name__ == "__main__":
    generar_y_fusionar()