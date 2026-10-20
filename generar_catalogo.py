# -*- coding: utf-8 -*-
"""
Generador de catálogo PDF - Distribuidora Bellingieri
- Incluye Modo de uso y Activos principales
- Omite el SKU
- Cabecera negra, marcos dorados gruesos y fusión automática con portada.
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
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from PIL import Image as PILImage

# Importamos la herramienta para unir PDFs
from pypdf import PdfWriter, PdfReader

# ============================================================
# CONFIGURACIÓN — EDITA AQUÍ
# ============================================================
EXCEL_PATH     = "catalogo.xlsx"
IMAGENES_DIR   = "imagenes"
LOGO_PATH      = "logo.png"  
PORTADA_PDF    = "portada.pdf" 
OUTPUT_PDF     = "Catalogo_Temporal.pdf" 
OUTPUT_FINAL   = "Catalogo_Distribuidora_Bellingieri_Final.pdf" 

COLOR_PRIMARIO   = colors.HexColor("#1F3864")
COLOR_TEXTO      = colors.HexColor("#222222")
COLOR_PRECIO     = colors.HexColor("#C00000")
COLOR_DORADO     = colors.HexColor("#C5A059") 

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
style_activos = ParagraphStyle(
    "activos", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=8, leading=10, 
    textColor=COLOR_DORADO, spaceAfter=3,
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
style_precio = ParagraphStyle(
    "precio", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=11, leading=13, textColor=COLOR_PRECIO,
)
style_cat_titulo = ParagraphStyle(
    "cat_titulo", parent=styles["Heading1"],
    fontName="Helvetica-Bold", fontSize=26, leading=30,
    textColor=COLOR_PRIMARIO, alignment=TA_LEFT,
)
style_cat_num = ParagraphStyle(
    "cat_num", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=14, leading=16,
    textColor=colors.HexColor("#888888"),
)

# ============================================================
# CARGA DE DATOS
# ============================================================
def cargar_datos(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No se encontró el archivo Excel: {path}")
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    df = df.fillna("")
    
    # Convertimos a string las columnas clave para evitar errores
    for col in ["Nombre imagen", "Modo de uso", "Activos principales"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    df["_orden"] = range(len(df))
    return df

def agrupar_por_categoria(df):
    grupos = []
    vistos = []
    for cat in df["Categoría"]:
        cat_limpia = str(cat).strip()
        if cat_limpia and cat_limpia not in vistos:
            vistos.append(cat_limpia)
    for cat in vistos:
        subset = df[df["Categoría"].astype(str).str.strip() == cat].sort_values("_orden")
        grupos.append((cat, subset.to_dict("records")))
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
    
    ancho_img = (ancho_celda * 0.30) - 0.5*cm # Reducimos un poco el ancho de la imagen para dar más espacio al texto
    alto_img = alto_celda - 0.6*cm 
    img = obtener_imagen_ajustada(img_path, ancho_img, alto_img)

    # --- TEXTOS ---
    nombre = Paragraph(str(row.get("Nombre del producto", "")), style_nombre)
    
    activos_txt = str(row.get("Activos principales", "")).strip()
    activos = Paragraph(f"<b>ACTIVOS:</b> {activos_txt}" if activos_txt and activos_txt.lower() != "nan" else "", style_activos)
    
    desc = Paragraph(str(row.get("Descripción", "")), style_desc)
    
    modo_txt = str(row.get("Modo de uso", "")).strip()
    # Solo mostramos el modo de uso si hay texto
    modo = Paragraph(f"<b>MODO DE USO:</b> {modo_txt}" if modo_txt and modo_txt.lower() != "nan" else "", style_modo)

    precio_txt = str(row.get("Precio", "")).strip()
    precio_str = f"${float(precio_txt):,.2f}" if precio_txt else "Consultar precio"
    precio = Paragraph(precio_str, style_precio)

    # Empaquetamos el texto en una tabla (columna derecha)
    texto_derecha = Table(
        [[nombre], [activos], [desc], [modo], [precio]],
        colWidths=[(ancho_celda * 0.70) - 0.5*cm]
    )
    texto_derecha.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "LEFT"), 
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0), 
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0), 
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))

    # Tabla principal (Imagen + Texto)
    contenido = Table(
        [[img, texto_derecha]], 
        colWidths=[ancho_celda * 0.30, ancho_celda * 0.70], 
        rowHeights=[alto_celda]
    )
    contenido.setStyle(TableStyle([
        ("ALIGN", (0,0), (0,0), "CENTER"), 
        ("VALIGN", (0,0), (0,0), "MIDDLE"), 
        ("ALIGN", (1,0), (1,0), "LEFT"),   
        ("VALIGN", (1,0), (1,0), "MIDDLE"), 
        ("LEFTPADDING", (0,0), (0,0), 12), 
        ("RIGHTPADDING", (0,0), (0,0), 0),
        ("LEFTPADDING", (1,0), (1,0), 10), 
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
    return [grid, PageBreak()]

# ============================================================
# ENCABEZADO (HEADER) Y PIE
# ============================================================
def header_footer(canvas, doc):
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
    grupos = agrupar_por_categoria(df)

    doc = BaseDocTemplate(
        OUTPUT_PDF, pagesize=A4,
        leftMargin=1.0*cm, rightMargin=1.0*cm,
        topMargin=3.5*cm, bottomMargin=1.8*cm, 
        title=f"{EMPRESA} - {SUBTITULO}", author=EMPRESA,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="cuerpo", frames=[frame], onPage=header_footer)])

    story = []
    total_cats = len(grupos)
    for idx, (cat, prods) in enumerate(grupos, start=1):
        story += [Spacer(1, 6*cm), Paragraph(f"{idx:02d} / {total_cats:02d}", style_cat_num), Spacer(1, 0.3*cm),
                  Paragraph(cat, style_cat_titulo), Spacer(1, 0.5*cm), PageBreak()]
        
        for i in range(0, len(prods), 4):
            story += pagina_productos(prods[i:i+4])

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