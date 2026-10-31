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

# ============================================================
# CONFIGURACIÓN
# ============================================================
EXCEL_PATH     = "catalogo.xlsx"
TIPOS_PATH     = "tipos.xlsx"
IMAGENES_DIR   = "imagenes"
LOGO_PATH      = "logo.png"  
PORTADA_PDF    = "portada.pdf" 
OUTPUT_PDF     = "Catalogo_Temporal.pdf" 
OUTPUT_FINAL   = "Catalogo_Distribuidora_Bellingieri_Final.pdf" 

COLOR_PRIMARIO   = colors.HexColor("#1F3864")
COLOR_TEXTO      = colors.HexColor("#222222")
COLOR_PRECIO     = colors.HexColor("#C00000")
COLOR_DORADO     = colors.HexColor("#C5A059") 
COLOR_CELESTE    = colors.HexColor("#00AEEF") 

EMPRESA   = "Distribuidora Bellingieri"
SUBTITULO = "Catálogo de Productos"

# ============================================================
# ESTILOS
# ============================================================
styles = getSampleStyleSheet()

style_nombre = ParagraphStyle(
    "nombre", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=10, leading=12, 
    textColor=COLOR_PRIMARIO, spaceAfter=2,
)
style_categoria_producto = ParagraphStyle(
    "categoria_producto", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=7, leading=8.5, 
    textColor=COLOR_DORADO, spaceAfter=3,
)
style_desc = ParagraphStyle(
    "desc", parent=styles["Normal"],
    fontName="Helvetica", fontSize=7.5, leading=9.5, 
    textColor=COLOR_TEXTO, spaceAfter=4,
)
style_modo = ParagraphStyle(
    "modo", parent=styles["Normal"],
    fontName="Helvetica", fontSize=7, leading=9, 
    textColor=colors.HexColor("#444444"), spaceAfter=4,
)
style_tipo_titulo = ParagraphStyle(
    "tipo_titulo", parent=styles["Heading1"],
    fontName="Helvetica-Bold", fontSize=36, leading=42,
    textColor=COLOR_PRIMARIO, alignment=TA_LEFT,
)

# Estilos para la tabla de precios
style_tabla_nombre = ParagraphStyle(
    "tabla_nombre", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=6, leading=7.5, textColor=colors.black,
    alignment=TA_CENTER
)
style_tabla_precio = ParagraphStyle(
    "tabla_precio", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=7, leading=8.5, textColor=COLOR_PRECIO,
    alignment=TA_CENTER
)

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================
def limpiar_texto(texto):
    if not texto: return ""
    return str(texto).replace("<br>", " ").replace("<br/>", " ").replace("<BR>", " ").replace("<BR/>", " ").strip()

def obtener_color_hex(color_str, default_color=COLOR_PRIMARIO):
    if not color_str or str(color_str).lower() == "nan": return default_color
    color_str = str(color_str).strip()
    if color_str.startswith("#"):
        try: return colors.HexColor(color_str)
        except: return default_color
    try: return getattr(colors, color_str.lower())
    except AttributeError:
        try:
            if len(color_str) == 6: return colors.HexColor(f"#{color_str}")
        except: pass
    return default_color

# ============================================================
# CARGA DE DATOS
# ============================================================
def cargar_datos(path):
    if not os.path.exists(path): raise FileNotFoundError(f"No se encontró: {path}")
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    if "Tipo" not in df.columns: raise ValueError("Falta columna 'Tipo'.")
    df = df.fillna("")
    if "Estatus" in df.columns:
        df["Estatus"] = df["Estatus"].astype(str).str.strip().str.upper()
        df = df[df["Estatus"] == "Y"]
    for col in ["Nombre imagen", "Modo de uso", "Activos principales", "Tipo", "Categoría", "Descripción", "Nombre del producto", "Presentacion 1", "Precio 1", "Presentacion 2", "Precio 2", "Presentacion 3", "Precio 3"]:
        if col in df.columns: df[col] = df[col].apply(limpiar_texto)
    df["_orden"] = range(len(df))
    return df

def cargar_tipos(path):
    if not os.path.exists(path): return {}
    try:
        df_tipos = pd.read_excel(path)
        df_tipos.columns = [c.strip() for c in df_tipos.columns]
        col_tipo = next((c for c in df_tipos.columns if c.lower() == "tipo"), None)
        col_color = next((c for c in df_tipos.columns if c.lower() == "color"), None)
        if not col_tipo or not col_color: return {}
        df_tipos[col_tipo] = df_tipos[col_tipo].astype(str).str.strip()
        df_tipos[col_color] = df_tipos[col_color].astype(str).str.strip()
        return dict(zip(df_tipos[col_tipo], df_tipos[col_color]))
    except: return {}

def agrupar_por_tipo(df):
    grupos, vistos = [], []
    for tipo in df["Tipo"]:
        t_limpio = str(tipo).strip()
        if t_limpio and t_limpio.lower() != "nan" and t_limpio not in vistos:
            vistos.append(t_limpio)
    for tipo in vistos:
        subset = df[df["Tipo"].astype(str).str.strip() == tipo].sort_values("_orden")
        if not subset.empty: grupos.append((tipo, subset.to_dict("records")))
    return grupos

# ============================================================
# CELDA DE PRODUCTO (Diseño para 2x2)
# ============================================================
def _placeholder(w, h):
    t = Table([["Sin imagen"]], colWidths=[w], rowHeights=[h])
    t.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#EEEEEE")), ("TEXTCOLOR", (0,0), (-1,-1), colors.grey), ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("FONTNAME", (0,0), (-1,-1), "Helvetica-Oblique")]))
    return t

def obtener_imagen_ajustada(img_path, max_w, max_h):
    if os.path.isfile(img_path):
        try:
            with PILImage.open(img_path) as im: w, h = im.size
            ratio = min(max_w / w, max_h / h)
            img = Image(img_path, width=w*ratio, height=h*ratio)
            img.hAlign = "CENTER"
            return img
        except: return _placeholder(max_w, max_h)
    return _placeholder(max_w, max_h)


def celda_producto(row, ancho_celda, alto_celda):
    nombre_imagen = str(row.get("Nombre imagen", "")).strip()
    img_path = os.path.join(IMAGENES_DIR, nombre_imagen) if nombre_imagen and nombre_imagen.lower() != "nan" else ""
    
    ancho_bloque_izq = ancho_celda * 0.40
    ancho_col_derecha = ancho_celda * 0.60
    
    alto_img_area = alto_celda * 0.72
    alto_tabla_area = alto_celda * 0.28
    
    img = obtener_imagen_ajustada(img_path, ancho_bloque_izq - 0.4*cm, alto_img_area - 0.4*cm)

    # --- CONSTRUCCIÓN DE LA LISTA DE PRECIOS (SIN TABLA) ---
    lineas_precios = []
    for i in range(1, 4):
        col_pres, col_prec = f"Presentacion {i}", f"Precio {i}"
        pres_txt = str(row.get(col_pres, "")).strip()
        prec_txt = str(row.get(col_prec, "")).strip()
        
        if pres_txt and pres_txt.lower() != "nan":
            prec_str = "Consultar"
            if prec_txt and prec_txt.lower() != "nan":
                try: prec_str = f"${float(prec_txt):,.2f}"
                except: prec_str = prec_txt
            
            lineas_precios.append(f"<b>{pres_txt}</b> - {prec_str}")
            
    if len(lineas_precios) > 0:
        texto_precios = "<br/>".join(lineas_precios)
        bloque_precios = Paragraph(texto_precios, style_tabla_precio)
        
        col_izquierda = Table(
            [[img], [bloque_precios]], 
            colWidths=[ancho_bloque_izq],
            rowHeights=[alto_img_area, alto_tabla_area]
        )
        col_izquierda.setStyle(TableStyle([
            ("ALIGN", (0,0), (-1,-1), "CENTER"), 
            ("VALIGN", (0,0), (0,0), "MIDDLE"),
            ("VALIGN", (0,1), (0,1), "TOP"), 
            ("LEFTPADDING", (0,0), (-1,-1), 0), ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0), ("BOTTOMPADDING", (0,0), (-1,-1), 0), 
        ]))
    else:
        col_izquierda = Table([[img]], colWidths=[ancho_bloque_izq], rowHeights=[alto_celda])
        col_izquierda.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))

    # --- TEXTOS DE LA DERECHA ---
    nombre = Paragraph(str(row.get("Nombre del producto", "")), style_nombre)
    
    categoria_txt = str(row.get("Categoría", "")).strip()
    categoria_prod = Paragraph(categoria_txt.upper(), style_categoria_producto) if categoria_txt and categoria_txt.lower() != "nan" else ""
    
    desc_txt = str(row.get("Descripción", "")).strip()
    if desc_txt and desc_txt.lower() != "nan":
        frases = desc_txt.replace(". ", "||").split("||")
        frases = [f.strip() for f in frases if f.strip()]
        if frases and not desc_txt.endswith("."):
            frases[-1] += "."
            
        beneficios_html = ""
        for frase in frases:
            beneficios_html += f"<font color='#28a745'>✔</font> {frase}<br/>"
            
        desc = Paragraph(beneficios_html, style_desc)
    else:
        desc = ""
        
    activos_txt = str(row.get("Activos principales", "")).strip()
    activos = Paragraph(f"<b>ACTIVOS:</b> {activos_txt}" if activos_txt and activos_txt.lower() != "nan" else "", style_modo)
    
    modo_txt = str(row.get("Modo de uso", "")).strip()
    modo = Paragraph(f"<b>MODO DE USO:</b> {modo_txt}" if modo_txt and modo_txt.lower() != "nan" else "", style_modo)

    margen_total = 15
    ancho_real_texto = ancho_col_derecha - margen_total

    texto_derecha = Table(
        [[nombre], [categoria_prod], [desc], [activos], [modo]],
        colWidths=[ancho_real_texto]
    )
    texto_derecha.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "LEFT"), ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0), ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0), 
        ("BOTTOMPADDING", (0,0), (-1,-1), 2), # Margen base para todas las filas
        
        # --- MÁRGENES PERSONALIZADOS ---
        ("BOTTOMPADDING", (0,1), (-1,1), 6), # 1. Margen entre Categoría y Descripción
        ("BOTTOMPADDING", (0,2), (-1,2), 6), # 2. Margen entre Descripción y Activos
        ("BOTTOMPADDING", (0,3), (-1,3), 6), # 3. Margen entre Activos y Modo de uso
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
        ("LEFTPADDING", (0,0), (0,0), 5), ("RIGHTPADDING", (0,0), (0,0), 0),
        ("LEFTPADDING", (1,0), (1,0), 10), ("RIGHTPADDING", (1,0), (1,0), 5), 
        ("TOPPADDING", (0,0), (-1,-1), 10), ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("BOX", (0,0), (-1,-1), 2.5, COLOR_DORADO), ("ROUNDEDCORNERS", [8, 8, 8, 8]),
    ]))
    return contenido

# ============================================================
# PÁGINA DE PRODUCTOS (Cuadrícula 2x2)
# ============================================================
def pagina_productos(productos):
    ancho_util = A4[0] - 2.0*cm 
    alto_util  = A4[1] - 5.5*cm 
    
    gap_x = 0.5*cm
    gap_y = 0.4*cm
    
    ancho_celda = (ancho_util - gap_x) / 2
    alto_celda = (alto_util - gap_y) / 2

    celdas = [celda_producto(p, ancho_celda, alto_celda) for p in productos]
    while len(celdas) < 4: celdas.append("")

    grid = Table(
        [
            [celdas[0], celdas[1]],
            [celdas[2], celdas[3]]
        ],
        colWidths=[ancho_celda, ancho_celda],
        rowHeights=[alto_celda, alto_celda]
    )
    
    grid.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), 
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("LEFTPADDING", (0,0), (-1,-1), gap_x/2), ("RIGHTPADDING", (0,0), (-1,-1), gap_x/2),
        ("TOPPADDING", (0,0), (-1,-1), gap_y/2), ("BOTTOMPADDING", (0,0), (-1,-1), gap_y/2),
    ]))
    return grid

# ============================================================
# ENCABEZADOS
# ============================================================
def header_limpio(canvas, doc):
    pass

def header_footer_dinamico(canvas, doc, bg_color):
    canvas.saveState()
    alto_header = 2.5*cm
    canvas.setFillColor(bg_color)
    canvas.rect(0, A4[1] - alto_header, A4[0], alto_header, fill=1, stroke=0)
    
    if os.path.exists(LOGO_PATH):
        try:
            with PILImage.open(LOGO_PATH) as im: w, h = im.size
            ratio = min(6*cm / w, 1.8*cm / h)
            logo_w, logo_h = w * ratio, h * ratio
            canvas.drawImage(LOGO_PATH, A4[0] - logo_w - 1.5*cm, A4[1] - (alto_header/2) - (logo_h/2), width=logo_w, height=logo_h, mask='auto')
        except Exception: pass

    tipo_actual = getattr(doc, '_tipo_actual', '')
    if tipo_actual:
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 20)
        canvas.drawString(1.5*cm, A4[1] - 1.5*cm, tipo_actual.upper())

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
    tipos_colores = cargar_tipos(TIPOS_PATH)
    grupos = agrupar_por_tipo(df)

    doc = BaseDocTemplate(
        OUTPUT_PDF, pagesize=A4,
        leftMargin=1.0*cm, rightMargin=1.0*cm,
        topMargin=3.5*cm, bottomMargin=1.8*cm, 
        title=f"{EMPRESA} - {SUBTITULO}", author=EMPRESA,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    
    doc.addPageTemplates([PageTemplate(id="limpia", frames=[frame], onPage=header_limpio)])

    story = []
    story.append(NextPageTemplate("limpia"))

    for idx, (tipo, prods) in enumerate(grupos, start=1):
        color_hex = tipos_colores.get(tipo, "#1F3864")
        color_fondo = obtener_color_hex(color_hex, COLOR_PRIMARIO)
        nombre_plantilla_prods = f"prods_{idx}"
        
        def crear_header_con_datos(t_name, bg_col):
            def custom_header(canvas, doc):
                doc._tipo_actual = t_name
                header_footer_dinamico(canvas, doc, bg_col)
            return custom_header

        doc.addPageTemplates([PageTemplate(id=nombre_plantilla_prods, frames=[frame], onPage=crear_header_con_datos(tipo, color_fondo))])
        
        story.append(Spacer(1, 5*cm))
        story.append(Paragraph(tipo.upper(), style_tipo_titulo))
        story.append(Spacer(1, 0.5*cm))
        story.append(Table([[""]], colWidths=[6*cm], rowHeights=[0.15*cm], style=TableStyle([("BACKGROUND", (0,0), (-1,-1), color_fondo)])))
        
        story.append(NextPageTemplate(nombre_plantilla_prods))
        story.append(PageBreak())

        for i in range(0, len(prods), 4):
            grid = pagina_productos(prods[i:i+4])
            story.append(grid)
            if i + 4 < len(prods):
                story.append(PageBreak())
            
        story.append(NextPageTemplate("limpia"))
        
    doc.build(story)
    print(f"📄 Cuerpo del catálogo generado temporalmente: {OUTPUT_PDF}")

    if os.path.exists(PORTADA_PDF):
        print(f"🖇️  Añadiendo la portada '{PORTADA_PDF}' al inicio...")
        writer = PdfWriter()
        reader_portada = PdfReader(PORTADA_PDF)
        for page in reader_portada.pages: writer.add_page(page)
        reader_catalogo = PdfReader(OUTPUT_PDF)
        for page in reader_catalogo.pages: writer.add_page(page)
        with open(OUTPUT_FINAL, "wb") as f: writer.write(f)
        print(f"✅ ¡Catálogo Final generado exitosamente!: {OUTPUT_FINAL}")
        os.remove(OUTPUT_PDF)
    else:
        print(f"⚠️ No se encontró '{PORTADA_PDF}'. El catálogo se generó sin portada.")
        os.rename(OUTPUT_PDF, OUTPUT_FINAL)
        print(f"✅ Catálogo Final (sin portada): {OUTPUT_FINAL}")

if __name__ == "__main__":
    generar_y_fusionar()