# -*- coding: utf-8 -*-
"""
Generador de catálogo PDF - Distribuidora Bellingieri
Uso: python generar_catalogo.py
"""

import os
import math
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
    Table, TableStyle, Image, PageBreak, NextPageTemplate, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image as PILImage
from pypdf import PdfWriter, PdfReader

# ============================================================
# CONFIGURACIÓN
# ============================================================
EXCEL_PATH     = "catalogo.xlsx"
TIPOS_PATH     = "tipos.xlsx"
IMAGENES_DIR   = "imagenes" 
ICONOS_DIR     = ""         
LOGO_PATH      = "logo.png"  
PORTADA_PDF    = "portada.pdf" 
OUTPUT_INDICE  = "Indice_Temporal.pdf"
OUTPUT_PRODS   = "Productos_Temporal.pdf"
OUTPUT_FINAL   = "Catalogo_Distribuidora_Bellingieri_Final.pdf" 

COLOR_PRIMARIO   = colors.HexColor("#1F3864")
COLOR_TEXTO      = colors.HexColor("#222222")
COLOR_PRECIO     = colors.HexColor("#C00000")
COLOR_DORADO     = colors.HexColor("#C5A059") 
COLOR_CELESTE    = colors.HexColor("#00AEEF") 
COLOR_GRIS_CAJA  = colors.HexColor("#E0E0E0")

EMPRESA   = "Distribuidora Bellingieri"
SUBTITULO = "Catálogo de Productos"

OFFSET_PAGINAS = 0

# ============================================================
# REGISTRO DE FUENTES (Open Sans Variables)
# ============================================================
try:
    # Registramos el archivo variable como 'OpenSans' y 'OpenSans-Bold'
    pdfmetrics.registerFont(TTFont('OpenSans', os.path.join('font', 'OpenSans-VariableFont_wdth,wght.ttf')))
    pdfmetrics.registerFont(TTFont('OpenSans-Bold', os.path.join('font', 'OpenSans-VariableFont_wdth,wght.ttf')))
    
    # Al registrar el mismo archivo con dos nombres, ReportLab usa el peso por defecto (Regular).
    # Para forzar el Bold en la fuente variable, usamos un truco de nombres.
    # Le decimos que 'OpenSans-Bold' es una variante de 'OpenSans' con el peso máximo.
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily('OpenSans', normal='OpenSans', bold='OpenSans-Bold')
    
    print("✅ Fuentes Open Sans (Variable) cargadas correctamente.")
except Exception as e:
    print(f"⚠️ No se pudieron cargar las fuentes: {e}. Se usará Helvetica por defecto.")

# ============================================================
# ESTILOS
# ============================================================
styles = getSampleStyleSheet()

style_nombre = ParagraphStyle("nombre", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=COLOR_PRIMARIO, spaceAfter=2)
style_categoria_producto = ParagraphStyle("categoria_producto", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7, leading=8.5, textColor=COLOR_DORADO, spaceAfter=3)
style_desc = ParagraphStyle("desc", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=9.5, textColor=COLOR_TEXTO, spaceAfter=4)
style_modo = ParagraphStyle("modo", parent=styles["Normal"], fontName="Helvetica", fontSize=7, leading=9, textColor=colors.HexColor("#444444"), spaceAfter=4)
style_tabla_nombre = ParagraphStyle("tabla_nombre", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6, leading=7.5, textColor=colors.black, alignment=TA_CENTER)
style_tabla_precio = ParagraphStyle("tabla_precio", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, leading=11, textColor=colors.black, alignment=TA_CENTER)

# Estilos Índice
style_index_texto_unico = ParagraphStyle(
    "index_texto_unico", parent=styles["Normal"],
    fontName="OpenSans-Bold", # Usamos la variante que acabamos de forzar
    fontSize=15, leading=18, 
    textColor=COLOR_PRIMARIO, 
    spaceAfter=20
)
style_index_tipo = ParagraphStyle("index_tipo", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=COLOR_TEXTO, alignment=TA_CENTER)
style_index_paginas = ParagraphStyle("index_paginas", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=9, textColor=colors.HexColor("#666666"), alignment=TA_CENTER)

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
    return df.to_dict("records")

def cargar_tipos(path):
    if not os.path.exists(path): return {}
    try:
        df_tipos = pd.read_excel(path)
        df_tipos.columns = [c.strip() for c in df_tipos.columns]
        
        col_tipo = next((c for c in df_tipos.columns if c.lower() == "tipo"), None)
        col_color = next((c for c in df_tipos.columns if c.lower() == "color"), None)
        col_img = next((c for c in df_tipos.columns if "imagen" in c.lower()), None)
        
        if not col_tipo or not col_color: return {}
        
        df_tipos = df_tipos.fillna("")
        df_tipos[col_tipo] = df_tipos[col_tipo].astype(str).str.strip()
        df_tipos[col_color] = df_tipos[col_color].astype(str).str.strip()
        if col_img:
            df_tipos[col_img] = df_tipos[col_img].astype(str).str.strip()
            
        resultado = {}
        for _, row in df_tipos.iterrows():
            tipo = row[col_tipo]
            color = row[col_color]
            imagen = row[col_img] if col_img else ""
            resultado[tipo] = {"color": color, "imagen": imagen}
            
        return resultado
    except Exception as e: 
        print(f"Error cargando tipos: {e}")
        return {}

# ============================================================
# CLASE PERSONALIZADA PARA DIBUJAR LA CUADRÍCULA
# ============================================================
class CuadriculaIndice(Flowable):
    def __init__(self, tipos_unicos, paginas_por_tipo, tipos_data, margen_izq_frame):
        super().__init__()
        self.tipos = tipos_unicos
        self.paginas = paginas_por_tipo
        self.data = tipos_data
        
        self.ancho_hoja = A4[0]
        
        self.ancho_caja = 4.5 * cm
        self.alto_caja = 4.2 * cm
        self.espacio_x = 0.8 * cm
        self.espacio_y = 0.5 * cm
        
        self.ancho_total = (self.ancho_caja * 3) + (self.espacio_x * 2)
        self.x_inicial = ((self.ancho_hoja - self.ancho_total) / 2) - margen_izq_frame
        
        self.filas = (len(self.tipos) + 2) // 3
        self.alto_total = (self.alto_caja * self.filas) + (self.espacio_y * (self.filas - 1))

    def wrap(self, availWidth, availHeight):
        return (availWidth, self.alto_total)

    def draw(self):
        canvas = self.canv

        for i, tipo in enumerate(self.tipos):
            col = i % 3
            fila = i // 3
            
            if fila == self.filas - 1 and len(self.tipos) % 3 == 1:
                col = 1
            
            x_pos = self.x_inicial + (col * (self.ancho_caja + self.espacio_x))
            y_pos = self.alto_total - (fila * (self.alto_caja + self.espacio_y)) - self.alto_caja

            paginas = self.paginas.get(tipo, [])
            if paginas:
                if len(paginas) == 1: rango_str = str(paginas[0])
                else:
                    es_consecutivo = all(paginas[j+1] == paginas[j] + 1 for j in range(len(paginas)-1))
                    if es_consecutivo: rango_str = f"{paginas[0]} - {paginas[-1]}"
                    else: rango_str = ", ".join(map(str, paginas))
            else: rango_str = "?"

            info = self.data.get(tipo, {})
            color_hex = info.get("color", "#1F3864")
            nombre_img = info.get("imagen", "")
            color_borde = obtener_color_hex(color_hex, COLOR_PRIMARIO)

            # Sombra
            canvas.setFillColor(COLOR_GRIS_CAJA)
            canvas.roundRect(x_pos + 2, y_pos - 2, self.ancho_caja, self.alto_caja, 6, stroke=0, fill=1)
            
            # Caja Blanca
            canvas.setFillColor(colors.white)
            canvas.setStrokeColor(COLOR_GRIS_CAJA)
            canvas.setLineWidth(0.5)
            canvas.roundRect(x_pos, y_pos, self.ancho_caja, self.alto_caja, 6, stroke=1, fill=1)
            
            # Línea superior
            canvas.setStrokeColor(color_borde)
            canvas.setLineWidth(4)
            canvas.line(x_pos + 1, y_pos + self.alto_caja, x_pos + self.ancho_caja - 1, y_pos + self.alto_caja)

            # Icono
            centro_x = x_pos + self.ancho_caja / 2
            if nombre_img and nombre_img.lower() != "nan":
                ruta = os.path.join(ICONOS_DIR, nombre_img)
                if os.path.isfile(ruta):
                    try:
                        canvas.drawImage(ruta, centro_x - (1.0*cm)/2, y_pos + self.alto_caja - 2.5*cm, width=1.0*cm, height=1.0*cm, mask='auto')
                    except: pass
            
            # Textos con tamaño aumentado
            canvas.setFillColor(COLOR_TEXTO)
            canvas.setFont("Helvetica-Bold", 9) 
            canvas.drawCentredString(centro_x, y_pos + 0.9*cm, tipo.upper())
            
            texto_pags = "Págs. "
            ancho_pags = canvas.stringWidth(texto_pags, "Helvetica", 8.5) 
            ancho_num = canvas.stringWidth(rango_str, "Helvetica-Bold", 12) 
            ancho_total_texto = ancho_pags + ancho_num
            
            inicio_x = centro_x - (ancho_total_texto / 2)
            
            canvas.setFillColor(colors.HexColor("#555555"))
            canvas.setFont("Helvetica", 8.5) 
            canvas.drawString(inicio_x, y_pos + 0.4*cm, texto_pags)
            
            canvas.setFillColor(COLOR_TEXTO)
            canvas.setFont("Helvetica-Bold", 12) 
            canvas.drawString(inicio_x + ancho_pags, y_pos + 0.4*cm, rango_str)

# ============================================================
# DISEÑO DE LA PÁGINA DE ÍNDICE
# ============================================================
def generar_pdf_indice(productos, tipos_data, output_path):
    doc = BaseDocTemplate(
        output_path, pagesize=A4,
        leftMargin=3.0*cm, rightMargin=3.0*cm,
        topMargin=3.0*cm, bottomMargin=2.0*cm, 
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="limpia", frames=[frame], onPage=lambda canvas, doc: None)])

    elementos = []

    # --- CABECERA CON LOGO Y TÍTULO GRANDE ---
    logo_img = ""
    if os.path.exists(LOGO_PATH):
        try:
            with PILImage.open(LOGO_PATH) as im: w, h = im.size
            max_h = 1.5*cm
            ratio = max_h / h
            logo_img = Image(LOGO_PATH, width=w*ratio, height=max_h)
            logo_img.hAlign = "LEFT"
        except Exception: pass

    style_catalogo_grande = ParagraphStyle(
        "catalogo_grande", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=28, leading=32, 
        textColor=COLOR_PRIMARIO, alignment=TA_RIGHT
    )
    texto_catalogo = Paragraph("CATÁLOGO", style_catalogo_grande)

    tabla_header = Table([[logo_img, texto_catalogo]], colWidths=[doc.width * 0.5, doc.width * 0.5])
    tabla_header.setStyle(TableStyle([
        ("VALIGN", (0,0), (0,0), "MIDDLE"), ("VALIGN", (1,0), (1,0), "MIDDLE"),
        ("ALIGN", (0,0), (0,0), "LEFT"), ("ALIGN", (1,0), (1,0), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 0), ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0), ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ("LINEBELOW", (0,0), (-1,-1), 1, COLOR_GRIS_CAJA),
    ]))
    
    elementos.append(tabla_header)
    elementos.append(Spacer(1, 1.5*cm))
    
    # --- TEXTO ÚNICO CENTRAL (Con Open Sans Variable - Bold) ---
    elementos.append(Paragraph("Encuentra la solución en segundos", style_index_texto_unico))
    elementos.append(Spacer(1, 0.5*cm))

    # Cálculo de páginas
    offset_indice = 2 if os.path.exists(PORTADA_PDF) else 1
    paginas_por_tipo = {}
    pagina_actual = offset_indice + 1
    
    for i in range(0, len(productos), 4):
        pagina = productos[i:i+4]
        tipo = pagina[0]["Tipo"]
        if tipo not in paginas_por_tipo: paginas_por_tipo[tipo] = []
        paginas_por_tipo[tipo].append(pagina_actual)
        pagina_actual += 1

    tipos_unicos = []
    for p in productos:
        if p["Tipo"] not in tipos_unicos: tipos_unicos.append(p["Tipo"])

    cuadricula = CuadriculaIndice(tipos_unicos, paginas_por_tipo, tipos_data, doc.leftMargin)
    elementos.append(cuadricula)
    
    doc.build(elementos)

# ============================================================
# CELDA DE PRODUCTO
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
    GAP_IMG_PRECIOS = 0.35 * cm
    
    img = obtener_imagen_ajustada(img_path, ancho_bloque_izq - 0.4*cm, alto_img_area - 0.4*cm)

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
            lineas_precios.append(f"<font color='#222222'><b>{pres_txt}</b></font> | <font color='#C00000'><b>{prec_str}</b></font>")
            
    if len(lineas_precios) > 0:
        texto_precios = "<br/>".join(lineas_precios)
        bloque_precios = Paragraph(texto_precios, style_tabla_precio)
        
        col_izquierda = Table([[img], [bloque_precios]], colWidths=[ancho_bloque_izq], rowHeights=[alto_img_area, alto_tabla_area])
        col_izquierda.setStyle(TableStyle([
            ("ALIGN", (0,0), (0,0), "CENTER"), ("ALIGN", (0,1), (0,1), "CENTER"),   
            ("VALIGN", (0,0), (0,0), "BOTTOM"), ("VALIGN", (0,1), (0,1), "TOP"), 
            ("LEFTPADDING", (0,0), (-1,-1), 0), ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (0,0), 0), ("BOTTOMPADDING", (0,0), (0,0), 0),
            ("TOPPADDING", (0,1), (0,1), GAP_IMG_PRECIOS), ("BOTTOMPADDING", (0,1), (0,1), 0), 
        ]))
    else:
        col_izquierda = Table([[img]], colWidths=[ancho_bloque_izq], rowHeights=[alto_celda])
        col_izquierda.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "BOTTOM")]))

    nombre = Paragraph(str(row.get("Nombre del producto", "")), style_nombre)
    categoria_txt = str(row.get("Categoría", "")).strip()
    categoria_prod = Paragraph(categoria_txt.upper(), style_categoria_producto) if categoria_txt and categoria_txt.lower() != "nan" else ""
    
    desc_txt = str(row.get("Descripción", "")).strip()
    if desc_txt and desc_txt.lower() != "nan":
        frases = desc_txt.replace(". ", "||").split("||")
        frases = [f.strip() for f in frases if f.strip()]
        if frases and not desc_txt.endswith("."): frases[-1] += "."
        beneficios_html = "".join([f"<font color='#28a745'>✔</font> {frase}<br/>" for frase in frases])
        desc = Paragraph(beneficios_html, style_desc)
    else:
        desc = ""
        
    activos_txt = str(row.get("Activos principales", "")).strip()
    activos = Paragraph(f"<b>ACTIVOS:</b> {activos_txt}" if activos_txt and activos_txt.lower() != "nan" else "", style_modo)
    modo_txt = str(row.get("Modo de uso", "")).strip()
    modo = Paragraph(f"<b>MODO DE USO:</b> {modo_txt}" if modo_txt and modo_txt.lower() != "nan" else "", style_modo)

    margen_total = 15
    ancho_real_texto = ancho_col_derecha - margen_total

    texto_derecha = Table([[nombre], [categoria_prod], [desc], [activos], [modo]], colWidths=[ancho_real_texto])
    texto_derecha.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "LEFT"), ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0), ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0), ("BOTTOMPADDING", (0,0), (-1,-1), 2), 
        ("BOTTOMPADDING", (0,1), (-1,1), 6), ("BOTTOMPADDING", (0,2), (-1,2), 6), ("BOTTOMPADDING", (0,3), (-1,3), 6), 
    ]))

    contenido = Table([[col_izquierda, texto_derecha]], colWidths=[ancho_bloque_izq, ancho_col_derecha], rowHeights=[alto_celda])
    contenido.setStyle(TableStyle([
        ("ALIGN", (0,0), (0,0), "CENTER"), ("VALIGN", (0,0), (0,0), "MIDDLE"), 
        ("ALIGN", (1,0), (1,0), "LEFT"), ("VALIGN", (1,0), (1,0), "MIDDLE"), 
        ("LEFTPADDING", (0,0), (0,0), 5), ("RIGHTPADDING", (0,0), (0,0), 0),
        ("LEFTPADDING", (1,0), (1,0), 10), ("RIGHTPADDING", (1,0), (1,0), 5), 
        ("TOPPADDING", (0,0), (-1,-1), 10), ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("BOX", (0,0), (-1,-1), 2.5, COLOR_DORADO), ("ROUNDEDCORNERS", [8, 8, 8, 8]),
    ]))
    return contenido

def pagina_productos(productos):
    ancho_util = A4[0] - 2.0*cm 
    alto_util  = A4[1] - 5.5*cm 
    gap_x, gap_y = 0.5*cm, 0.4*cm
    ancho_celda = (ancho_util - gap_x) / 2
    alto_celda = (alto_util - gap_y) / 2

    celdas = [celda_producto(p, ancho_celda, alto_celda) for p in productos]
    while len(celdas) < 4: celdas.append("")

    grid = Table([[celdas[0], celdas[1]], [celdas[2], celdas[3]]], colWidths=[ancho_celda, ancho_celda], rowHeights=[alto_celda, alto_celda])
    grid.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("LEFTPADDING", (0,0), (-1,-1), gap_x/2), ("RIGHTPADDING", (0,0), (-1,-1), gap_x/2),
        ("TOPPADDING", (0,0), (-1,-1), gap_y/2), ("BOTTOMPADDING", (0,0), (-1,-1), gap_y/2),
    ]))
    return grid

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
    
    num_pagina_real = doc.page + OFFSET_PAGINAS
    canvas.drawRightString(A4[0]-1.5*cm, 0.9*cm, f"Página {num_pagina_real}")
    canvas.restoreState()

def generar_pdf_productos(productos, tipos_data, output_path):
    doc = BaseDocTemplate(
        output_path, pagesize=A4,
        leftMargin=1.0*cm, rightMargin=1.0*cm,
        topMargin=3.5*cm, bottomMargin=1.8*cm, 
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    
    story = []

    for i in range(0, len(productos), 4):
        pagina = productos[i:i+4]
        tipo_pagina = pagina[0]["Tipo"]
        
        color_hex = tipos_data.get(tipo_pagina, {}).get("color", "#1F3864")
        color_fondo = obtener_color_hex(color_hex, COLOR_PRIMARIO)
        nombre_plantilla = f"pag_{i}"
        
        def crear_header_con_datos(t_name, bg_col):
            def custom_header(canvas, doc):
                doc._tipo_actual = t_name
                header_footer_dinamico(canvas, doc, bg_col)
            return custom_header

        doc.addPageTemplates([PageTemplate(id=nombre_plantilla, frames=[frame], onPage=crear_header_con_datos(tipo_pagina, color_fondo))])
        
        story.append(NextPageTemplate(nombre_plantilla))
        if i > 0:
            story.append(PageBreak())

        grid = pagina_productos(pagina)
        story.append(grid)

    doc._firstPageTemplateIndex = 1
    doc.build(story)

# ============================================================
# GENERACIÓN Y FUSIÓN FINAL
# ============================================================
def generar_y_fusionar():
    global OFFSET_PAGINAS
    
    print("📊 Cargando datos...")
    productos = cargar_datos(EXCEL_PATH)
    tipos_data = cargar_tipos(TIPOS_PATH)

    OFFSET_PAGINAS = 2 if os.path.exists(PORTADA_PDF) else 1

    print("📄 Generando Índice...")
    generar_pdf_indice(productos, tipos_data, OUTPUT_INDICE)

    print("📄 Generando Productos...")
    generar_pdf_productos(productos, tipos_data, OUTPUT_PRODS)

    print("🖇️  Fusionando todo (Portada + Índice + Productos)...")
    writer = PdfWriter()

    if os.path.exists(PORTADA_PDF):
        reader_portada = PdfReader(PORTADA_PDF)
        for page in reader_portada.pages: 
            writer.add_page(page)

    if os.path.exists(OUTPUT_INDICE):
        reader_indice = PdfReader(OUTPUT_INDICE)
        for page in reader_indice.pages: 
            writer.add_page(page)
        
    if os.path.exists(OUTPUT_PRODS):
        reader_prods = PdfReader(OUTPUT_PRODS)
        for page in reader_prods.pages: 
            writer.add_page(page)

    with open(OUTPUT_FINAL, "wb") as f:
        writer.write(f)
        
    print(f"✅ ¡Catálogo Final generado exitosamente!: {OUTPUT_FINAL}")

    for tmp in [OUTPUT_INDICE, OUTPUT_PRODS]:
        if os.path.exists(tmp): 
            os.remove(tmp)

if __name__ == "__main__":
    generar_y_fusionar()