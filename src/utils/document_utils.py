# ====================================================================
# --- BLOQUE 0: Imports y Dependencias de Python-Docx ---
# ====================================================================
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_TwipsMeasure, Twips

# ====================================================================
# --- BLOQUE 1: Parche Técnico Librearías Temporales (TwipsMeasure) ---
# ====================================================================
# Patch docx
original_convert_from_xml = ST_TwipsMeasure.convert_from_xml

@classmethod
def patch_convert_from_xml(cls, str_value):
    try:
        return Twips(int(str_value))
    except ValueError:
        try:
            return Twips(int(float(str_value)))
        except:
            return original_convert_from_xml(str_value)

ST_TwipsMeasure.convert_from_xml = patch_convert_from_xml

# ====================================================================
# --- BLOQUE 2: Utilidades de Modificación y Estilos de Tabla ---
# ====================================================================
def set_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tblPr.append(borders)

def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_table_margins(table, top=0, bottom=0, left=10, right=10):
    tblPr = table._tbl.tblPr
    tblCellMar = parse_xml(f'''
    <w:tblCellMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:top w:w="{top}" w:type="dxa"/>
        <w:left w:w="{left}" w:type="dxa"/>
        <w:bottom w:w="{bottom}" w:type="dxa"/>
        <w:right w:w="{right}" w:type="dxa"/>
    </w:tblCellMar>
    ''')
    tblPr.append(tblCellMar)

# ====================================================================
# ====================================================================
# --- BLOQUE 2.5: Medidas Predeterminadas de Referencia (CERT-COM-217) ---
# ====================================================================
# Medidas calibradas tomadas de CERT-COM-217-CHALLAPAMPA:
# - Párrafo de firmas contiguo al último párrafo de texto (sin párrafos vacíos intermedios).
# - Espaciado de párrafo: space_before=0, space_after=0, line_spacing=1.0, alignment=CENTER.
# - Firma 1 (Izquierda): cx=1507663 (118.71 pt), cy=1012138 (79.70 pt), posH=1495425, posV=476250 (~37.5 pt)
# - Firma 2 (Derecha):   cx=1665286 (131.12 pt), cy=1030891 (81.17 pt), posH=3533775, posV=473040 (~37.25 pt)
# - Márgenes de envoltura de ancla: distT=0, distB=0
MEDIDAS_REFERENCIA_FIRMAS = {
    'firma_izq': {
        'cx': 1507663,
        'cy': 1012138,
        'pos_h': 1495425,
        'pos_v': 476250,
    },
    'firma_der': {
        'cx': 1665286,
        'cy': 1030891,
        'pos_h': 3533775,
        'pos_v': 473040,
    }
}

def ajustar_posicion_y_tamano_firmas(doc, num_items):
    """
    Ajusta dinámicamente las firmas en la plantilla Word (.docx) tomando como
    referencia predeterminada las medidas exactas de CERT-COM-217-CHALLAPAMPA.
    Ubica las firmas de forma limpia a la distancia calibrada del último párrafo,
    evitando que se sobrepongan al pie de página. Si una tabla contiene un volumen
    inusual de filas (>= 5), compacta adaptativamente el espacio respetando un
    límite legible para no distorsionar el documento, permitiendo edición manual si excede.
    """
    try:
        # 1. Localizar el párrafo de firmas (buscando desde el final hacia arriba)
        sig_p = None
        sig_idx = -1
        total_p = len(doc.paragraphs)
        for i in range(total_p - 1, -1, -1):
            if i < total_p // 2:
                break
            p = doc.paragraphs[i]
            drawings = p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
            sig_drawings = []
            for d in drawings:
                ext = d.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
                if ext is not None:
                    try:
                        cy = int(ext.get('cy', 0))
                        # Las firmas tienen alto > 20 pt (254,000 EMUs) y cy < 4,000,000 (no marcas de agua)
                        if 254000 < cy < 4000000:
                            sig_drawings.append(d)
                    except (ValueError, TypeError):
                        pass
            if len(sig_drawings) >= 1 and ('[[TABLA_NOTAS]]' not in p.text):
                sig_p = p
                sig_idx = i
                break

        if sig_p is None or sig_idx == -1:
            return

        # 2. Eliminar párrafos vacíos posteriores al párrafo de firmas
        paras_posteriores_a_eliminar = []
        for i in range(sig_idx + 1, len(doc.paragraphs)):
            p = doc.paragraphs[i]
            drawings = p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
            if not p.text.strip() and not drawings:
                paras_posteriores_a_eliminar.append(p)
        for p in paras_posteriores_a_eliminar:
            try:
                p._element.getparent().remove(p._element)
            except Exception:
                pass

        # 3. Eliminar párrafos vacíos inmediatamente anteriores al párrafo de firmas
        # (para que el párrafo de firmas quede contiguo al último párrafo de texto, igual que en CERT-COM-217)
        paras_anteriores_a_eliminar = []
        for i in range(sig_idx - 1, -1, -1):
            p = doc.paragraphs[i]
            drawings = p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
            if not p.text.strip() and not drawings:
                paras_anteriores_a_eliminar.append(p)
            else:
                break
        for p in paras_anteriores_a_eliminar:
            try:
                p._element.getparent().remove(p._element)
            except Exception:
                pass

        # 4. Formato del párrafo de firmas idéntico a CERT-COM-217
        sig_p.paragraph_format.space_before = Pt(0)
        sig_p.paragraph_format.space_after = Pt(0)
        sig_p.paragraph_format.line_spacing = 1.0
        sig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 5. Escala y distancia predeterminada (referencia CERT-COM-217-CHALLAPAMPA)
        # En caso estándar (hasta 3-4 ítems), usa exactamente el 100% de las medidas de CERT-COM-217.
        if num_items <= 3:
            scale = 1.0
            scale_v = 1.0
        elif num_items == 4:
            scale = 0.96
            scale_v = 0.88
        elif num_items <= 6:
            scale = 0.88
            scale_v = 0.72
            # Compactar párrafos vacíos anteriores si la tabla es grande
            for p in doc.paragraphs:
                if not p.text.strip() and not p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'):
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = Pt(3)
        else:
            # 7 o más ítems: límite legible
            scale = 0.80
            scale_v = 0.55
            for p in doc.paragraphs:
                if not p.text.strip() and not p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'):
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = Pt(2)

        # 6. Obtener dibujos y ordenarlos de izquierda a derecha según su posH
        drawings = sig_p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
        
        def get_pos_h(d):
            posH = d.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}positionH')
            if posH is not None:
                off = posH.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}posOffset')
                if off is not None and off.text:
                    try:
                        return int(off.text)
                    except ValueError:
                        pass
            return 0

        drawings.sort(key=get_pos_h)

        configs = [MEDIDAS_REFERENCIA_FIRMAS['firma_izq'], MEDIDAS_REFERENCIA_FIRMAS['firma_der']]

        for idx, d in enumerate(drawings[:2]):
            cfg = configs[idx]

            # Eliminar márgenes de envoltura en anchor
            for anchor in d.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor'):
                anchor.set('distT', '0')
                anchor.set('distB', '0')

            # Posición horizontal predeterminada (posOffset en column)
            posH = d.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}positionH')
            if posH is not None:
                off = posH.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}posOffset')
                if off is not None:
                    off.text = str(cfg['pos_h'])

            # Distancia vertical predeterminada respecto al párrafo (posOffset en paragraph)
            posV = d.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}positionV')
            if posV is not None:
                off = posV.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}posOffset')
                if off is not None:
                    off.text = str(int(cfg['pos_v'] * scale_v))

            # Tamaño predeterminado (cx, cy)
            new_cx = str(int(cfg['cx'] * scale))
            new_cy = str(int(cfg['cy'] * scale))
            for el in d.iter():
                if 'cx' in el.attrib and 'cy' in el.attrib:
                    el.attrib['cx'] = new_cx
                    el.attrib['cy'] = new_cy

    except Exception as e:
        print(f"Aviso en ajuste de firmas: {e}")

# ====================================================================
# --- BLOQUE 3: Lógica Principal de Inyección Documental ---
# ====================================================================
def inyectar_tabla_en_docx(doc_io, data_items):
    doc = Document(doc_io)
    target_paragraph = None
    for p in doc.paragraphs:
        if '[[TABLA_NOTAS]]' in p.text:
            target_paragraph = p
            break
            
    if target_paragraph:
        target_paragraph.text = target_paragraph.text.replace('[[TABLA_NOTAS]]', '')
        for p in doc.paragraphs[:5]:
            if "CERTIFICADO" in p.text.upper():
                p.paragraph_format.space_after = Pt(0)
        
        table = doc.add_table(rows=1, cols=7)
        try:
            table.style = 'Table Grid'
        except:
            set_borders(table)
            
        table.autofit = False
        table.allow_autofit = False
        set_table_margins(table, top=72, bottom=72, left=30, right=30)

        widths = [Inches(0.75), Inches(0.75), Inches(0.75), Inches(3.0), Inches(0.75), Inches(0.75), Inches(0.75)]
        for i, col in enumerate(table.columns):
            col.width = widths[i]
        
        encabezados = ['Fecha', 'Placa', 'N° Guía', 'Descripción', 'Cantidad', 'Medida', 'Peso (Kg)']
        hdr_cells = table.rows[0].cells
        for i, nombre in enumerate(encabezados):
            cell = hdr_cells[i]
            cell.text = nombre
            cell.width = widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background(cell, "70ad47")
            
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1
                run = p.runs[0] if p.runs else p.add_run(nombre)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.name = 'Calibri'
                run.font.size = Pt(9)
        
        for item in data_items:
            row_cells = table.add_row().cells
            vals = [
                str(item.get('fecha_origen', '')),
                str(item.get('placa_origen', '')),
                str(item.get('guia_origen', '')),
                str(item.get('desc', '')),
                str(item.get('cant', '')),
                str(item.get('um', '')).upper(),
                str(item.get('peso', ''))
            ]
            
            for idx, valor in enumerate(vals):
                cell = row_cells[idx]
                cell.text = valor
                cell.width = widths[idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1
                    run = p.runs[0] if p.runs else p.add_run(valor)
                    run.font.name = 'Calibri'
                    run.font.size = Pt(9)

        tbl, p = table._tbl, target_paragraph._p
        p.addnext(tbl)

    # Ajuste adaptativo de firmas para evitar superposición con el pie de página
    ajustar_posicion_y_tamano_firmas(doc, len(data_items))

    new_buffer = io.BytesIO()
    doc.save(new_buffer)
    return new_buffer.getvalue()

def sustituir_certificado_en_pdf(pdf_unificado_bytes, nuevo_cert_pdf_bytes, num_paginas_reemplazar=1):
    """
    Sustituye la(s) primera(s) página(s) de un PDF unificado (el certificado desactualizado)
    por las páginas del nuevo certificado, conservando intactas todas las páginas posteriores (las guías).
    
    :param pdf_unificado_bytes: bytes del PDF consolidado original.
    :param nuevo_cert_pdf_bytes: bytes del nuevo certificado en formato PDF.
    :param num_paginas_reemplazar: número de páginas iniciales a sustituir (por defecto 1).
    :return: bytes del nuevo PDF unificado.
    """
    from pypdf import PdfReader, PdfWriter

    reader_unido = PdfReader(io.BytesIO(pdf_unificado_bytes))
    reader_nuevo = PdfReader(io.BytesIO(nuevo_cert_pdf_bytes))
    writer = PdfWriter()

    # 1. Agregar todas las páginas del nuevo certificado
    for p in reader_nuevo.pages:
        writer.add_page(p)

    # 2. Agregar las páginas restantes del PDF unificado original (las guías)
    total_unido = len(reader_unido.pages)
    paginas_a_saltar = min(num_paginas_reemplazar, total_unido)
    for p in reader_unido.pages[paginas_a_saltar:]:
        writer.add_page(p)

    out_io = io.BytesIO()
    writer.write(out_io)
    out_io.seek(0)
    return out_io.getvalue()
