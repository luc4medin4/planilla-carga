"""
Planilla de Carga — Beccacece Hnos SA
Generador automático v3.2 | Streamlit + ReportLab
"""
import io, math, hashlib
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas

# ─── CONSTANTS ──────────────────────────────────────────────────────────────
PAGE_W, PAGE_H = A4
MARGIN = 12 * mm
CW = PAGE_W - 2 * MARGIN           # content width ≈ 527 pt

DARK_BLUE   = colors.HexColor('#1a3a6b')
MED_BLUE    = colors.HexColor('#2e5fa3')
LIGHT_GRAY  = colors.HexColor('#EFEFEF')
BG_GRAY     = colors.HexColor('#F4F4F4')
DARK_ROW    = colors.HexColor('#AAAAAA')
MED_ROW     = colors.HexColor('#D8D8D8')
AMBER       = colors.HexColor('#FFE0B2')
RED_ALERT   = colors.HexColor('#C0392B')
FOOT_GRAY   = colors.HexColor('#555555')
BORDER_CLR  = colors.HexColor('#AAAAAA')
HDR_BG      = colors.HexColor('#E8E8E8')
YELLOW_RTE  = colors.HexColor('#FFF3A0')
WM_GRAY     = colors.HexColor('#888888')

EXCLUDED_SKUS = {2730, 2731, 2776, 2780, 5192}
EXCLUDED_PATS = ['Q CERVEZAS', 'Q PLAS', 'BOT 1/1', 'BOT AMBAR', 'ARACELI']
CANCHA_ORDER  = ['CANCHA I', 'CANCHA II', 'CANCHA III', 'CANCHA IV', 'CANCHA V']

LEMAS = [
    "Cada bulto bien puesto es una entrega perfecta. Dale con todo!",
    "El picking preciso empieza con vos. Conta, verifica, confia.",
    "Hoy es dia de hacer historia en el deposito. A romperla!",
    "Rapido no es apurado. Rapido es preciso y sin errores.",
    "El mejor pickero no es el mas veloz - es el mas exacto.",
    "Tu equipo cuenta con vos. Cada bulto importa.",
    "Zapatos abrochados, espalda cuidada, mente enfocada. Arrancamos!",
    "Paleta perfecta = cliente feliz = empresa fuerte. Todo empieza aca.",
    "Cero errores de picking es el estandar. Vos podes.",
    "Si dudas del SKU, verifica. Un segundo de control evita una hora de correccion.",
    "La diferencia entre bueno y excelente esta en los detalles. Se excelente.",
    "Arrancar bien es terminar mejor. Que empiece el picking!",
    "Trabajo limpio, paleta ordenada, conciencia tranquila.",
    "Sos parte del corazon de la operacion. El deposito te necesita al 100%.",
    "Pensa antes de mover. El orden en la paleta es orden en la entrega.",
    "Tecnica correcta al levantar = cuerpo sano manana tambien.",
    "Verifica el SKU, verifica la cantidad, verifica el lote. Tres checks, cero errores!",
    "Cada noche de picking sin errores construye la reputacion de todos.",
    "El deposito funciona porque vos funcionas. Gracias por el compromiso.",
    "Paleta bien armada, camion bien cargado, cliente bien atendido.",
    "El cansancio es real. El orgullo del trabajo bien hecho tambien.",
    "Reporta lo que esta mal antes de moverlo. Un reporte a tiempo vale oro.",
    "Somos un equipo. Si uno falla, fallamos todos. Si uno gana, ganamos todos.",
    "Segui la planilla al pie de la letra. Esta hecha para que llegues sin errores.",
    "Esta noche vas a hacer 500 movimientos correctos. Demostralo!",
    "Seguridad primero: EPI puesto, pasillo libre, mente clara.",
    "Los mejores pickeros no nacen - se hacen en noches como esta.",
    "Picking nocturno = concentracion maxima. El dia descansa, vos construis.",
    "Cada producto en su lugar es una promesa cumplida al cliente.",
    "Ultima caja, mismo cuidado que la primera. Asi se hace el trabajo bien hecho.",
]

# ─── DATA LOADING ───────────────────────────────────────────────────────────

def is_envase(sku: int, desc: str) -> bool:
    if sku in EXCLUDED_SKUS:
        return True
    d = str(desc).upper()
    for p in EXCLUDED_PATS:
        if p in d:
            return True
    if d.strip() == 'RET':
        return True
    return False

def load_car(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    df['Artículo'] = pd.to_numeric(df['Artículo'], errors='coerce')
    df = df.dropna(subset=['Artículo'])
    df['Artículo'] = df['Artículo'].astype(int)
    df['Bultos']   = pd.to_numeric(df['Bultos'],   errors='coerce').fillna(0)
    df['Unids']    = pd.to_numeric(df['Unids'],     errors='coerce').fillna(0)
    df = df[~df.apply(lambda r: is_envase(r['Artículo'], r['Descripción Artículo']), axis=1)]
    return df

def load_frescura(file):
    xls = pd.ExcelFile(file)

    # API sheet
    api = pd.read_excel(xls, sheet_name='API')
    api = api[['Artículo', 'CANCHA.1', 'almacén', 'detalle']].copy()
    api.columns = ['sku', 'cancha', 'alm_id', 'alm_det']
    api['sku'] = pd.to_numeric(api['sku'], errors='coerce')
    api = api.dropna(subset=['sku'])
    api['sku'] = api['sku'].astype(int)
    api = api.drop_duplicates(subset=['sku'])

    # DDM sheet
    ddm = pd.read_excel(xls, sheet_name='DDM')
    ddm = ddm[['ARTÍCULO', 'BULTOS X PALLET']].copy()
    ddm.columns = ['sku', 'bxp']
    ddm['sku'] = pd.to_numeric(ddm['sku'], errors='coerce')
    ddm = ddm.dropna(subset=['sku'])
    ddm['sku'] = ddm['sku'].astype(int)
    ddm = ddm.drop_duplicates(subset=['sku'])

    # Frescura sheet
    fr = pd.read_excel(xls, sheet_name='Frescura')
    fr = fr[['Cód', 'FECHA DE VENC.', 'Status']].copy()
    fr.columns = ['sku', 'fecha', 'status']
    fr['sku'] = pd.to_numeric(fr['sku'], errors='coerce')
    fr = fr.dropna(subset=['sku'])
    fr['sku'] = fr['sku'].astype(int)
    # FEFO: sort by date, keep first (earliest)
    fr = fr.sort_values('fecha').drop_duplicates(subset=['sku'])

    return api, ddm, fr

# ─── BUSINESS LOGIC ─────────────────────────────────────────────────────────

def normalize_cancha(val) -> str:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return 'SIN CANCHA'
    s = str(val).strip().upper()
    if s in ('', '0', 'NAN'):
        return 'SIN CANCHA'
    MAP = {'CANCHA I': 'CANCHA I', 'CANCHA II': 'CANCHA II',
           'CANCHA III': 'CANCHA III', 'CANCHA IV': 'CANCHA IV',
           'CANCHA V': 'CANCHA V', 'MKPL': 'CANCHA V', 'MERCH': 'CANCHA V'}
    for k, v in MAP.items():
        if s == k or s.startswith(k):
            return v
    return 'SIN CANCHA'

def get_lema(transport, idx=0) -> str:
    h = int(hashlib.md5(str(transport).encode()).hexdigest(), 16)
    return LEMAS[(h + idx) % len(LEMAS)]

def pallet_reading(val) -> tuple:
    """(valor_minimo_str, valor_entero_str, lectura)"""
    if val is None or val == 0.0:
        return ('0.00', '0 paletas', 'Sin carga de picking en esta cancha')
    vm = f'{val:.2f}'
    n  = int(val)
    fr = val - n
    if val < 1.0:
        return (vm, '1 paleta', 'Menos de una paleta')
    if val == 1.0:
        return (vm, '1 paleta', '1 paleta exacta')
    if 1.0 < val < 1.5:
        return (vm, '2 paletas', 'Mas de una paleta')
    if val == 1.5:
        return (vm, '2 paletas', '1 paleta y media')
    if 1.5 < val < 2.0:
        return (vm, '2 paletas', 'Casi 2 paletas')
    # n >= 2
    if fr == 0.0:
        return (vm, f'{n} paletas', f'{n} paletas exactas')
    if 0.0 < fr < 0.5:
        return (vm, f'{n+1} paletas', f'Mas de {n} paletas')
    if fr == 0.5:
        return (vm, f'{n+1} paletas', f'{n} paletas y media')
    return (vm, f'{n+1} paletas', f'Casi {n+1} paletas')

def process_reparto(rep_df, api, ddm, fr) -> list:
    """Build enriched row list for one reparto."""
    rows = []
    for _, r in rep_df.iterrows():
        sku = int(r['Artículo'])
        blts_raw = float(r['Bultos'])
        unids    = float(r['Unids'])
        desc     = str(r['Descripción Artículo'])

        # API lookup
        a = api[api['sku'] == sku]
        if len(a):
            cancha_raw = a.iloc[0]['cancha']
            alm_id     = a.iloc[0]['alm_id']
            alm_det    = str(a.iloc[0]['alm_det']) if pd.notna(a.iloc[0]['alm_det']) else 'sin datos'
        else:
            cancha_raw, alm_id, alm_det = None, None, 'sin datos'

        cancha = normalize_cancha(cancha_raw)
        try:
            alm_id_int = int(float(alm_id)) if pd.notna(alm_id) else None
        except:
            alm_id_int = None

        # BXP lookup
        d = ddm[ddm['sku'] == sku]
        bxp = float(d.iloc[0]['bxp']) if len(d) and pd.notna(d.iloc[0]['bxp']) else None

        # Paleta pura
        if bxp and bxp > 0:
            pal_ent  = int(blts_raw / bxp)
            blts_pick = blts_raw - pal_ent * bxp
            has_pal  = pal_ent >= 1
            miss_bxp = False
        else:
            pal_ent, blts_pick, has_pal, miss_bxp = 0, blts_raw, False, True

        # Frescura / color
        skip_color = (sku == 7038) or (alm_id_int == 46)
        f = fr[fr['sku'] == sku]
        if len(f) and not skip_color:
            fv = f.iloc[0]['fecha']
            st_ = str(f.iloc[0]['status']).strip().upper()
            if pd.isna(fv):
                fecha_s, row_st = '', 'NONE'
            else:
                fecha_s = pd.Timestamp(fv).strftime('%d/%m/%y')
                row_st  = st_ if st_ in ('RED', 'YELLOW', 'GREEN') else 'NONE'
        else:
            fecha_s, row_st = '', 'NONE'

        rows.append({
            'sku': sku, 'desc': desc, 'blts_raw': blts_raw,
            'unids': unids, 'blts_pick': blts_pick, 'pal_ent': pal_ent,
            'has_pal': has_pal, 'miss_bxp': miss_bxp, 'bxp': bxp,
            'cancha': cancha, 'alm_id': alm_id_int, 'alm_det': alm_det,
            'fecha_s': fecha_s, 'row_st': row_st,
        })
    return rows

def compute_pall_value(rows, cancha) -> float:
    total = 0.0
    for r in rows:
        if r['cancha'] == cancha and r['bxp'] and r['bxp'] > 0:
            total += r['blts_pick'] / r['bxp']
    return total

# ─── PDF DRAWING HELPERS ─────────────────────────────────────────────────────

def ry(y_from_top):
    """Convert top-down Y to ReportLab bottom-up Y."""
    return PAGE_H - y_from_top

def draw_watermark(c):
    c.saveState()
    c.setFillColor(WM_GRAY)
    c.setFillAlpha(0.09)
    c.setFont('Helvetica-Bold', 72)
    c.translate(PAGE_W / 2, PAGE_H / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, 'BECCACECE HNOS')
    c.restoreState()

def draw_rect_filled(c, x, y_top, w, h, fill_color, stroke=None):
    c.setFillColor(fill_color)
    if stroke:
        c.setStrokeColor(stroke)
        c.rect(x, ry(y_top + h), w, h, fill=1, stroke=1)
    else:
        c.rect(x, ry(y_top + h), w, h, fill=1, stroke=0)

def draw_text(c, x, y_top, text, font='Helvetica', size=8,
              color=colors.black, align='left', max_w=None):
    c.setFont(font, size)
    c.setFillColor(color)
    if max_w and c.stringWidth(text, font, size) > max_w:
        # Truncate
        while text and c.stringWidth(text + '…', font, size) > max_w:
            text = text[:-1]
        text += '…'
    if align == 'center':
        c.drawCentredString(x, ry(y_top), text)
    elif align == 'right':
        c.drawRightString(x, ry(y_top), text)
    else:
        c.drawString(x, ry(y_top), text)

def draw_title_bar(c, y_top, numero, fecha_str) -> float:
    """Dark blue title bar. Returns height consumed."""
    h = 36
    draw_rect_filled(c, MARGIN, y_top, CW, h, DARK_BLUE)
    # Row 1: PLANILLA DE CARGA (centered)
    draw_text(c, MARGIN + CW/2, y_top + 10, 'PLANILLA DE CARGA',
              'Helvetica-Bold', 11, colors.white, 'center')
    # Row 2: Reparto | Fecha
    rep_txt  = f'Reparto Nro: {numero}'
    date_txt = f'Fecha: {fecha_str}'
    draw_text(c, MARGIN + 8,     y_top + 26, rep_txt,  'Helvetica-Bold', 8, colors.white)
    draw_text(c, MARGIN + CW - 8, y_top + 26, date_txt, 'Helvetica-Bold', 8, colors.white, 'right')
    return h

def draw_transport_line(c, y_top, transport, chofer) -> float:
    h = 16
    draw_rect_filled(c, MARGIN, y_top, CW, h, colors.HexColor('#EEF2F8'))
    txt = f'Transporte: {transport} — {chofer}   |   Depósito: 001 — CASA CENTRAL'
    draw_text(c, MARGIN + 6, y_top + 11, txt, 'Helvetica-Bold', 8)
    return h

def draw_partida_regreso(c, y_top) -> float:
    h = 16
    txt = 'F. y H. Est. de Partida: ___________________________   |   F. y H. Est. de Regreso: ___________________________'
    draw_text(c, MARGIN + 6, y_top + 11, txt, 'Helvetica', 8)
    return h

def draw_control_carga(c, y_top) -> float:
    h = 52
    box_h = 38
    # Label
    draw_text(c, MARGIN + CW/2, y_top + 10, 'CONTROL DE CARGA',
              'Helvetica-Bold', 9, colors.black, 'center')
    # Box with two columns
    bx = MARGIN
    by = y_top + 12
    half = CW / 2
    # Outer border
    c.setStrokeColor(BORDER_CLR)
    c.setLineWidth(0.5)
    c.rect(bx, ry(by + box_h), CW, box_h, fill=0, stroke=1)
    # Middle divider
    c.line(bx + half, ry(by), bx + half, ry(by + box_h))
    # Labels
    draw_text(c, bx + half/2,       by + 8, 'Firma del PICKERO',     'Helvetica-Bold', 7.5, colors.black, 'center')
    draw_text(c, bx + half + half/2, by + 8, 'Firma del CONTROLADOR', 'Helvetica-Bold', 7.5, colors.black, 'center')
    return h

def draw_lema(c, y_top, lema_txt) -> float:
    h = 18
    draw_rect_filled(c, MARGIN, y_top, CW, h, DARK_BLUE)
    draw_text(c, MARGIN + CW/2, y_top + 12, f'"{lema_txt}"',
              'Helvetica-Oblique', 8, colors.white, 'center', CW - 16)
    return h

def draw_cancha_header(c, y_top, cancha_name) -> float:
    h = 20
    draw_rect_filled(c, MARGIN, y_top, CW, h, MED_BLUE)
    draw_text(c, MARGIN + 10, y_top + 13, f'◀  {cancha_name}',
              'Helvetica-Bold', 10, colors.white)
    return h

def draw_sin_cancha_header(c, y_top) -> float:
    h = 20
    draw_rect_filled(c, MARGIN, y_top, CW, h, RED_ALERT)
    draw_text(c, MARGIN + CW/2, y_top + 13, '⚠  SIN CANCHA ASIGNADA — REVISAR',
              'Helvetica-Bold', 10, colors.white, 'center')
    return h

def draw_almacen_header(c, y_top, alm_id, alm_det) -> float:
    h = 16
    if alm_id is not None:
        alm_label = f'Almacén {alm_id} — {alm_det}'
    else:
        alm_label = f'Almacén: SIN DATOS — sin asignación en base de datos'
    draw_rect_filled(c, MARGIN, y_top, CW, h, colors.HexColor('#DDEEFF'))
    draw_text(c, MARGIN + 6, y_top + 11, alm_label, 'Helvetica-Bold', 8)
    return h

def draw_table_header(c, y_top, cols, widths) -> float:
    h = 14
    draw_rect_filled(c, MARGIN, y_top, CW, h, HDR_BG, BORDER_CLR)
    c.setStrokeColor(BORDER_CLR)
    c.setLineWidth(0.4)
    x = MARGIN
    for col, w in zip(cols, widths):
        c.rect(x, ry(y_top + h), w, h, fill=0, stroke=1)
        draw_text(c, x + w/2, y_top + 9.5, col, 'Helvetica-Bold', 7, colors.black, 'center')
        x += w
    return h

def draw_product_row(c, y_top, row_data, widths) -> float:
    """Draw one product row. row_data = dict with sku,desc,fecha_s,row_st,blts_pick,unids,has_pal,miss_bxp"""
    h = 14
    sku_str   = str(row_data['sku'])
    desc_str  = str(row_data['desc'])
    blts_str  = str(int(row_data['blts_pick'])) if row_data['blts_pick'] == int(row_data['blts_pick']) else f"{row_data['blts_pick']:.0f}"
    unids_str = str(int(row_data['unids']))
    status    = row_data['row_st']
    fecha_s   = row_data['fecha_s']

    # Row background based on status
    if status == 'RED':
        bg = DARK_ROW
    elif status == 'YELLOW':
        bg = MED_ROW
    elif row_data.get('is_sin_cancha'):
        bg = AMBER
    else:
        bg = colors.white

    # Paleta pura: gray bg if has_pal
    if row_data.get('has_pal') and not row_data.get('is_sin_cancha'):
        bg = colors.HexColor('#EFEFEF') if status not in ('RED', 'YELLOW') else bg

    draw_rect_filled(c, MARGIN, y_top, CW, h, bg, BORDER_CLR)
    c.setStrokeColor(BORDER_CLR)
    c.setLineWidth(0.4)

    x = MARGIN
    cell_data = []
    # SKU col
    sku_disp = f"{sku_str}(!)" if row_data.get('miss_bxp') else sku_str
    cell_data.append((sku_disp, 'center'))
    # Desc col
    cell_data.append((desc_str, 'left'))
    # Venc col
    if fecha_s:
        ind = '■' if status == 'RED' else ('▲' if status == 'YELLOW' else '○')
        venc_disp = f'{ind} {fecha_s}'
    else:
        venc_disp = ''
    cell_data.append((venc_disp, 'center'))
    # Bultos col
    cell_data.append((blts_str, 'center'))
    # Unids col
    cell_data.append((unids_str, 'center'))

    font = 'Helvetica-Bold' if status == 'RED' else 'Helvetica'
    for (txt, align), w in zip(cell_data, widths):
        c.rect(x, ry(y_top + h), w, h, fill=0, stroke=1)
        max_w = w - 4
        if align == 'center':
            draw_text(c, x + w/2, y_top + 9.5, txt, font, 7, colors.black, 'center', max_w)
        else:
            draw_text(c, x + 3, y_top + 9.5, txt, font, 7, colors.black, 'left', max_w)
        # Underline for paleta pura
        if row_data.get('has_pal') and align == 'center' and txt == blts_str:
            tw = c.stringWidth(txt, font, 7)
            cx = x + w/2
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.6)
            c.line(cx - tw/2, ry(y_top + h) + 1, cx + tw/2, ry(y_top + h) + 1)
            c.setLineWidth(0.4)
        x += w
    return h

def draw_total_row(c, y_top, total_blts, total_unids, widths) -> float:
    h = 14
    # Merge first 3 cols for "TOTAL ALMACÉN"
    merged_w = widths[0] + widths[1] + widths[2]
    draw_rect_filled(c, MARGIN, y_top, CW, h, colors.HexColor('#E8E8E8'), BORDER_CLR)
    c.setStrokeColor(BORDER_CLR)
    c.setLineWidth(0.4)
    c.rect(MARGIN, ry(y_top + h), merged_w, h, fill=0, stroke=1)
    draw_text(c, MARGIN + merged_w/2, y_top + 9.5, 'T O T A L  A L M A C É N',
              'Helvetica-Bold', 7.5, colors.black, 'center')
    x = MARGIN + merged_w
    for val, w in [(total_blts, widths[3]), (total_unids, widths[4])]:
        c.rect(x, ry(y_top + h), w, h, fill=0, stroke=1)
        s = str(int(val)) if val == int(val) else f'{val:.0f}'
        draw_text(c, x + w/2, y_top + 9.5, s, 'Helvetica-Bold', 7.5, colors.black, 'center')
        x += w
    return h

def draw_espacio_asignado(c, y_top, cancha_name, pall_val) -> float:
    h = 58
    draw_rect_filled(c, MARGIN, y_top, CW, h, BG_GRAY, BORDER_CLR)
    # Title
    draw_text(c, MARGIN + CW/2, y_top + 10, f'ESPACIO ASIGNADO — {cancha_name}',
              'Helvetica-Bold', 8, colors.black, 'center')
    # Sub-table
    col_w = [CW/3, CW/3, CW/3]
    hdrs  = ['Valor mínimo', 'Valor entero', 'Lectura']
    bx    = MARGIN
    by    = y_top + 14
    row_h = 13
    # Header row
    c.setStrokeColor(BORDER_CLR)
    c.setLineWidth(0.4)
    for hdr, w in zip(hdrs, col_w):
        c.setFillColor(HDR_BG)
        c.rect(bx, ry(by + row_h), w, row_h, fill=1, stroke=1)
        draw_text(c, bx + w/2, by + 9, hdr, 'Helvetica-Bold', 7, colors.black, 'center')
        bx += w
    # Data row
    vm, ve, lec = pallet_reading(pall_val)
    bx = MARGIN
    by2 = by + row_h
    vals = [vm, ve, lec]
    for val, w in zip(vals, col_w):
        c.setFillColor(colors.white)
        c.rect(bx, ry(by2 + row_h), w, row_h, fill=1, stroke=1)
        draw_text(c, bx + w/2, by2 + 9, val, 'Helvetica', 7.5, colors.black, 'center')
        bx += w
    return h

def draw_route_cancha1(c, y_top, dest_cancha) -> float:
    if not dest_cancha:
        return 0
    h = 16
    draw_rect_filled(c, MARGIN, y_top, CW, h, YELLOW_RTE, BORDER_CLR)
    draw_text(c, MARGIN + CW/2, y_top + 11, f'▶ DIRIGIR PALETA A: {dest_cancha}',
              'Helvetica-Bold', 8.5, colors.black, 'center')
    return h

def draw_route_receive(c, y_top) -> float:
    h = 16
    draw_rect_filled(c, MARGIN, y_top, CW, h, LIGHT_GRAY, BORDER_CLR)
    draw_text(c, MARGIN + CW/2, y_top + 11, '◀ RECIBE PALETA DESDE: CANCHA I',
              'Helvetica-Bold', 8.5, colors.black, 'center')
    return h

def draw_footer(c, date_str, page_num):
    y_bottom = PAGE_H - MARGIN + 2
    # Separator line
    c.setStrokeColor(colors.HexColor('#CCCCCC'))
    c.setLineWidth(0.5)
    c.line(MARGIN, ry(y_bottom - 4), MARGIN + CW, ry(y_bottom - 4))
    # Line 1: centered institutional
    draw_text(c, MARGIN + CW/2, y_bottom + 6,
              'Almacen Digital 3.0  |  Beccacece Hnos SA  |  Distribuidor Oficial CMQ — Desde 1963',
              'Helvetica', 7, FOOT_GRAY, 'center')
    # Line 2: date left, page right
    draw_text(c, MARGIN, y_bottom + 14, date_str, 'Helvetica', 7, FOOT_GRAY)
    draw_text(c, MARGIN + CW, y_bottom + 14, f'Página: {page_num}', 'Helvetica', 7, FOOT_GRAY, 'right')

def draw_no_load_msg(c, y_top, cancha) -> float:
    h = 20
    draw_text(c, MARGIN + CW/2, y_top + 13, f'— Sin carga de {cancha} para este camión —',
              'Helvetica-Oblique', 8, colors.HexColor('#888888'), 'center')
    return h

# ─── PAGE GENERATORS ─────────────────────────────────────────────────────────

COL_WIDTHS = [50, CW - 50 - 60 - 52 - 52, 60, 52, 52]  # SKU, Desc, Venc, Bultos, Unids

def draw_cancha_page(c, reparto_rows, cancha, is_first, numero, transport,
                     chofer, lema, fecha_str, date_str, page_num,
                     pall_values, route_dest=None, receives_route=False):
    draw_watermark(c)
    y = MARGIN

    y += draw_title_bar(c, y, numero, fecha_str)
    y += 2
    y += draw_transport_line(c, y, transport, chofer)
    y += 2

    if is_first:
        y += draw_partida_regreso(c, y)
        y += 2
        y += draw_control_carga(c, y)
        y += 4

    y += draw_lema(c, y, lema)
    y += 4
    y += draw_cancha_header(c, y, cancha)
    y += 2

    # Filter rows for this cancha, group by almacén
    c_rows = [r for r in reparto_rows if r['cancha'] == cancha]

    if not c_rows:
        y += draw_no_load_msg(c, y, cancha)
    else:
        # Group by almacén
        alm_groups = {}
        alm_order  = []
        for r in c_rows:
            key = (r['alm_id'], r['alm_det'])
            if key not in alm_groups:
                alm_groups[key] = []
                alm_order.append(key)
            alm_groups[key].append(r)

        # Sort each group by blts_raw DESC (stable sort preserves CAR order on tie)
        for key in alm_order:
            alm_groups[key].sort(key=lambda r: -r['blts_raw'])

        # Table header
        y += draw_table_header(c, y, ['SKU', 'Descripción', 'Venc.', 'Bultos', 'Unids'], COL_WIDTHS)

        for key in alm_order:
            alm_id, alm_det = key
            y += draw_almacen_header(c, y, alm_id, alm_det)
            tot_blts = tot_unids = 0
            for r in alm_groups[key]:
                y += draw_product_row(c, y, r, COL_WIDTHS)
                tot_blts += r['blts_pick']
                tot_unids += r['unids']
            y += draw_total_row(c, y, tot_blts, tot_unids, COL_WIDTHS)
            y += 2

    # Footer elements — placed from bottom upward
    # Reserve space: footer=28, espacio=60, route=18
    footer_top   = PAGE_H - MARGIN - 28
    espacio_h    = 60
    route_h      = 18 if (route_dest or receives_route) else 0

    espacio_top  = footer_top - espacio_h - route_h - 4

    pv = pall_values.get(cancha, 0.0)
    draw_espacio_asignado(c, espacio_top, cancha, pv)

    if route_dest:
        draw_route_cancha1(c, espacio_top + espacio_h + 2, route_dest)
    elif receives_route:
        draw_route_receive(c, espacio_top + espacio_h + 2)

    draw_footer(c, date_str, page_num)

def draw_sin_cancha_page(c, row, numero, transport, chofer, lema,
                          fecha_str, date_str, page_num):
    draw_watermark(c)
    y = MARGIN

    y += draw_title_bar(c, y, numero, fecha_str)
    y += 2
    y += draw_transport_line(c, y, transport, chofer)
    y += 2
    y += draw_lema(c, y, lema)
    y += 4
    y += draw_sin_cancha_header(c, y)
    y += 2
    # Almacén: SIN DATOS
    y += draw_almacen_header(c, y, None, 'sin asignación en base de datos')
    y += 2

    # Single product row with amber bg
    row_copy = dict(row)
    row_copy['is_sin_cancha'] = True
    y += draw_table_header(c, y, ['SKU', 'Descripción', 'Venc.', 'Bultos', 'Unids'], COL_WIDTHS)
    y += draw_product_row(c, y, row_copy, COL_WIDTHS)
    y += 6

    # Note box
    note_txt = f"SKU {row['sku']} no tiene cancha asignada en Frescura / API. Validar con Jefe de Almacén antes de cargar."
    c.setFillColor(colors.HexColor('#FFF3E0'))
    c.setStrokeColor(colors.HexColor('#FFB300'))
    c.setLineWidth(0.6)
    c.rect(MARGIN, ry(y + 22), CW, 22, fill=1, stroke=1)
    draw_text(c, MARGIN + CW/2, y + 14, note_txt,
              'Helvetica', 7.5, colors.HexColor('#C0392B'), 'center', CW - 10)
    y += 26

    draw_footer(c, date_str, page_num)

# ─── MAIN GENERATOR ──────────────────────────────────────────────────────────

def generate_pdf(car_df, api, ddm, fr) -> bytes:
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)

    # Determine report date from CAR
    if 'Fecha Mvto' in car_df.columns:
        raw_date = car_df['Fecha Mvto'].dropna().iloc[0] if len(car_df) else datetime.today()
        try:
            fecha_dt = pd.Timestamp(raw_date)
            fecha_str = fecha_dt.strftime('%d/%m/%Y')
            date_str  = fecha_dt.strftime('%d/%m/%Y')
        except:
            fecha_str = date_str = datetime.today().strftime('%d/%m/%Y')
    else:
        fecha_str = date_str = datetime.today().strftime('%d/%m/%Y')

    page_num     = 0
    stats        = {'repartos': [], 'total_pages': 0, 'sin_cancha_skus': [],
                    'red': [], 'yellow': [], 'pallet_applied': [], 'miss_bxp': [],
                    'no_fecha': [], 'excluded_rows': 0}

    repartos = car_df.groupby('Número', sort=False)

    for rep_idx, (numero, rep_df) in enumerate(repartos):
        transport = str(rep_df['Transporte'].iloc[0])
        chofer    = str(rep_df['Descripción Transporte'].iloc[0])
        lema      = get_lema(transport, rep_idx)

        stats['repartos'].append({'numero': numero, 'transport': transport, 'chofer': chofer})

        # Process rows
        rows = process_reparto(rep_df, api, ddm, fr)

        # Compute pall values per cancha
        pall_values = {ch: compute_pall_value(rows, ch) for ch in CANCHA_ORDER}

        # Routing: CANCHA I fraccional?
        pv1    = pall_values['CANCHA I']
        frac1  = pv1 - int(pv1)
        if frac1 > 0.001:
            pv2 = pall_values['CANCHA II']
            pv4 = pall_values['CANCHA IV']
            if pv2 == 0 and pv4 == 0:
                route_dest = None
                route_recv = None
            elif pv4 > pv2:
                route_dest = 'CANCHA IV'
                route_recv = 'CANCHA IV'
            else:
                route_dest = 'CANCHA II'
                route_recv = 'CANCHA II'
        else:
            route_dest = None
            route_recv = None

        # Gather SIN CANCHA rows
        sin_cancha_rows = [r for r in rows if r['cancha'] == 'SIN CANCHA']

        # Draw 5 cancha pages
        for ci, cancha in enumerate(CANCHA_ORDER):
            page_num += 1
            c.showPage() if page_num > 1 else None  # no-op first page
            is_first = (ci == 0)
            receives = (cancha == route_recv and not is_first)
            draw_cancha_page(
                c, rows, cancha, is_first, numero, transport, chofer,
                lema, fecha_str, date_str, page_num, pall_values,
                route_dest if is_first else None,
                receives
            )

        # Draw SIN CANCHA pages
        for r in sin_cancha_rows:
            page_num += 1
            c.showPage()
            draw_sin_cancha_page(c, r, numero, transport, chofer, lema,
                                  fecha_str, date_str, page_num)
            stats['sin_cancha_skus'].append(f"{r['sku']} — {r['desc']}")

        # Collect stats
        for r in rows:
            if r['row_st'] == 'RED':     stats['red'].append(r['sku'])
            if r['row_st'] == 'YELLOW':  stats['yellow'].append(r['sku'])
            if r['has_pal']:             stats['pallet_applied'].append(r['sku'])
            if r['miss_bxp']:            stats['miss_bxp'].append(r['sku'])
            if not r['fecha_s']:         stats['no_fecha'].append(r['sku'])

    c.save()
    buf.seek(0)
    stats['total_pages'] = page_num
    return buf.read(), stats

# ─── STREAMLIT UI ────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title='Planilla de Carga — Beccacece Hnos SA',
        page_icon='📦',
        layout='centered',
    )

    st.markdown("""
    <style>
    .stApp { background: #f8f9fb; }
    .title-box { background:#1a3a6b; color:white; padding:18px 24px; border-radius:8px; margin-bottom:20px; }
    .title-box h2 { margin:0; font-size:22px; }
    .title-box p  { margin:4px 0 0; font-size:13px; opacity:0.85; }
    .stat-box { background:white; border:1px solid #dde; border-radius:6px; padding:12px 16px; margin:6px 0; font-size:13px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class='title-box'>
      <h2>📦 Planilla de Carga — Generador Automático</h2>
      <p>Beccacece Hnos SA &nbsp;|&nbsp; Almacén Digital 3.0 &nbsp;|&nbsp; v3.2</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        car_file = st.file_uploader('📄 Archivo CAR (exportado de ChessERP)', type=['xlsx'])
    with col2:
        fr_file  = st.file_uploader('🌡️ Frescura 3.0 (actualizado del día)', type=['xlsx'])

    if car_file and fr_file:
        with st.spinner('⚙️ Procesando archivos…'):
            try:
                car_df       = load_car(car_file)
                api, ddm, fr = load_frescura(fr_file)
                pdf_bytes, stats = generate_pdf(car_df, api, ddm, fr)

                st.success(f'✅ PDF generado — {stats["total_pages"]} páginas | {len(stats["repartos"])} repartos')

                # Download button
                fname = f'Planilla_Carga_{datetime.today().strftime("%d%m%Y")}.pdf'
                st.download_button(
                    label='⬇️  Descargar PDF',
                    data=pdf_bytes,
                    file_name=fname,
                    mime='application/pdf',
                    use_container_width=True,
                )

                # Validation panel
                with st.expander('📋 Validación pre-generación', expanded=True):
                    def fmt_list(lst):
                        u = list(dict.fromkeys(lst))
                        return ', '.join(str(x) for x in u[:10]) + (f' + {len(u)-10} más' if len(u) > 10 else '') if u else 'ninguno'

                    rep_ids = ' | '.join(str(r['numero']) for r in stats['repartos'])
                    st.markdown(f"""
```
VALIDACIÓN PRE-GENERACIÓN
──────────────────────────────────────────────────────
✅ Repartos procesados : {len(stats['repartos'])}  (IDs: {rep_ids})
✅ Páginas totales     : {stats['total_pages']}  (5 canchas × {len(stats['repartos'])} repartos + {len(set(stats['sin_cancha_skus']))} hojas SIN CANCHA)
✅ Marca de agua       : BECCACECE HNOS, 8-10% opacidad, sin logo
✅ ESPACIO ASIGNADO    : en cada hoja de cancha
✅ Línea de transporte : en TODAS las páginas
⚠️  SKUs sin fecha Frescura  : {fmt_list(stats['no_fecha'])}
■  SKUs alerta RED          : {fmt_list(stats['red'])}
▲  SKUs alerta YELLOW       : {fmt_list(stats['yellow'])}
📦 SKUs con paleta pura     : {fmt_list(stats['pallet_applied'])}
⚠️  SKUs sin BXP en DDM     : {fmt_list(stats['miss_bxp'])}
⚠️  SKUs SIN CANCHA         : {fmt_list(stats['sin_cancha_skus'])}
──────────────────────────────────────────────────────
```
                    """)

            except Exception as e:
                st.error(f'❌ Error al procesar: {e}')
                st.exception(e)
    else:
        st.info('👆 Subí los dos archivos para generar la Planilla de Carga en PDF.')

        st.markdown('---')
        st.markdown('**Archivos necesarios:**')
        st.markdown('- `CAR_DDMM.xlsx` — exportado de ChessERP del día')
        st.markdown('- `Frescura_3.0_Beccacece_Hnos.xlsx` — hoja actualizada del día')
        st.markdown('**El PDF incluye:** exclusión automática de envases · FEFO · paleta pura · semáforo frescura · canchas I-V · hojas SIN CANCHA · marca de agua · pie institucional')

if __name__ == '__main__':
    main()
