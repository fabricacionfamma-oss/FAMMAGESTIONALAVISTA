import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import tempfile
import os
import calendar
from fpdf import FPDF
from datetime import timedelta

# ==========================================
# 0. CONFIGURACIÓN Y ESTILOS
# ==========================================
st.set_page_config(page_title="Reportes FAMMA", layout="wide", page_icon="📊")

st.markdown("""
<style>
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; }
    .stButton>button { height: 3rem; font-size: 16px; font-weight: bold; }
    .header-style { font-size: 26px; font-weight: bold; margin-bottom: 5px; color: #1F2937; }
</style>
""", unsafe_allow_html=True)

col_title, col_btn = st.columns([4, 1])
with col_title:
    st.markdown('<div class="header-style">📄 Reportes PDF - FAMMA</div>', unsafe_allow_html=True)
    st.write("Seleccione los parámetros para descargar los reportes matemáticamente exactos al Dashboard Web.")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# 1. FUNCIONES AUXILIARES Y CLASE PDF
# ==========================================
class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area; self.fecha_str = fecha_str; self.theme_color = theme_color

    def add_gradient_background(self):
        r1, g1, b1 = 240, 242, 246
        r2, g2, b2 = 215, 220, 225
        h = self.h; w = self.w
        for i in range(int(h * 2)):
            ratio = i / (h * 2)
            r = int(r1 + (r2 - r1) * ratio); g = int(g1 + (g2 - g1) * ratio); b = int(b1 + (b2 - b1) * ratio)
            self.set_fill_color(r, g, b); self.rect(0, i / 2, w, 0.5, 'F')

    def rounded_rect(self, x, y, w, h, r, style=''):
        k = self.k; hp = self.h
        op = 'f' if style == 'F' else 'B' if style in ['FD', 'DF'] else 'S'
        MyArc = 4/3 * ((2 ** 0.5) - 1)
        self._out(f'{(x + r) * k:.2f} {(hp - y) * k:.2f} m')
        xc = x + w - r; yc = y + r
        self._out(f'{xc * k:.2f} {(hp - y) * k:.2f} l')
        self._out(f'{(xc + r * MyArc) * k:.2f} {(hp - y) * k:.2f} {(x + w) * k:.2f} {(hp - yc + r * MyArc) * k:.2f} {(x + w) * k:.2f} {(hp - yc) * k:.2f} c')
        yc = y + h - r
        self._out(f'{(x + w) * k:.2f} {(hp - yc) * k:.2f} l')
        self._out(f'{(x + w) * k:.2f} {(hp - yc - r * MyArc) * k:.2f} {(xc + r * MyArc) * k:.2f} {(hp - y - h) * k:.2f} {xc * k:.2f} {(hp - y - h) * k:.2f} c')
        xc = x + r
        self._out(f'{xc * k:.2f} {(hp - y - h) * k:.2f} l')
        self._out(f'{(xc - r * MyArc) * k:.2f} {(hp - y - h) * k:.2f} {x * k:.2f} {(hp - yc - r * MyArc) * k:.2f} {x * k:.2f} {(hp - yc) * k:.2f} c')
        yc = y + r
        self._out(f'{x * k:.2f} {(hp - yc) * k:.2f} l')
        self._out(f'{x * k:.2f} {(hp - yc + r * MyArc) * k:.2f} {(xc - r * MyArc) * k:.2f} {(hp - y) * k:.2f} {xc * k:.2f} {(hp - y) * k:.2f} c')
        self._out(op)

    def draw_panel(self, x, y, w, h, r=3, bg_color=(255,255,255)):
        self.set_fill_color(210, 210, 210); self.rounded_rect(x + 1.5, y + 1.5, w, h, r, style='F')
        self.set_fill_color(*bg_color); self.set_draw_color(180, 180, 180); self.rounded_rect(x, y, w, h, r, style='DF')

    def draw_kpi_panel(self, x, y, w, h, r=3, bg_color=None):
        bg = bg_color if bg_color else self.theme_color
        self.set_fill_color(200, 200, 200); self.rounded_rect(x + 1.5, y + 1.5, w, h, r, style='F')
        self.set_fill_color(*bg); self.rounded_rect(x, y, w, h, r, style='F')

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

def save_chart(fig, w=600, h=300):
    fig.update_layout(width=w, height=h, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig.write_image(tmp.name, engine="kaleido", scale=2.5); return tmp.name

def format_percentage(val):
    if val > 1.5: val /= 100.0
    return val

# ==========================================
# 2. CARGA Y PROCESAMIENTO DE DATOS EXACTOS
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d 00:00:00')
        fin_str = fecha_fin.strftime('%Y-%m-%d 23:59:59')
        
        # 1. LECTURA EXACTA DEL DASHBOARD (Tablas pre-calculadas de Fumiscor)
        q_global = f"SELECT Month, MAX(COALESCE(Performance, 0)) as Perf, MAX(COALESCE(Availability, 0)) as Disp, MAX(COALESCE(Quality, 0)) as Cal, MAX(COALESCE(Oee, 0)) as OEE, SUM(COALESCE(ProductiveTime, 0)) as T_Op, SUM(COALESCE(DownTime, 0)) as T_Par, SUM(COALESCE(Good, 0)) as Buenas, SUM(COALESCE(Rework, 0)) as RT, SUM(COALESCE(Scrap, 0)) as Scrap FROM PROD_M_06 WHERE Year = {anio} AND Month <= {mes} GROUP BY Month"
        q_factory = f"SELECT Factory, Month, MAX(COALESCE(Performance, 0)) as Perf, MAX(COALESCE(Availability, 0)) as Disp, MAX(COALESCE(Quality, 0)) as Cal, MAX(COALESCE(Oee, 0)) as OEE, SUM(COALESCE(ProductiveTime, 0)) as T_Op, SUM(COALESCE(DownTime, 0)) as T_Par, SUM(COALESCE(Good, 0)) as Buenas, SUM(COALESCE(Rework, 0)) as RT, SUM(COALESCE(Scrap, 0)) as Scrap FROM PROD_M_05 WHERE Year = {anio} AND Month <= {mes} GROUP BY Factory, Month"
        q_line = f"SELECT Factory, Line as Linea, Month, MAX(COALESCE(Performance, 0)) as Perf, MAX(COALESCE(Availability, 0)) as Disp, MAX(COALESCE(Quality, 0)) as Cal, MAX(COALESCE(Oee, 0)) as OEE, SUM(COALESCE(ProductiveTime, 0)) as T_Op, SUM(COALESCE(DownTime, 0)) as T_Par, SUM(COALESCE(Good, 0)) as Buenas, SUM(COALESCE(Rework, 0)) as RT, SUM(COALESCE(Scrap, 0)) as Scrap FROM PROD_M_04 WHERE Year = {anio} AND Month <= {mes} GROUP BY Factory, Line, Month"
        q_machine = f"SELECT c.Name as Máquina, p.Factory, p.Line as Linea, p.Month, MAX(COALESCE(p.Performance, 0)) as Perf, MAX(COALESCE(p.Availability, 0)) as Disp, MAX(COALESCE(p.Quality, 0)) as Cal, MAX(COALESCE(p.Oee, 0)) as OEE, SUM(COALESCE(p.ProductiveTime, 0)) as T_Op, SUM(COALESCE(p.DownTime, 0)) as T_Par, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as RT, SUM(COALESCE(p.Scrap, 0)) as Scrap FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY c.Name, p.Factory, p.Line, p.Month"
        
        # 2. EVENTOS: Cruza CELDA y MANTENIMIENTO ANDON
        q_event = f"""
            SELECT e.Id as Evento_Id, c.Name as Máquina, e.Line as Linea, e.Factory, 
                   e.Started as Inicio, e.Finish as Fin, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], 
                   t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4], 
                   op_celda.Name as Operador_Celda, op_req.Name as Operador_Req, op_resp.Name as Operador_Resp
            FROM EVENT_01 e
            LEFT JOIN CELL c ON e.CellId = c.CellId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            LEFT JOIN EVENT_OPERATOR_01 eo ON e.Id = eo.EventId
            LEFT JOIN OPERATOR op_celda ON eo.OperatorId = op_celda.OperatorId
            LEFT JOIN ANDON_01 a ON e.CellId = a.CellId AND e.Started = a.Started
            LEFT JOIN OPERATOR op_req ON a.RequesterOperatorId = op_req.OperatorId
            LEFT JOIN OPERATOR op_resp ON a.ResponserOperatorId = op_resp.OperatorId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        
        # 3. PIEZAS
        q_piezas_hist = f"SELECT c.Name as Máquina, p.Line as Linea, p.Factory, p.Month, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as RT, SUM(COALESCE(p.Scrap, 0)) as Scrap FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY c.Name, p.Line, p.Factory, p.Month"
        q_piezas_mes = f"SELECT c.Name as Máquina, p.Line as Linea, p.Factory, COALESCE(pr.Code, 'S/C') as Pieza, SUM(COALESCE(p.Scrap, 0)) as Scrap, SUM(COALESCE(p.Rework, 0)) as RT FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId LEFT JOIN PRODUCT pr ON p.ProductId = pr.ProductId WHERE p.Year = {anio} AND p.Month = {mes} GROUP BY c.Name, p.Line, p.Factory, pr.Code"

        # Ejecutamos las consultas
        df_g = conn.query(q_global).fillna(0)
        df_f = conn.query(q_factory).fillna(0)
        df_l = conn.query(q_line).fillna(0)
        df_m = conn.query(q_machine).fillna(0)
        df_ev = conn.query(q_event)
        df_ph = conn.query(q_piezas_hist).fillna(0)
        df_pm = conn.query(q_piezas_mes).fillna(0)

        # Calculamos los ponderadores de las métricas maestras
        def calc_nums(df):
            if df.empty: return df
            for c in ['Perf', 'Disp', 'Cal', 'OEE', 'T_Op', 'T_Par', 'Buenas', 'RT', 'Scrap']:
                if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            if 'Buenas' in df.columns: df['Totales'] = df['Buenas'] + df['RT'] + df['Scrap']
            if 'T_Op' in df.columns: df['T_Plan'] = df['T_Op'] + df['T_Par']
            if 'Perf' in df.columns:
                df['Perf_Num'] = df['Perf'] * df['T_Op']
                df['Disp_Num'] = df['Disp'] * df['T_Plan']
                df['Cal_Num'] = df['Cal'] * df['Totales']
                df['OEE_Num'] = df['OEE'] * df['T_Plan']
            return df
            
        df_g = calc_nums(df_g); df_f = calc_nums(df_f); df_l = calc_nums(df_l); df_m = calc_nums(df_m); df_ph = calc_nums(df_ph)

        # --- LIMPIEZA INTELIGENTE DE EVENTOS ---
        if not df_ev.empty:
            df_ev['Tiempo (Min)'] = pd.to_numeric(df_ev['Tiempo (Min)'], errors='coerce').fillna(0)
            df_ev['Operador_Celda'] = df_ev['Operador_Celda'].fillna('').astype(str)
            df_ev['Operador_Req'] = df_ev['Operador_Req'].fillna('').astype(str)
            df_ev['Operador_Resp'] = df_ev['Operador_Resp'].fillna('').astype(str)
            df_ev['Linea'] = df_ev['Linea'].fillna('SIN LINEA').astype(str)
            df_ev['Factory'] = df_ev['Factory'].fillna('SIN FABRICA').astype(str)

            cols_grupo = [c for c in df_ev.columns if c not in ['Operador_Celda', 'Operador_Req', 'Operador_Resp']]

            def agrupar_nombres(ops):
                n = [str(x).strip() for x in ops.unique() if pd.notna(x) and str(x).strip() != '']
                return ' / '.join(n)

            df_ev = df_ev.groupby(cols_grupo, dropna=False).agg({'Operador_Celda': agrupar_nombres, 'Operador_Req': agrupar_nombres, 'Operador_Resp': agrupar_nombres}).reset_index()

            cols_niveles = [c for c in df_ev.columns if 'Nivel Evento' in c]

            def categorizar_estado(row):
                t = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                if 'PRODUCCION' in t or 'PRODUCCIÓN' in t: return 'Producción'
                if 'PROYECTO' in t: return 'Proyecto'
                if 'BAÑO' in t or 'BANO' in t or 'REFRIGERIO' in t: return 'Descanso'
                if 'PARADA PROGRAMADA' in t: return 'Parada Programada'
                return 'Falla/Gestión'

            def clasificar_macro(row):
                t = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                for cat in ["MANTENIMIENTO", "MATRICERIA", "DISPOSITIVOS", "TECNOLOGIA", "GESTION", "LOGISTICA", "CALIDAD"]:
                    if cat in t: return cat.capitalize()
                return 'Otra Falla/Gestión'

            def obtener_detalle_final(row):
                niveles = [str(row.get(c, '')) for c in cols_niveles]
                validos = [n.strip() for n in niveles if n.strip() and n.strip().lower() not in ['none', 'nan', 'null']]
                if not validos: return "Sin detalle en sistema"
                ultimo = validos[-1].upper()
                if row.get('Estado_Global', '') == 'Falla/Gestión' and row.get('Categoria_Macro', '') != 'Otra Falla/Gestión':
                    return f"[{row.get('Categoria_Macro', '').upper()}] {ultimo}"
                return validos[-1]

            def determinar_operador_final(row):
                resp = row['Operador_Resp']; req = row['Operador_Req']; cel = row['Operador_Celda']
                if resp:
                    reales = [n.strip() for n in resp.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                if req:
                    reales = [n.strip() for n in req.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                if cel:
                    reales = [n.strip() for n in cel.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                return '-'
                
            df_ev['Estado_Global'] = df_ev.apply(categorizar_estado, axis=1)
            df_ev['Categoria_Macro'] = df_ev.apply(clasificar_macro, axis=1)
            df_ev['Detalle_Final'] = df_ev.apply(obtener_detalle_final, axis=1)
            df_ev['Operador'] = df_ev.apply(determinar_operador_final, axis=1)

            df_ev = df_ev[df_ev['Estado_Global'] != 'Proyecto'].copy()

        return df_g, df_f, df_l, df_m, df_ev, df_ph, df_pm
    except Exception as e: 
        st.error(f"Error SQL: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Helper para segmentar los DataFrames según el PDF que estemos armando
def slice_data_for_target(target, area, df_f, df_l, df_m, df_ev, df_ph, df_pm):
    if target == 'GENERAL':
        mask_f = df_f['Factory'].str.contains(area, case=False, na=False)
        df_kpi = df_f[mask_f].copy()
        df_e = df_ev[df_ev['Factory'].str.contains(area, case=False, na=False)].copy()
        df_h = df_ph[df_ph['Factory'].str.contains(area, case=False, na=False)].copy()
        df_p = df_pm[df_pm['Factory'].str.contains(area, case=False, na=False)].copy()
    elif area.upper() == 'ESTAMPADO':
        # Es una LINEA particular
        mask_l = (df_l['Factory'].str.contains(area, case=False, na=False)) & (df_l['Linea'] == target)
        df_kpi = df_l[mask_l].copy()
        df_e = df_ev[(df_ev['Factory'].str.contains(area, case=False, na=False)) & (df_ev['Linea'] == target)].copy()
        df_h = df_ph[(df_ph['Factory'].str.contains(area, case=False, na=False)) & (df_ph['Linea'] == target)].copy()
        df_p = df_pm[(df_pm['Factory'].str.contains(area, case=False, na=False)) & (df_pm['Linea'] == target)].copy()
    else:
        # Es Soldadura -> CELDAS o PRP (Agrupación custom)
        mask_m = df_m['Factory'].str.contains(area, case=False, na=False)
        if target == 'CELDAS':
            mask_m = mask_m & df_m['Máquina'].str.upper().str.contains('CEL', na=False)
            df_e = df_ev[(df_ev['Factory'].str.contains(area, case=False, na=False)) & (df_ev['Máquina'].str.upper().str.contains('CEL', na=False))].copy()
            df_h = df_ph[(df_ph['Factory'].str.contains(area, case=False, na=False)) & (df_ph['Máquina'].str.upper().str.contains('CEL', na=False))].copy()
            df_p = df_pm[(df_pm['Factory'].str.contains(area, case=False, na=False)) & (df_pm['Máquina'].str.upper().str.contains('CEL', na=False))].copy()
        else:
            mask_m = mask_m & (df_m['Máquina'].str.upper().str.contains('PRP', na=False) | df_m['Máquina'].str.upper().str.contains('SOLD', na=False))
            df_e = df_ev[(df_ev['Factory'].str.contains(area, case=False, na=False)) & (df_ev['Máquina'].str.upper().str.contains('PRP', na=False) | df_ev['Máquina'].str.upper().str.contains('SOLD', na=False))].copy()
            df_h = df_ph[(df_ph['Factory'].str.contains(area, case=False, na=False)) & (df_ph['Máquina'].str.upper().str.contains('PRP', na=False) | df_ph['Máquina'].str.upper().str.contains('SOLD', na=False))].copy()
            df_p = df_pm[(df_pm['Factory'].str.contains(area, case=False, na=False)) & (df_pm['Máquina'].str.upper().str.contains('PRP', na=False) | df_pm['Máquina'].str.upper().str.contains('SOLD', na=False))].copy()
            
        df_maq = df_m[mask_m]
        # Ponderación manual para grupos custom
        df_kpi = df_maq.groupby('Month')[['T_Op', 'T_Par', 'T_Plan', 'Totales', 'Perf_Num', 'Disp_Num', 'Cal_Num', 'OEE_Num']].sum().reset_index()
        df_kpi['Perf'] = df_kpi['Perf_Num'] / df_kpi['T_Op'].replace(0, 1)
        df_kpi['Disp'] = df_kpi['Disp_Num'] / df_kpi['T_Plan'].replace(0, 1)
        df_kpi['Cal'] = df_kpi['Cal_Num'] / df_kpi['Totales'].replace(0, 1)
        df_kpi['OEE'] = df_kpi['OEE_Num'] / df_kpi['T_Plan'].replace(0, 1)

    return df_kpi, df_e, df_h, df_p

# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (DISPONIBILIDAD POR ÁREA)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, mes, df_f, df_l, df_m, df_ev, df_ph, df_pm):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0); theme_hex = '#%02x%02x%02x' % theme_color
    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    
    # Detección dinámica de páginas
    if area.upper() == 'ESTAMPADO':
        lineas_activas = df_l[(df_l['Factory'].str.contains(area, case=False, na=False)) & (df_l['Month'] == mes)]['Linea'].dropna().unique().tolist()
        lineas_activas = [str(x).strip() for x in lineas_activas if str(x).strip() not in ['', 'None', 'nan']]
        paginas = ['GENERAL'] + sorted(lineas_activas)
    else:
        paginas = ['GENERAL', 'CELDAS', 'PRP']

    for target in paginas:
        df_kpi, df_e, df_h, df_p = slice_data_for_target(target, area, df_f, df_l, df_m, df_ev, df_ph, df_pm)
        if target != 'GENERAL' and df_kpi.empty and df_e.empty: continue # Salta hojas sin datos
            
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C', fill=True)

        row_actual = df_kpi[df_kpi['Month'] == mes]
        if not row_actual.empty:
            v_oee = format_percentage(row_actual['OEE'].values[0])
            v_perf = format_percentage(row_actual['Perf'].values[0])
            v_disp = format_percentage(row_actual['Disp'].values[0])
            v_cal = format_percentage(row_actual['Cal'].values[0])
        else:
            v_oee = v_perf = v_disp = v_cal = 0

        kpis = {"OEE": {"val": v_oee, "min": 0.75, "max": 0.85}, "PERFORMANCE": {"val": v_perf, "min": 0.80, "max": 0.90}, "DISPONIBILIDAD": {"val": v_disp, "min": 0.75, "max": 0.85}, "CALIDAD": {"val": v_cal, "min": 0.75, "max": 0.85}}
        for i, (lbl, data) in enumerate(kpis.items()):
            v = data["val"]; c_min = data["min"]; c_max = data["max"]
            if v < c_min: bg_col = (231, 76, 60); txt_col = 255
            elif v < c_max: bg_col = (241, 196, 15); txt_col = 0 
            else: bg_col = (46, 204, 113); txt_col = 255

            x = 10 + (i * 68.5)
            pdf.draw_kpi_panel(x, y_kpi:=25, 65, 20, bg_color=bg_col)
            pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(txt_col); pdf.cell(65, 6, lbl, 0, 1, 'C')
            pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{v*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        def add_trend_bar(df_in, col, title, x_pos, y_pos, min_t, max_t):
            if df_in.empty: return
            df_g = df_in.copy()
            if col == 'OEE': df_g['Val'] = df_g['OEE']
            elif col == 'PERFORMANCE': df_g['Val'] = df_g['Perf']
            elif col == 'DISPONIBILIDAD': df_g['Val'] = df_g['Disp']
            elif col == 'CALIDAD': df_g['Val'] = df_g['Cal']
            
            df_g['Val'] = df_g['Val'].apply(format_percentage)
            
            # YTD (Acumulado)
            ytd_v = 0
            if col == 'OEE': ytd_v = df_g['OEE_Num'].sum() / df_g['T_Plan'].sum() if df_g['T_Plan'].sum() > 0 else 0
            elif col == 'PERFORMANCE': ytd_v = df_g['Perf_Num'].sum() / df_g['T_Op'].sum() if df_g['T_Op'].sum() > 0 else 0
            elif col == 'DISPONIBILIDAD': ytd_v = df_g['Disp_Num'].sum() / df_g['T_Plan'].sum() if df_g['T_Plan'].sum() > 0 else 0
            elif col == 'CALIDAD': ytd_v = df_g['Cal_Num'].sum() / df_g['Totales'].sum() if df_g['Totales'].sum() > 0 else 0
            ytd_v = format_percentage(ytd_v)

            def get_c(v): return '#E74C3C' if v < min_t else ('#F1C40F' if v < max_t else '#2ECC71')
            df_g['Mes_Str'] = df_g['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
            df_g['Color'] = df_g['Val'].apply(get_c)
            df_g = pd.concat([df_g, pd.DataFrame([{'Month': 99, 'Mes_Str': 'Acum.', 'Val': ytd_v, 'Color': get_c(ytd_v)}])], ignore_index=True)

            max_y = df_g['Val'].max() if not df_g.empty else 1
            upper_limit = max(1.1, max_y * 1.3) if not pd.isna(max_y) else 1.1

            fig = go.Figure(data=[go.Bar(x=df_g['Mes_Str'], y=df_g['Val'], marker=dict(color=df_g['Color'], line=dict(color='rgba(0,0,0,0.8)', width=2)), text=df_g['Val'], texttemplate='<b>%{text:.1%}</b>', textposition='outside', opacity=0.85)])
            fig.add_hline(y=min_t, line_dash="dash", line_color="#E74C3C", annotation_text=f"<b>{min_t*100:.0f}%</b>", annotation_font_color='black')
            fig.add_hline(y=max_t, line_dash="dash", line_color="#2ECC71", annotation_text=f"<b>{max_t*100:.0f}%</b>", annotation_font_color='black')
            if len(df_g) > 1: fig.add_vline(x=len(df_g) - 1.5, line_width=2, line_dash="dash", line_color="rgba(0,0,0,0.4)")
            
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(t=30, b=20, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, upper_limit], visible=False), xaxis_title="")
            fig.update_traces(textfont=dict(color='black', size=11, family="Arial"), cliponaxis=False)
            img = save_chart(fig, 600, 220); pdf.image(img, x=x_pos+2, y=y_pos+2, w=132); os.remove(img)

        pdf.draw_panel(10, 48, 136, 52); pdf.draw_panel(149, 48, 138, 52)
        add_trend_bar(df_kpi, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48, 0.75, 0.85)
        add_trend_bar(df_kpi, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48, 0.80, 0.90) 
        
        pdf.draw_panel(10, 102, 136, 52); pdf.draw_panel(149, 102, 138, 52)
        add_trend_bar(df_kpi, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 102, 0.75, 0.85)
        add_trend_bar(df_kpi, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 102, 0.75, 0.85)
        
        pdf.draw_panel(10, 156, 136, 45); pdf.draw_panel(149, 156, 138, 45)
        pdf.set_xy(10, 156); pdf.set_font("Times", 'B', 11); pdf.set_text_color(0); pdf.cell(136, 6, "TOP 5 FALLOS", border=0, ln=True, align='C')
        
        df_f = df_e[df_e['Estado_Global'] == 'Falla/Gestión'] if not df_e.empty else pd.DataFrame()
        if not df_f.empty and df_f['Tiempo (Min)'].sum() > 0:
            excluir = ['BAÑO', 'BANO', 'REFRIGERIO', 'DESCANSO']
            df_f_puras = df_f[~df_f['Detalle_Final'].str.upper().apply(lambda x: any(excl in x for excl in excluir))]
            top5 = df_f_puras.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
            
            pdf.set_xy(10, 162); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
            pdf.cell(76, 5, "FALLO", border=1, fill=True); pdf.cell(30, 5, "MINUTOS", border=1, align='C', fill=True); pdf.cell(30, 5, "% TOTAL", border=1, align='C', ln=True, fill=True)
            pdf.set_font("Arial", '', 8); pdf.set_text_color(0); pdf.set_fill_color(255, 255, 255)
            
            t_total = df_f['Tiempo (Min)'].sum()
            for _, r in top5.iterrows():
                pdf.set_x(10); pdf.cell(76, 6, clean_text(str(r['Detalle_Final']))[:45], border=1, fill=True)
                pdf.cell(30, 6, f"{r['Tiempo (Min)']:.0f}", border=1, align='C', fill=True)
                pdf.cell(30, 6, f"{(r['Tiempo (Min)']/t_total)*100:.1f}%", border=1, align='C', ln=True, fill=True)
            
            df_macro = df_f.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
            df_macro['%'] = df_macro['Tiempo (Min)'] / t_total; df_macro['Y'] = "Pérdidas"
            
            fig_stack = px.bar(df_macro, x='%', y='Y', color='Categoria_Macro', orientation='h', color_discrete_sequence=px.colors.qualitative.Safe)
            fig_stack.update_traces(texttemplate='<b>%{x:.1%}</b>', textposition='inside', marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.9, textfont=dict(color='black', size=11))
            fig_stack.update_layout(barmode='stack', title=dict(text="<b>PROPORCIÓN DE PÉRDIDAS ÁREAS MACRO (100%)</b>", font=dict(family="Times", size=13, color="black")), xaxis=dict(visible=False, range=[0, 1]), yaxis=dict(visible=False), margin=dict(t=30, b=5, l=10, r=10), legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5, title="", font=dict(size=10)))
            img_stack = save_chart(fig_stack, 600, 180); pdf.image(img_stack, 151, 158, 134); os.remove(img_stack)
            
    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO (CALIDAD)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, mes, anio, hs_rt, df_f, df_l, df_m, df_ev, df_ph, df_pm):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0); theme_hex = '#%02x%02x%02x' % theme_color
    scrap_c = '#002147' if area.upper() == "ESTAMPADO" else '#722F37'; rt_c = theme_hex
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    
    if area.upper() == 'ESTAMPADO':
        lineas_activas = df_l[(df_l['Factory'].str.contains(area, case=False, na=False)) & (df_l['Month'] == mes)]['Linea'].dropna().unique().tolist()
        lineas_activas = [str(x).strip() for x in lineas_activas if str(x).strip() not in ['', 'None', 'nan']]
        paginas = ['GENERAL'] + sorted(lineas_activas)
    else:
        paginas = ['GENERAL', 'CELDAS', 'PRP']

    for target in paginas:
        df_kpi, df_e, df_h, df_p = slice_data_for_target(target, area, df_f, df_l, df_m, df_ev, df_ph, df_pm)
        if target != 'GENERAL' and df_h.empty and df_p.empty: continue
            
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(20, 6, "MES", 1, 0, 'C', fill=True); pdf.cell(20, 6, "AÑO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "AREA", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(20, 6, str(mes), 1, 0, 'C', fill=True); pdf.cell(20, 6, str(anio), 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C', fill=True)

        if df_h.empty: continue
        
        df_ev_chart = df_h.groupby('Month')[['Buenas', 'Scrap', 'RT', 'Totales']].sum().reset_index()
        df_ev_chart['Totales_Div'] = df_ev_chart['Totales'].apply(lambda x: x if x > 0 else 1)
        df_ev_chart['% Scrap'] = ((df_ev_chart['Scrap'] / df_ev_chart['Totales_Div']) * 100).round(2)
        df_ev_chart['% RT'] = ((df_ev_chart['RT'] / df_ev_chart['Totales_Div']) * 100).round(2)
        df_ev_chart['Mes_Str'] = df_ev_chart['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})

        f1 = px.bar(df_ev_chart, x='Mes_Str', y='Totales', color_discrete_sequence=[theme_hex]); f1.update_traces(texttemplate='<b>%{y:.3s}</b>')
        f2 = px.bar(df_ev_chart, x='Mes_Str', y='% Scrap', color_discrete_sequence=[theme_hex]); f2.update_traces(texttemplate='<b>%{y:.2f}%</b>')
        f3 = px.bar(df_ev_chart, x='Mes_Str', y='% RT', color_discrete_sequence=[theme_hex]); f3.update_traces(texttemplate='<b>%{y:.2f}%</b>')
        
        titles = ["PIEZAS PRODUCIDAS MES A MES", "% DE SCRAP MES A MES", "% DE RT MES A MES"]
        for i, f in enumerate([f1, f2, f3]): 
            max_y = df_ev_chart['Totales'].max() if i==0 else (df_ev_chart['% Scrap'].max() if i==1 else df_ev_chart['% RT'].max())
            upper_limit = max(0.2, max_y * 1.3)
            f.update_yaxes(range=[0, upper_limit])
            f.update_layout(title=dict(text=f"<b>{titles[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=10, t=30, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis=dict(visible=False))
            f.update_traces(textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

        h_box = 60; pdf.draw_panel(10, 22, 135, h_box); pdf.draw_panel(10, 85, 135, h_box); pdf.draw_panel(10, 148, 135, h_box)
        i1 = save_chart(f1, w=550, h=260); pdf.image(i1, x=11, y=23, w=133, h=h_box-2); os.remove(i1)
        i2 = save_chart(f2, w=550, h=260); pdf.image(i2, x=11, y=86, w=133, h=h_box-2); os.remove(i2)
        i3 = save_chart(f3, w=550, h=260); pdf.image(i3, x=11, y=149, w=133, h=h_box-2); os.remove(i3)

        h_br = 83.5; pdf.draw_panel(150, 22, 135, h_br); pdf.draw_panel(150, 108.5, 135, h_br)
        if not df_p.empty:
            t_s = df_p.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index().sort_values('Scrap', ascending=True)
            t_rt = df_p.groupby('Pieza')['RT'].sum().nlargest(5).reset_index().sort_values('RT', ascending=True)
            
            f4 = px.bar(t_s, x='Scrap', y='Pieza', orientation='h', color_discrete_sequence=[scrap_c])
            f5 = px.bar(t_rt, x='RT', y='Pieza', orientation='h', color_discrete_sequence=[rt_c])
            
            titles_right = ["TOP 5 SCRAP POR PIEZA", "TOP 5 RT POR PIEZA"]
            for i, f in enumerate([f4, f5]):
                max_x = t_s['Scrap'].max() if i==0 else t_rt['RT'].max()
                upper_limit = max_x * 1.3 if max_x > 0 else 1
                f.update_xaxes(range=[0, upper_limit])
                f.update_layout(title=dict(text=f"<b>{titles_right[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=30, t=35, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False), yaxis=dict(title="", automargin=True, tickfont=dict(color='black', size=10)))
                f.update_traces(texttemplate='<b>%{x}</b>', textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

            i4 = save_chart(f4, w=550, h=330); pdf.image(i4, x=151, y=23, w=133, h=h_br-2); os.remove(i4)
            i5 = save_chart(f5, w=550, h=330); pdf.image(i5, x=151, y=109.5, w=133, h=h_br-2); os.remove(i5)
            
        if target == 'GENERAL':
            pdf.draw_panel(150, 196, 135, 12, 2, (240,240,240)); pdf.set_xy(150, 196); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0); pdf.cell(67.5, 12, "HS DE RT", 0, 0, 'C')
            pdf.draw_panel(217.5, 196, 67.5, 12, 2, (255,255,255)); pdf.set_xy(217.5, 196); pdf.cell(67.5, 12, f"{hs_rt:.1f}", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')


# ==========================================
# 5. MOTOR: OEE GENERAL DE PLANTA (GLOBAL)
# ==========================================
def crear_pdf_oee_general(label_reporte, df_g, mes):
    theme_color = (44, 62, 80)
    pdf = ReportePDF("GLOBAL PLANTA FAMMA", label_reporte, theme_color)
    pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()

    pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
    pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA FAMMA - GENERAL", 1, 0, 'C', fill=True); pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
    pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
    pdf.cell(40, 6, label_reporte, 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "DISPONIBILIDAD GLOBAL", 1, 1, 'C', fill=True)

    row_actual = df_g[df_g['Month'] == mes]
    if not row_actual.empty:
        v_oee = format_percentage(row_actual['OEE'].values[0])
        v_perf = format_percentage(row_actual['Perf'].values[0])
        v_disp = format_percentage(row_actual['Disp'].values[0])
        v_cal = format_percentage(row_actual['Cal'].values[0])
    else:
        v_oee = v_perf = v_disp = v_cal = 0

    kpis = {"OEE": {"val": v_oee, "min": 0.75, "max": 0.85}, "PERFORMANCE": {"val": v_perf, "min": 0.80, "max": 0.90}, "DISPONIBILIDAD": {"val": v_disp, "min": 0.75, "max": 0.85}, "CALIDAD": {"val": v_cal, "min": 0.75, "max": 0.85}}
    for i, (lbl, data) in enumerate(kpis.items()):
        v = data["val"]; c_min = data["min"]; c_max = data["max"]
        if v < c_min: bg_col = (231, 76, 60); txt_col = 255
        elif v < c_max: bg_col = (241, 196, 15); txt_col = 0
        else: bg_col = (46, 204, 113); txt_col = 255

        x = 10 + (i * 68.5)
        pdf.draw_kpi_panel(x, y_kpi:=25, 65, 20, bg_color=bg_col)
        pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(txt_col); pdf.cell(65, 6, lbl, 0, 1, 'C')
        pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{v*100:.1f}%", 0, 0, 'C')

    pdf.set_text_color(0)

    def add_global_trend_bar(df_in, col, title, x_pos, y_pos, min_t, max_t, w_panel, h_panel):
        if df_in.empty: return
        df_plot = df_in.copy()
        
        if col == 'OEE': df_plot['Val'] = df_plot['OEE']
        elif col == 'PERFORMANCE': df_plot['Val'] = df_plot['Perf']
        elif col == 'DISPONIBILIDAD': df_plot['Val'] = df_plot['Disp']
        elif col == 'CALIDAD': df_plot['Val'] = df_plot['Cal']
        df_plot['Val'] = df_plot['Val'].apply(format_percentage)

        # Acumulado
        ytd_v = 0
        if col == 'OEE': ytd_v = df_plot['OEE_Num'].sum() / df_plot['T_Plan'].sum() if df_plot['T_Plan'].sum() > 0 else 0
        elif col == 'PERFORMANCE': ytd_v = df_plot['Perf_Num'].sum() / df_plot['T_Op'].sum() if df_plot['T_Op'].sum() > 0 else 0
        elif col == 'DISPONIBILIDAD': ytd_v = df_plot['Disp_Num'].sum() / df_plot['T_Plan'].sum() if df_plot['T_Plan'].sum() > 0 else 0
        elif col == 'CALIDAD': ytd_v = df_plot['Cal_Num'].sum() / df_plot['Totales'].sum() if df_plot['Totales'].sum() > 0 else 0
        ytd_v = format_percentage(ytd_v)

        def get_c(v): return '#E74C3C' if v < min_t else ('#F1C40F' if v < max_t else '#2ECC71')
        df_plot['Mes_Str'] = df_plot['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
        df_plot['Color'] = df_plot['Val'].apply(get_c)

        df_plot = pd.concat([df_plot, pd.DataFrame([{'Month': 99, 'Mes_Str': 'Acum.', 'Val': ytd_v, 'Color': get_c(ytd_v)}])], ignore_index=True)

        max_y = df_plot['Val'].max() if not df_plot.empty else 1
        upper_limit = max(1.1, max_y * 1.3) if not pd.isna(max_y) else 1.1

        fig = go.Figure(data=[go.Bar(x=df_plot['Mes_Str'], y=df_plot['Val'], marker=dict(color=df_plot['Color'], line=dict(color='rgba(0,0,0,0.8)', width=2)), text=df_plot['Val'], texttemplate='<b>%{text:.1%}</b>', textposition='outside', opacity=0.85)])
        fig.add_hline(y=min_t, line_dash="dash", line_color="#E74C3C", annotation_text=f"<b>{min_t*100:.0f}%</b>", annotation_font_color='black')
        fig.add_hline(y=max_t, line_dash="dash", line_color="#2ECC71", annotation_text=f"<b>{max_t*100:.0f}%</b>", annotation_font_color='black')
        if len(df_plot) > 1: fig.add_vline(x=len(df_plot) - 1.5, line_width=2, line_dash="dash", line_color="rgba(0,0,0,0.4)")

        fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(t=30, b=20, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, upper_limit], visible=False), xaxis_title="")
        fig.update_traces(textfont=dict(color='black', size=11, family="Arial"), cliponaxis=False)

        img = save_chart(fig, 600, 320)
        pdf.image(img, x=x_pos+2, y=y_pos+2, w=w_panel-4, h=h_panel-4)
        os.remove(img)

    pdf.draw_panel(10, 50, 136, 72); pdf.draw_panel(149, 50, 138, 72)
    add_global_trend_bar(df_g, 'OEE', 'OEE (%) - GLOBAL PLANTA', 10, 50, 0.75, 0.85, 136, 72)
    add_global_trend_bar(df_g, 'PERFORMANCE', 'PERFORMANCE (%) - GLOBAL PLANTA', 149, 50, 0.80, 0.90, 138, 72)

    pdf.draw_panel(10, 126, 136, 72); pdf.draw_panel(149, 126, 138, 72)
    add_global_trend_bar(df_g, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - GLOBAL PLANTA', 10, 126, 0.75, 0.85, 136, 72)
    add_global_trend_bar(df_g, 'CALIDAD', 'CALIDAD (%) - GLOBAL PLANTA', 149, 126, 0.75, 0.85, 138, 72)

    return pdf.output(dest='S').encode('latin-1')


# ==========================================
# 6. INTERFAZ STREAMLIT
# ==========================================
st.write("### 1. Seleccione el Período (Mensual)")
col1, col2 = st.columns(2)
today = pd.to_datetime("today").date()
with col1: 
    m_sel = st.selectbox("Mes", range(1, 13), index=today.month-1)
with col2: 
    a_sel = st.selectbox("Año", [2024, 2025, 2026], index=2)

ini = pd.to_datetime(f"{a_sel}-{m_sel}-01")
fin = pd.to_datetime(f"{a_sel}-{m_sel}-{calendar.monthrange(a_sel, m_sel)[1]}")
lab = f"{m_sel}/{a_sel}"

df_g, df_f, df_l, df_m, df_ev, df_ph, df_pm = fetch_data_from_db(ini, fin, m_sel, a_sel)

st.write("### 2. Datos Manuales (Informe Productivo)")
hs_rt = st.number_input("Horas de RT:", min_value=0.0, max_value=1000.0, value=0.0, step=1.0)

st.divider()
st.write("### 3. Descargar Reportes")
c_g, c_d, c_p = st.columns(3)

with c_g:
    st.markdown("#### 🌍 OEE General de Planta")
    if st.button("Generar OEE Global FAMMA", use_container_width=True):
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF General", crear_pdf_oee_general(lab, df_g, m_sel), "OEE_General_Planta.pdf")

with c_d:
    st.markdown("#### ⚙️ Informe de Disponibilidad (OEE)")
    if st.button("Disponibilidad ESTAMPADO", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Estampado", crear_pdf_gestion_a_la_vista("Estampado", lab, m_sel, df_f, df_l, df_m, df_ev, df_ph, df_pm), "Disp_Estampado.pdf")
    if st.button("Disponibilidad SOLDADURA", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Soldadura", crear_pdf_gestion_a_la_vista("Soldadura", lab, m_sel, df_f, df_l, df_m, df_ev, df_ph, df_pm), "Disp_Soldadura.pdf")

with c_p:
    st.markdown("#### 🏭 Informe Productivo (Calidad)")
    if st.button("Productivo ESTAMPADO", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Estampado", crear_pdf_informe_productivo("Estampado", lab, m_sel, a_sel, hs_rt, df_f, df_l, df_m, df_ev, df_ph, df_pm), "Prod_Estampado.pdf")
    if st.button("Productivo SOLDADURA", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Soldadura", crear_pdf_informe_productivo("Soldadura", lab, m_sel, a_sel, hs_rt, df_f, df_l, df_m, df_ev, df_ph, df_pm), "Prod_Soldadura.pdf")
