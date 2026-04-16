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
# 0. CONFIGURACIÓN Y CONSTANTES FAMMA
# ==========================================
st.set_page_config(page_title="Reportes FAMMA", layout="wide", page_icon="📊")

# DICCIONARIO OFICIAL FAMMA (Coincidencia exacta con la Base de Datos)
MAQUINAS_MAP = {
    # --- ESTAMPADO ---
    "LINEA 1.2": "LINEA 1.2",
    "LINEA 1.4": "LINEA 1.4",
    "LINEA 1.5": "LINEA 1.5",
    "LINEA 2": "LINEA 2",
    "LINEA 3": "LINEA 3",
    "LINEA 4": "LINEA 4",
    
    # --- SOLDADURA ---
    "Cell 1 Famma": "CELDAS", "Cell 2 Famma": "CELDAS", "Cell 3 Famma": "CELDAS",
    "Cell 4 Famma": "CELDAS", "Cell 5 Famma": "CELDAS", "Cell 6 Famma": "CELDAS",
    "Cell 7 Famma": "CELDAS", "Cell 8 Famma": "CELDAS", "Cell 9 Famma": "CELDAS",
    "Cell 10 Famma": "CELDAS", "Cell 11 Famma": "CELDAS", "Cell 12 Famma": "CELDAS",
    "Cell 13 Famma": "CELDAS", "Cell 14 Famma": "CELDAS", "Cell 15A Famma": "CELDAS",
    "Cell 15B Famma": "CELDAS", "Cell 16 Famma": "CELDAS", "Cell 17 Famma": "CELDAS",
    
    "PRP 1": "PRP", "PRP 2": "PRP", "PRP 3": "PRP",
    "PRP 4": "PRP", "PRP 5": "PRP", "PRP 6": "PRP",
    
    "MIG 1": "MIG", "MIG 2": "MIG",
}

GRUPOS_ESTAMPADO = ['LINEA 1.2', 'LINEA 1.4', 'LINEA 1.5', 'LINEA 2', 'LINEA 3', 'LINEA 4']
GRUPOS_SOLDADURA = ['CELDAS', 'PRP', 'MIG']

# ==========================================
# 1. FUNCIONES AUXILIARES Y PDF
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

# ==========================================
# 2. CARGA DE DATOS 
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        
        ini_str = fecha_ini.strftime('%Y-%m-%d 00:00:00')
        fin_str = fecha_fin.strftime('%Y-%m-%d 23:59:59')
        
        # --- 1. OEE Y PORCENTAJES ---
        q_oee_m = f"SELECT c.Name as Máquina, p.Performance as Perf_Num, p.Availability as Disp_Num, p.Quality as Cal_Num, p.Oee as OEE_Num, COALESCE(p.ProductiveTime, 0) as T_Operativo, (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month = {mes}"
        df_oee = conn.query(q_oee_m).fillna(0)

        # --- 2. PIEZAS ---
        q_pcs_m = f"SELECT c.Name as Máquina, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, SUM(COALESCE(p.Scrap, 0)) as Observadas FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month = {mes} GROUP BY c.Name"
        df_pcs = conn.query(q_pcs_m).fillna(0)

        if not df_oee.empty and not df_pcs.empty:
            df_metrics = pd.merge(df_oee, df_pcs, on='Máquina', how='outer').fillna(0)
        else:
            df_metrics = df_oee if not df_oee.empty else df_pcs

        # --- 3. TOP 5 PIEZAS (SCRAP/RT) ---
        q_top_m = f"SELECT c.Name as Máquina, COALESCE(pr.Code, 'S/C') as Pieza, SUM(COALESCE(p.Scrap, 0)) as Scrap, SUM(COALESCE(p.Rework, 0)) as RT FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId LEFT JOIN PRODUCT pr ON p.ProductId = pr.ProductId WHERE p.Year = {anio} AND p.Month = {mes} GROUP BY c.Name, pr.Code"
        df_piezas = conn.query(q_top_m).fillna(0)

        # --- 4. TENDENCIAS MENSUALES ---
        q_trend_oee_m = f"SELECT p.Month, c.Name as Máquina, p.Performance as Perf_Num, p.Availability as Disp_Num, p.Quality as Cal_Num, p.Oee as OEE_Num, COALESCE(p.ProductiveTime, 0) as T_Operativo, (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes}"
        q_trend_pcs_m = f"SELECT p.Month, c.Name as Máquina, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, SUM(COALESCE(p.Scrap, 0)) as Observadas, SUM(COALESCE(p.Good, 0) + COALESCE(p.Rework, 0) + COALESCE(p.Scrap, 0)) as Totales FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY p.Month, c.Name"
        
        df_t_oee = conn.query(q_trend_oee_m).fillna(0)
        df_t_pcs = conn.query(q_trend_pcs_m).fillna(0)

        if not df_t_oee.empty and not df_t_pcs.empty:
            df_trend = pd.merge(df_t_pcs, df_t_oee, on=['Month', 'Máquina'], how='outer').fillna(0)
        else:
            df_trend = df_t_pcs if not df_t_pcs.empty else df_t_oee

        # --- 5. EVENTOS: LOGICA AVANZADA DE RECONOCIMIENTO ---
        q_event = f"""
            SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], 
                   t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4],
                   t5.Name as [Nivel Evento 5], t6.Name as [Nivel Evento 6]
            FROM EVENT_01 e 
            LEFT JOIN CELL c ON e.CellId = c.CellId 
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId 
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId 
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId 
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            LEFT JOIN EVENTTYPE t5 ON e.EventTypeLevel5 = t5.EventTypeId
            LEFT JOIN EVENTTYPE t6 ON e.EventTypeLevel6 = t6.EventTypeId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        df_raw = conn.query(q_event)

        # === BLINDAJE Y CONVERSIÓN ===
        cols_metrics = ['Buenas', 'Retrabajo', 'Observadas', 'Perf_Num', 'Disp_Num', 'Cal_Num', 'OEE_Num', 'T_Operativo', 'T_Planificado']
        for c in cols_metrics:
            if c not in df_metrics.columns: df_metrics[c] = 0
            df_metrics[c] = pd.to_numeric(df_metrics[c], errors='coerce').fillna(0)

        for col in ['Month', 'Perf_Num', 'Disp_Num', 'Cal_Num', 'OEE_Num', 'Buenas', 'Retrabajo', 'Observadas', 'Totales', 'T_Operativo', 'T_Planificado']:
            if col not in df_trend.columns: df_trend[col] = 0
            df_trend[col] = pd.to_numeric(df_trend[col], errors='coerce').fillna(0)

        if df_raw.empty: 
            df_raw = pd.DataFrame(columns=['Máquina', 'Tiempo (Min)', 'Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4', 'Estado_Global', 'Categoria_Macro', 'Detalle_Final'])
        else:
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)
            cols_niveles = ['Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4', 'Nivel Evento 5', 'Nivel Evento 6']
            for col in cols_niveles:
                if col in df_raw.columns: df_raw[col] = df_raw[col].fillna('').astype(str)
                else: df_raw[col] = ''

            # Lógica extraída del modelo para clasificar inteligentemente 
            def categorizar_estado(row):
                texto_completo = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                if 'PRODUCCION' in texto_completo or 'PRODUCCIÓN' in texto_completo: return 'Producción'
                if 'PROYECTO' in texto_completo: return 'Proyecto'
                if 'BAÑO' in texto_completo or 'BANO' in texto_completo or 'REFRIGERIO' in texto_completo or 'DESCANSO' in texto_completo: return 'Descanso'
                if 'PARADA PROGRAMADA' in texto_completo or 'SMED' in texto_completo: return 'Parada Programada'
                return 'Falla/Gestión'

            def clasificar_macro(row):
                texto_completo = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                categorias_clave = ["MANTENIMIENTO", "MATRICERIA", "DISPOSITIVOS", "TECNOLOGIA", "GESTION", "LOGISTICA", "CALIDAD"]
                for cat in categorias_clave:
                    if cat in texto_completo:
                        return cat.capitalize()
                return 'Otra Falla/Gestión'

            def obtener_detalle_final(row):
                niveles = [str(row.get(c, '')) for c in cols_niveles]
                validos = [n.strip() for n in niveles if n.strip() and n.strip().lower() not in ['none', 'nan', 'null']]
                
                if not validos: return "Sin detalle en sistema"
                return validos[-1].upper()
                
            df_raw['Estado_Global'] = df_raw.apply(categorizar_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(clasificar_macro, axis=1)
            df_raw['Detalle_Final'] = df_raw.apply(obtener_detalle_final, axis=1)

        return df_metrics, df_raw, df_trend, df_piezas
    except Exception as e: 
        st.error(f"Error SQL Crítico: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (OEE + FALLAS)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw, df_trend):
    if area.upper() == "ESTAMPADO":
        theme_color = (15, 76, 129); grupos_area = GRUPOS_ESTAMPADO
    elif area.upper() == "SOLDADURA":
        theme_color = (211, 84, 0); grupos_area = GRUPOS_SOLDADURA
    else:
        theme_color = (40, 40, 40); grupos_area = GRUPOS_ESTAMPADO + GRUPOS_SOLDADURA

    mapa_limpio = {str(k).strip().upper(): str(v).strip().upper() for k, v in MAQUINAS_MAP.items()}
    grupos_area_upper = [g.upper() for g in grupos_area]
    
    df_m = df_metrics_pdf.copy(); df_t = df_trend.copy(); df_r = df_pdf_raw.copy()
        
    for d in [df_m, df_t, df_r]: 
        if not d.empty and 'Máquina' in d.columns:
            d['Grupo'] = d['Máquina'].astype(str).str.strip().str.upper().map(mapa_limpio).fillna('OTRO')
        else: d['Grupo'] = 'OTRO'
            
    df_m = df_m[df_m['Grupo'].isin(grupos_area_upper)]
    df_t = df_t[df_t['Grupo'].isin(grupos_area_upper)]
    df_r = df_r[df_r['Grupo'].isin(grupos_area_upper)]
    
    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    paginas = ['GENERAL'] if area.upper() == "GLOBAL" else ['GENERAL'] + [g for g in grupos_area_upper if g in df_m['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        
        if target == 'GENERAL':
            df_m_target = df_m; df_t_target = df_t; df_r_target = df_r
        else:
            df_m_target = df_m[df_m['Grupo'] == target]
            df_t_target = df_t[df_t['Grupo'] == target]
            df_r_target = df_r[df_r['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True)
        pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}" if area.upper() != "GLOBAL" else "PLANTA GLOBAL FAMMA - RESUMEN GENERAL", 1, 0, 'C', fill=True)
        pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
        
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C', fill=True)

        valid_m = df_m_target.copy() if not df_m_target.empty else pd.DataFrame()

        t_plan = valid_m['T_Planificado'].sum() if not valid_m.empty else 0
        t_op = valid_m['T_Operativo'].sum() if not valid_m.empty else 0
        
        if t_plan > 0:
            v_oee = (valid_m['OEE_Num'] * valid_m['T_Planificado']).sum() / t_plan
            v_disp = (valid_m['Disp_Num'] * valid_m['T_Planificado']).sum() / t_plan
        else:
            v_oee = valid_m['OEE_Num'].mean() if not valid_m.empty else 0
            v_disp = valid_m['Disp_Num'].mean() if not valid_m.empty else 0

        if t_op > 0:
            v_perf = (valid_m['Perf_Num'] * valid_m['T_Operativo']).sum() / t_op
            v_cal = (valid_m['Cal_Num'] * valid_m['T_Operativo']).sum() / t_op
        else:
            v_perf = valid_m['Perf_Num'].mean() if not valid_m.empty else 0
            v_cal = valid_m['Cal_Num'].mean() if not valid_m.empty else 0
        
        if v_oee > 1.5 or v_perf > 1.5 or v_disp > 1.5:
            v_oee /= 100.0; v_perf /= 100.0; v_disp /= 100.0; v_cal /= 100.0
            
        kpis = {
            "OEE": {"val": v_oee, "min": 0.75, "max": 0.85},
            "PERFORMANCE": {"val": v_perf, "min": 0.80, "max": 0.90},
            "DISPONIBILIDAD": {"val": v_disp, "min": 0.75, "max": 0.85},
            "CALIDAD": {"val": v_cal, "min": 0.75, "max": 0.85}
        }
        
        for i, (lbl, data) in enumerate(kpis.items()):
            v = data["val"]
            if v < data["min"]: bg_col, txt_col = (231, 76, 60), 255
            elif v < data["max"]: bg_col, txt_col = (241, 196, 15), 0
            else: bg_col, txt_col = (46, 204, 113), 255

            x = 10 + (i * 68.5)
            pdf.draw_kpi_panel(x, y_kpi:=25, 65, 20, bg_color=bg_col)
            pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(txt_col); pdf.cell(65, 6, lbl, 0, 1, 'L')
            pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{v*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        def add_trend_bar(df_in, col, title, x_pos, y_pos, min_t, max_t, draw_large=False):
            if df_in.empty: return
            res_mensual = []
            for m, grp in df_in.groupby('Month'):
                s_plan = grp['T_Planificado'].sum(); s_op = grp['T_Operativo'].sum()
                if col == 'OEE': val = (grp['OEE_Num'] * grp['T_Planificado']).sum() / s_plan if s_plan > 0 else grp['OEE_Num'].mean()
                elif col == 'PERFORMANCE': val = (grp['Perf_Num'] * grp['T_Operativo']).sum() / s_op if s_op > 0 else grp['Perf_Num'].mean()
                elif col == 'DISPONIBILIDAD': val = (grp['Disp_Num'] * grp['T_Planificado']).sum() / s_plan if s_plan > 0 else grp['Disp_Num'].mean()
                elif col == 'CALIDAD': val = (grp['Cal_Num'] * grp['T_Operativo']).sum() / s_op if s_op > 0 else grp['Cal_Num'].mean()
                res_mensual.append({'Month': int(m), 'Val': val})
                
            if not res_mensual: return
            df_g = pd.DataFrame(res_mensual)
            if df_g['Val'].max() > 1.5: df_g['Val'] /= 100.0

            s_plan_tot = df_in['T_Planificado'].sum(); s_op_tot = df_in['T_Operativo'].sum()
            ytd_v = 0
            if col == 'OEE': ytd_v = (df_in['OEE_Num'] * df_in['T_Planificado']).sum() / s_plan_tot if s_plan_tot > 0 else df_in['OEE_Num'].mean()
            elif col == 'PERFORMANCE': ytd_v = (df_in['Perf_Num'] * df_in['T_Operativo']).sum() / s_op_tot if s_op_tot > 0 else df_in['Perf_Num'].mean()
            elif col == 'DISPONIBILIDAD': ytd_v = (df_in['Disp_Num'] * df_in['T_Planificado']).sum() / s_plan_tot if s_plan_tot > 0 else df_in['Disp_Num'].mean()
            elif col == 'CALIDAD': ytd_v = (df_in['Cal_Num'] * df_in['T_Operativo']).sum() / s_op_tot if s_op_tot > 0 else df_in['Cal_Num'].mean()
            if ytd_v > 1.5: ytd_v /= 100.0

            def get_c(v): return '#E74C3C' if v < min_t else ('#F1C40F' if v < max_t else '#2ECC71')
            df_g['Mes_Str'] = df_g['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
            df_g['Color'] = df_g['Val'].apply(get_c)
            ytd_row = pd.DataFrame([{'Month': 99, 'Mes_Str': 'Acum.', 'Val': ytd_v, 'Color': get_c(ytd_v)}])
            df_g = pd.concat([df_g, ytd_row], ignore_index=True)

            max_y = df_g['Val'].max() if not df_g.empty else 1
            upper_limit = max(1.1, max_y * 1.3) if not pd.isna(max_y) else 1.1

            fig = go.Figure(data=[go.Bar(x=df_g['Mes_Str'], y=df_g['Val'], marker=dict(color=df_g['Color'], line=dict(color='rgba(0,0,0,0.8)', width=2)), text=df_g['Val'], texttemplate='<b>%{text:.1%}</b>', textposition='outside', opacity=0.85)])
            fig.add_hline(y=min_t, line_dash="dash", line_color="#E74C3C", annotation_text=f"<b>{min_t*100:.0f}%</b>", annotation_font_color='black')
            fig.add_hline(y=max_t, line_dash="dash", line_color="#2ECC71", annotation_text=f"<b>{max_t*100:.0f}%</b>", annotation_font_color='black')
            if len(df_g) > 1: fig.add_vline(x=len(df_g) - 1.5, line_width=2, line_dash="dash", line_color="rgba(0,0,0,0.4)")
            
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(t=30, b=20, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, upper_limit], visible=False), xaxis_title="")
            fig.update_traces(textfont=dict(color='black', size=11, family="Arial"), cliponaxis=False)
            
            w_img, h_img = (600, 300) if draw_large else (600, 220)
            w_pdf = 132 if not draw_large else 134
            img = save_chart(fig, w_img, h_img); pdf.image(img, x=x_pos+2, y=y_pos+2, w=w_pdf); os.remove(img)

        # === DIBUJO DE LOS GRÁFICOS INFERIORES Y FALLAS ===
        if area.upper() == "GLOBAL":
            pdf.draw_panel(10, 48, 136, 75); pdf.draw_panel(149, 48, 138, 75)
            add_trend_bar(df_t_target, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48, 0.75, 0.85, draw_large=True)
            add_trend_bar(df_t_target, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48, 0.80, 0.90, draw_large=True) 
            
            pdf.draw_panel(10, 126, 136, 75); pdf.draw_panel(149, 126, 138, 75)
            add_trend_bar(df_t_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 126, 0.75, 0.85, draw_large=True)
            add_trend_bar(df_t_target, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 126, 0.75, 0.85, draw_large=True)
        else:
            pdf.draw_panel(10, 48, 136, 52); pdf.draw_panel(149, 48, 138, 52)
            add_trend_bar(df_t_target, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48, 0.75, 0.85)
            add_trend_bar(df_t_target, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48, 0.80, 0.90) 
            
            pdf.draw_panel(10, 102, 136, 52); pdf.draw_panel(149, 102, 138, 52)
            add_trend_bar(df_t_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 102, 0.75, 0.85)
            add_trend_bar(df_t_target, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 102, 0.75, 0.85)
            
            # --- TOP 5 FALLOS Y GRÁFICO ACUMULADO 100% ---
            pdf.draw_panel(10, 156, 136, 45); pdf.draw_panel(149, 156, 138, 45)
            pdf.set_xy(10, 156); pdf.set_font("Times", 'B', 11); pdf.set_text_color(0)
            
            titulo_fallos = "TOP 5 FALLOS GENERALES" if target == 'GENERAL' else f"TOP 5 FALLOS - {target}"
            pdf.cell(136, 6, titulo_fallos, border=0, ln=True, align='C')
            
            df_f = df_r_target[df_r_target['Estado_Global'] == 'Falla/Gestión'] if not df_r_target.empty else pd.DataFrame()
            
            if not df_f.empty and df_f['Tiempo (Min)'].sum() > 0:
                excluir = ['BAÑO', 'BANO', 'REFRIGERIO', 'DESCANSO']
                mask_puras = ~df_f['Detalle_Final'].str.upper().apply(lambda x: any(excl in x for excl in excluir))
                df_f_puras = df_f[mask_puras]
                
                top5 = df_f_puras.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
                
                pdf.set_xy(10, 162); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
                # Redistribuimos columnas para ganar espacio en el detalle de la falla
                pdf.cell(100, 5, "FALLO", border=1, fill=True); pdf.cell(18, 5, "MIN", border=1, align='C', fill=True); pdf.cell(18, 5, "%", border=1, align='C', ln=True, fill=True)
                pdf.set_font("Arial", '', 7.5); pdf.set_text_color(0); pdf.set_fill_color(255, 255, 255)
                
                t_total = df_f['Tiempo (Min)'].sum()
                for _, r in top5.iterrows():
                    pdf.set_x(10); pdf.cell(100, 6, clean_text(str(r['Detalle_Final']))[:70], border=1, fill=True)
                    pdf.cell(18, 6, f"{r['Tiempo (Min)']:.0f}", border=1, align='C', fill=True)
                    pdf.cell(18, 6, f"{(r['Tiempo (Min)']/t_total)*100:.1f}%", border=1, align='C', ln=True, fill=True)
                
                # Gráfico Barra Horizontal 100% por Categoria_Macro con Leyenda dinámica (sin etiquetas internas superpuestas)
                df_macro = df_f.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
                df_macro['%'] = df_macro['Tiempo (Min)'] / t_total
                df_macro['Y'] = "Pérdidas"
                # Creamos el texto de la leyenda incorporando el porcentaje
                df_macro['Leyenda'] = df_macro.apply(lambda r: f"{r['Categoria_Macro']} ({r['%']:.1%})", axis=1)
                
                fig_stack = px.bar(df_macro, x='%', y='Y', color='Leyenda', orientation='h', color_discrete_sequence=px.colors.qualitative.Safe)
                fig_stack.update_traces(marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.9)
                fig_stack.update_layout(
                    barmode='stack', 
                    title=dict(text="<b>PROPORCIÓN DE PÉRDIDAS POR ÁREA (100%)</b>", font=dict(family="Times", size=11, color="black")), 
                    xaxis=dict(visible=False, range=[0, 1]), 
                    yaxis=dict(visible=False), 
                    margin=dict(t=25, b=30, l=10, r=10), 
                    legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5, title="", font=dict(size=9))
                )
                
                img_stack = save_chart(fig_stack, 600, 180); pdf.image(img_stack, 151, 158, 134); os.remove(img_stack)
            else:
                pdf.set_xy(10, 175); pdf.set_font("Arial", 'I', 10); pdf.set_text_color(100)
                pdf.cell(136, 6, "Excelente. No hay fallos registrados en este período.", 0, 1, 'C')
                pdf.set_xy(149, 175); pdf.cell(138, 6, "Sin datos de fallas para graficar.", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO (CALIDAD)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, mes_sel, anio_sel, hs_rt):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    theme_hex = '#%02x%02x%02x' % theme_color
    scrap_c = '#002147' if area.upper() == "ESTAMPADO" else '#722F37' 
    rt_c = theme_hex
    grupos = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    
    mapa_limpio = {str(k).strip().upper(): str(v).strip().upper() for k, v in MAQUINAS_MAP.items()}
    grupos_upper = [g.upper() for g in grupos]
    
    df_t = df_trend.copy(); df_p = df_piezas.copy()
    for d in [df_t, df_p]: 
        if not d.empty and 'Máquina' in d.columns: d['Grupo'] = d['Máquina'].astype(str).str.strip().str.upper().map(mapa_limpio).fillna('OTRO')
        else: d['Grupo'] = 'OTRO'

    df_t = df_t[df_t['Grupo'].isin(grupos_upper)]; df_p = df_p[df_p['Grupo'].isin(grupos_upper)]
    paginas = ['GENERAL'] + [g for g in grupos_upper if g in df_t['Grupo'].unique()]

    # Definir valores de objetivo exactos
    if area.upper() == "ESTAMPADO":
        target_scrap = 0.50
        target_rt = 2.00
    else:
        target_scrap = 0.30
        target_rt = 2.00

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        
        if target == 'GENERAL':
            df_t_target = df_t; df_p_target = df_p
        else:
            df_t_target = df_t[df_t['Grupo'] == target]; df_p_target = df_p[df_p['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(20, 6, "MES", 1, 0, 'C', fill=True); pdf.cell(20, 6, "AÑO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "AREA", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(20, 6, str(mes_sel), 1, 0, 'C', fill=True); pdf.cell(20, 6, str(anio_sel), 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C', fill=True)

        if df_t_target.empty: continue
        for col in ['Buenas', 'Observadas', 'Retrabajo', 'Totales']:
            if col in df_t_target.columns: df_t_target[col] = pd.to_numeric(df_t_target[col], errors='coerce').fillna(0)

        df_ev = df_t_target.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
        if 'Month' in df_ev.columns: df_ev['Month'] = df_ev['Month'].astype(int)
        
        df_ev['Totales_Div'] = df_ev['Totales'].apply(lambda x: x if x > 0 else 1)
        df_ev['% Scrap'] = ((df_ev['Observadas'] / df_ev['Totales_Div']) * 100).round(2)
        df_ev['% RT'] = ((df_ev['Retrabajo'] / df_ev['Totales_Div']) * 100).round(2)
        meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
        df_ev['Mes_Str'] = df_ev['Month'].map(meses_map)

        # LÓGICA DE COLORES DE BARRAS SI PASA EL OBJETIVO
        df_ev['Color_Scrap'] = df_ev['% Scrap'].apply(lambda x: '#E74C3C' if x > target_scrap else theme_hex)
        df_ev['Color_RT'] = df_ev['% RT'].apply(lambda x: '#E74C3C' if x > target_rt else theme_hex)

        f1 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['Totales'], marker_color=theme_hex, text=df_ev['Totales'], texttemplate='<b>%{text:.3s}</b>')])
        f2 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['% Scrap'], marker_color=df_ev['Color_Scrap'], text=df_ev['% Scrap'], texttemplate='<b>%{text:.2f}%</b>')])
        f3 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['% RT'], marker_color=df_ev['Color_RT'], text=df_ev['% RT'], texttemplate='<b>%{text:.2f}%</b>')])
        
        titles = ["PIEZAS PRODUCIDAS MES A MES", "% DE SCRAP MES A MES", "% DE RT MES A MES"]
        for i, f in enumerate([f1, f2, f3]): 
            max_y = df_ev['Totales'].max() if i==0 else (df_ev['% Scrap'].max() if i==1 else df_ev['% RT'].max())
            
            if i == 0: 
                upper_limit = max_y * 1.3 if max_y > 0 else 1
            else: 
                current_target = target_scrap if i == 1 else target_rt
                upper_limit = max(0.2, max_y * 1.3, current_target * 1.5)
                f.add_hline(y=current_target, line_dash="dash", line_width=2, line_color="#E74C3C", annotation_text=f"<b>Obj: {current_target}%</b>", annotation_font_color='black')
                
            f.update_yaxes(range=[0, upper_limit])
            f.update_layout(title=dict(text=f"<b>{titles[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=10, t=30, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis=dict(visible=False))
            f.update_traces(textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

        h_box = 60; pdf.draw_panel(10, 22, 135, h_box); pdf.draw_panel(10, 85, 135, h_box); pdf.draw_panel(10, 148, 135, h_box)
        i1 = save_chart(f1, w=550, h=260); pdf.image(i1, x=11, y=23, w=133, h=h_box-2); os.remove(i1)
        i2 = save_chart(f2, w=550, h=260); pdf.image(i2, x=11, y=86, w=133, h=h_box-2); os.remove(i2)
        i3 = save_chart(f3, w=550, h=260); pdf.image(i3, x=11, y=149, w=133, h=h_box-2); os.remove(i3)

        h_br = 83.5; pdf.draw_panel(150, 22, 135, h_br); pdf.draw_panel(150, 108.5, 135, h_br)
        if not df_p_target.empty:
            t_s = df_p_target.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index().sort_values('Scrap', ascending=True)
            t_rt = df_p_target.groupby('Pieza')['RT'].sum().nlargest(5).reset_index().sort_values('RT', ascending=True)
            
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
            
        # El recuadro de HS DE RT solo se renderiza si estamos en la hoja "GENERAL" y el área es "ESTAMPADO"
        if target == 'GENERAL' and area.upper() == 'ESTAMPADO':
            pdf.draw_panel(150, 196, 135, 12, 2, (240,240,240)); pdf.set_xy(150, 196); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0); pdf.cell(67.5, 12, "HS DE RT", 0, 0, 'C')
            pdf.draw_panel(217.5, 196, 67.5, 12, 2, (255,255,255)); pdf.set_xy(217.5, 196); pdf.cell(67.5, 12, f"{hs_rt:.1f}", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Reportes FAMMA")
st.divider()

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

with st.spinner("Conectando con la base de datos de FAMMA..."):
    df_m, df_r, df_t, df_p = fetch_data_from_db(ini, fin, m_sel, a_sel)

with st.expander("🔍 Previsualizar Datos Crudos Extraídos"):
    st.write("**Datos de Máquinas y OEE:**", df_m.head(5) if not df_m.empty else "Vacío")
    st.write("**Datos de Piezas Totales:**", f"{df_p['Scrap'].sum() if not df_p.empty else 0} Scrap | {df_p['RT'].sum() if not df_p.empty else 0} RT")
    st.write("**Eventos extraídos de la BD (Fallas/Paradas):**", df_r.head(10) if not df_r.empty else "Vacío - No hay registros en este mes/año en la tabla EVENT_01")

st.write("### 2. Datos Manuales (Informe Productivo)")
hs_rt = st.number_input("Horas de RT (Solo válido para Estampado General):", min_value=0.0, max_value=1000.0, value=0.0, step=1.0)

st.divider()
st.write("### 3. Preparar y Descargar Reportes")
c_d, c_p, c_g = st.columns(3)

# Lógica de carga On-Demand usando st.session_state
with c_d:
    st.markdown("#### ⚙️ Disponibilidad (OEE)")
    if not df_m.empty:
        if st.button("⚙️ Preparar PDF Estampado"):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_oee_est'] = crear_pdf_gestion_a_la_vista("Estampado", lab, df_m, df_r, df_t)
        if 'pdf_oee_est' in st.session_state:
            st.download_button("📥 Bajar PDF Estampado", data=st.session_state['pdf_oee_est'], file_name="Disp_Estampado.pdf", mime="application/pdf")
            
        st.write("---")
        
        if st.button("⚙️ Preparar PDF Soldadura"):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_oee_sol'] = crear_pdf_gestion_a_la_vista("Soldadura", lab, df_m, df_r, df_t)
        if 'pdf_oee_sol' in st.session_state:
            st.download_button("📥 Bajar PDF Soldadura", data=st.session_state['pdf_oee_sol'], file_name="Disp_Soldadura.pdf", mime="application/pdf")
    else:
        st.error("No hay datos.")

with c_p:
    st.markdown("#### 🏭 Informe Productivo (Calidad)")
    if not df_t.empty:
        if st.button("🏭 Preparar Prod. Estampado"):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_prod_est'] = crear_pdf_informe_productivo("Estampado", lab, df_t, df_p, m_sel, a_sel, hs_rt)
        if 'pdf_prod_est' in st.session_state:
            st.download_button("📥 Bajar Prod. Estampado", data=st.session_state['pdf_prod_est'], file_name="Prod_Estampado.pdf", mime="application/pdf")
        
        st.write("---")
        
        if st.button("🏭 Preparar Prod. Soldadura"):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_prod_sol'] = crear_pdf_informe_productivo("Soldadura", lab, df_t, df_p, m_sel, a_sel, hs_rt)
        if 'pdf_prod_sol' in st.session_state:
            st.download_button("📥 Bajar Prod. Soldadura", data=st.session_state['pdf_prod_sol'], file_name="Prod_Soldadura.pdf", mime="application/pdf")
    else:
        st.error("No hay datos.")

with c_g:
    st.markdown("#### 🌎 Reporte Maestro")
    if not df_m.empty:
        if st.button("🌎 Preparar PDF Global"):
            with st.spinner("Generando documento maestro..."):
                st.session_state['pdf_oee_glob'] = crear_pdf_gestion_a_la_vista("GLOBAL", lab, df_m, df_r, df_t)
        if 'pdf_oee_glob' in st.session_state:
            st.download_button("📥 Bajar PDF Global", data=st.session_state['pdf_oee_glob'], file_name="Disp_Global_Famma.pdf", mime="application/pdf")
    else:
        st.error("No hay datos.")
