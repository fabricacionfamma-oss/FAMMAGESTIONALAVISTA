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

# DICCIONARIO OFICIAL FAMMA
MAQUINAS_MAP = {
    # --- ESTAMPADO ---
    "LINEA 1.2": "LINEA 1.2", "LINEA 1.4": "LINEA 1.4", "LINEA 1.5": "LINEA 1.5",
    "LINEA 2": "LINEA 2", "LINEA 3": "LINEA 3", "LINEA 4": "LINEA 4",
    
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
        
        q_oee_m = f"SELECT c.Name as Máquina, p.Performance as Perf_Num, p.Availability as Disp_Num, p.Quality as Cal_Num, p.Oee as OEE_Num, COALESCE(p.ProductiveTime, 0) as T_Operativo, (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month = {mes}"
        df_oee = conn.query(q_oee_m).fillna(0)

        q_pcs_m = f"SELECT c.Name as Máquina, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, SUM(COALESCE(p.Scrap, 0)) as Observadas FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month = {mes} GROUP BY c.Name"
        df_pcs = conn.query(q_pcs_m).fillna(0)

        df_metrics = pd.merge(df_oee, df_pcs, on='Máquina', how='outer').fillna(0) if not df_oee.empty and not df_pcs.empty else (df_oee if not df_oee.empty else df_pcs)

        q_top_m = f"SELECT c.Name as Máquina, COALESCE(pr.Code, 'S/C') as Pieza, SUM(COALESCE(p.Scrap, 0)) as Scrap, SUM(COALESCE(p.Rework, 0)) as RT FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId LEFT JOIN PRODUCT pr ON p.ProductId = pr.ProductId WHERE p.Year = {anio} AND p.Month = {mes} GROUP BY c.Name, pr.Code"
        df_piezas = conn.query(q_top_m).fillna(0)

        q_trend_oee_m = f"SELECT p.Month, c.Name as Máquina, p.Performance as Perf_Num, p.Availability as Disp_Num, p.Quality as Cal_Num, p.Oee as OEE_Num, COALESCE(p.ProductiveTime, 0) as T_Operativo, (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes}"
        q_trend_pcs_m = f"SELECT p.Month, c.Name as Máquina, SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, SUM(COALESCE(p.Scrap, 0)) as Observadas, SUM(COALESCE(p.Good, 0) + COALESCE(p.Rework, 0) + COALESCE(p.Scrap, 0)) as Totales FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY p.Month, c.Name"
        
        df_t_oee = conn.query(q_trend_oee_m).fillna(0)
        df_t_pcs = conn.query(q_trend_pcs_m).fillna(0)
        df_trend = pd.merge(df_t_pcs, df_t_oee, on=['Month', 'Máquina'], how='outer').fillna(0) if not df_t_oee.empty and not df_t_pcs.empty else (df_t_pcs if not df_t_pcs.empty else df_t_oee)

        q_event = f"""
            SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4], t5.Name as [Nivel Evento 5], t6.Name as [Nivel Evento 6]
            FROM EVENT_01 e JOIN CELL c ON e.CellId = c.CellId 
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId 
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            LEFT JOIN EVENTTYPE t5 ON e.EventTypeLevel5 = t5.EventTypeId LEFT JOIN EVENTTYPE t6 ON e.EventTypeLevel6 = t6.EventTypeId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        df_raw = conn.query(q_event)

        # Lógica de Parseo de Eventos (Inversa)
        if not df_raw.empty:
            def parse_event_tree(row):
                niveles = [str(row.get(f'Nivel Evento {i}', '')).strip().upper() for i in range(1, 7)]
                validos = [n for n in niveles if n and n not in ['NONE', 'NAN', 'NULL']]
                if not validos: return 'Falla/Gestión', 'Otra Falla/Gestión', 'Sin detalle'
                
                texto = " > ".join(validos)
                estado = 'Falla/Gestión'
                if any(x in texto for x in ['BAÑO', 'BANO', 'REFRIGERIO', 'DESCANSO']): estado = 'Descanso'
                elif any(x in texto for x in ['PARADA PROGRAMADA', 'SMED']): estado = 'Parada Programada'
                elif 'PRODUCCION' in validos[0]: estado = 'Producción'
                
                macro = 'Otra Falla/Gestión'
                areas = {'MANTENIMIENTO': 'Mantenimiento', 'MATRICERIA': 'Matricería', 'GESTION': 'Gestión', 'LOGISTICA': 'Logística', 'CALIDAD': 'Calidad', 'TECNOLOGIA': 'Tecnología'}
                for n in reversed(validos):
                    for k, v in areas.items():
                        if k in n: macro = v; break
                    if macro != 'Otra Falla/Gestión': break
                
                return estado, macro, f"[{macro.upper()}] {validos[-1]}" if macro != 'Otra Falla/Gestión' else validos[-1]

            df_raw[['Estado_Global', 'Categoria_Macro', 'Detalle_Final']] = df_raw.apply(lambda r: pd.Series(parse_event_tree(r)), axis=1)

        return df_metrics, df_raw, df_trend, df_piezas
    except Exception as e:
        st.error(f"Error SQL: {e}"); return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (CORREGIDO)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw, df_trend):
    if area.upper() == "ESTAMPADO": theme_color = (15, 76, 129); grupos_area = GRUPOS_ESTAMPADO
    elif area.upper() == "SOLDADURA": theme_color = (211, 84, 0); grupos_area = GRUPOS_SOLDADURA
    else: theme_color = (40, 40, 40); grupos_area = GRUPOS_ESTAMPADO + GRUPOS_SOLDADURA

    mapa_limpio = {str(k).strip().upper(): str(v).strip().upper() for k, v in MAQUINAS_MAP.items()}
    grupos_area_upper = [g.upper() for g in grupos_area]
    
    df_m = df_metrics_pdf.copy(); df_t = df_trend.copy(); df_r = df_pdf_raw.copy()
    for d in [df_m, df_t, df_r]: 
        if not d.empty and 'Máquina' in d.columns: d['Grupo'] = d['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('OTRO')
    
    df_m = df_m[df_m['Grupo'].isin(grupos_area_upper)]; df_t = df_t[df_t['Grupo'].isin(grupos_area_upper)]; df_r = df_r[df_r['Grupo'].isin(grupos_area_upper)]
    
    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    paginas = ['GENERAL'] if area.upper() == "GLOBAL" else ['GENERAL'] + [g for g in grupos_area_upper if g in df_m['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        df_m_t = df_m if target == 'GENERAL' else df_m[df_m['Grupo'] == target]
        df_t_t = df_t if target == 'GENERAL' else df_t[df_t['Grupo'] == target]
        df_r_t = df_r if target == 'GENERAL' else df_r[df_r['Grupo'] == target]

        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', True); pdf.cell(40, 6, "INFORME", 1, 1, 'C', True)
        pdf.set_fill_color(255, 255, 255); pdf.set_text_color(0); pdf.set_font("Arial", '', 10)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C', True); pdf.cell(197, 6, "EMPRESA: FAMMA", 1, 0, 'C', True); pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C', True)

        # Calculo de KPI Promedio Ponderado
        tp = df_m_t['T_Planificado'].sum(); to = df_m_t['T_Operativo'].sum()
        v_oee = (df_m_t['OEE_Num'] * df_m_t['T_Planificado']).sum() / tp if tp > 0 else 0
        v_disp = (df_m_t['Disp_Num'] * df_m_t['T_Planificado']).sum() / tp if tp > 0 else 0
        v_perf = (df_m_t['Perf_Num'] * df_m_t['T_Operativo']).sum() / to if to > 0 else 0
        v_cal = (df_m_t['Cal_Num'] * df_m_t['T_Operativo']).sum() / to if to > 0 else 0
        if v_oee > 1.5: v_oee /= 100; v_disp /= 100; v_perf /= 100; v_cal /= 100

        # KPI CONFIG (OBJETIVOS SOLICITADOS)
        kpis = {"OEE": (v_oee, 0.75), "PERFORMANCE": (v_perf, 0.90), "DISPONIBILIDAD": (v_disp, 0.88), "CALIDAD": (v_cal, 0.95)}
        
        for i, (lbl, (v, obj)) in enumerate(kpis.items()):
            bg = (46, 204, 113) if v >= obj else (231, 76, 60) # Verde si cumple, Rojo si no
            pdf.draw_kpi_panel(x := 10 + (i * 68.5), 25, 65, 20, bg_color=bg)
            pdf.set_xy(x, 27); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(255); pdf.cell(65, 6, lbl, 0, 1, 'L')
            pdf.set_xy(x, 33); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{v*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        # Gráficos de Tendencia con Línea de Objetivo
        def add_trend(df_in, col, title, x_pos, y_pos, target_val, large=False):
            if df_in.empty: return
            res = []
            for m, grp in df_in.groupby('Month'):
                tp_m, to_m = grp['T_Planificado'].sum(), grp['T_Operativo'].sum()
                if col == 'OEE': val = (grp['OEE_Num'] * grp['T_Planificado']).sum() / tp_m if tp_m > 0 else 0
                elif col == 'PERFORMANCE': val = (grp['Perf_Num'] * grp['T_Operativo']).sum() / to_m if to_m > 0 else 0
                elif col == 'DISPONIBILIDAD': val = (grp['Disp_Num'] * grp['T_Planificado']).sum() / tp_m if tp_m > 0 else 0
                else: val = (grp['Cal_Num'] * grp['T_Operativo']).sum() / to_m if to_m > 0 else 0
                res.append({'M': int(m), 'V': val/100 if val > 1.5 else val})
            
            df_g = pd.DataFrame(res); df_g['Mes'] = df_g['M'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
            df_g['C'] = df_g['V'].apply(lambda v: '#2ECC71' if v >= target_val else '#E74C3C')
            
            fig = go.Figure(go.Bar(x=df_g['Mes'], y=df_g['V'], marker_color=df_g['C'], text=df_g['V'], texttemplate='<b>%{text:.1%}</b>', textposition='outside'))
            fig.add_hline(y=target_val, line_dash="dash", line_color="#2ECC71", line_width=2, annotation_text=f"<b>Obj: {target_val*100:.0f}%</b>")
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(size=13)), margin=dict(t=35, b=20, l=10, r=10), yaxis=dict(visible=False, range=[0, max(1.1, df_g['V'].max()*1.2)]))
            img = save_chart(fig, 600, 300 if large else 220); pdf.image(img, x_pos+2, y_pos+2, 134 if large else 132); os.remove(img)

        # Posicionamiento de paneles y gráficos
        if area.upper() == "GLOBAL":
            pdf.draw_panel(10, 48, 136, 75); pdf.draw_panel(149, 48, 138, 75); add_trend(df_t_t, 'OEE', 'OEE (%)', 10, 48, 0.75, True); add_trend(df_t_t, 'PERFORMANCE', 'PERFORMANCE (%)', 150, 48, 0.90, True)
            pdf.draw_panel(10, 126, 136, 75); pdf.draw_panel(149, 126, 138, 75); add_trend(df_t_t, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%)', 10, 126, 0.88, True); add_trend(df_t_t, 'CALIDAD', 'CALIDAD (%)', 150, 126, 0.95, True)
        else:
            pdf.draw_panel(10, 48, 136, 52); pdf.draw_panel(149, 48, 138, 52); add_trend(df_t_t, 'OEE', 'OEE (%)', 10, 48, 0.75); add_trend(df_t_t, 'PERFORMANCE', 'PERFORMANCE (%)', 150, 48, 0.90)
            pdf.draw_panel(10, 102, 136, 52); pdf.draw_panel(149, 102, 138, 52); add_trend(df_t_t, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%)', 10, 102, 0.88); add_trend(df_t_t, 'CALIDAD', 'CALIDAD (%)', 150, 102, 0.95)
            
            # Top Fallos
            pdf.draw_panel(10, 156, 136, 45); pdf.draw_panel(149, 156, 138, 45)
            df_f = df_r_t[df_r_t['Estado_Global'] == 'Falla/Gestión']
            if not df_f.empty:
                top5 = df_f.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
                pdf.set_xy(10, 158); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
                pdf.cell(100, 5, "FALLO", 1, 0, 'L', True); pdf.cell(18, 5, "MIN", 1, 0, 'C', True); pdf.cell(18, 5, "%", 1, 1, 'C', True)
                pdf.set_font("Arial", '', 7.5); pdf.set_text_color(0); tt = df_f['Tiempo (Min)'].sum()
                for _, r in top5.iterrows():
                    pdf.set_x(10); pdf.cell(100, 6, clean_text(r['Detalle_Final'])[:65], 1); pdf.cell(18, 6, f"{r['Tiempo (Min)']:.0f}", 1, 0, 'C'); pdf.cell(18, 6, f"{(r['Tiempo (Min)']/tt)*100:.1f}%", 1, 1, 'C')
                
                df_macro = df_f.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
                fig_s = px.bar(df_macro, x='Tiempo (Min)', y=['Pérdidas']*len(df_macro), color='Categoria_Macro', orientation='h', color_discrete_sequence=px.colors.qualitative.Safe)
                fig_s.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2), margin=dict(t=10, b=10), xaxis=dict(visible=False), yaxis=dict(visible=False))
                imgs = save_chart(fig_s, 600, 180); pdf.image(imgs, 151, 158, 134); os.remove(imgs)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, mes_sel, anio_sel, hs_rt):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    target_scrap = 0.50 if area.upper() == "ESTAMPADO" else 0.30
    target_rt = 2.00
    
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    mapa = {str(k).strip().upper(): str(v).strip().upper() for k, v in MAQUINAS_MAP.items()}
    df_t = df_trend.copy(); df_p = df_piezas.copy()
    for d in [df_t, df_p]: d['Grupo'] = d['Máquina'].str.strip().str.upper().map(mapa).fillna('OTRO')
    
    grupos = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    paginas = ['GENERAL'] + [g.upper() for g in grupos if g.upper() in df_t['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.add_gradient_background()
        df_t_t = df_t if target == 'GENERAL' else df_t[df_t['Grupo'] == target]
        df_p_t = df_p if target == 'GENERAL' else df_p[df_p['Grupo'] == target]

        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, f"MES: {mes_sel}", 1, 0, 'C', True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', True); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C', True)
        
        df_ev = df_t_t.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
        df_ev['% Scrap'] = (df_ev['Observadas'] / df_ev['Totales'].replace(0, 1)) * 100
        df_ev['% RT'] = (df_ev['Retrabajo'] / df_ev['Totales'].replace(0, 1)) * 100
        df_ev['Mes'] = df_ev['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})

        def add_prod_chart(y_col, title, y_pos, target_line, is_pct=True):
            color = '#E74C3C' if is_pct and df_ev[y_col].iloc[-1] > target_line else '#0F4C81'
            fig = go.Figure(go.Bar(x=df_ev['Mes'], y=df_ev[y_col], marker_color=color, text=df_ev[y_col], texttemplate='<b>%{text:.2f}%</b>' if is_pct else '<b>%{text:.3s}</b>', textposition='outside'))
            if target_line: fig.add_hline(y=target_line, line_dash="dash", line_color="#E74C3C", annotation_text=f"Obj: {target_line}%")
            fig.update_layout(title=dict(text=f"<b>{title}</b>"), margin=dict(t=30, b=20, l=10, r=10), yaxis=dict(visible=False, range=[0, max(target_line*1.5, df_ev[y_col].max()*1.3) if target_line else None]))
            img = save_chart(fig, 550, 260); pdf.image(img, 11, y_pos, 133); os.remove(img)

        pdf.draw_panel(10, 22, 135, 60); add_prod_chart('Totales', 'PIEZAS TOTALES', 23, None, False)
        pdf.draw_panel(10, 85, 135, 60); add_prod_chart('% Scrap', '% SCRAP', 86, target_scrap)
        pdf.draw_panel(10, 148, 135, 60); add_prod_chart('% RT', '% RE-TRABAJO', 149, target_rt)

        # Top Scrap/RT por Pieza
        if not df_p_t.empty:
            pdf.draw_panel(150, 22, 135, 83); pdf.draw_panel(150, 108, 135, 83)
            ts = df_p_t.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index().sort_values('Scrap')
            tr = df_p_t.groupby('Pieza')['RT'].sum().nlargest(5).reset_index().sort_values('RT')
            f4 = px.bar(ts, x='Scrap', y='Pieza', orientation='h', title="TOP 5 SCRAP", color_discrete_sequence=['#002147'])
            f5 = px.bar(tr, x='RT', y='Pieza', orientation='h', title="TOP 5 RT", color_discrete_sequence=['#D35400'])
            for i, f in enumerate([f4, f5]):
                f.update_layout(margin=dict(t=35, b=20, l=10, r=30), xaxis=dict(visible=False), yaxis=dict(title=""))
                f.update_traces(texttemplate='<b>%{x}</b>', textposition='outside')
                img = save_chart(f, 550, 330); pdf.image(img, 151, 23 if i==0 else 109, 133); os.remove(img)

        if target == 'GENERAL' and area.upper() == 'ESTAMPADO':
            pdf.draw_panel(150, 196, 135, 12, 2, (240,240,240))
            pdf.set_xy(150, 196); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0); pdf.cell(67, 12, "HS RE-TRABAJO", 0, 0, 'C'); pdf.cell(67, 12, f"{hs_rt:.1f}", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Reportes FAMMA")
st.write("### 1. Período")
c1, c2 = st.columns(2)
with c1: m_sel = st.selectbox("Mes", range(1, 13), index=pd.Timestamp.now().month-1)
with c2: a_sel = st.selectbox("Año", [2024, 2025, 2026], index=2)

mes_str = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}[m_sel]
ini = pd.to_datetime(f"{a_sel}-{m_sel}-01"); fin = ini + pd.offsets.MonthEnd(0)

with st.spinner("Cargando datos..."):
    df_m, df_r, df_t, df_p = fetch_data_from_db(ini, fin, m_sel, a_sel)

st.write("### 2. Datos Manuales")
hs_rt = st.number_input("Horas de RT (Estampado):", 0.0, 1000.0, 0.0)

st.divider()
st.write("### 3. Descargas")
cd, cp, cg = st.columns(3)

with cd:
    st.subheader("⚙️ OEE")
    if st.button("Preparar OEE Estampado"): st.session_state['oee_e'] = crear_pdf_gestion_a_la_vista("Estampado", f"{m_sel}/{a_sel}", df_m, df_r, df_t)
    if 'oee_e' in st.session_state: st.download_button("📥 Bajar Estampado", st.session_state['oee_e'], f"FAMMA_OEE_ESTAMPADO_{mes_str}.pdf")
    if st.button("Preparar OEE Soldadura"): st.session_state['oee_s'] = crear_pdf_gestion_a_la_vista("Soldadura", f"{m_sel}/{a_sel}", df_m, df_r, df_t)
    if 'oee_s' in st.session_state: st.download_button("📥 Bajar Soldadura", st.session_state['oee_s'], f"FAMMA_OEE_SOLDADURA_{mes_str}.pdf")

with cp:
    st.subheader("🏭 Producción")
    if st.button("Preparar Prod Estampado"): st.session_state['pr_e'] = crear_pdf_informe_productivo("Estampado", f"{m_sel}/{a_sel}", df_t, df_p, m_sel, a_sel, hs_rt)
    if 'pr_e' in st.session_state: st.download_button("📥 Bajar Prod Estampado", st.session_state['pr_e'], f"FAMMA_PROD_ESTAMPADO_{mes_str}.pdf")
    if st.button("Preparar Prod Soldadura"): st.session_state['pr_s'] = crear_pdf_informe_productivo("Soldadura", f"{m_sel}/{a_sel}", df_t, df_p, m_sel, a_sel, hs_rt)
    if 'pr_s' in st.session_state: st.download_button("📥 Bajar Prod Soldadura", st.session_state['pr_s'], f"FAMMA_PROD_SOLDADURA_{mes_str}.pdf")

with cg:
    st.subheader("🌎 Global")
    if st.button("Preparar Reporte Global"): st.session_state['glob'] = crear_pdf_gestion_a_la_vista("GLOBAL", f"{m_sel}/{a_sel}", df_m, df_r, df_t)
    if 'glob' in st.session_state: st.download_button("📥 Bajar Global", st.session_state['glob'], f"FAMMA_GENERAL_{mes_str}.pdf")
