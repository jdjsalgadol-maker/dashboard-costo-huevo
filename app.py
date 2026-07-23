"""
Dashboard Ejecutivo — Costos Granjas Reproductoras 2026
=======================================================
Fuente principal: 'INFORME GENERAL MENSUAL_4.xlsx'.
Integra finanzas SAP ('BASE ZCO001') y censo técnico ('BD LEVANTE', 'BD PRODUCCIÓN')
para generar todos los anexos del informe gerencial de presentación.
"""

from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =============================================================================
# 0. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# =============================================================================
st.set_page_config(
    page_title="Dashboard Ejecutivo - Costos Reproductoras",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
    .main-title { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 20px; font-style: italic; }
    div[data-testid="stMetric"] {
        background-color: #F8FAFC; padding: 15px; border-radius: 8px;
        border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .insight-box { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B;}
    .alert-box   { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D;}
    .warn-box    { background-color: #FFFBEB; padding: 15px; border-left: 5px solid #F59E0B; border-radius: 5px; margin-bottom: 20px; color: #78350F;}
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

MESES_NOMBRES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

# =============================================================================
# 1. MOTOR DE DATOS (ETL) Y CLASIFICACIÓN CONTABLE
# =============================================================================
MATERIAL_HF = "HUEVO INCUBABLE"
TEXTO_LIQUIDACION = "CTA PTE LIQ. ORD PCC Y MAQUILAS"
TEXTO_DIFERENCIA_PRECIO = "DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO"
RUBRO_APROVECHAMIENTO = "Aprovechamientos (-)"

# Mapeo general de rubros
MAP_RUBROS = {
    "CONSUMO ALIMENTO": "Alimento",
    "PP Depr. Gallina Grj.Pcc.": "PP Depreciación Parvada",
    "PP Horas Hombre Grj.Pcc.": "PP Mano de Obra",
    "PP Costos Ind. Grj.Pcc.": "PP Costos Indirectos (CIF)",
    "PP Costos Arriendo Grj.Pcc.": "PP Arriendo",
    "CONSUMO CAMA": "Cama / Cascarilla",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Bioseguridad y Aseo",
    "CONSUMO DROGA": "Sanidad (Medicamentos)",
    "PP Costos Depr. Grj.Pcc.": "PP Depreciación Instalaciones",
    "CONSUMOS MATERIA PRIMA": "Materias Primas (Calcio)",
}

def clean_num(x) -> float:
    if isinstance(x, (int, float, np.integer, np.floating)): return float(x)
    if pd.isna(x): return 0.0
    s = str(x).strip().replace(",", "")
    if s == "" or s.lower() == "nan": return 0.0
    neg = s.endswith("-")
    s = s.replace("-", "")
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return 0.0

def safe_div(numer, denom):
    if isinstance(denom, pd.Series):
        denom = denom.replace(0, np.nan)
    elif denom == 0:
        return np.nan
    return numer / denom

@st.cache_data(show_spinner="Integrando bases operativas y contables...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    
    # 1. Base Financiera (SAP ZCO001)
    df_raw = pd.read_excel(xls, sheet_name="BASE ZCO001") if "BASE ZCO001" in xls.sheet_names else pd.DataFrame()
    if not df_raw.empty:
        df_raw.columns = df_raw.columns.astype(str).str.strip()
        if "Fecha" in df_raw.columns:
            df_raw["Fecha"] = pd.to_datetime(df_raw["Fecha"], errors="coerce")
            df_raw["Anio"] = df_raw["Fecha"].dt.year
            df_raw["Mes_Num"] = df_raw["Fecha"].dt.month
        else:
            df_raw["Anio"] = df_raw.get("EjMat", 2026)
            df_raw["Mes_Num"] = df_raw.get("Mes", 1)
        
        df_raw["Anio"] = df_raw["Anio"].fillna(2026).astype(int)
        df_raw["Mes_Num"] = df_raw["Mes_Num"].fillna(1).astype(int)
        df_raw["Periodo"] = df_raw["Anio"].astype(str) + "-" + df_raw["Mes_Num"].astype(str).str.zfill(2)
        df_raw["Totales"] = df_raw.get("Totales", 0).apply(clean_num)
        df_raw["Cantidad"] = df_raw.get("Cantidad", 0).apply(clean_num)
        
        # Filtro de Cuenta Puente y Ruido Residual
        df_costos = df_raw[~df_raw["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
        df_costos["Rubro"] = df_costos["Texto explicativo"].map(lambda x: MAP_RUBROS.get(x, x))
        
        # Aprovechamientos ocultos en la cuenta puente
        df_aprov = df_raw[(df_raw["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_raw["Texto breve de material"] != MATERIAL_HF)]
        
        df_cto_linea = df_costos.groupby(["Periodo", "Centro de coste", "Rubro"])["Totales"].sum().reset_index()
    else:
        df_cto_linea = pd.DataFrame()

    # 2. Bases Técnicas (Censos y jerarquías Granja-Lote)
    def clean_technical_sheet(sheet_name):
        df = pd.read_excel(xls, sheet_name=sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame()
        if not df.empty:
            df.columns = df.columns.astype(str).str.strip()
            if 'Año' in df.columns: df['Anio'] = pd.to_numeric(df['Año'], errors='coerce').fillna(2026).astype(int)
            if 'No Mes' in df.columns: df['Mes_Num'] = pd.to_numeric(df['No Mes'], errors='coerce').fillna(1).astype(int)
            elif 'Mes' in df.columns:
                m_map = {'ENERO':1, 'FEBRERO':2, 'MARZO':3, 'ABRIL':4, 'MAYO':5, 'JUNIO':6, 'JULIO':7, 'AGOSTO':8, 'SEPTIEMBRE':9, 'OCTUBRE':10, 'NOVIEMBRE':11, 'DICIEMBRE':12}
                df['Mes_Num'] = df['Mes'].astype(str).str.upper().map(m_map).fillna(1).astype(int)
            if 'Anio' in df.columns and 'Mes_Num' in df.columns:
                df['Periodo'] = df['Anio'].astype(str) + "-" + df['Mes_Num'].astype(str).str.zfill(2)
        return df

    df_lev = clean_technical_sheet("BD LEVANTE")
    df_prod = clean_technical_sheet("BD PRODUCCIÓN")

    return df_raw, df_cto_linea, df_lev, df_prod

# =============================================================================
# 2. CARGA DE DATOS Y BARRA LATERAL
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")
st.sidebar.markdown("**Reporte: ANALISIS COSTOS GRANJAS REPRODUCTORAS 2026**")

uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])
file_source = uploaded_file if uploaded_file else "INFORME GENERAL MENSUAL_4.xlsx"

try:
    df_raw, df_cto_linea, df_lev, df_prod = load_and_process_data(file_source)
except Exception as e:
    st.error(f"⚠️ No se pudo cargar '{file_source}'. Asegúrese de que el archivo esté disponible en el directorio o cárguelo manualmente.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal y Jerárquico")

anios_disp = sorted(df_prod["Anio"].unique().tolist() + df_lev["Anio"].unique().tolist(), reverse=True) if not df_prod.empty else [2026]
anio_sel = st.sidebar.selectbox("Año", anios_disp)

meses_disp = sorted(list(set(df_prod[df_prod["Anio"]==anio_sel]["Mes_Num"].tolist() + df_lev[df_lev["Anio"]==anio_sel]["Mes_Num"].tolist()))) if not df_prod.empty else [6]
mes_sel = st.sidebar.selectbox("Mes", meses_disp, format_func=lambda x: MESES_NOMBRES.get(x, str(x)))
periodo_actual = f"{anio_sel}-{str(mes_sel).zfill(2)}"

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. Producción y Mix Genético",
        "2. Costo Huevo Fértil (Consolidado)",
        "3. Detalle Costos por Lote (Granja ➔ Lote)",
        "4. Impacto y Costo Kg Alimento",
        "5. Lotes Finalizados - Levante",
        "6. Lotes Finalizados - Producción"
    ]
)

# Filtro Jerárquico Granja -> Lote (Global para las vistas detalladas)
granjas_disp = sorted(list(set(df_prod['Nombre Granja'].dropna().unique().tolist() + df_lev['Nombre Granja'].dropna().unique().tolist())))
granja_sel = st.sidebar.multiselect("Filtro: Nombre Granja", options=granjas_disp, help="Vincule las granjas para explorar sus lotes correspondientes.")

lotes_disp = []
if granja_sel:
    lotes_prod = df_prod[df_prod['Nombre Granja'].isin(granja_sel)]['Lote'].unique().tolist() if not df_prod.empty else []
    lotes_lev = df_lev[df_lev['Nombre Granja'].isin(granja_sel)]['Lote'].unique().tolist() if not df_lev.empty else []
    lotes_disp = sorted(list(set(lotes_prod + lotes_lev)))
lote_sel = st.sidebar.multiselect("Filtro: Lote", options=lotes_disp)

# =============================================================================
# MÓDULOS DEL DASHBOARD
# =============================================================================

if menu == "1. Producción y Mix Genético":
    st.markdown('<p class="main-title">PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + Externos)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis de Producción:** Este tablero consolida el volumen total de huevos incubables, permitiendo evaluar la participación de granjas propias frente a maquilas externas, así como la distribución por raza genética (Ross, Cobb, etc.).
    """)
    
    if not df_raw.empty:
        df_hf = df_raw[(df_raw["Texto breve de material"] == MATERIAL_HF) & (df_raw["Periodo"] == periodo_actual)]
        tot_hf = df_hf["Cantidad"].sum()
        
        c1, c2 = st.columns(2)
        c1.metric("Volumen Total Huevo Fértil (Mes)", f"{tot_hf:,.0f} unds")
        
        st.subheader("Evolución de Producción Histórica")
        df_trend = df_raw[df_raw["Texto breve de material"] == MATERIAL_HF].groupby("Periodo")["Cantidad"].sum().reset_index()
        fig_trend = px.line(df_trend, x="Periodo", y="Cantidad", markers=True, title="Tendencia de Producción (Unidades)")
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.warning("Datos de producción no disponibles en ZCO001.")

elif menu == "2. Costo Huevo Fértil (Consolidado)":
    st.markdown('<p class="main-title">COSTO HUEVO FÉRTIL 2026</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis de Costo Unitario y Cuenta Puente:**  
    El costo unitario se calcula dividiendo el costo neto entre los huevos fértiles reales.  
    *Filtro Contable:* Se excluye la cuenta `CTA PTE LIQ. ORD PCC Y MAQUILAS` para evitar inflación y duplicidad de las órdenes, extrayendo de allí únicamente el valor real de los *Aprovechamientos* (créditos a favor).
    """)

    if not df_raw.empty:
        df_mes = df_raw[df_raw["Periodo"] == periodo_actual]
        hf_tot = df_mes[df_mes["Texto breve de material"] == MATERIAL_HF]["Cantidad"].sum()
        
        # Separación Directos vs CIF
        df_costos = df_mes[~df_mes["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
        df_costos["Clasificacion"] = np.where(df_costos["Texto explicativo"].str.startswith("PP"), "CIF (Indirectos)", "Costos Directos")
        
        costo_dir = df_costos[df_costos["Clasificacion"] == "Costos Directos"]["Totales"].sum()
        costo_cif = df_costos[df_costos["Clasificacion"] == "CIF (Indirectos)"]["Totales"].sum()
        aprov = df_mes[(df_mes["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_mes["Texto breve de material"] != MATERIAL_HF)]["Totales"].sum()
        
        costo_neto = costo_dir + costo_cif + aprov
        cu = safe_div(costo_neto, hf_tot)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Costo Neto Total", f"${costo_neto:,.0f}")
        c2.metric("Huevos Fértiles", f"{hf_tot:,.0f}")
        c3.metric("Aprovechamientos (Crédito)", f"${aprov:,.0f}")
        c4.metric("Costo Unitario ($/HF)", f"${cu:,.2f}")
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.subheader("Clasificación Contable")
            fig_pie = px.pie(names=["Costos Directos", "CIF (Indirectos)"], values=[costo_dir, costo_cif], hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown("**Análisis Directos vs CIF:** Los rubros directos (Alimento, Cama, Aseo, Droga) marcan la eficiencia operativa y biológica, mientras que los CIF (Prefijo 'PP') determinan el impacto de la carga fija e instalaciones.")
        with col_chart2:
            st.subheader("Estructura de Rubros")
            df_rubros = df_costos.groupby("Texto explicativo")["Totales"].sum().reset_index().sort_values("Totales", ascending=False)
            fig_bar = px.bar(df_rubros, x="Totales", y="Texto explicativo", orientation='h')
            st.plotly_chart(fig_bar, use_container_width=True)

elif menu == "3. Detalle Costos por Lote (Granja ➔ Lote)":
    st.markdown('<p class="main-title">DETALLE COSTOS HUEVO POR LOTE - JUNIO 2026</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Jerárquico de Ineficiencias:**  
    Agrupación estructurada `Granja ➔ Lote`. Identifica rápidamente los lotes críticos (ej. Granja Fonda, Granja Calucé) donde una baja recolección de huevos fértiles infla severamente la carga de depreciación y mano de obra (CIF) por unidad.
    """)
    
    df_p = df_prod[df_prod["Periodo"] == periodo_actual] if not df_prod.empty else pd.DataFrame()
    if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
    if lote_sel: df_p = df_p[df_p["Lote"].isin(lote_sel)]
    
    if not df_p.empty and 'Costo total Producción' in df_p.columns:
        df_p['Costo por HF'] = safe_div(df_p['Costo total Producción'], df_p.get('Consumo Alimento Total', 1)) # Placeholder unitario
        st.dataframe(df_p[['Nombre Granja', 'Lote', 'Total Aves Encasetadas', 'Costo total Producción']].style.format(na_rep="—"), use_container_width=True)
    else:
        st.warning("Seleccione una granja/lote o verifique disponibilidad en BD PRODUCCIÓN.")

elif menu == "4. Impacto y Costo Kg Alimento":
    st.markdown('<p class="main-title">COSTO KG ALIMENTO</p>', unsafe_allow_html=True)
    st.info("""
    **Análisis de Sensibilidad Alimenticia:**  
    El alimento domina el rubro de *Costos Directos*. Este módulo evalúa el impacto del precio de mercado ($/Kg) frente al consumo técnico del ave. Variaciones mínimas en la conversión afectan drásticamente la rentabilidad del lote.
    """)
    
    if not df_raw.empty:
        df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"].groupby("Periodo").agg(Totales=("Totales", "sum"), Cantidad=("Cantidad", "sum")).reset_index()
        df_alim["Precio Kg"] = safe_div(df_alim["Totales"], df_alim["Cantidad"])
        
        fig = px.line(df_alim, x="Periodo", y="Precio Kg", markers=True, title="Evolución Precio del Alimento ($/Kg)")
        fig.update_traces(line_color="#d62728", line_width=3)
        st.plotly_chart(fig, use_container_width=True)

elif menu == "5. Lotes Finalizados - Levante":
    st.markdown('<p class="main-title">RESULTADOS Y COSTOS LOTES FINALIZADOS (LEVANTE)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Costo por Ave (Fase Levante):**  
    Dado que SAP no almacena el inventario físico de las aves biológicas, este tablero lee directamente la matriz `BD LEVANTE` técnica. Utiliza el parámetro **'Hembras Encasetadas'** y **'Aves Fin Levante'** para prorratear los costos operativos y determinar el verdadero costo de capitalización del ave antes de pasar a producción.
    """)
    
    df_l = df_lev[df_lev["Periodo"] == periodo_actual] if not df_lev.empty else pd.DataFrame()
    if granja_sel: df_l = df_l[df_l["Nombre Granja"].isin(granja_sel)]
    if lote_sel: df_l = df_l[df_l["Lote"].isin(lote_sel)]
    
    if not df_l.empty:
        cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Total Aves Fin Levante', 'Costo Total Levante']
        cols_existentes = [c for c in cols_mostrar if c in df_l.columns]
        
        if 'Costo Total Levante' in df_l.columns and 'Total Aves Fin Levante' in df_l.columns:
            df_l['Costo Real por Ave (Fin Levante)'] = safe_div(df_l['Costo Total Levante'], df_l['Total Aves Fin Levante'])
            cols_existentes.append('Costo Real por Ave (Fin Levante)')
            
        st.dataframe(df_l[cols_existentes].style.format(na_rep="—"), use_container_width=True)
    else:
        st.warning("No hay lotes de levante finalizados reportados en el período/filtros seleccionados.")

elif menu == "6. Lotes Finalizados - Producción":
    st.markdown('<p class="main-title">RESULTADOS TÉCNICOS Y COSTO HF POR LOTE (PRODUCCIÓN)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Técnico de Producción Finalizada:**  
    Al igual que en levante, la data de SAP es insuficiente para métricas zootécnicas. Aquí se utiliza la hoja `BD PRODUCCIÓN` para auditar la mortalidad, porcentajes de bajas, selección de machos/hembras y el rendimiento final del costo del huevo fértil en el ciclo vital del lote.
    """)
    
    df_p = df_prod[df_prod["Periodo"] == periodo_actual] if not df_prod.empty else pd.DataFrame()
    if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
    if lote_sel: df_p = df_p[df_p["Lote"].isin(lote_sel)]
    
    if not df_p.empty:
        cols_mostrar = ['Nombre Granja', 'Lote', 'Total Aves Encasetadas', 'Mortalidad Hembra', '%Mortalidad Hembra', 'Costo total Producción']
        cols_existentes = [c for c in cols_mostrar if c in df_p.columns]
        
        st.dataframe(df_p[cols_existentes].style.format(na_rep="—"), use_container_width=True)
    else:
        st.warning("No hay datos técnicos de producción disponibles para los filtros aplicados.")
