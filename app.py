"""
Dashboard Ejecutivo — Costos Granjas Reproductoras 2026
=======================================================
Fuente principal: 'INFORME GENERAL MENSUAL_4.xlsx'.
Este script procesa nativamente las hojas 'BD', 'BD CTO LINEA', 'BD LEVANTE' y 'BD PRODUCCIÓN'
para generar todos los anexos del informe gerencial en PowerPoint.
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

def safe_div(numer, denom):
    if isinstance(denom, pd.Series):
        denom = denom.replace(0, np.nan)
    elif denom == 0:
        return np.nan
    return numer / denom

# =============================================================================
# 1. MOTOR DE DATOS (ETL) NATIVO PARA INFORME GENERAL MENSUAL
# =============================================================================
@st.cache_data(show_spinner="Integrando bases operativas y contables (SAP)...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    
    def normalize_dates(df, date_col='fecha'):
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            df["Anio"] = df[date_col].dt.year.fillna(2026).astype(int)
            df["Mes_Num"] = df[date_col].dt.month.fillna(1).astype(int)
            df["Periodo"] = df["Anio"].astype(str) + "-" + df["Mes_Num"].astype(str).str.zfill(2)
        return df

    # 1. Base Macro Financiera (BD)
    df_bd = pd.read_excel(xls, sheet_name="BD") if "BD" in xls.sheet_names else pd.DataFrame()
    if not df_bd.empty:
        df_bd.columns = df_bd.columns.astype(str).str.strip()
        df_bd = normalize_dates(df_bd, 'Fecha')
        for c in ['Producción HF', 'Externos HF', 'Costo total', 'Alimento ', 'KG cons Alimento']:
            if c in df_bd.columns:
                df_bd[c] = pd.to_numeric(df_bd[c], errors='coerce').fillna(0)

    # 2. Base Detalle por Lote/Línea (BD CTO LINEA)
    df_cto = pd.read_excel(xls, sheet_name="BD CTO LINEA") if "BD CTO LINEA" in xls.sheet_names else pd.DataFrame()
    if not df_cto.empty:
        df_cto.columns = df_cto.columns.astype(str).str.strip()
        df_cto = normalize_dates(df_cto, 'fecha')
        cols_num = ['ALIMENTO', 'CAMA', 'DROGA', 'MATERIA PRIMA', 'ELEMENTOS DE ASEO Y DESINFECCION', 
                    'ARRIENDO', 'DEPRECIACIÓN CONST Y EDIF', 'INDIRECTOS', 'DEPRECIACIÓN GALLINA', 
                    'MANO DE OBRA', 'SUBPRODUCTOS', 'Huevos fértiles', 'TOTAL COSTO', 'consumo alimento kg']
        for c in cols_num:
            if c in df_cto.columns:
                df_cto[c] = pd.to_numeric(df_cto[c], errors='coerce').fillna(0)
        
        # Agrupación CIF vs Directos en la base de detalle
        dir_cols = ['ALIMENTO', 'CAMA', 'DROGA', 'MATERIA PRIMA', 'ELEMENTOS DE ASEO Y DESINFECCION']
        cif_cols = ['DEPRECIACIÓN CONST Y EDIF', 'INDIRECTOS', 'DEPRECIACIÓN GALLINA', 'MANO DE OBRA', 'ARRIENDO']
        df_cto['Costos Directos'] = df_cto[[c for c in dir_cols if c in df_cto.columns]].sum(axis=1)
        df_cto['Costos CIF'] = df_cto[[c for c in cif_cols if c in df_cto.columns]].sum(axis=1)

    # 3. Producción por Raza (BD PN HF RAZA)
    df_raza = pd.read_excel(xls, sheet_name="BD PN HF RAZA") if "BD PN HF RAZA" in xls.sheet_names else pd.DataFrame()
    if not df_raza.empty:
        df_raza.columns = df_raza.columns.astype(str).str.strip()
        df_raza = normalize_dates(df_raza, 'FECHA')

    # 4. Bases Técnicas (Levante y Producción)
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

    return df_bd, df_cto, df_raza, df_lev, df_prod

# =============================================================================
# 2. CARGA DE DATOS Y BARRA LATERAL
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")
st.sidebar.markdown("**Reporte: ANALISIS COSTOS GRANJAS REPRODUCTORAS 2026**")

uploaded_file = st.sidebar.file_uploader("Subir INFORME GENERAL MENSUAL_4.xlsx", type=["xlsx", "xls"])

try:
    if uploaded_file:
        df_bd, df_cto, df_raza, df_lev, df_prod = load_and_process_data(uploaded_file)
    else:
        # Fallback to local path for seamless testing
        df_bd, df_cto, df_raza, df_lev, df_prod = load_and_process_data("INFORME GENERAL MENSUAL_4.xlsx")
except Exception as e:
    st.error("⚠️ Por favor, carga tu archivo 'INFORME GENERAL MENSUAL_4.xlsx' en el panel lateral para iniciar el dashboard.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal y Jerárquico")

anios_disp = sorted(df_cto["Anio"].unique().tolist(), reverse=True) if not df_cto.empty else [2026]
anio_sel = st.sidebar.selectbox("Año", anios_disp)

meses_disp = sorted(df_cto[df_cto["Anio"]==anio_sel]["Mes_Num"].unique().tolist()) if not df_cto.empty else [6]
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

# Filtro Jerárquico Granja -> Lote
st.sidebar.markdown("---")
st.sidebar.markdown("**Filtros Jerárquicos Técnicos (Aplica a Pestañas 3, 5 y 6)**")
granjas_disp = sorted(list(set(df_prod['Nombre Granja'].dropna().unique().tolist() + df_lev['Nombre Granja'].dropna().unique().tolist()))) if not df_prod.empty else []
granja_sel = st.sidebar.multiselect("Filtro: Nombre Granja", options=granjas_disp, help="Agrupa la información respetando la jerarquía Granja ➔ Lote.")

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
    
    if not df_raza.empty:
        df_r = df_raza[df_raza["Periodo"] == periodo_actual]
        tot_propios = df_r["PROPIOS"].sum() if "PROPIOS" in df_r.columns else 0
        tot_ext = df_r["EXTERNOS"].sum() if "EXTERNOS" in df_r.columns else 0
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Volumen Granjas Propias", f"{tot_propios:,.0f} unds")
        c2.metric("Volumen Externos (Maquilas)", f"{tot_ext:,.0f} unds")
        c3.metric("Total Huevo Fértil (Mes)", f"{(tot_propios + tot_ext):,.0f} unds")
        
        col1, col2 = st.columns([2,1])
        with col1:
            st.subheader("Evolución Mensual (Propios vs Externos)")
            df_trend = df_raza.groupby("Periodo")[["PROPIOS", "EXTERNOS"]].sum().reset_index()
            fig_trend = px.line(df_trend, x="Periodo", y=["PROPIOS", "EXTERNOS"], markers=True)
            st.plotly_chart(fig_trend, use_container_width=True)
        with col2:
            st.subheader(f"Mix Genético ({MESES_NOMBRES.get(mes_sel)})")
            df_pie = df_r.groupby("RAZA")[["PROPIOS", "EXTERNOS"]].sum().sum(axis=1).reset_index(name="Volumen")
            fig_pie = px.pie(df_pie, names="RAZA", values="Volumen", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.warning("Datos de producción por raza no disponibles en la hoja BD PN HF RAZA.")

elif menu == "2. Costo Huevo Fértil (Consolidado)":
    st.markdown('<p class="main-title">COSTO HUEVO FÉRTIL 2026</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis de Costo Unitario y Clasificación Contable:**  
    En concordancia con las mejores prácticas contables, los costos han sido categorizados visualmente en **Directos** (Alimento, Cama, Droga, Aseo, Materia Prima) y **CIF** (rubros como PP Depr. Gallina, Mano de Obra, Arriendo).  
    *Cuenta Puente:* El ruido de la cuenta `CTA PTE LIQ. ORD PCC Y MAQUILAS` está excluido del costo bruto, extrayendo matemáticamente de allí los **Aprovechamientos (SUBPRODUCTOS)** para obtener un costo neto 100% puro y libre de duplicidades.
    """)

    if not df_cto.empty:
        df_mes = df_cto[df_cto["Periodo"] == periodo_actual]
        hf_tot = df_mes["Huevos fértiles"].sum()
        
        costo_dir = df_mes["Costos Directos"].sum()
        costo_cif = df_mes["Costos CIF"].sum()
        aprov = df_mes["SUBPRODUCTOS"].sum() # Aprovechamientos
        
        costo_neto = costo_dir + costo_cif + aprov
        cu = safe_div(costo_neto, hf_tot)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Costo Neto Consolidado", f"${costo_neto:,.0f}")
        c2.metric("Huevos Fértiles (Línea)", f"{hf_tot:,.0f}")
        c3.metric("Aprovechamientos (Crédito Real)", f"${aprov:,.0f}")
        c4.metric("Costo Unitario ($/HF)", f"${cu:,.2f}")
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.subheader("Clasificación Contable")
            fig_pie = px.pie(names=["Costos Directos", "CIF (Indirectos / PP)"], values=[costo_dir, costo_cif], hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_chart2:
            st.subheader("Estructura Top Rubros")
            cols_rubros = ['ALIMENTO', 'DEPRECIACIÓN GALLINA', 'MANO DE OBRA', 'INDIRECTOS', 'DEPRECIACIÓN CONST Y EDIF', 'ELEMENTOS DE ASEO Y DESINFECCION', 'CAMA', 'DROGA']
            vals = [df_mes[c].sum() for c in cols_rubros if c in df_mes.columns]
            df_rubros = pd.DataFrame({"Rubro": [c for c in cols_rubros if c in df_mes.columns], "Valor": vals}).sort_values("Valor", ascending=False)
            fig_bar = px.bar(df_rubros, x="Valor", y="Rubro", orientation='h', text_auto=".2s")
            st.plotly_chart(fig_bar, use_container_width=True)

elif menu == "3. Detalle Costos por Lote (Granja ➔ Lote)":
    st.markdown('<p class="main-title">DETALLE COSTOS HUEVO POR LOTE - JUNIO 2026</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Jerárquico de Lotes:**  
    Dado que un usuario nuevo no conoce los lotes de memoria, este reporte agrupa automáticamente bajo la estructura **Granja ➔ Lote**. Permite auditar qué lotes específicos están inflando el promedio de la granja (causa raíz frecuente: caídas en curva de postura que disparan la Depreciación Gallina por huevo).
    """)
    
    if not df_prod.empty and not df_cto.empty:
        df_p = df_prod[df_prod["Periodo"] == periodo_actual].copy()
        
        if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_p = df_p[df_p["Lote"].isin(lote_sel)]
        
        if not df_p.empty:
            df_p["Lote"] = df_p["Lote"].astype(str)
            df_c = df_cto[df_cto["Periodo"] == periodo_actual].copy()
            df_c["lote"] = df_c["lote"].astype(str)
            
            # Cruzar datos técnicos con financieros
            df_merge = pd.merge(df_p, df_c, left_on="Lote", right_on="lote", how="inner")
            
            if not df_merge.empty:
                df_merge["Costo Unitario ($/HF)"] = safe_div(df_merge["TOTAL COSTO"], df_merge["Huevos fértiles"])
                cols_visuales = ['Nombre Granja', 'Lote', 'Total Aves Encasetadas', 'Huevos fértiles', 'TOTAL COSTO', 'Costo Unitario ($/HF)']
                
                st.dataframe(df_merge[cols_visuales].sort_values("Costo Unitario ($/HF)", ascending=False).style.format({
                    "Total Aves Encasetadas": "{:,.0f}",
                    "Huevos fértiles": "{:,.0f}",
                    "TOTAL COSTO": "${:,.0f}",
                    "Costo Unitario ($/HF)": "${:,.2f}"
                }).background_gradient(subset=["Costo Unitario ($/HF)"], cmap="Reds"), use_container_width=True)
            else:
                st.warning("No se logró hacer cruce entre la Granja/Lote y la tabla financiera de costos para este período.")
        else:
            st.warning("No hay datos técnicos de Producción para la granja seleccionada.")
    else:
        st.warning("No se dispone de las bases BD PRODUCCIÓN o BD CTO LINEA.")

elif menu == "4. Impacto y Costo Kg Alimento":
    st.markdown('<p class="main-title">COSTO KG ALIMENTO E IMPACTO DIRECTO</p>', unsafe_allow_html=True)
    st.info("""
    **Sensibilidad de Materias Primas:**  
    El rubro ALIMENTO está clasificado como el mayor Costo Directo de la operación. Esta pestaña audita la relación entre la inversión ($) y los kilogramos consumidos, evaluando la conversión alimenticia y el precio por Kg, factores vitales para anticipar desviaciones presupuestales.
    """)
    
    if not df_cto.empty:
        df_alim = df_cto.groupby("Periodo").agg(Costo_Alimento=("ALIMENTO", "sum"), Kg_Consumidos=("consumo alimento kg", "sum"), Huevos=("Huevos fértiles", "sum")).reset_index()
        df_alim["Precio por Kg ($)"] = safe_div(df_alim["Costo_Alimento"], df_alim["Kg_Consumidos"])
        df_alim["Gramos Alimento / Huevo"] = safe_div(df_alim["Kg_Consumidos"] * 1000, df_alim["Huevos"])
        
        col1, col2 = st.columns(2)
        with col1:
            fig_precio = px.line(df_alim, x="Periodo", y="Precio por Kg ($)", markers=True, title="Evolución Precio Kg Alimento")
            fig_precio.update_traces(line_color="#d62728", line_width=3)
            st.plotly_chart(fig_precio, use_container_width=True)
        with col2:
            fig_conv = px.line(df_alim, x="Periodo", y="Gramos Alimento / Huevo", markers=True, title="Conversión Alimenticia (g/Huevo)")
            fig_conv.update_traces(line_color="#2ca02c", line_width=3)
            st.plotly_chart(fig_conv, use_container_width=True)

elif menu == "5. Lotes Finalizados - Levante":
    st.markdown('<p class="main-title">RESULTADOS Y COSTOS LOTES FINALIZADOS (LEVANTE)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Técnico - Costo por Ave Valla (Censo Técnico):**  
    Como identificamos, la información contable pura (ZCO001) carece del censo de aves vivas. Para resolver esto y obtener el Costo por Ave, hemos integrado la matriz técnica `BD LEVANTE`. De aquí extraemos métricas críticas como "Hembras Encasetadas" y el inventario de finalización para generar los unitarios reales.
    """)
    
    if not df_lev.empty:
        df_l = df_lev[df_lev["Periodo"] == periodo_actual].copy()
        if granja_sel: df_l = df_l[df_l["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_l = df_l[df_l["Lote"].isin(lote_sel)]
        
        if not df_l.empty:
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Total Aves Encasetadas']
            
            # Agregar columnas técnicas si existen
            if 'Total Aves Fin Levante' in df_l.columns: cols_mostrar.append('Total Aves Fin Levante')
            if 'Costo Total Levante' in df_l.columns: 
                cols_mostrar.append('Costo Total Levante')
                if 'Total Aves Fin Levante' in df_l.columns:
                    df_l['Costo por Ave Finalizada'] = safe_div(df_l['Costo Total Levante'], df_l['Total Aves Fin Levante'])
                    cols_mostrar.append('Costo por Ave Finalizada')
                    
            existentes = [c for c in cols_mostrar if c in df_l.columns]
            st.dataframe(df_l[existentes].style.format(na_rep="—", formatter={
                "Hembras Encasetadas": "{:,.0f}", "Total Aves Encasetadas": "{:,.0f}", 
                "Total Aves Fin Levante": "{:,.0f}", "Costo Total Levante": "${:,.0f}", 
                "Costo por Ave Finalizada": "${:,.2f}"
            }), use_container_width=True)
        else:
            st.warning("No hay lotes de Levante reportados en el período y granja seleccionados.")
    else:
        st.error("Hoja BD LEVANTE no encontrada o vacía.")

elif menu == "6. Lotes Finalizados - Producción":
    st.markdown('<p class="main-title">RESULTADOS TÉCNICOS Y COSTO HF POR LOTE (PRODUCCIÓN)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Métricas Zootécnicas de Producción Finalizada:**  
    Este módulo utiliza `BD PRODUCCIÓN` para suplir la falta de datos de censo en el reporte SAP estándar. Permite evaluar el Costo Total del Lote cruzado directamente con el comportamiento biológico de la parvada (Mortalidad, aves fin producción y edad de sacrificio).
    """)
    
    if not df_prod.empty:
        df_p = df_prod[df_prod["Periodo"] == periodo_actual].copy()
        if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_p = df_p[df_p["Lote"].isin(lote_sel)]
        
        if not df_p.empty:
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Mortalidad Hembra', '%Mortalidad Hembra']
            if 'Costo total Producción' in df_p.columns: cols_mostrar.append('Costo total Producción')
            if 'Total Aves Fin Producción' in df_p.columns: cols_mostrar.append('Total Aves Fin Producción')
            
            existentes = [c for c in cols_mostrar if c in df_p.columns]
            st.dataframe(df_p[existentes].style.format(na_rep="—", formatter={
                "Hembras Encasetadas": "{:,.0f}", "Mortalidad Hembra": "{:,.0f}", 
                "%Mortalidad Hembra": "{:.1%}", "Costo total Producción": "${:,.0f}", 
                "Total Aves Fin Producción": "{:,.0f}"
            }), use_container_width=True)
        else:
            st.warning("No hay datos técnicos de producción disponibles para los filtros aplicados.")
    else:
        st.error("Hoja BD PRODUCCIÓN no encontrada o vacía.")
