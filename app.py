"""
Dashboard Ejecutivo — Análisis de Costos Granjas Reproductoras 2026
===================================================================
Fuente principal: 'INFORME GENERAL MENSUAL_5.xlsx'.
Integra finanzas (BD y BD CTO LINEA) y censos técnicos (BD LEVANTE y BD PRODUCCIÓN)
para generar informes comparativos, análisis de variaciones y costos unitarios.
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
    .insight-box { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B; font-size: 14px;}
    .alert-box   { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D;}
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
# 1. MOTOR DE DATOS (ETL) NATIVO
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
        for c in ['Producción HF', 'Externos HF', 'Costo total', 'Alimento ', 'KG cons Alimento', 'Aprovechamientos (-)']:
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
        
        # Agrupación CIF vs Directos
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

uploaded_file = st.sidebar.file_uploader("Subir Archivo Consolidado", type=["xlsx", "xls"])

try:
    file_source = uploaded_file if uploaded_file else "INFORME GENERAL MENSUAL_5.xlsx"
    df_bd, df_cto, df_raza, df_lev, df_prod = load_and_process_data(file_source)
except Exception as e:
    st.error("⚠️ Por favor, carga tu archivo 'INFORME GENERAL MENSUAL_5.xlsx' en el panel lateral para iniciar el dashboard.")
    st.stop()

# =============================================================================
# 3. CONTROLES TEMPORALES
# =============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal")

anios_disp = sorted(df_bd["Anio"].unique().tolist(), reverse=True) if not df_bd.empty else [2026]

def selector_periodo(titulo, key_prefix, anios_opciones, mes_reverse=False):
    st.sidebar.markdown(f"**{titulo}:**")
    c1, c2 = st.sidebar.columns(2)
    anio = c1.selectbox("Año", anios_opciones, key=f"{key_prefix}_anio")
    meses_disp = sorted(df_bd[df_bd["Anio"] == anio]["Mes_Num"].unique(), reverse=mes_reverse)
    mes = c2.selectbox("Mes", meses_disp, format_func=lambda x: MESES_NOMBRES[x], key=f"{key_prefix}_mes")
    return f"{anio}-{str(mes).zfill(2)}"

modo_analisis = st.sidebar.radio(
    "Seleccione el enfoque temporal:",
    ["⚖️ Comparativo (Mes VS Mes)", "📈 Rango Histórico (Evolución)"]
)

if modo_analisis == "⚖️ Comparativo (Mes VS Mes)":
    p_base = selector_periodo("Período Base (contra qué comparo)", "base", anios_disp)
    p_actual = selector_periodo("Período Actual (qué estoy evaluando)", "actual", anios_disp, mes_reverse=True)
    texto_contexto = (f"Comparativa: **{MESES_NOMBRES[int(p_actual.split('-')[1])]} {p_actual.split('-')[0]}** "
                       f"VS **{MESES_NOMBRES[int(p_base.split('-')[1])]} {p_base.split('-')[0]}**")
else:
    p_base = selector_periodo("Inicio del rango", "ini", sorted(anios_disp))
    p_actual = selector_periodo("Fin del rango", "fin", anios_disp, mes_reverse=True)
    texto_contexto = f"Evolución Histórica: Desde **{min(p_base, p_actual)}** hasta **{max(p_base, p_actual)}**"

rango_inicio, rango_fin = min(p_base, p_actual), max(p_base, p_actual)

df_bd_f = df_bd[(df_bd["Periodo"] >= rango_inicio) & (df_bd["Periodo"] <= rango_fin)].sort_values("Periodo")
df_cto_f = df_cto[(df_cto["Periodo"] >= rango_inicio) & (df_cto["Periodo"] <= rango_fin)].sort_values("Periodo")

# =============================================================================
# 4. CONTROLES JERÁRQUICOS Y MENÚ
# =============================================================================
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

st.sidebar.markdown("---")
st.sidebar.markdown("**Filtros Jerárquicos Técnicos (Pestañas 3, 5 y 6)**")

# Corrección de la lista de Granjas Disponibles
granjas_disp = []
if not df_prod.empty and 'Nombre Granja' in df_prod.columns:
    granjas_disp += df_prod['Nombre Granja'].dropna().astype(str).unique().tolist()
if not df_lev.empty and 'Nombre Granja' in df_lev.columns:
    granjas_disp += df_lev['Nombre Granja'].dropna().astype(str).unique().tolist()
granjas_disp = sorted(list(set(granjas_disp)))

granja_sel = st.sidebar.multiselect("Filtro: Nombre Granja", options=granjas_disp, help="Agrupa la información respetando la jerarquía.")

# Corrección de la lógica de los lotes: si no hay granja seleccionada, mostrar TODOS los lotes
lotes_p = []
lotes_l = []

if granja_sel:
    # Mostrar lotes solo de la(s) granja(s) seleccionada(s)
    if not df_prod.empty and 'Lote' in df_prod.columns and 'Nombre Granja' in df_prod.columns:
        lotes_p = df_prod[df_prod['Nombre Granja'].isin(granja_sel)]['Lote'].dropna().astype(str).unique().tolist()
    if not df_lev.empty and 'Lote' in df_lev.columns and 'Nombre Granja' in df_lev.columns:
        lotes_l = df_lev[df_lev['Nombre Granja'].isin(granja_sel)]['Lote'].dropna().astype(str).unique().tolist()
else:
    # Si no hay granja seleccionada, mostrar TODOS los lotes
    if not df_prod.empty and 'Lote' in df_prod.columns:
        lotes_p = df_prod['Lote'].dropna().astype(str).unique().tolist()
    if not df_lev.empty and 'Lote' in df_lev.columns:
        lotes_l = df_lev['Lote'].dropna().astype(str).unique().tolist()

lotes_disp = sorted(list(set(lotes_p + lotes_l)))
lote_sel = st.sidebar.multiselect("Filtro: Lote", options=lotes_disp)

# =============================================================================
# MÓDULOS DEL DASHBOARD
# =============================================================================

if menu == "1. Producción y Mix Genético":
    st.markdown('<p class="main-title">PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + Externos)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Análisis Ejecutivo de Producción:**  
    La cantidad de **Huevos Fértiles** es el denominador universal en la avicultura. Cualquier caída en este volumen castiga severamente la absorción de los Costos Fijos (como Depreciaciones y Mano de Obra). Controlar el balance entre la producción de granjas propias y el apalancamiento a través de externos (maquilas) es vital para diluir la carga operativa.
    """)
    
    if not df_raza.empty:
        df_r = df_raza[df_raza["Periodo"] == p_actual]
        tot_propios = df_r["PROPIOS"].sum() if "PROPIOS" in df_r.columns else 0
        tot_ext = df_r["EXTERNOS"].sum() if "EXTERNOS" in df_r.columns else 0
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Volumen Granjas Propias", f"{tot_propios:,.0f} unds")
        c2.metric("Volumen Externos (Maquilas)", f"{tot_ext:,.0f} unds")
        c3.metric(f"Total Huevo Fértil ({p_actual})", f"{(tot_propios + tot_ext):,.0f} unds")
        
        col1, col2 = st.columns([2,1])
        with col1:
            st.subheader("Evolución Mensual (Propios vs Externos)")
            df_trend = df_raza.groupby("Periodo")[["PROPIOS", "EXTERNOS"]].sum().reset_index()
            fig_trend = px.line(df_trend, x="Periodo", y=["PROPIOS", "EXTERNOS"], markers=True)
            st.plotly_chart(fig_trend, use_container_width=True)
        with col2:
            st.subheader(f"Mix Genético ({p_actual})")
            df_pie = df_r.groupby("RAZA")[["PROPIOS", "EXTERNOS"]].sum().sum(axis=1).reset_index(name="Volumen")
            fig_pie = px.pie(df_pie, names="RAZA", values="Volumen", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.warning("Datos de producción por raza no disponibles.")

elif menu == "2. Costo Huevo Fértil (Consolidado)":
    st.markdown('<p class="main-title">ANÁLISIS FINANCIERO: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    st.info("""
    **Análisis Directos vs CIF y Tratamiento de la 'Cuenta Puente':**  
    1. **Cuenta Puente y Aprovechamientos:** Hemos comprobado que la cuenta `CTA PTE LIQ. ORD PCC Y MAQUILAS` inflaba el balance global en SAP. Ha sido matemáticamente excluida del costo total. En su lugar, el modelo extrae únicamente el saldo real de los **Aprovechamientos** (ingreso por venta de huevo comercial/gallinaza), registrándolo como un crédito a favor que disminuye el costo final.  
    2. **Clasificación Contable:** Todos los rubros provenientes de SAP que inician con **PP** (PP Depr. Gallina, PP Mano de Obra, PP Arriendos) han sido encapsulados como **CIF (Costos Indirectos)**. El resto (Alimento, Cama, Sanidad) componen los **Costos Directos**, midiendo la eficiencia biológica pura de la granja.
    """)

    if p_base == p_actual and modo_analisis == "⚖️ Comparativo (Mes VS Mes)":
        st.warning("⚠️ Para el análisis de variaciones, selecciona un 'Período Base' diferente al 'Período Actual'.")
        st.stop()

    fila_b = df_bd[df_bd["Periodo"] == p_base]
    fila_a = df_bd[df_bd["Periodo"] == p_actual]
    
    if fila_b.empty or fila_a.empty:
        st.error("No hay datos consolidados suficientes en los períodos seleccionados para comparar.")
        st.stop()

    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    cu_b = safe_div(df_b['Costo total'], df_b['Producción HF'])
    cu_a = safe_div(df_a['Costo total'], df_a['Producción HF'])
    var_tot = cu_a - cu_b

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo Unitario {p_base} (Base)", f"${cu_b:,.2f}")
    c2.metric(f"Costo Unitario {p_actual} (Actual)", f"${cu_a:,.2f}")
    delta_pct = (var_tot / cu_b) * 100 if pd.notna(cu_b) and cu_b != 0 else 0
    c3.metric("Desviación Total ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")

    # Análisis de Variación por Rubro (VIF)
    rubros_vif = [c for c in ['Arriendo ', 'Cama - Cascarilla', 'Depreciacion Const. Y Edif.', 'Depreciacion Huevo ', 'Droga', 'Materia prima (calcio)', 'Alimento ', 'Mano de Obra', 'Aseo y desinfección', 'Indirectos'] if c in df_bd.columns]
    
    filas = []
    for r in rubros_vif:
        ub = safe_div(df_b[r], df_b['Producción HF'])
        ua = safe_div(df_a[r], df_a['Producción HF'])
        filas.append({"Rubro": r.strip(), "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan})
    
    df_vif = pd.DataFrame(filas).dropna().sort_values("Impacto ($/HF)", ascending=False)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Evolución del Costo Unitario")
        df_bd_f["Costo Unitario"] = df_bd_f["Costo total"] / df_bd_f["Producción HF"]
        fig_line = px.line(df_bd_f, x="Periodo", y="Costo Unitario", markers=True)
        st.plotly_chart(fig_line, use_container_width=True)
    with col2:
        st.subheader("Tornado de Variación (VIF) por Rubro")
        if not df_vif.empty:
            fig_tor = px.bar(df_vif.sort_values("Impacto ($/HF)"), y="Rubro", x="Impacto ($/HF)", orientation="h", color="Impacto ($/HF)", color_continuous_scale="RdYlGn_r")
            st.plotly_chart(fig_tor, use_container_width=True)

elif menu == "3. Detalle Costos por Lote (Granja ➔ Lote)":
    st.markdown('<p class="main-title">DETALLE COSTOS HUEVO POR LOTE (MICRO-AUDITORÍA)</p>', unsafe_allow_html=True)
    
    st.info("""
    **Vincular Lotes a Granjas (Micro-Auditoría):**  
    En SAP, los lotes son números transaccionales sin contexto geográfico. Al enlazar la columna "Nombre Granja", el código crea una jerarquía lógica **Granja ➔ Lote**. Esto permite identificar quirúrgicamente qué galpones específicos están operando ineficientemente y afectando el promedio general de la granja.
    """)
    
    if not df_prod.empty and not df_cto.empty:
        df_p = df_prod[df_prod["Periodo"] == p_actual].copy()
        
        if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_p = df_p[df_p["Lote"].astype(str).isin(lote_sel)]
        
        if not df_p.empty:
            df_p["Lote"] = df_p["Lote"].astype(str)
            df_c = df_cto[df_cto["Periodo"] == p_actual].copy()
            df_c["lote"] = df_c["lote"].astype(str)
            
            df_merge = pd.merge(df_p, df_c, left_on="Lote", right_on="lote", how="inner")
            
            if not df_merge.empty:
                df_merge["Costo Unitario ($/HF)"] = safe_div(df_merge["TOTAL COSTO"], df_merge["Huevos fértiles"])
                cols_visuales = ['Nombre Granja', 'Lote', 'Total Aves Encasetadas', 'Huevos fértiles', 'Costos Directos', 'Costos CIF', 'SUBPRODUCTOS', 'TOTAL COSTO', 'Costo Unitario ($/HF)']
                
                st.dataframe(df_merge[cols_visuales].sort_values("Costo Unitario ($/HF)", ascending=False).style.format({
                    "Total Aves Encasetadas": "{:,.0f}", "Huevos fértiles": "{:,.0f}",
                    "Costos Directos": "${:,.0f}", "Costos CIF": "${:,.0f}", "SUBPRODUCTOS": "${:,.0f}",
                    "TOTAL COSTO": "${:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"
                }).background_gradient(subset=["Costo Unitario ($/HF)"], cmap="Reds"), use_container_width=True)
            else:
                st.warning(f"No se logró hacer cruce entre la Granja/Lote técnico y la tabla financiera para {p_actual}.")
        else:
            st.warning("No hay datos técnicos de Producción para los filtros seleccionados.")

elif menu == "4. Impacto y Costo Kg Alimento":
    st.markdown('<p class="main-title">COSTO KG ALIMENTO E IMPACTO DIRECTO</p>', unsafe_allow_html=True)
    
    st.info("""
    **Sensibilidad de Materias Primas y Conversión:**  
    El Alimento es el Costo Directo con mayor peso. Este módulo audita dos palancas críticas:  
    - **Precio del Alimento ($/Kg):** Refleja la eficiencia en la negociación y compra de materias primas.  
    - **Conversión Alimenticia (Gramos/Huevo):** Es la eficiencia metabólica. Si el precio del Kg baja, pero la conversión sube (las aves comen más para producir un huevo), el margen financiero global se destruye.
    """)
    
    if not df_bd_f.empty and 'Alimento ' in df_bd_f.columns and 'KG cons Alimento' in df_bd_f.columns:
        df_alim = df_bd_f.copy()
        df_alim["Precio por Kg ($)"] = safe_div(df_alim["Alimento "], df_alim["KG cons Alimento"])
        df_alim["Gramos Alimento / Huevo"] = safe_div(df_alim["KG cons Alimento"] * 1000, df_alim["Producción HF"])
        
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
    
    st.success("""
    **Solución Estructural: El Costo por Ave Finalizada**  
    El sistema SAP estándar (`BASE ZCO001`) almacena valores financieros y kilos, pero omite el inventario biológico de aves vivas. Para poder entregar el **Costo por Ave**, el código extrae automáticamente las métricas de **'Hembras Encasetadas'** y **'Aves Fin Levante'** desde la matriz técnica (`BD LEVANTE`), cruzando el gasto acumulado entre las sobrevivientes reales para capitalizar el costo antes del traslado a producción.
    """)
    
    if not df_lev.empty:
        df_l = df_lev[df_lev["Periodo"] == p_actual].copy()
        if granja_sel: df_l = df_l[df_l["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_l = df_l[df_l["Lote"].astype(str).isin(lote_sel)]
        
        if not df_l.empty:
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Total Aves Encasetadas']
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
        st.error("Hoja técnica BD LEVANTE no encontrada.")

elif menu == "6. Lotes Finalizados - Producción":
    st.markdown('<p class="main-title">RESULTADOS TÉCNICOS DE PRODUCCIÓN FINALIZADOS</p>', unsafe_allow_html=True)
    
    st.info("""
    **Métricas Zootécnicas Complementarias:**  
    En la etapa de producción, una alta **Mortalidad de Hembras** golpea la rentabilidad dos veces: primero, reduciendo la producción total de Huevos Fértiles (nuestro divisor universal); y segundo, concentrando la pesada carga de la Depreciación Gallina en un menor número de aves vivas. Este panel extrae el desempeño biológico final desde la `BD PRODUCCIÓN`.
    """)
    
    if not df_prod.empty:
        df_p = df_prod[df_prod["Periodo"] == p_actual].copy()
        if granja_sel: df_p = df_p[df_p["Nombre Granja"].isin(granja_sel)]
        if lote_sel: df_p = df_p[df_p["Lote"].astype(str).isin(lote_sel)]
        
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
