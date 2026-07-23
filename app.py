"""
Dashboard Ejecutivo — Costo del Huevo Fértil y Análisis Gerencial
=================================================================
Fuente: export SAP 'BASE ZCO001' + Matrices Técnicas de Producción.
Este script reconcilia exactamente los informes gerenciales y añade 
un motor de inteligencia de negocios para diagnosticar ineficiencias,
analizar variables clave y proponer soluciones operativas por cada pestaña.
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
    page_title="Dashboard Ejecutivo - Costos Huevo Fértil",
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
    .insight-box { background-color: #F0F9FF; padding: 15px; border-left: 5px solid #0284C7; border-radius: 5px; margin-bottom: 20px; color: #075985; font-size: 14px;}
    .alert-box   { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D;}
    .success-box { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B;}
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

MESES_NOMBRES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

# =============================================================================
# 1. MOTOR DE DATOS (ETL) BLINDADO Y SEGURO
# =============================================================================
MATERIAL_HF = "HUEVO INCUBABLE"
TEXTO_LIQUIDACION = "CTA PTE LIQ. ORD PCC Y MAQUILAS"
TEXTO_DIFERENCIA_PRECIO = "DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO"
RUBRO_APROVECHAMIENTO = "Aprovechamientos (-)"

MAP_RUBROS = {
    "CONSUMO ALIMENTO": "Alimento",
    "PP Depr. Gallina Grj.Pcc.": "Depreciación Parvada",
    "PP Horas Hombre Grj.Pcc.": "Mano de Obra",
    "PP Costos Ind. Grj.Pcc.": "Costos Indirectos (CIF)",
    "PP Costos Arriendo Grj.Pcc.": "Arriendo",
    "CONSUMO CAMA": "Cama / Cascarilla",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Bioseguridad y Aseo",
    "CONSUMO DROGA": "Sanidad (Medicamentos)",
    "PP Costos Depr. Grj.Pcc.": "Depreciación Instalaciones",
    "CONSUMOS MATERIA PRIMA": "Materias Primas (Calcio)",
}

def clean_num(x) -> float:
    if isinstance(x, (int, float, np.integer, np.floating)): return float(x)
    if pd.isna(x): return 0.0
    s = str(x).strip()
    if s == "" or s.lower() == "nan": return 0.0
    neg = s.endswith("-")
    s = s.replace("-", "").replace(",", "")
    try: v = float(s)
    except ValueError: return 0.0
    return -v if neg else v

def safe_div(numer, denom):
    if isinstance(denom, pd.Series): denom = denom.replace(0, np.nan)
    elif denom == 0: denom = np.nan
    return numer / denom

def clean_lote(s):
    if pd.isna(s): return np.nan
    s = str(s).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    return s

def find_default_excel():
    for carpeta in (Path("data/raw"), Path(".")):
        if carpeta.exists():
            encontrados = sorted(f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$"))
            if encontrados: return encontrados[0]
    return None

@st.cache_data(show_spinner="Procesando datos SAP y reconciliando información técnica...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if "BASE ZCO001" not in xls.sheet_names:
        st.error("⚠️ Error Crítico: no se encontró la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

    # --- EXTRACCIÓN TÉCNICA SEGURA (MAPEO DE GRANJAS) ---
    lote_to_granja = {}
    df_lev, df_prod = pd.DataFrame(), pd.DataFrame()
    
    def extract_tech(sheet):
        if sheet in xls.sheet_names:
            d = pd.read_excel(xls, sheet_name=sheet)
            d.columns = d.columns.astype(str).str.strip()
            if 'Lote' in d.columns: d['Lote_str'] = d['Lote'].apply(clean_lote)
            if 'Año' in d.columns: d['Anio'] = pd.to_numeric(d['Año'], errors='coerce').fillna(2026).astype(int)
            if 'No Mes' in d.columns: d['Mes_Num'] = pd.to_numeric(d['No Mes'], errors='coerce').fillna(1).astype(int)
            elif 'Mes' in d.columns:
                m_map = {'ENERO':1, 'FEBRERO':2, 'MARZO':3, 'ABRIL':4, 'MAYO':5, 'JUNIO':6, 'JULIO':7, 'AGOSTO':8, 'SEPTIEMBRE':9, 'OCTUBRE':10, 'NOVIEMBRE':11, 'DICIEMBRE':12}
                d['Mes_Num'] = d['Mes'].astype(str).str.upper().map(m_map).fillna(1).astype(int)
            if 'Anio' in d.columns and 'Mes_Num' in d.columns:
                d['Periodo'] = d['Anio'].astype(str) + "-" + d['Mes_Num'].astype(str).str.zfill(2)
            return d
        return pd.DataFrame()

    df_lev = extract_tech("BD LEVANTE")
    df_prod = extract_tech("BD PRODUCCIÓN")

    for temp_df in [df_prod, df_lev]:
        if not temp_df.empty and 'Lote_str' in temp_df.columns and 'Nombre Granja' in temp_df.columns:
            valid = temp_df.dropna(subset=['Lote_str', 'Nombre Granja'])
            lote_to_granja.update(dict(zip(valid['Lote_str'], valid['Nombre Granja'])))

    # --- EXTRACCIÓN FINANCIERA SAP (BASE MAESTRA) ---
    df_raw = pd.read_excel(xls, sheet_name="BASE ZCO001")
    df_raw.columns = df_raw.columns.astype(str).str.strip()

    if "Fecha" in df_raw.columns:
        df_raw["Fecha"] = pd.to_datetime(df_raw["Fecha"], errors="coerce")
        df_raw["Anio"] = df_raw["Fecha"].dt.year
        df_raw["Mes_Num"] = df_raw["Fecha"].dt.month
        if df_raw["Anio"].isna().any() and "EjMat" in df_raw.columns:
            df_raw["Anio"] = df_raw["Anio"].fillna(df_raw["EjMat"])
            df_raw["Mes_Num"] = df_raw["Mes_Num"].fillna(df_raw["Mes"])
    else:
        df_raw["Anio"] = df_raw["EjMat"]
        df_raw["Mes_Num"] = df_raw["Mes"]

    df_raw["Anio"] = df_raw["Anio"].astype(int)
    df_raw["Mes_Num"] = df_raw["Mes_Num"].astype(int)
    df_raw["Periodo"] = df_raw["Anio"].astype(str) + "-" + df_raw["Mes_Num"].astype(str).str.zfill(2)
    df_raw["Totales"] = df_raw["Totales"].apply(clean_num)
    df_raw["Cantidad"] = df_raw["Cantidad"].apply(clean_num)

    if 'Lote' in df_raw.columns:
        df_raw['Lote_str'] = df_raw['Lote'].apply(clean_lote)
        df_raw['Granja'] = df_raw['Lote_str'].map(lote_to_granja).fillna("Sin Granja Asignada")
    else:
        df_raw['Granja'] = "Sin Granja Asignada"

    # Consolidación Financiera
    df_hf = df_raw[df_raw["Texto breve de material"] == MATERIAL_HF]
    hf_mes = df_hf.groupby("Periodo")["Cantidad"].sum()

    df_costos = df_raw[~df_raw["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
    df_costos["Rubro"] = df_costos["Texto explicativo"].map(lambda x: MAP_RUBROS.get(x, x))
    costos_piv = df_costos.groupby(["Periodo", "Rubro"])["Totales"].sum().unstack(fill_value=0)

    df_aprov = df_raw[(df_raw["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_raw["Texto breve de material"] != MATERIAL_HF)]
    aprov_mes = df_aprov.groupby("Periodo")["Totales"].sum()

    df_res = costos_piv.copy()
    df_res[RUBRO_APROVECHAMIENTO] = aprov_mes.reindex(df_res.index).fillna(0)
    df_res["Costo Total"] = df_res.sum(axis=1)
    df_res["Huevos Fértiles"] = hf_mes.reindex(df_res.index)
    df_res["Costo Huevo Fértil"] = safe_div(df_res["Costo Total"], df_res["Huevos Fértiles"])

    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"]
    alim_kg = df_alim.groupby("Periodo")["Cantidad"].sum()
    df_res["Consumo Alimento Kg"] = alim_kg.reindex(df_res.index)
    df_res["Precio Kg Alimento"] = safe_div(df_res["Alimento"], df_res["Consumo Alimento Kg"]) if "Alimento" in df_res.columns else np.nan
    df_res["Gramos Alimento/Huevo"] = safe_div(df_res["Consumo Alimento Kg"] * 1000, df_res["Huevos Fértiles"])

    df_res = df_res.reset_index().sort_values("Periodo")
    df_res["Anio"] = df_res["Periodo"].str.split("-").str[0].astype(int)
    df_res["Mes_Num"] = df_res["Periodo"].str.split("-").str[1].astype(int)

    rubros = [r for r in list(MAP_RUBROS.values()) + [RUBRO_APROVECHAMIENTO] if r in df_res.columns]

    return df_res, rubros, df_raw, df_lev, df_prod

# =============================================================================
# 2. CARGA DE DATOS
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")
uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df, rubros_items, df_raw, df_lev, df_prod = load_and_process_data(uploaded_file)
else:
    archivo_default = find_default_excel()
    if archivo_default is None:
        st.warning("⚠️ Se requiere un archivo Excel para inicializar el sistema.")
        st.stop()
    df, rubros_items, df_raw, df_lev, df_prod = load_and_process_data(str(archivo_default))

# =============================================================================
# 3. BARRA LATERAL — CONTROLES DE TIEMPO
# =============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal")

anios_disp = sorted(df["Anio"].unique(), reverse=True)

def selector_periodo(titulo, key_prefix, anios_opciones, mes_reverse=False):
    st.sidebar.markdown(f"**{titulo}:**")
    c1, c2 = st.sidebar.columns(2)
    anio = c1.selectbox("Año", anios_opciones, key=f"{key_prefix}_anio")
    meses_disp = sorted(df[df["Anio"] == anio]["Mes_Num"].unique(), reverse=mes_reverse)
    mes = c2.selectbox("Mes", meses_disp, format_func=lambda x: MESES_NOMBRES[x], key=f"{key_prefix}_mes")
    return f"{anio}-{str(mes).zfill(2)}"

modo_analisis = st.sidebar.radio("Seleccione el enfoque:", ["⚖️ Comparativo (Mes VS Mes)", "📈 Rango Histórico (Evolución)"])

if modo_analisis == "⚖️ Comparativo (Mes VS Mes)":
    p_base = selector_periodo("Período Base (contra qué comparo)", "base", anios_disp)
    p_actual = selector_periodo("Período Actual (qué estoy evaluando)", "actual", anios_disp, mes_reverse=True)
    texto_contexto = (f"Comparativa Estratégica: **{MESES_NOMBRES[int(p_actual.split('-')[1])]} {p_actual.split('-')[0]}** "
                       f"VS **{MESES_NOMBRES[int(p_base.split('-')[1])]} {p_base.split('-')[0]}**")
else:
    p_base = selector_periodo("Inicio del rango", "ini", sorted(anios_disp))
    p_actual = selector_periodo("Fin del rango", "fin", anios_disp, mes_reverse=True)
    texto_contexto = f"Evolución Histórica: Desde **{min(p_base, p_actual)}** hasta **{max(p_base, p_actual)}**"

rango_inicio, rango_fin = min(p_base, p_actual), max(p_base, p_actual)
df_filtrado_graficas = df[(df["Periodo"] >= rango_inicio) & (df["Periodo"] <= rango_fin)].sort_values("Periodo")
df_raw_graficas = df_raw[(df_raw["Periodo"] >= rango_inicio) & (df_raw["Periodo"] <= rango_fin)]

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. Análisis de Producción Global",
        "2. Mix Genético y Comportamiento",
        "3. Costo Consolidado (Huevo Fértil)",
        "4. Auditoría Micro por Granja/Lote",
        "5. Costo Kg Alimento y Eficiencia",
        "6. Resultados Técnicos Levante",
        "7. Resultados Técnicos Producción"
    ]
)

# =============================================================================
# MÓDULOS DEL DASHBOARD (CADA UNO ES UN INFORME ANALÍTICO COMPLETO)
# =============================================================================

if menu == "1. Análisis de Producción Global":
    st.markdown('<p class="main-title">1. PRODUCCIÓN CONSOLIDADA DE HUEVO FÉRTIL</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Variables de Producción:</b><br>
        El volumen de Huevos Fértiles es el "denominador universal" de la rentabilidad avícola. Una alta producción diluye automáticamente la carga pesada de los Costos Fijos (Depreciación de las Aves, Instalaciones y Mano de Obra). <br>
        <i>💡 Soluciones ante caídas:</i> Si el volumen propio cae, la estrategia de apalancamiento mediante "Maquilas" (producción externa) es vital para cumplir con los despachos a planta de incubación sin asumir los costos directos de levante.
    </div>
    """, unsafe_allow_html=True)

    df_hf_f = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    tot_propios = df_hf_f["Cantidad"].sum()
    n_meses_rango = df_filtrado_graficas["Periodo"].nunique()
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Volumen Total Período (Propios)", f"{tot_propios:,.0f} HF")
    c2.metric("Lotes Activos Reportados", f"{df_hf_f['Lote_str'].nunique()}")
    c3.metric("Promedio Mensual de Producción", f"{tot_propios / max(n_meses_rango, 1):,.0f} HF/mes")
    
    st.markdown("---")
    st.subheader("📈 Evolución Histórica del Volumen Operativo")
    df_ts = df_hf_f.groupby("Periodo")["Cantidad"].sum().reset_index()
    fig_ts = px.line(df_ts, x="Periodo", y="Cantidad", markers=True)
    fig_ts.update_traces(line_color="#1E3A8A", line_width=3)
    st.plotly_chart(fig_ts, use_container_width=True)

elif menu == "2. Mix Genético y Comportamiento":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÓN Y COSTO POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Desempeño Genético:</b><br>
        Las distintas líneas genéticas (Ej. ROSS vs COBB) presentan comportamientos asimétricos en el pico de postura y en la viabilidad del ave adulta. Un costo unitario alto en una línea específica generalmente indica una rápida caída en la curva de producción post-semana 50.<br>
        <i>💡 Soluciones:</i> Si una línea genética es persistentemente más costosa por huevo, se debe evaluar la rentabilidad final considerando no solo el costo en granja, sino el % de incubabilidad y nacimientos en la planta.
    </div>
    """, unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)]
    df_costo_lin = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov_lin = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    c_lin = df_costo_lin.groupby("linea")["Totales"].sum()
    a_lin = df_aprov_lin.groupby("linea")["Totales"].sum()
    h_lin = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby("linea")["Cantidad"].sum()
    
    df_gen = pd.DataFrame({"Costo Base": c_lin, "Aprovechamientos": a_lin, "Huevos Fértiles": h_lin}).fillna(0)
    df_gen["Costo Total ($)"] = df_gen["Costo Base"] + df_gen["Aprovechamientos"]
    df_gen["Costo Unitario ($/HF)"] = df_gen["Costo Total ($)"] / df_gen["Huevos Fértiles"]
    df_gen = df_gen[df_gen["Huevos Fértiles"] > 0].reset_index()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🧬 Participación por Raza en la Producción")
        fig_pie = px.pie(df_gen, values="Huevos Fértiles", names="linea", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
    with col2:
        st.subheader(f"📊 Costo Unitario Promedio por Línea ({p_actual})")
        fig_bar = px.bar(df_gen, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".1f")
        st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(df_gen[["linea", "Huevos Fértiles", "Costo Total ($)", "Costo Unitario ($/HF)"]].style.format({
        "Costo Total ($)": "${:,.0f}", "Huevos Fértiles": "{:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"
    }), use_container_width=True)

elif menu == "3. Costo Consolidado (Huevo Fértil)":
    st.markdown('<p class="main-title">3. ANÁLISIS INTEGRAL: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis del período {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)
    
    if p_base == p_actual:
        st.warning("⚠️ Selecciona el Modo Comparativo con dos períodos distintos en la barra lateral para ver la variación.")
        st.stop()
        
    fila_b = df[df["Periodo"] == p_base]
    fila_a = df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty: st.stop()
    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    var_tot = df_a["Costo Huevo Fértil"] - df_b["Costo Huevo Fértil"]

    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Estructura Financiera:</b><br>
        El Costo del Huevo Fértil se rige por la "Regla del 80/20", donde el <b>Alimento</b> y la <b>Depreciación de la Gallina</b> dominan la matriz.<br>
        <i>💡 Soluciones Integrales:</i> Para neutralizar variaciones en el tornado de rubros, se requiere: 1) Estricto control de básculas para evitar desperdicio de alimento, y 2) Identificación temprana de galpones enfermos que elevan artificialmente la amortización del ave al no producir.
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo Promedio {p_base} (Base)", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo Promedio {p_actual} (Actual)", f"${df_a['Costo Huevo Fértil']:,.2f}")
    delta_pct = (var_tot / df_b["Costo Huevo Fértil"]) * 100 if df_b["Costo Huevo Fértil"] else 0
    c3.metric("Desviación Total ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")
    
    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos Fértiles"], df_a["Huevos Fértiles"]
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({"Rubro": r, "Costo Unit Base ($)": ub, "Costo Unit Actual ($)": ua, "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("📈 Evolución Histórica de Costo ($)")
        fig_line_c = px.line(df_filtrado_graficas, x="Periodo", y="Costo Huevo Fértil", markers=True)
        fig_line_c.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_line_c, use_container_width=True)
    with col2:
        st.subheader("📊 Tornado de Desviación por Rubro (VIF)")
        if not df_vif.empty:
            fig_tor = px.bar(df_vif.sort_values("Impacto ($/HF)"), y="Rubro", x="Impacto ($/HF)", orientation="h", color="Impacto ($/HF)", color_continuous_scale="Reds" if var_tot > 0 else "Greens")
            st.plotly_chart(fig_tor, use_container_width=True)

elif menu == "4. Auditoría Micro por Granja/Lote":
    st.markdown('<p class="main-title">4. DIAGNÓSTICO MICRO-OPERATIVO POR GRANJA Y LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría Quirúrgica de Costos para el período: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)].copy()
    if df_l.empty:
        st.warning("No hay registros financieros en SAP para el mes seleccionado.")
        st.stop()

    df_costo_lote = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov_lote = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    # Agrupación Segura garantizando que no se pierdan lotes sin Granja
    c_lote = df_costo_lote.groupby(["Granja", "Lote_str"])["Totales"].sum().reset_index(name="Costo Base")
    a_lote = df_aprov_lote.groupby(["Granja", "Lote_str"])["Totales"].sum().reset_index(name="Aprovechamientos")
    h_lote = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby(["Granja", "Lote_str"])["Cantidad"].sum().reset_index(name="Huevos Fértiles")

    df_m = h_lote.merge(c_lote, on=["Granja", "Lote_str"], how="left").merge(a_lote, on=["Granja", "Lote_str"], how="left").fillna(0)
    df_m["Costo Total"] = df_m["Costo Base"] + df_m["Aprovechamientos"]
    df_m["Costo Unitario"] = df_m["Costo Total"] / df_m["Huevos Fértiles"]
    df_m = df_m[df_m["Huevos Fértiles"] > 0].reset_index(drop=True)
    df_m = df_m.rename(columns={"Lote_str": "Lote"})
    
    if df_m.empty:
        st.warning("Existen costos, pero no se reportaron huevos fértiles producidos en este período.")
        st.stop()
        
    promedio_mes = df_m["Costo Total"].sum() / df_m["Huevos Fértiles"].sum()
    df_m = df_m.sort_values("Costo Unitario", ascending=False)
    lote_critico = df_m.iloc[0]

    st.markdown(f"""
    <div class="alert-box">
        <b>🚨 Auditoría de Ineficiencia en Campo (Causa Raíz):</b><br>
        El comportamiento financiero revela que el <b>Lote {lote_critico['Lote']}</b> ({lote_critico['Granja']}) está arrastrando la rentabilidad de la compañía con un costo de <b>${lote_critico['Costo Unitario']:,.1f} COP/Huevo</b> (frente al promedio de ${promedio_mes:,.1f}).<br>
        <i>💡 Análisis Técnico:</i> Lotes que cuadruplican el costo promedio suelen estar al final de su vida útil productiva (baja postura = divisor pequeño = costo disparado). <br>
        <i>Acción Inmediata:</i> Proceder al retiro comercial (descarte) programado o revisión clínica si el lote es joven.
    </div>
    """, unsafe_allow_html=True)

    st.subheader("🏢 Consolidado Operativo por Granja")
    df_granja = df_m.groupby("Granja").agg(
        Lotes_Operativos=("Lote", lambda x: ", ".join(x.unique())),
        Total_Huevos=("Huevos Fértiles", "sum"),
        Costo_Operativo=("Costo Total", "sum")
    ).reset_index()
    df_granja["Costo Promedio ($/HF)"] = safe_div(df_granja["Costo_Operativo"], df_granja["Total_Huevos"])
    st.dataframe(df_granja.style.format({
        "Total_Huevos": "{:,.0f} unds", "Costo_Operativo": "${:,.0f}", "Costo Promedio ($/HF)": "${:,.2f}"
    }).background_gradient(subset=["Costo Promedio ($/HF)"], cmap="YlOrRd"), use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Auditoría Exacta: Desglose Lote a Lote")
    st.dataframe(df_m[["Granja", "Lote", "Huevos Fértiles", "Costo Total", "Costo Unitario"]].style.format({
        "Huevos Fértiles": "{:,.0f} unds", "Costo Total": "${:,.0f}", "Costo Unitario": "${:,.2f}"
    }).background_gradient(subset=["Costo Unitario"], cmap="Reds"), use_container_width=True)

elif menu == "5. Costo Kg Alimento y Eficiencia":
    st.markdown('<p class="main-title">5. IMPACTO DEL COSTO DE ALIMENTO Y CONVERSIÓN</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Eficiencia Nutricional:</b><br>
        El alimento representa más del 40% del costo productivo. Un precio por kilogramo bajo en compras es inútil si la <b>Conversión Alimenticia (gramos por huevo)</b> sube drásticamente, lo cual ocurre por problemas de temperatura, formulación pobre o desperdicios físicos en comederos.<br>
        <i>💡 Soluciones:</i> Cruzar siempre el Precio vs. el Gramaje. Optimizar calibración de comederos automáticos y realizar auditorías de calidad de la materia prima en planta.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📉 Negociación: Precio Alimento ($/Kg)")
        fig_a = px.line(df_filtrado_graficas, x="Periodo", y="Precio Kg Alimento", markers=True)
        fig_a.update_traces(line_color="#d62728", line_width=3)
        st.plotly_chart(fig_a, use_container_width=True)
    with col2:
        st.subheader("🧬 Biología: Conversión (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x="Periodo", y="Gramos Alimento/Huevo", markers=True)
        fig_g.update_traces(line_color="#2ca02c", line_width=3)
        st.plotly_chart(fig_g, use_container_width=True)

    st.subheader("🎛️ Simulador Predictivo de Rentabilidad")
    st.markdown(f"Escenario aplicado al cierre operativo de **{p_actual}**:")
    df_a = df[df["Periodo"] == p_actual].iloc[0] if not df[df["Periodo"] == p_actual].empty else pd.Series()
    
    if not df_a.empty and pd.notna(df_a.get("Precio Kg Alimento")):
        s1, s2, s3 = st.columns(3)
        ahorro_alim = s1.number_input("1. Mejora en Compra ($/Kg):", value=50, step=10)
        mejora_conv = s2.number_input("2. Bajar desperdicio (g/huevo):", value=15, step=5)
        mejora_post = s3.number_input("3. Subir pico de postura en (%):", value=3.0, step=0.5)
        
        precio_actual_kg = df_a["Precio Kg Alimento"]
        nuevo_p_alim = max(precio_actual_kg - ahorro_alim, 500)
        nuevo_cons_kg = max(df_a["Consumo Alimento Kg"] - ((mejora_conv / 1000) * df_a["Huevos Fértiles"]), 0)
        nuevo_costo_alim = nuevo_cons_kg * nuevo_p_alim
        
        nuevos_hf = df_a["Huevos Fértiles"] * (1 + (mejora_post / 100))
        costo_sin_alimento = df_a["Costo Total"] - df_a.get("Alimento", 0)
        nuevo_costo_total = costo_sin_alimento + nuevo_costo_alim
        nuevo_costo_huevo = nuevo_costo_total / nuevos_hf if nuevos_hf > 0 else 0
        ahorro_un = df_a["Costo Huevo Fértil"] - nuevo_costo_huevo
        
        st.success(f"🎯 **Ahorro Estratégico:** Con esta eficiencia conjunta, el costo global bajaría de **${df_a['Costo Huevo Fértil']:,.1f}** a **${nuevo_costo_huevo:,.1f} COP/Huevo**. ¡Un impacto directo de **${ahorro_un:,.1f} netos a la utilidad** por unidad!")
    else:
        st.info("No hay datos suficientes de Alimento en el mes actual para correr el simulador.")

elif menu == "6. Resultados Técnicos Levante":
    st.markdown('<p class="main-title">6. DESEMPEÑO ZOOTÉCNICO Y CAPITALIZACIÓN EN LEVANTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Evaluación de Lotes en Crianza para: <b>{p_actual}</b></p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Etapa de Levante:</b><br>
        La fase de levante es un "centro de costos puros" (sin ingresos operativos). El éxito financiero aquí se mide dividiendo la inversión acumulada entre las **Aves Vivas Finalizadas**. Altas mortalidades en las primeras 5 semanas destruyen el valor del activo, ya que encarecen exponencialmente el costo promedio de las aves sobrevivientes que llegarán a producción.<br>
        <i>💡 Soluciones Preventivas:</i> Auditar planes de vacunación, condiciones de bioseguridad y pesos corporales uniformes para maximizar la viabilidad del traslado.
    </div>
    """, unsafe_allow_html=True)
    
    if not df_lev.empty:
        df_l = df_lev[df_lev["Periodo"] == p_actual].copy()
        if not df_l.empty:
            st.subheader("🏢 Desempeño Consolidado de Crianza por Granja")
            matriz_lev = df_l.groupby("Nombre Granja").agg(
                Lotes_Asociados=("Lote_str", lambda x: ", ".join(x.dropna().unique())),
                Hembras_Encasetadas=("Hembras Encasetadas", "sum"),
                Aves_Finalizadas=("Total Aves Fin Levante", "sum") if "Total Aves Fin Levante" in df_l.columns else ("Hembras Encasetadas", "sum"),
                Costo_Total_Levante=("Costo Total Levante", "sum") if "Costo Total Levante" in df_l.columns else ("Hembras Encasetadas", "sum")
            ).reset_index()
            matriz_lev["Costo Unitario (Activo Capitalizado)"] = safe_div(matriz_lev["Costo_Total_Levante"], matriz_lev["Aves_Finalizadas"])
            st.dataframe(matriz_lev.style.format({
                "Hembras_Encasetadas": "{:,.0f}", "Aves_Finalizadas": "{:,.0f}", 
                "Costo_Total_Levante": "${:,.0f}", "Costo Unitario (Activo Capitalizado)": "${:,.2f}"
            }), use_container_width=True)
            
            st.markdown("---")
            st.subheader("📋 Desglose Técnico por Lote (Levante)")
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas']
            if 'Total Aves Fin Levante' in df_l.columns: cols_mostrar.append('Total Aves Fin Levante')
            if 'Costo Total Levante' in df_l.columns: 
                cols_mostrar.append('Costo Total Levante')
                if 'Total Aves Fin Levante' in df_l.columns:
                    df_l['Costo Unitario Ave'] = safe_div(df_l['Costo Total Levante'], df_l['Total Aves Fin Levante'])
                    cols_mostrar.append('Costo Unitario Ave')
            existentes = [c for c in cols_mostrar if c in df_l.columns]
            st.dataframe(df_l[existentes].style.format(na_rep="—", formatter={
                "Hembras Encasetadas": "{:,.0f}", "Total Aves Fin Levante": "{:,.0f}", 
                "Costo Total Levante": "${:,.0f}", "Costo Unitario Ave": "${:,.2f}"
            }), use_container_width=True)
        else:
            st.info("No hay reportes técnicos de Levante procesados para el período seleccionado.")
    else:
        st.warning("Hoja de captura BD LEVANTE no está presente o no tiene formato válido.")

elif menu == "7. Resultados Técnicos Producción":
    st.markdown('<p class="main-title">7. MÉTRICAS ZOOTÉCNICAS FINALES EN PRODUCCIÓN</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Evaluación de Lotes Activos para: <b>{p_actual}</b></p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="insight-box">
        <b>📝 Análisis Ejecutivo - Mortalidad y Biología de Producción:</b><br>
        En la fase final de producción, la <b>Mortalidad Acumulada de la Hembra</b> es la métrica de alarma. Cada punto porcentual de aves muertas no solo frena el ingreso diario esperado, sino que fuerza a la compañía a amortizar aceleradamente a las gallinas restantes.<br>
        <i>💡 Soluciones:</i> Si la mortalidad o selección es anormal, implementar planes de choque veterinarios inmediatos. Las pérdidas en esta fase representan el golpe financiero más doloroso debido a los meses de crianza ya invertidos.
    </div>
    """, unsafe_allow_html=True)
    
    if not df_prod.empty:
        df_p = df_prod[df_prod["Periodo"] == p_actual].copy()
        if not df_p.empty:
            st.subheader("🏢 Consolidado Técnico por Granja (Mortalidad General)")
            matriz_prod_fin = df_p.groupby("Nombre Granja").agg(
                Lotes_Asociados=("Lote_str", lambda x: ", ".join(x.dropna().unique())),
                Hembras_Encasetadas=("Hembras Encasetadas", "sum"),
                Mortalidad=("Mortalidad Hembra", "sum") if "Mortalidad Hembra" in df_p.columns else ("Hembras Encasetadas", "sum"),
                Costo_Produccion=("Costo total Producción", "sum") if "Costo total Producción" in df_p.columns else ("Hembras Encasetadas", "sum")
            ).reset_index()
            matriz_prod_fin["% Mortalidad Real"] = safe_div(matriz_prod_fin["Mortalidad"], matriz_prod_fin["Hembras_Encasetadas"])
            st.dataframe(matriz_prod_fin.style.format({
                "Hembras_Encasetadas": "{:,.0f}", "Mortalidad": "{:,.0f}", 
                "Costo_Produccion": "${:,.0f}", "% Mortalidad Real": "{:.2%}"
            }), use_container_width=True)

            st.markdown("---")
            st.subheader("📋 Desglose Biológico Lote a Lote")
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Mortalidad Hembra', '%Mortalidad Hembra']
            if 'Costo total Producción' in df_p.columns: cols_mostrar.append('Costo total Producción')
            existentes = [c for c in cols_mostrar if c in df_p.columns]
            st.dataframe(df_p[existentes].style.format(na_rep="—", formatter={
                "Hembras Encasetadas": "{:,.0f}", "Mortalidad Hembra": "{:,.0f}", 
                "%Mortalidad Hembra": "{:.2%}", "Costo total Producción": "${:,.0f}"
            }), use_container_width=True)
        else:
            st.info("No hay datos técnicos de Producción registrados en este período específico.")
    else:
        st.warning("Hoja de captura BD PRODUCCIÓN no está presente o no tiene formato válido.")
