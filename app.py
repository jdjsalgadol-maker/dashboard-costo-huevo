"""
Dashboard Ejecutivo — Costo del Huevo Fértil
=============================================
Fuente: export SAP 'BASE ZCO001' + Matrices Técnicas. 
Mantiene intacta la lógica de reconciliación financiera original y 
agrega consolidación automática por Granja basada en los informes de Levante y Producción.
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
# 1. MOTOR DE DATOS (ETL) - MANTIENE LÓGICA SAP EXACTA Y CRUZA GRANJAS
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

RECOMENDACIONES_RUBRO = {
    "Alimento": (
        "⚠️ <b>Factor Crítico - Alimento:</b> <i>¿Por qué subió?</i> Incremento en precio de "
        "mercado de materias primas o deterioro en la conversión alimenticia. <i>¿Qué hacer?</i> "
        "Evaluar formulación nutricional, revisar desperdicios en comederos y negociar compras a futuro."
    ),
    "Depreciación Parvada": (
        "⚠️ <b>Factor Crítico - Edad / Postura:</b> <i>¿Por qué subió?</i> La amortización contable "
        "por huevo se disparó. Ocurre cuando los lotes superan las 60 semanas y el % de postura cae. "
        "<i>¿Qué hacer?</i> Acelerar programa de descartes en lotes improductivos."
    ),
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
    """Estandariza los lotes para que los cruces numéricos vs texto no fallen"""
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

@st.cache_data(show_spinner="Procesando datos y reconciliando SAP con Granjas...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if "BASE ZCO001" not in xls.sheet_names:
        st.error("⚠️ Error Crítico: no se encontró la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

    # 1. Extracción Técnica para Mapeo de Granjas
    def extract_tech_data(sheet_name):
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.columns = df.columns.astype(str).str.strip()
            if 'Año' in df.columns: df['Anio'] = pd.to_numeric(df['Año'], errors='coerce').fillna(2026).astype(int)
            if 'No Mes' in df.columns: df['Mes_Num'] = pd.to_numeric(df['No Mes'], errors='coerce').fillna(1).astype(int)
            elif 'Mes' in df.columns:
                m_map = {'ENERO':1, 'FEBRERO':2, 'MARZO':3, 'ABRIL':4, 'MAYO':5, 'JUNIO':6, 'JULIO':7, 'AGOSTO':8, 'SEPTIEMBRE':9, 'OCTUBRE':10, 'NOVIEMBRE':11, 'DICIEMBRE':12}
                df['Mes_Num'] = df['Mes'].astype(str).str.upper().map(m_map).fillna(1).astype(int)
            if 'Anio' in df.columns and 'Mes_Num' in df.columns:
                df['Periodo'] = df['Anio'].astype(str) + "-" + df['Mes_Num'].astype(str).str.zfill(2)
            if 'Lote' in df.columns: df['Lote_str'] = df['Lote'].apply(clean_lote)
            return df
        return pd.DataFrame()

    df_lev = extract_tech_data("BD LEVANTE")
    df_prod = extract_tech_data("BD PRODUCCIÓN")

    lote_to_granja = {}
    for temp_df in [df_prod, df_lev]:
        if not temp_df.empty and 'Lote_str' in temp_df.columns and 'Nombre Granja' in temp_df.columns:
            valid = temp_df.dropna(subset=['Lote_str', 'Nombre Granja'])
            lote_to_granja.update(dict(zip(valid['Lote_str'], valid['Nombre Granja'])))

    # 2. Extracción Financiera SAP (BASE ZCO001)
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

    # Inyección de la Granja en la base SAP
    if 'Lote' in df_raw.columns:
        df_raw['Lote_str'] = df_raw['Lote'].apply(clean_lote)
        df_raw['Granja'] = df_raw['Lote_str'].map(lote_to_granja).fillna("Sin Granja Asignada")
    else:
        df_raw['Granja'] = "Sin Granja Asignada"

    # Consolidación Financiera (Costo Huevo Fértil)
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
    
    if "Alimento" in df_res.columns: df_res["Precio Kg Alimento"] = safe_div(df_res["Alimento"], df_res["Consumo Alimento Kg"])
    else: df_res["Precio Kg Alimento"] = np.nan
        
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
# 3. BARRA LATERAL — FILTROS DE TIEMPO Y MENÚ (SIN FILTROS JERÁRQUICOS)
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
        "1. Producción",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil",
        "4. Detalle Costos por Lote (y Granjas)",
        "5. Detalle Costos por Línea",
        "6. Costo Kg Alimento",
        "7. Lotes Finalizados - Levante",
        "8. Lotes Finalizados - Producción"
    ]
)

# =============================================================================
# 4. INTELIGENCIA DE NEGOCIO (texto automático)
# =============================================================================
def generar_diagnostico_costos(df_impactos, var_total, p_actual, p_base):
    if df_impactos.empty: return ""
    if var_total <= 0:
        rubro_exito = df_impactos.iloc[0]["Rubro"]
        return (f'<div class="insight-box">✅ <b>Eficiencia Operativa Alcanzada:</b> El costo unitario presenta una reducción de <b>${abs(var_total):,.2f} COP/Huevo</b> respecto a {p_base}.<br><i>Causa Raíz:</i> El rubro que más impulsó este ahorro fue <b>{rubro_exito}</b>. Mantener controles actuales.</div>')
    rubro_critico = df_impactos.iloc[-1]["Rubro"]
    impacto_critico = df_impactos.iloc[-1]["Impacto ($/HF)"]
    pct_explicado = (impacto_critico / var_total) * 100 if var_total else 0
    recomendacion = RECOMENDACIONES_RUBRO.get(rubro_critico, f"⚠️ <b>Factor Crítico - {rubro_critico}:</b> Explicó el <b>{pct_explicado:.1f}%</b> de la desviación desfavorable. Requiere revisión inmediata.")
    return (f'<div class="alert-box">🚨 <b>Alerta de Desviación Financiera (+${var_total:,.2f} COP/Huevo):</b><br>{recomendacion}</div>')

# =============================================================================
# MÓDULOS DEL DASHBOARD
# =============================================================================
if menu == "1. Producción":
    st.markdown('<p class="main-title">1. PRODUCCIÓN CONSOLIDADA DE HUEVO FÉRTIL</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    df_hf_f = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    tot_propios = df_hf_f["Cantidad"].sum()
    n_meses_rango = df_filtrado_graficas["Periodo"].nunique()
    c1, c2, c3 = st.columns(3)
    c1.metric("Volumen Total Período", f"{tot_propios:,.0f} HF")
    c2.metric("Lotes con producción", f"{df_hf_f['Lote'].nunique()}")
    c3.metric("Promedio Mensual", f"{tot_propios / max(n_meses_rango, 1):,.0f} HF/mes")
    st.markdown("---")
    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("📈 Evolución Histórica de Producción")
        df_ts = df_hf_f.groupby("Periodo")["Cantidad"].sum().reset_index()
        fig_ts = px.line(df_ts, x="Periodo", y="Cantidad", markers=True)
        fig_ts.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_ts, use_container_width=True)
    with col_r:
        st.subheader("🧬 Mix Genético")
        df_gen = df_hf_f.groupby("linea")["Cantidad"].sum().reset_index()
        if not df_gen.empty:
            fig_pie = px.pie(df_gen, values="Cantidad", names="linea", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÓN MENSUAL POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    df_hf = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    piv_prod = df_hf.groupby(["Periodo", "linea"])["Cantidad"].sum().unstack(fill_value=0)
    if not piv_prod.empty:
        fig_bar = px.bar(piv_prod.reset_index(), x="Periodo", y=piv_prod.columns, text_auto=".2s")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(piv_prod.style.format("{:,.0f}"), use_container_width=True)

elif menu == "3. Costo Huevo Fértil":
    st.markdown('<p class="main-title">3. ANÁLISIS INTEGRAL: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)
    if p_base == p_actual:
        st.warning("⚠️ Selecciona un Rango de Meses o el Modo Comparativo con dos períodos distintos en la barra lateral.")
        st.stop()
    fila_b = df[df["Periodo"] == p_base]
    fila_a = df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty: st.stop()
    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    var_tot = df_a["Costo Huevo Fértil"] - df_b["Costo Huevo Fértil"]
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {p_base} (Base)", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo {p_actual} (Actual)", f"${df_a['Costo Huevo Fértil']:,.2f}")
    delta_pct = (var_tot / df_b["Costo Huevo Fértil"]) * 100 if df_b["Costo Huevo Fértil"] else 0
    c3.metric("Desviación Total ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")
    
    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos Fértiles"], df_a["Huevos Fértiles"]
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({"Rubro": r, "Costo Unit Base ($)": ub, "Costo Unit Actual ($)": ua, "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)
    st.markdown(generar_diagnostico_costos(df_vif.sort_values("Impacto ($/HF)"), var_tot, p_actual, p_base), unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        fig_line_c = px.line(df_filtrado_graficas, x="Periodo", y="Costo Huevo Fértil", markers=True)
        fig_line_c.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_line_c, use_container_width=True)
    with col2:
        if not df_vif.empty:
            fig_tor = px.bar(df_vif.sort_values("Impacto ($/HF)"), y="Rubro", x="Impacto ($/HF)", orientation="h", color="Impacto ($/HF)", color_continuous_scale="Reds" if var_tot > 0 else "Greens")
            st.plotly_chart(fig_tor, use_container_width=True)

elif menu == "4. Detalle Costos por Lote (y Granjas)":
    st.markdown('<p class="main-title">4. DIAGNÓSTICO MICRO-OPERATIVO POR LOTE Y GRANJA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)]
    if df_l.empty:
        st.warning("No hay registros en SAP para el mes seleccionado.")
        st.stop()

    df_costo_lote = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov_lote = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    # Agrupar respetando la Granja que mapeamos en el ETL
    c_lote = df_costo_lote.groupby(["Granja", "Lote_str"])["Totales"].sum().reset_index(name="Costo Base")
    a_lote = df_aprov_lote.groupby(["Granja", "Lote_str"])["Totales"].sum().reset_index(name="Aprovechamientos")
    h_lote = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby(["Granja", "Lote_str"])["Cantidad"].sum().reset_index(name="Huevos Fértiles")

    df_m = h_lote.merge(c_lote, on=["Granja", "Lote_str"], how="left").merge(a_lote, on=["Granja", "Lote_str"], how="left").fillna(0)
    df_m["Costo Total"] = df_m["Costo Base"] + df_m["Aprovechamientos"]
    df_m["Costo Unitario"] = df_m["Costo Total"] / df_m["Huevos Fértiles"]
    df_m = df_m[df_m["Huevos Fértiles"] > 0].reset_index(drop=True)
    df_m = df_m.rename(columns={"Lote_str": "Lote"})
    
    # 1. Matriz de Granjas Automática
    st.subheader("🏢 Consolidado Operativo por Granja")
    df_granja = df_m.groupby("Granja").agg(
        Lotes_Asociados=("Lote", lambda x: ", ".join(x.unique())),
        Huevos_Fértiles=("Huevos Fértiles", "sum"),
        Costo_Total=("Costo Total", "sum")
    ).reset_index()
    df_granja["Costo Promedio ($/HF)"] = safe_div(df_granja["Costo_Total"], df_granja["Huevos_Fértiles"])
    st.dataframe(df_granja.style.format({
        "Huevos_Fértiles": "{:,.0f} unds", "Costo_Total": "${:,.0f}", "Costo Promedio ($/HF)": "${:,.2f}"
    }), use_container_width=True)
    
    st.markdown("---")
    # 2. Desglose por Lote
    col_grafico, col_tabla = st.columns([3, 2])
    with col_grafico:
        st.subheader("📊 Ranking de Ineficiencia por Lote")
        promedio_mes = df_m["Costo Total"].sum() / df_m["Huevos Fértiles"].sum()
        fig_bar_lote = px.bar(df_m.sort_values("Costo Unitario", ascending=False), x="Lote", y="Costo Unitario", text_auto=".1f", color="Costo Unitario", color_continuous_scale="RdYlGn_r")
        fig_bar_lote.add_hline(y=promedio_mes, line_dash="dot", annotation_text=f"Promedio: ${promedio_mes:,.1f}", line_color="black")
        st.plotly_chart(fig_bar_lote, use_container_width=True)

    with col_tabla:
        st.subheader("📋 Auditoría Lote a Lote")
        st.dataframe(df_m[["Granja", "Lote", "Huevos Fértiles", "Costo Total", "Costo Unitario"]].sort_values("Costo Unitario", ascending=False).style.format({
            "Huevos Fértiles": "{:,.0f} unds", "Costo Total": "${:,.0f}", "Costo Unitario": "${:,.2f}"
        }).background_gradient(subset=["Costo Unitario"], cmap="Reds"), use_container_width=True)

elif menu == "5. Detalle Costos por Línea":
    st.markdown('<p class="main-title">5. EVALUACIÓN FINANCIERA POR GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
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
        st.subheader("📈 Evolución Histórica por Raza")
        df_ts_g = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF].groupby(["Periodo", "linea"])["Cantidad"].sum().reset_index()
        fig_ts = px.line(df_ts_g, x="Periodo", y="Cantidad", color="linea", markers=True)
        st.plotly_chart(fig_ts, use_container_width=True)
    with col2:
        st.subheader(f"📊 Costo Unitario Promedio ({p_actual})")
        fig_bar = px.bar(df_gen, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".1f")
        st.plotly_chart(fig_bar, use_container_width=True)
    st.dataframe(df_gen[["linea", "Huevos Fértiles", "Costo Total ($)", "Costo Unitario ($/HF)"]].style.format({
        "Costo Total ($)": "${:,.0f}", "Huevos Fértiles": "{:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"
    }), use_container_width=True)

elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">6. IMPACTO DEL COSTO DE ALIMENTO</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Precio Alimento ($/Kg)")
        fig_a = px.line(df_filtrado_graficas, x="Periodo", y="Precio Kg Alimento", markers=True)
        fig_a.update_traces(line_color="#d62728", line_width=3)
        st.plotly_chart(fig_a, use_container_width=True)
    with col2:
        st.subheader("📈 Conversión (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x="Periodo", y="Gramos Alimento/Huevo", markers=True)
        fig_g.update_traces(line_color="#2ca02c", line_width=3)
        st.plotly_chart(fig_g, use_container_width=True)

elif menu == "7. Lotes Finalizados - Levante":
    st.markdown('<p class="main-title">7. RESULTADOS Y COSTOS LOTES FINALIZADOS (LEVANTE)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)
    
    if not df_lev.empty:
        df_l = df_lev[df_lev["Periodo"] == p_actual].copy()
        if not df_l.empty:
            st.subheader("🏢 Consolidado de Levante por Granja")
            matriz_lev = df_l.groupby("Nombre Granja").agg(
                Lotes_Asociados=("Lote_str", lambda x: ", ".join(x.dropna().unique())),
                Hembras_Encasetadas=("Hembras Encasetadas", "sum"),
                Aves_Finalizadas=("Total Aves Fin Levante", "sum") if "Total Aves Fin Levante" in df_l.columns else ("Hembras Encasetadas", "sum"),
                Costo_Total_Levante=("Costo Total Levante", "sum") if "Costo Total Levante" in df_l.columns else ("Hembras Encasetadas", "sum")
            ).reset_index()
            matriz_lev["Costo por Ave Finalizada"] = safe_div(matriz_lev["Costo_Total_Levante"], matriz_lev["Aves_Finalizadas"])
            st.dataframe(matriz_lev.style.format({
                "Hembras_Encasetadas": "{:,.0f}", "Aves_Finalizadas": "{:,.0f}", 
                "Costo_Total_Levante": "${:,.0f}", "Costo por Ave Finalizada": "${:,.2f}"
            }), use_container_width=True)
            
            st.markdown("---")
            st.subheader("📋 Desglose Técnico de Levante por Lote")
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
            st.info("No hay lotes de Levante reportados para el período seleccionado.")
    else:
        st.warning("Hoja BD LEVANTE no procesable o vacía.")

elif menu == "8. Lotes Finalizados - Producción":
    st.markdown('<p class="main-title">8. RESULTADOS TÉCNICOS FINALIZADOS (PRODUCCIÓN)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)
    
    if not df_prod.empty:
        df_p = df_prod[df_prod["Periodo"] == p_actual].copy()
        if not df_p.empty:
            st.subheader("🏢 Consolidado Técnico por Granja (Mortalidad y Costo)")
            matriz_prod_fin = df_p.groupby("Nombre Granja").agg(
                Lotes_Asociados=("Lote_str", lambda x: ", ".join(x.dropna().unique())),
                Hembras_Encasetadas=("Hembras Encasetadas", "sum"),
                Mortalidad=("Mortalidad Hembra", "sum") if "Mortalidad Hembra" in df_p.columns else ("Hembras Encasetadas", "sum"),
                Costo_Produccion=("Costo total Producción", "sum") if "Costo total Producción" in df_p.columns else ("Hembras Encasetadas", "sum")
            ).reset_index()
            matriz_prod_fin["% Mortalidad Global"] = safe_div(matriz_prod_fin["Mortalidad"], matriz_prod_fin["Hembras_Encasetadas"])
            st.dataframe(matriz_prod_fin.style.format({
                "Hembras_Encasetadas": "{:,.0f}", "Mortalidad": "{:,.0f}", 
                "Costo_Produccion": "${:,.0f}", "% Mortalidad Global": "{:.2%}"
            }), use_container_width=True)

            st.markdown("---")
            st.subheader("📋 Desglose Técnico por Lote")
            cols_mostrar = ['Nombre Granja', 'Lote', 'Hembras Encasetadas', 'Mortalidad Hembra', '%Mortalidad Hembra']
            if 'Costo total Producción' in df_p.columns: cols_mostrar.append('Costo total Producción')
            existentes = [c for c in cols_mostrar if c in df_p.columns]
            st.dataframe(df_p[existentes].style.format(na_rep="—", formatter={
                "Hembras Encasetadas": "{:,.0f}", "Mortalidad Hembra": "{:,.0f}", 
                "%Mortalidad Hembra": "{:.2%}", "Costo total Producción": "${:,.0f}"
            }), use_container_width=True)
        else:
            st.info("No hay datos técnicos de Producción para el período seleccionado.")
    else:
        st.warning("Hoja BD PRODUCCIÓN no procesable o vacía.")
