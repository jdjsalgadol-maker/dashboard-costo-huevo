"""
Dashboard Ejecutivo — Costo del Huevo Fértil
=============================================
Fuente: export SAP 'BASE ZCO001'. Reconciliado con informes gerenciales
de Reproductoras (Producción, Costo Huevo Fértil, Detalle Granja/Lote/Línea, Costo Kg Alimento).

Mejoras incorporadas:
  - Jerarquía operativa: Granja ('Nombre 1') ➔ Lote.
  - Clasificación contable: Costos Directos vs. Costos Indirectos de Fabricación (CIF - prefijo "PP").
  - Filtro exacto de la cuenta puente y aprovechamientos.
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
# 1. MOTOR DE DATOS (ETL)
# =============================================================================
MATERIAL_HF = "HUEVO INCUBABLE"
TEXTO_LIQUIDACION = "CTA PTE LIQ. ORD PCC Y MAQUILAS"
TEXTO_DIFERENCIA_PRECIO = "DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO"
RUBRO_APROVECHAMIENTO = "Aprovechamientos (-)"

MAP_RUBROS = {
    "CONSUMO ALIMENTO": "Alimento (Directo)",
    "CONSUMO CAMA": "Cama / Cascarilla (Directo)",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Bioseguridad y Aseo (Directo)",
    "CONSUMO DROGA": "Sanidad / Medicamentos (Directo)",
    "CONSUMOS MATERIA PRIMA": "Materias Primas / Calcio (Directo)",
}

RECOMENDACIONES_RUBRO = {
    "Alimento (Directo)": (
        "⚠️ <b>Factor Crítico - Alimento:</b> Incremento en precio de materias primas "
        "o desviación en la conversión alimenticia. Evaluar formulación y control de mermas en comederos."
    ),
    "Depreciación Parvada (CIF)": (
        "⚠️ <b>Factor Crítico - Edad / Postura:</b> La amortización por ave se incrementa cuando "
        "los lotes superan el ciclo óptimo o cae la postura. Acelerar programa de reemplazos."
    ),
}


def clean_num(x) -> float:
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    neg = s.endswith("-")
    s = s.replace("-", "").replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def safe_div(numer, denom):
    if isinstance(denom, pd.Series):
        denom = denom.replace(0, np.nan)
    elif denom == 0:
        denom = np.nan
    return numer / denom


def find_default_excel():
    for carpeta in (Path("data/raw"), Path(".")):
        if carpeta.exists():
            encontrados = sorted(f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$"))
            if encontrados:
                return encontrados[0]
    return None


@st.cache_data(show_spinner="Procesando datos y estructurando jerarquías...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if "BASE ZCO001" not in xls.sheet_names:
        st.error("⚠️ Error Crítico: no se encontró la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

    df_raw = pd.read_excel(xls, sheet_name="BASE ZCO001")
    df_raw.columns = df_raw.columns.astype(str).str.strip()

    # ---- Fechas / periodo ----
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

    # ---- Limpieza numérica ----
    df_raw["Totales"] = df_raw["Totales"].apply(clean_num)
    df_raw["Cantidad"] = df_raw["Cantidad"].apply(clean_num)

    # ---- Granja Normalizada ('Nombre 1') ----
    if "Nombre 1" in df_raw.columns:
        df_raw["Granja"] = df_raw["Nombre 1"].fillna("Granja General").str.strip()
    else:
        df_raw["Granja"] = "Granja General"

    # ---- Clasificación de Rubros (Directos vs CIF por prefijo 'PP') ----
    def clasificar_rubro(texto_explicativo):
        t = str(texto_explicativo).strip()
        if t in MAP_RUBROS:
            return MAP_RUBROS[t]
        if t.startswith("PP "):
            # Limpiar nombre del CIF
            clean_name = t.replace("PP ", "").replace("Grj.Pcc.", "").strip()
            return f"{clean_name} (CIF)"
        return t

    df_raw["Rubro_Clasificado"] = df_raw["Texto explicativo"].apply(clasificar_rubro)

    # ---- Huevos fértiles ----
    df_hf = df_raw[df_raw["Texto breve de material"] == MATERIAL_HF]
    hf_mes = df_hf.groupby("Periodo")["Cantidad"].sum()

    # ---- Costos directos e indirectos (excluyendo cuenta puente y diferencia precio) ----
    df_costos = df_raw[~df_raw["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
    costos_piv = df_costos.groupby(["Periodo", "Rubro_Clasificado"])["Totales"].sum().unstack(fill_value=0)

    # ---- Aprovechamiento (dentro de cuenta de liquidación) ----
    df_aprov = df_raw[(df_raw["Texto explicativo"] == TEXTO_LIQUIDACION) &
                       (df_raw["Texto breve de material"] != MATERIAL_HF)]
    aprov_mes = df_aprov.groupby("Periodo")["Totales"].sum()

    df_res = costos_piv.copy()
    df_res[RUBRO_APROVECHAMIENTO] = aprov_mes.reindex(df_res.index).fillna(0)

    df_res["Costo Total"] = df_res.sum(axis=1)
    df_res["Huevos Fértiles"] = hf_mes.reindex(df_res.index)
    df_res["Costo Huevo Fértil"] = safe_div(df_res["Costo Total"], df_res["Huevos Fértiles"])

    # ---- Alimento Kg ----
    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"]
    alim_kg = df_alim.groupby("Periodo")["Cantidad"].sum()
    df_res["Consumo Alimento Kg"] = alim_kg.reindex(df_res.index)
    
    col_alim_key = [c for c in df_res.columns if "Alimento" in c and "Directo" in c]
    if col_alim_key:
        df_res["Precio Kg Alimento"] = safe_div(df_res[col_alim_key[0]], df_res["Consumo Alimento Kg"])
    else:
        df_res["Precio Kg Alimento"] = np.nan
        
    df_res["Gramos Alimento/Huevo"] = safe_div(df_res["Consumo Alimento Kg"] * 1000, df_res["Huevos Fértiles"])

    df_res = df_res.reset_index().sort_values("Periodo")
    df_res["Anio"] = df_res["Periodo"].str.split("-").str[0].astype(int)
    df_res["Mes_Num"] = df_res["Periodo"].str.split("-").str[1].astype(int)

    rubros = [c for c in df_res.columns if c not in ["Periodo", "Costo Total", "Huevos Fértiles", "Costo Huevo Fértil", "Consumo Alimento Kg", "Precio Kg Alimento", "Gramos Alimento/Huevo", "Anio", "Mes_Num"]]

    return df_res, rubros, df_raw


# =============================================================================
# 2. CARGA DE DATOS
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")
uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df, rubros_items, df_raw = load_and_process_data(uploaded_file)
else:
    archivo_default = find_default_excel()
    if archivo_default is None:
        st.warning("⚠️ Se requiere un archivo Excel para inicializar el sistema. "
                   "Cárgalo en la barra lateral o colócalo en `data/raw/`.")
        st.stop()
    df, rubros_items, df_raw = load_and_process_data(str(archivo_default))

if df.empty:
    st.error("El archivo no contiene datos procesables en la hoja 'BASE ZCO001'.")
    st.stop()

# =============================================================================
# 3. BARRA LATERAL — FILTROS DE TIEMPO Y GRANJA
# =============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal y Ubicación")

anios_disp = sorted(df["Anio"].unique(), reverse=True)


def selector_periodo(titulo, key_prefix, anios_opciones, mes_reverse=False):
    st.sidebar.markdown(f"**{titulo}:**")
    c1, c2 = st.sidebar.columns(2)
    anio = c1.selectbox("Año", anios_opciones, key=f"{key_prefix}_anio")
    meses_disp = sorted(df[df["Anio"] == anio]["Mes_Num"].unique(), reverse=mes_reverse)
    mes = c2.selectbox("Mes", meses_disp, format_func=lambda x: MESES_NOMBRES[x], key=f"{key_prefix}_mes")
    return f"{anio}-{str(mes).zfill(2)}"


modo_analisis = st.sidebar.radio(
    "Seleccione el enfoque:",
    ["⚖️ Comparativo (Mes VS Mes)", "📈 Rango Histórico (Evolución)"],
)

if modo_analisis == "⚖️ Comparativo (Mes VS Mes)":
    p_base = selector_periodo("Período Base", "base", anios_disp)
    p_actual = selector_periodo("Período Actual", "actual", anios_disp, mes_reverse=True)
    texto_contexto = (f"Comparativa Estratégica: **{MESES_NOMBRES[int(p_actual.split('-')[1])]} {p_actual.split('-')[0]}** "
                       f"VS **{MESES_NOMBRES[int(p_base.split('-')[1])]} {p_base.split('-')[0]}**")
else:
    p_base = selector_periodo("Inicio del rango", "ini", sorted(anios_disp))
    p_actual = selector_periodo("Fin del rango", "fin", anios_disp, mes_reverse=True)
    texto_contexto = f"Evolución Histórica: Desde **{min(p_base, p_actual)}** hasta **{max(p_base, p_actual)}**"

rango_inicio, rango_fin = min(p_base, p_actual), max(p_base, p_actual)
df_filtrado_graficas = df[(df["Periodo"] >= rango_inicio) & (df["Periodo"] <= rango_fin)].sort_values("Periodo")
df_raw_graficas = df_raw[(df_raw["Periodo"] >= rango_inicio) & (df_raw["Periodo"] <= rango_fin)]

# Filtro opcional por Granja
granjas_disp = ["Todas las Granjas"] + sorted(df_raw_graficas["Granja"].unique().tolist())
granja_seleccionada = st.sidebar.selectbox("Filtro Granja:", granjas_disp)

if granja_seleccionada != "Todas las Granjas":
    df_raw_graficas = df_raw_graficas[df_raw_graficas["Granja"] == granja_seleccionada]

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. Producción",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil",
        "4. Detalle Costos por Lote (Granja ➔ Lote)",
        "5. Detalle Costos por Línea",
        "6. Costo Kg Alimento",
    ],
)


# =============================================================================
# 4. INTELIGENCIA DE NEGOCIO
# =============================================================================
def generar_diagnostico_costos(df_impactos, var_total, p_actual, p_base):
    if df_impactos.empty:
        return ""
    if var_total <= 0:
        rubro_exito = df_impactos.iloc[0]["Rubro"]
        return (
            f'<div class="insight-box">✅ <b>Eficiencia Operativa:</b> Reducción de '
            f'<b>${abs(var_total):,.2f} COP/Huevo</b> frente a {p_base}. '
            f'Impulsado principalmente por el rubro: <b>{rubro_exito}</b>.</div>'
        )
    rubro_critico = df_impactos.iloc[-1]["Rubro"]
    impacto_critico = df_impactos.iloc[-1]["Impacto ($/HF)"]
    pct_explicado = (impacto_critico / var_total) * 100 if var_total else 0
    recomendacion = RECOMENDACIONES_RUBRO.get(
        rubro_critico,
        f"⚠️ <b>Factor Crítico - {rubro_critico}:</b> Explicó el <b>{pct_explicado:.1f}%</b> de la desviación desfavorable."
    )
    return f'<div class="alert-box">🚨 <b>Alerta de Desviación Financiera (+${var_total:,.2f} COP/Huevo):</b><br>{recomendacion}</div>'


# =============================================================================
# 5. PÁGINA 1 — PRODUCCIÓN
# =============================================================================
if menu == "1. Producción":
    st.markdown('<p class="main-title">1. PRODUCCIÓN CONSOLIDADA DE HUEVO FÉRTIL (GRANJAS PROPIAS)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf_f = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    tot_propios = df_hf_f["Cantidad"].sum()
    n_meses_rango = df_filtrado_graficas["Periodo"].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Volumen Total Período", f"{tot_propios:,.0f} HF")
    c2.metric("Granjas Activas", f"{df_hf_f['Granja'].nunique()}")
    c3.metric("Promedio Mensual", f"{tot_propios / max(n_meses_rango, 1):,.0f} HF/mes")

    st.markdown("---")
    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("📈 Evolución Histórica de Producción")
        df_ts = df_hf_f.groupby("Periodo")["Cantidad"].sum().reset_index()
        fig_ts = px.line(df_ts, x="Periodo", y="Cantidad", markers=True, title="Volumen en el Rango")
        fig_ts.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_ts, use_container_width=True)
    with col_r:
        st.subheader("🧬 Mix Genético")
        df_gen = df_hf_f.groupby("linea")["Cantidad"].sum().reset_index()
        if not df_gen.empty:
            fig_pie = px.pie(df_gen, values="Cantidad", names="linea", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)


# =============================================================================
# 6. PÁGINA 2 — PRODUCCIÓN MES A MES POR LÍNEA
# =============================================================================
elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÓN MENSUAL POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    piv_prod = df_hf.groupby(["Periodo", "linea"])["Cantidad"].sum().unstack(fill_value=0)

    if not piv_prod.empty:
        fig_bar = px.bar(piv_prod.reset_index(), x="Periodo", y=piv_prod.columns, title="Barras Apiladas por Raza", text_auto=".2s")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(piv_prod.style.format("{:,.0f}"), use_container_width=True)


# =============================================================================
# 7. PÁGINA 3 — COSTO HUEVO FÉRTIL
# =============================================================================
elif menu == "3. Costo Huevo Fértil":
    st.markdown('<p class="main-title">3. ANÁLISIS INTEGRAL: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)

    if p_base == p_actual:
        st.warning("⚠️ Selecciona un Rango de Meses o Modo Comparativo con dos períodos distintos en la barra lateral.")
        st.stop()

    fila_b = df[df["Periodo"] == p_base]
    fila_a = df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty:
        st.error("Datos insuficientes para comparar.")
        st.stop()

    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    var_tot = df_a["Costo Huevo Fértil"] - df_b["Costo Huevo Fértil"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {p_base}", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo {p_actual}", f"${df_a['Costo Huevo Fértil']:,.2f}")
    c3.metric("Desviación ($/HF)", f"${var_tot:+,.2f}", delta_color="inverse")

    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos Fértiles"], df_a["Huevos Fértiles"]
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({"Rubro": r, "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)

    if not df_vif.empty:
        st.markdown(generar_diagnostico_costos(df_vif.sort_values("Impacto ($/HF)"), var_tot, p_actual, p_base), unsafe_allow_html=True)

    st.dataframe(df_vif.style.format({"Impacto ($/HF)": "${:+,.2f}"}), use_container_width=True)


# =============================================================================
# 8. PÁGINA 4 — DETALLE COSTOS POR LOTE (GRANJA ➔ LOTE)
# =============================================================================
elif menu == "4. Detalle Costos por Lote (Granja ➔ Lote)":
    st.markdown('<p class="main-title">4. DIAGNÓSTICO MICRO-OPERATIVO: GRANJA ➔ LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    df_l = df_raw[df_raw["Periodo"] == p_actual]
    if granja_seleccionada != "Todas las Granjas":
        df_l = df_l[df_l["Granja"] == granja_seleccionada]

    df_c = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].groupby(["Granja", "Lote"])["Totales"].sum()
    df_aprov_lote = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) &
                          (df_l["Texto breve de material"] != MATERIAL_HF)].groupby(["Granja", "Lote"])["Totales"].sum()
    df_h = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby(["Granja", "Lote"])["Cantidad"].sum()

    df_m = pd.DataFrame({"Costo Directo": df_c, "Aprovechamiento": df_aprov_lote, "Huevos Fértiles": df_h}).fillna(0)
    df_m = df_m[df_m["Huevos Fértiles"] > 0].copy()
    df_m["Costo Total"] = df_m["Costo Directo"] + df_m["Aprovechamiento"]
    df_m["Costo Unitario"] = df_m["Costo Total"] / df_m["Huevos Fértiles"]
    df_m = df_m.reset_index()

    if df_m.empty:
        st.warning("No hay registros para la selección actual.")
        st.stop()

    df_m["Lote"] = df_m["Lote"].astype(str)
    df_m = df_m.sort_values("Costo Unitario", ascending=False)
    promedio_mes = df_m["Costo Total"].sum() / df_m["Huevos Fértiles"].sum()

    st.subheader("📋 Matriz Jerárquica Granja ➔ Lote")
    st.dataframe(
        df_m.style.format({
            "Huevos Fértiles": "{:,.0f} unds", "Costo Total": "${:,.0f}", "Costo Unitario": "${:,.2f}",
        }).background_gradient(subset=["Costo Unitario"], cmap="Reds"),
        use_container_width=True,
    )

    st.subheader("📊 Costo Unitario por Lote y Granja")
    fig_hier = px.bar(
        df_m, x="Lote", y="Costo Unitario", color="Granja", text_auto=".1f",
        title=f"Costo por Lote agrupado por Granja - {p_actual}"
    )
    fig_hier.add_hline(y=promedio_mes, line_dash="dot", annotation_text=f"Promedio: ${promedio_mes:,.1f}", line_color="black")
    st.plotly_chart(fig_hier, use_container_width=True)


# =============================================================================
# 9. PÁGINA 5 — DETALLE COSTOS POR LÍNEA
# =============================================================================
elif menu == "5. Detalle Costos por Línea":
    st.markdown('<p class="main-title">5. EVALUACIÓN FINANCIERA POR GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_c = df_raw_graficas[~df_raw_graficas["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov = df_raw_graficas[(df_raw_graficas["Texto explicativo"] == TEXTO_LIQUIDACION) &
                                (df_raw_graficas["Texto breve de material"] != MATERIAL_HF)]
    df_h = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]

    c_lin = df_c[df_c["Periodo"] == p_actual].groupby("linea")["Totales"].sum()
    aprov_lin = df_aprov[df_aprov["Periodo"] == p_actual].groupby("linea")["Totales"].sum()
    h_lin = df_h[df_h["Periodo"] == p_actual].groupby("linea")["Cantidad"].sum()

    df_gen = pd.DataFrame({"Costo Directo ($)": c_lin, "Aprovechamiento ($)": aprov_lin, "Huevos Fértiles": h_lin}).fillna(0)
    df_gen = df_gen[df_gen["Huevos Fértiles"] > 0].copy()
    df_gen["Costo Total ($)"] = df_gen["Costo Directo ($)"] + df_gen["Aprovechamiento ($)"]
    df_gen["Costo Unitario ($/HF)"] = df_gen["Costo Total ($)"] / df_gen["Huevos Fértiles"]
    df_gen = df_gen.reset_index().rename(columns={"index": "linea"})

    if not df_gen.empty:
        fig_bar = px.bar(df_gen, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".1f")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(df_gen.style.format({"Costo Total ($)": "${:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"}), use_container_width=True)


# =============================================================================
# 10. PÁGINA 6 — COSTO KG ALIMENTO
# =============================================================================
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">6. IMPACTO DEL COSTO DE ALIMENTO</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Evolución Precio Alimento ($/Kg)")
        fig_a = px.line(df_filtrado_graficas, x="Periodo", y="Precio Kg Alimento", markers=True)
        st.plotly_chart(fig_a, use_container_width=True)
    with col2:
        st.subheader("📈 Evolución Conversión (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x="Periodo", y="Gramos Alimento/Huevo", markers=True)
        st.plotly_chart(fig_g, use_container_width=True)
