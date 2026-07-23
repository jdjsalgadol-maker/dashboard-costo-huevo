import os
import io
from pathlib import Path
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# =============================================================================
# 0. CONFIGURACIÃ“N DE PÃGINA Y ESTILOS
# =============================================================================
st.set_page_config(
    page_title="Dashboard Ejecutivo - Costos Huevo FÃ©rtil",
    page_icon="ðŸ”",
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
    "CONSUMO ALIMENTO": "Alimento",
    "PP Depr. Gallina Grj.Pcc.": "DepreciaciÃ³n Parvada",
    "PP Horas Hombre Grj.Pcc.": "Mano de Obra",
    "PP Costos Ind. Grj.Pcc.": "Costos Indirectos (CIF)",
    "PP Costos Arriendo Grj.Pcc.": "Arriendo",
    "CONSUMO CAMA": "Cama / Cascarilla",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Bioseguridad y Aseo",
    "CONSUMO DROGA": "Sanidad (Medicamentos)",
    "PP Costos Depr. Grj.Pcc.": "DepreciaciÃ³n Instalaciones",
    "CONSUMOS MATERIA PRIMA": "Materias Primas (Calcio)",
}

RECOMENDACIONES_RUBRO = {
    "Alimento": (
        "âš ï¸ <b>Factor CrÃ­tico - Alimento:</b> <i>Â¿Por quÃ© subiÃ³?</i> Incremento en precio de "
        "mercado de materias primas o deterioro en la conversiÃ³n alimenticia. <i>Â¿QuÃ© hacer?</i> "
        "Evaluar formulaciÃ³n nutricional, revisar desperdicios en comederos y negociar compras a futuro."
    ),
    "DepreciaciÃ³n Parvada": (
        "âš ï¸ <b>Factor CrÃ­tico - Edad / Postura:</b> <i>Â¿Por quÃ© subiÃ³?</i> La amortizaciÃ³n contable "
        "por huevo se disparÃ³. Ocurre cuando los lotes superan las 60 semanas y el % de postura cae. "
        "<i>Â¿QuÃ© hacer?</i> Acelerar programa de descartes en lotes improductivos."
    ),
}

def clean_num(x) -> float:
    """Parsea el formato negativo particular de SAP"""
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
    """DivisiÃ³n segura: evita divisiÃ³n por cero."""
    if isinstance(denom, pd.Series):
        denom = denom.replace(0, np.nan)
    elif denom == 0:
        denom = np.nan
    return numer / denom

def find_default_excel():
    """Busca archivo Excel en la ruta local."""
    for carpeta in (Path("data/raw"), Path(".")):
        if carpeta.exists():
            encontrados = sorted(f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$"))
            if encontrados:
                return encontrados[0]
    return None

@st.cache_data(show_spinner="Procesando datos y reconciliando SAP...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if "BASE ZCO001" not in xls.sheet_names:
        st.error("âš ï¸ Error CrÃ­tico: no se encontrÃ³ la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

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

    # Huevos FÃ©rtiles
    df_hf = df_raw[df_raw["Texto breve de material"] == MATERIAL_HF]
    hf_mes = df_hf.groupby("Periodo")["Cantidad"].sum()

    # Costos Directos 
    df_costos = df_raw[~df_raw["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
    df_costos["Rubro"] = df_costos["Texto explicativo"].map(lambda x: MAP_RUBROS.get(x, x))
    costos_piv = df_costos.groupby(["Periodo", "Rubro"])["Totales"].sum().unstack(fill_value=0)

    # APROVECHAMIENTO: La clave de la reconciliaciÃ³n
    df_aprov = df_raw[(df_raw["Texto explicativo"] == TEXTO_LIQUIDACION) &
                       (df_raw["Texto breve de material"] != MATERIAL_HF)]
    aprov_mes = df_aprov.groupby("Periodo")["Totales"].sum()

    df_res = costos_piv.copy()
    df_res[RUBRO_APROVECHAMIENTO] = aprov_mes.reindex(df_res.index).fillna(0)

    df_res["Costo Total"] = df_res.sum(axis=1)
    df_res["Huevos FÃ©rtiles"] = hf_mes.reindex(df_res.index)
    df_res["Costo Huevo FÃ©rtil"] = safe_div(df_res["Costo Total"], df_res["Huevos FÃ©rtiles"])

    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"]
    alim_kg = df_alim.groupby("Periodo")["Cantidad"].sum()
    df_res["Consumo Alimento Kg"] = alim_kg.reindex(df_res.index)
    
    if "Alimento" in df_res.columns:
        df_res["Precio Kg Alimento"] = safe_div(df_res["Alimento"], df_res["Consumo Alimento Kg"])
    else:
        df_res["Precio Kg Alimento"] = np.nan
        
    df_res["Gramos Alimento/Huevo"] = safe_div(df_res["Consumo Alimento Kg"] * 1000, df_res["Huevos FÃ©rtiles"])

    df_res = df_res.reset_index().sort_values("Periodo")
    df_res["Anio"] = df_res["Periodo"].str.split("-").str[0].astype(int)
    df_res["Mes_Num"] = df_res["Periodo"].str.split("-").str[1].astype(int)

    rubros = [r for r in list(MAP_RUBROS.values()) + [RUBRO_APROVECHAMIENTO] if r in df_res.columns]

    return df_res, rubros, df_raw

# =============================================================================
# 2. CARGA DE DATOS
# =============================================================================
st.sidebar.title("ðŸ” BI AvÃ­cola Gerencial")
uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df, rubros_items, df_raw = load_and_process_data(uploaded_file)
else:
    archivo_default = find_default_excel()
    if archivo_default is None:
        st.warning("âš ï¸ Se requiere un archivo Excel para inicializar el sistema. "
                   "CÃ¡rgalo en la barra lateral o colÃ³calo en `data/raw/`.")
        st.stop()
    df, rubros_items, df_raw = load_and_process_data(str(archivo_default))

if df.empty:
    st.error("El archivo no contiene datos procesables en la hoja 'BASE ZCO001'.")
    st.stop()

# =============================================================================
# 3. BARRA LATERAL â€” FILTROS DE TIEMPO
# =============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("ðŸ“… Control Temporal")

anios_disp = sorted(df["Anio"].unique(), reverse=True)

def selector_periodo(titulo, key_prefix, anios_opciones, mes_reverse=False):
    st.sidebar.markdown(f"**{titulo}:**")
    c1, c2 = st.sidebar.columns(2)
    anio = c1.selectbox("AÃ±o", anios_opciones, key=f"{key_prefix}_anio")
    meses_disp = sorted(df[df["Anio"] == anio]["Mes_Num"].unique(), reverse=mes_reverse)
    mes = c2.selectbox("Mes", meses_disp, format_func=lambda x: MESES_NOMBRES[x], key=f"{key_prefix}_mes")
    return f"{anio}-{str(mes).zfill(2)}"

modo_analisis = st.sidebar.radio(
    "Seleccione el enfoque:",
    ["âš–ï¸ Comparativo (Mes VS Mes)", "ðŸ“ˆ Rango HistÃ³rico (EvoluciÃ³n)"],
)

if modo_analisis == "âš–ï¸ Comparativo (Mes VS Mes)":
    p_base = selector_periodo("PerÃ­odo Base (contra quÃ© comparo)", "base", anios_disp)
    p_actual = selector_periodo("PerÃ­odo Actual (quÃ© estoy evaluando)", "actual", anios_disp, mes_reverse=True)
    texto_contexto = (f"Comparativa EstratÃ©gica: **{MESES_NOMBRES[int(p_actual.split('-')[1])]} {p_actual.split('-')[0]}** "
                       f"VS **{MESES_NOMBRES[int(p_base.split('-')[1])]} {p_base.split('-')[0]}**")
else:
    p_base = selector_periodo("Inicio del rango", "ini", sorted(anios_disp))
    p_actual = selector_periodo("Fin del rango", "fin", anios_disp, mes_reverse=True)
    texto_contexto = f"EvoluciÃ³n HistÃ³rica: Desde **{min(p_base, p_actual)}** hasta **{max(p_base, p_actual)}**"

rango_inicio, rango_fin = min(p_base, p_actual), max(p_base, p_actual)
df_filtrado_graficas = df[(df["Periodo"] >= rango_inicio) & (df["Periodo"] <= rango_fin)].sort_values("Periodo")
df_raw_graficas = df_raw[(df_raw["Periodo"] >= rango_inicio) & (df_raw["Periodo"] <= rango_fin)]

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. ProducciÃ³n",
        "2. ProducciÃ³n Mes a Mes por LÃ­nea",
        "3. Costo Huevo FÃ©rtil",
        "4. Detalle Costos por Lote",
        "5. Detalle Costos por LÃ­nea",
        "6. Costo Kg Alimento",
    ],
)

# =============================================================================
# 4. INTELIGENCIA DE NEGOCIO (texto automÃ¡tico)
# =============================================================================
def generar_diagnostico_costos(df_impactos, var_total, p_actual, p_base):
    if df_impactos.empty:
        return ""

    if var_total <= 0:
        rubro_exito = df_impactos.iloc[0]["Rubro"]
        return (
            f'<div class="insight-box">âœ… <b>Eficiencia Operativa Alcanzada:</b> El costo unitario '
            f'presenta una reducciÃ³n de <b>${abs(var_total):,.2f} COP/Huevo</b> respecto a {p_base}.<br>'
            f'<i>Causa RaÃ­z:</i> El rubro que mÃ¡s impulsÃ³ este ahorro fue <b>{rubro_exito}</b>. '
            f'Mantener controles actuales.</div>'
        )

    rubro_critico = df_impactos.iloc[-1]["Rubro"]
    impacto_critico = df_impactos.iloc[-1]["Impacto ($/HF)"]
    pct_explicado = (impacto_critico / var_total) * 100 if var_total else 0

    recomendacion = RECOMENDACIONES_RUBRO.get(
        rubro_critico,
        f"âš ï¸ <b>Factor CrÃ­tico - {rubro_critico}:</b> ExplicÃ³ el <b>{pct_explicado:.1f}%</b> de la "
        f"desviaciÃ³n desfavorable. Requiere revisiÃ³n inmediata de presupuestos.",
    )
    return (
        f'<div class="alert-box">ðŸš¨ <b>Alerta de DesviaciÃ³n Financiera '
        f'(+${var_total:,.2f} COP/Huevo):</b><br>{recomendacion}</div>'
    )

# =============================================================================
# 5. PÃGINA 1 â€” PRODUCCIÃ“N
# =============================================================================
if menu == "1. ProducciÃ³n":
    st.markdown('<p class="main-title">1. PRODUCCIÃ“N CONSOLIDADA DE HUEVO FÃ‰RTIL (GRANJAS PROPIAS)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf_f = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    tot_propios = df_hf_f["Cantidad"].sum()
    n_meses_rango = df_filtrado_graficas["Periodo"].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Volumen Total PerÃ­odo (Propios)", f"{tot_propios:,.0f} HF")
    c2.metric("Lotes con producciÃ³n", f"{df_hf_f['Lote'].nunique()}")
    c3.metric("Promedio Mensual", f"{tot_propios / max(n_meses_rango, 1):,.0f} HF/mes")

    st.markdown("""
    <div class="warn-box">â„¹ï¸ Esta base contiene Ãºnicamente granjas <b>propias</b>. La producciÃ³n de
    terceros ("Externos"/maquila) no estÃ¡ incluida en el archivo de origen, por lo que no se muestra
    aquÃ­ para evitar reportar una cifra estimada como si fuera un dato real.</div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("ðŸ“ˆ EvoluciÃ³n HistÃ³rica de ProducciÃ³n")
        df_ts = df_hf_f.groupby("Periodo")["Cantidad"].sum().reset_index()
        fig_ts = px.line(df_ts, x="Periodo", y="Cantidad", markers=True,
                          title="Comportamiento del Volumen en el Rango Seleccionado")
        fig_ts.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_ts, use_container_width=True)

    with col_r:
        st.subheader("ðŸ§¬ Mix GenÃ©tico")
        df_gen = df_hf_f.groupby("linea")["Cantidad"].sum().reset_index()
        if not df_gen.empty:
            fig_pie = px.pie(df_gen, values="Cantidad", names="linea", hole=0.4, title="ParticipaciÃ³n por Raza")
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Sin datos de producciÃ³n por lÃ­nea en el perÃ­odo seleccionado.")

# =============================================================================
# 6. PÃGINA 2 â€” PRODUCCIÃ“N MES A MES POR LÃNEA
# =============================================================================
elif menu == "2. ProducciÃ³n Mes a Mes por LÃ­nea":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÃ“N MENSUAL POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF]
    piv_prod = df_hf.groupby(["Periodo", "linea"])["Cantidad"].sum().unstack(fill_value=0)

    if piv_prod.empty:
        st.info("Sin datos de producciÃ³n para el perÃ­odo seleccionado.")
    else:
        fig_bar = px.bar(piv_prod.reset_index(), x="Periodo", y=piv_prod.columns,
                          title="Comportamiento Mensual (Barras Apiladas)", text_auto=".2s")
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("Matriz Exacta de ProducciÃ³n (Unidades):")
        st.dataframe(piv_prod.style.format("{:,.0f}"), use_container_width=True)

# =============================================================================
# 7. PÃGINA 3 â€” COSTO HUEVO FÃ‰RTIL (VIF, EVOLUTIVA & SIMULADOR)
# =============================================================================
elif menu == "3. Costo Huevo FÃ©rtil":
    st.markdown('<p class="main-title">3. ANÃLISIS INTEGRAL: COSTO HUEVO FÃ‰RTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">AnÃ¡lisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)

    if p_base == p_actual:
        st.warning("âš ï¸ Selecciona un Rango de Meses o el Modo Comparativo con dos perÃ­odos distintos en la barra lateral.")
        st.stop()

    fila_b = df[df["Periodo"] == p_base]
    fila_a = df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty:
        st.error("No hay datos suficientes en los perÃ­odos seleccionados para comparar.")
        st.stop()

    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]

    if pd.isna(df_b["Costo Huevo FÃ©rtil"]) or pd.isna(df_a["Costo Huevo FÃ©rtil"]):
        st.error("Alguno de los dos perÃ­odos no tiene huevos fÃ©rtiles registrados; "
                 "no es posible calcular el costo unitario.")
        st.stop()

    var_tot = df_a["Costo Huevo FÃ©rtil"] - df_b["Costo Huevo FÃ©rtil"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {p_base} (Base)", f"${df_b['Costo Huevo FÃ©rtil']:,.2f}")
    c2.metric(f"Costo {p_actual} (Actual)", f"${df_a['Costo Huevo FÃ©rtil']:,.2f}")
    delta_pct = (var_tot / df_b["Costo Huevo FÃ©rtil"]) * 100 if df_b["Costo Huevo FÃ©rtil"] else 0
    c3.metric("DesviaciÃ³n Total ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")

    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos FÃ©rtiles"], df_a["Huevos FÃ©rtiles"]
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({
            "Rubro": r, "Costo Unit Base ($)": ub, "Costo Unit Actual ($)": ua,
            "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan
        })
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)

    st.markdown(generar_diagnostico_costos(df_vif.sort_values("Impacto ($/HF)"), var_tot, p_actual, p_base), unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("ðŸ“ˆ EvoluciÃ³n HistÃ³rica de Costo ($)")
        fig_line_c = px.line(df_filtrado_graficas, x="Periodo", y="Costo Huevo FÃ©rtil", markers=True)
        fig_line_c.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_line_c, use_container_width=True)
    with col2:
        st.subheader("ðŸ“Š Tornado de DesviaciÃ³n por Rubro")
        if not df_vif.empty:
            fig_tor = px.bar(df_vif.sort_values("Impacto ($/HF)"), y="Rubro", x="Impacto ($/HF)", orientation="h",
                             color="Impacto ($/HF)", color_continuous_scale="Reds" if var_tot > 0 else "Greens")
            st.plotly_chart(fig_tor, use_container_width=True)

    st.markdown("---")
    st.subheader("ðŸ—“ï¸ Matriz de EvoluciÃ³n Mensual del Costo (Rango Seleccionado)")
    df_evolucion = df_filtrado_graficas[["Periodo", "Huevos FÃ©rtiles", "Costo Total", "Costo Huevo FÃ©rtil"]].copy()
    df_evolucion["VariaciÃ³n ($/HF)"] = df_evolucion["Costo Huevo FÃ©rtil"].diff()
    df_evolucion["% VariaciÃ³n"] = df_evolucion["Costo Huevo FÃ©rtil"].pct_change() * 100

    st.dataframe(df_evolucion.style.format({
        "Huevos FÃ©rtiles": "{:,.0f}", "Costo Total": "${:,.0f}", "Costo Huevo FÃ©rtil": "${:,.2f}",
        "VariaciÃ³n ($/HF)": "${:+,.2f}", "% VariaciÃ³n": "{:+.2f}%"
    }), use_container_width=True)

    st.markdown("---")
    st.subheader("ðŸ“‹ Matriz Comparativa (VariaciÃ³n VIF Directa)")
    st.dataframe(df_vif.style.format({
        "Costo Unit Base ($)": "${:,.2f}", "Costo Unit Actual ($)": "${:,.2f}", "Impacto ($/HF)": "${:+,.2f}"
    }), use_container_width=True)

    st.markdown("---")
    st.subheader("ðŸŽ›ï¸ Modulo Predictivo: Estrategias de ReducciÃ³n de Costo")
    st.markdown(f"Evaluando sobre la base operativa de **{p_actual}**:")
    
    s1, s2, s3 = st.columns(3)
    ahorro_alim = s1.number_input("1. NegociaciÃ³n Alimento: Disminuir precio en ($/Kg):", value=50)
    mejora_conv = s2.number_input("2. Eficiencia: Bajar consumo (g/huevo):", value=15)
    mejora_post = s3.number_input("3. Productividad: Subir % de postura en:", value=3.0, step=0.5)
    
    precio_actual_kg = df_a["Precio Kg Alimento"] if pd.notna(df_a["Precio Kg Alimento"]) else 0
    nuevo_p_alim = max(precio_actual_kg - ahorro_alim, 500)
    nuevo_cons_kg = max(df_a["Consumo Alimento Kg"] - ((mejora_conv / 1000) * df_a["Huevos FÃ©rtiles"]), 0)
    nuevo_costo_alim = nuevo_cons_kg * nuevo_p_alim
    
    nuevos_hf = df_a["Huevos FÃ©rtiles"] * (1 + (mejora_post / 100))
    costo_sin_alimento = df_a["Costo Total"] - df_a.get("Alimento", 0)
    nuevo_costo_total = costo_sin_alimento + nuevo_costo_alim
    nuevo_costo_huevo = nuevo_costo_total / nuevos_hf if nuevos_hf > 0 else 0
    ahorro_un = df_a["Costo Huevo FÃ©rtil"] - nuevo_costo_huevo
    
    st.success(f"ðŸŽ¯ **ProyecciÃ³n Exitosa:** Ejecutando estas tres acciones conjuntas, el costo bajarÃ­a de **${df_a['Costo Huevo FÃ©rtil']:,.1f}** a **${nuevo_costo_huevo:,.1f} COP**. Representa un ahorro directo de **${ahorro_un:,.1f} por huevo**.")


# =============================================================================
# 8. PÃGINA 4 â€” DETALLE COSTOS POR LOTE
# =============================================================================
elif menu == "4. Detalle Costos por Lote":
    st.markdown('<p class="main-title">4. DIAGNÃ“STICO MICRO-OPERATIVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">AuditorÃ­a del perÃ­odo de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)]
    
    # Costos Lote (Aplica la misma lÃ³gica que el consolidado para asegurar el match)
    df_costo_lote = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov_lote = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    c_lote = df_costo_lote.groupby("Lote")["Totales"].sum()
    a_lote = df_aprov_lote.groupby("Lote")["Totales"].sum()
    h_lote = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby("Lote")["Cantidad"].sum()

    df_m = pd.DataFrame({"Costo Base": c_lote, "Aprovechamientos": a_lote, "Huevos FÃ©rtiles": h_lote}).fillna(0)
    df_m["Costo Total"] = df_m["Costo Base"] + df_m["Aprovechamientos"]
    df_m["Costo Unitario"] = df_m["Costo Total"] / df_m["Huevos FÃ©rtiles"]
    df_m = df_m[df_m["Huevos FÃ©rtiles"] > 0].reset_index()

    if df_m.empty:
        st.warning("No hay registros de lotes para el mes seleccionado.")
        st.stop()

    df_m["Lote"] = df_m["Lote"].astype(int).astype(str)
    df_m = df_m.sort_values("Costo Unitario", ascending=False)

    lote_critico = df_m.iloc[0]
    lote_eficiente = df_m.iloc[-1]
    promedio_mes = df_m["Costo Total"].sum() / df_m["Huevos FÃ©rtiles"].sum()

    st.markdown(f"""
    <div class="alert-box">
        ðŸš¨ <b>AuditorÃ­a de Ineficiencia (Causa RaÃ­z):</b><br>
        El <b>Lote {lote_critico['Lote']}</b> es el punto crÃ­tico de la operaciÃ³n este mes, con un costo unitario de <b>${lote_critico['Costo Unitario']:,.1f} COP/HF</b> (muy por encima del promedio de ${promedio_mes:,.1f}).<br>
        <i>Â¿Por quÃ© sucede esto?</i> Usualmente un volumen muy bajo de producciÃ³n provoca que los costos fijos (depreciaciÃ³n gallina/M.O.) se dividan entre muy pocas unidades. <b>Evaluar descarte por curva de postura.</b><br><br>
        âœ… El <b>Lote {lote_eficiente['Lote']}</b> sostiene la rentabilidad con un costo Ã³ptimo de <b>${lote_eficiente['Costo Unitario']:,.1f} COP/HF</b>.
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Lote CrÃ­tico (Mayor Costo)", f"Lote {lote_critico['Lote']}", f"${lote_critico['Costo Unitario']:,.1f}/HF", delta_color="inverse")
    c2.metric("Promedio Ponderado Mes", "General", f"${promedio_mes:,.1f}/HF", delta_color="off")
    c3.metric("Lote MÃ¡s Eficiente", f"Lote {lote_eficiente['Lote']}", f"${lote_eficiente['Costo Unitario']:,.1f}/HF")

    st.markdown("---")
    col_grafico, col_tabla = st.columns([3, 2])

    with col_grafico:
        st.subheader("ðŸ“Š Ranking de Ineficiencia por Lote")
        fig_bar_lote = px.bar(
            df_m, x="Lote", y="Costo Unitario", text_auto=".1f",
            color="Costo Unitario", color_continuous_scale="RdYlGn_r",
            title=f"Costo Unitario Comparativo - {p_actual}"
        )
        fig_bar_lote.add_hline(y=promedio_mes, line_dash="dot", annotation_text=f"Promedio: ${promedio_mes:,.1f}", line_color="black")
        fig_bar_lote.update_traces(textposition="outside")
        fig_bar_lote.update_layout(yaxis_title="Costo Unitario ($ COP)")
        st.plotly_chart(fig_bar_lote, use_container_width=True)

    with col_tabla:
        st.subheader("ðŸ“‹ Matriz de AuditorÃ­a Lotes")
        df_tabla = df_m[["Lote", "Huevos FÃ©rtiles", "Costo Total", "Costo Unitario"]].copy()
        st.dataframe(df_tabla.style.format({
            "Huevos FÃ©rtiles": "{:,.0f} unds", "Costo Total": "${:,.0f}", "Costo Unitario": "${:,.2f}"
        }), use_container_width=True)

# =============================================================================
# 9. PÃGINA 5 â€” DETALLE COSTOS POR LÃNEA
# =============================================================================
elif menu == "5. Detalle Costos por LÃ­nea":
    st.markdown('<p class="main-title">5. EVALUACIÃ“N FINANCIERA POR GENÃ‰TICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)]
    
    df_costo_lin = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov_lin = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    c_lin = df_costo_lin.groupby("linea")["Totales"].sum()
    a_lin = df_aprov_lin.groupby("linea")["Totales"].sum()
    h_lin = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby("linea")["Cantidad"].sum()

    df_gen = pd.DataFrame({"Costo Base": c_lin, "Aprovechamientos": a_lin, "Huevos FÃ©rtiles": h_lin}).fillna(0)
    df_gen["Costo Total ($)"] = df_gen["Costo Base"] + df_gen["Aprovechamientos"]
    df_gen["Costo Unitario ($/HF)"] = df_gen["Costo Total ($)"] / df_gen["Huevos FÃ©rtiles"]
    df_gen = df_gen[df_gen["Huevos FÃ©rtiles"] > 0].reset_index()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("ðŸ“ˆ EvoluciÃ³n HistÃ³rica por Raza")
        df_ts_g = df_raw_graficas[df_raw_graficas["Texto breve de material"] == MATERIAL_HF].groupby(["Periodo", "linea"])["Cantidad"].sum().reset_index()
        fig_ts = px.line(df_ts_g, x="Periodo", y="Cantidad", color="linea", markers=True)
        fig_ts.update_layout(yaxis_range=[df_ts_g["Cantidad"].min()*0.9, df_ts_g["Cantidad"].max()*1.15])
        st.plotly_chart(fig_ts, use_container_width=True)

    with col2:
        st.subheader(f"ðŸ“Š Costo Unitario Promedio por Raza ({p_actual})")
        fig_bar = px.bar(df_gen, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".1f")
        st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(df_gen[["linea", "Huevos FÃ©rtiles", "Costo Total ($)", "Costo Unitario ($/HF)"]].style.format({
        "Costo Total ($)": "${:,.0f}", "Huevos FÃ©rtiles": "{:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"
    }), use_container_width=True)

# =============================================================================
# 10. PÃGINA 6 â€” COSTO KG ALIMENTO
# =============================================================================
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">6. IMPACTO DEL COSTO DE ALIMENTO</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="insight-box">
        ðŸ’¡ <b>Sensibilidad Financiera:</b> El alimento pondera la mayor parte del costo total operativo. Monitorear las fluctuaciones del precio del kilogramo y la conversiÃ³n alimenticia (gramos consumidos por huevo) es el pilar de la rentabilidad avÃ­cola.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("ðŸ“ˆ EvoluciÃ³n Precio Alimento ($/Kg)")
        fig_a = px.line(df_filtrado_graficas, x="Periodo", y="Precio Kg Alimento", markers=True)
        fig_a.update_traces(line_color="#d62728", line_width=3)
        fig_a.update_layout(yaxis_range=[df_filtrado_graficas["Precio Kg Alimento"].min()*0.9, df_filtrado_graficas["Precio Kg Alimento"].max()*1.1])
        st.plotly_chart(fig_a, use_container_width=True)

    with col2:
        st.subheader("ðŸ“ˆ EvoluciÃ³n ConversiÃ³n (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x="Periodo", y="Gramos Alimento/Huevo", markers=True)
        fig_g.update_traces(line_color="#2ca02c", line_width=3)
        fig_g.update_layout(yaxis_range=[df_filtrado_graficas["Gramos Alimento/Huevo"].min()*0.9, df_filtrado_graficas["Gramos Alimento/Huevo"].max()*1.1])
        st.plotly_chart(fig_g, use_container_width=True)

    st.subheader("ðŸ“‹ Matriz Comparativa Interanual (Alimento):")
    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"].copy()
    res_alim = df_alim.groupby(["Anio", "Mes_Num"]).apply(
        lambda x: x["Totales"].sum() / x["Cantidad"].sum() if x["Cantidad"].sum() > 0 else 0
    ).unstack(level=0)

    if 2025 in res_alim.columns and 2026 in res_alim.columns:
        res_df = res_alim[[2025, 2026]].dropna().reset_index()
        res_df["% VAR"] = ((res_df[2026] - res_df[2025]) / res_df[2025]) * 100
        st.dataframe(res_df.style.format({2025: "${:,.1f}", 2026: "${:,.1f}", "% VAR": "{:+.2f}%"}), use_container_width=True)
    else:
        st.info("La tabla comparativa 2025 vs 2026 requiere datos de ambos aÃ±os procesados en la matriz base.")
