from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =============================================================================
# 0. CONFIGURACIÓN DE PÁGINA Y ESTILOS EJECUTIVOS
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
# 1. MOTOR DE DATOS (ETL) PARA INFORME GENERAL MENSUAL
# =============================================================================
RUBRO_APROVECHAMIENTO = "Aprovechamientos (-)"

MAP_RUBROS_BD = {
    "Alimento": "Alimento",
    "Cama - Cascarilla": "Cama / Cascarilla",
    "Droga": "Sanidad (Medicamentos)",
    "Materia prima (calcio)": "Materias Primas (Calcio)",
    "Aseo y desinfección": "Bioseguridad y Aseo",
    "Arriendo": "Arriendo",
    "Depreciacion Const. Y Edif.": "Depreciación Instalaciones",
    "Indirectos": "Costos Indirectos (CIF)",
    "Depreciacion Huevo": "Depreciación Parvada",
    "Mano de Obra": "Mano de Obra",
    "Aprovechamientos (-)": RUBRO_APROVECHAMIENTO
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

@st.cache_data(show_spinner="Procesando datos gerenciales...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if "BD" not in xls.sheet_names:
        st.error("⚠️ Error Crítico: no se encontró la hoja 'BD' en el archivo.")
        st.stop()

    # Cargar base macro (BD)
    df_bd = pd.read_excel(xls, sheet_name="BD")
    df_bd.columns = df_bd.columns.astype(str).str.strip()

    df_bd["Fecha"] = pd.to_datetime(df_bd["Fecha"], errors="coerce")
    df_bd["Anio"] = df_bd["Fecha"].dt.year
    df_bd["Mes_Num"] = df_bd["Fecha"].dt.month
    df_bd["Periodo"] = df_bd["Anio"].astype(str) + "-" + df_bd["Mes_Num"].astype(str).str.zfill(2)

    # Limpieza numérica de rubros en BD
    cols_costos = [
        'Arriendo', 'Cama - Cascarilla', 'Depreciacion Const. Y Edif.', 
        'Depreciacion Huevo', 'Droga', 'Materia prima (calcio)', 
        'Alimento', 'Mano de Obra', 'Aseo y desinfección', 'Indirectos', 
        'Aprovechamientos (-)', 'Costo total', 'Producción HF', 'KG cons Alimento', 'Externos HF'
    ]
    for c in cols_costos:
        if c in df_bd.columns:
            df_bd[c] = pd.to_numeric(df_bd[c], errors="coerce").fillna(0)

    # Construir dataframe resumen mensual
    df_res = pd.DataFrame()
    df_res["Periodo"] = df_bd["Periodo"]
    df_res["Anio"] = df_bd["Anio"]
    df_res["Mes_Num"] = df_bd["Mes_Num"]

    for col_orig, col_dest in MAP_RUBROS_BD.items():
        if col_orig in df_bd.columns:
            df_res[col_dest] = df_bd[col_orig]

    df_res["Costo Total"] = df_bd["Costo total"]
    df_res["Huevos Fértiles"] = df_bd["Producción HF"]
    df_res["Costo Huevo Fértil"] = safe_div(df_res["Costo Total"], df_res["Huevos Fértiles"])

    df_res["Consumo Alimento Kg"] = df_bd["KG cons Alimento"] if "KG cons Alimento" in df_bd.columns else 0
    df_res["Precio Kg Alimento"] = safe_div(df_res.get("Alimento", 0), df_res["Consumo Alimento Kg"])
    df_res["Gramos Alimento/Huevo"] = safe_div(df_res["Consumo Alimento Kg"] * 1000, df_res["Huevos Fértiles"])

    df_res = df_res.sort_values("Periodo").reset_index(drop=True)

    # Cargar base micro (BD CTO LINEA y BD PN HF RAZA si existen)
    df_cto = pd.read_excel(xls, sheet_name="BD CTO LINEA") if "BD CTO LINEA" in xls.sheet_names else pd.DataFrame()
    if not df_cto.empty:
        df_cto.columns = df_cto.columns.astype(str).str.strip()
        df_cto["fecha"] = pd.to_datetime(df_cto["fecha"], errors="coerce")
        df_cto["Anio"] = df_cto["fecha"].dt.year
        df_cto["Mes_Num"] = df_cto["fecha"].dt.month
        df_cto["Periodo"] = df_cto["Anio"].astype(str) + "-" + df_cto["Mes_Num"].astype(str).str.zfill(2)

    df_raza = pd.read_excel(xls, sheet_name="BD PN HF RAZA") if "BD PN HF RAZA" in xls.sheet_names else pd.DataFrame()

    rubros = [r for r in list(MAP_RUBROS_BD.values()) if r in df_res.columns]

    return df_res, rubros, df_bd, df_cto, df_raza

# =============================================================================
# 2. CARGA DE DATOS EN LA APLICACIÓN
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")
uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df, rubros_items, df_bd, df_cto, df_raza = load_and_process_data(uploaded_file)
else:
    archivo_default = find_default_excel()
    if archivo_default is None:
        st.warning("⚠️ Se requiere el archivo 'INFORME GENERAL MENSUAL.xlsx' para inicializar el sistema.")
        st.stop()
    df, rubros_items, df_bd, df_cto, df_raza = load_and_process_data(str(archivo_default))

if df.empty:
    st.error("El archivo no contiene datos procesables.")
    st.stop()

# =============================================================================
# 3. BARRA LATERAL — FILTROS DE TIEMPO
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
df_bd_graficas = df_bd[(df_bd["Periodo"] >= rango_inicio) & (df_bd["Periodo"] <= rango_fin)]

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. Producción",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil",
        "4. Detalle Costos por Lote",
        "5. Detalle Costos por Línea",
        "6. Costo Kg Alimento",
    ],
)

# =============================================================================
# 4. INTELIGENCIA DE NEGOCIO (texto automático)
# =============================================================================
def generar_diagnostico_costos(df_impactos, var_total, p_actual, p_base):
    if df_impactos.empty:
        return ""
    if var_total <= 0:
        rubro_exito = df_impactos.iloc[0]["Rubro"]
        return (
            f'<div class="insight-box">✅ <b>Eficiencia Operativa Alcanzada:</b> El costo unitario '
            f'presenta una reducción de <b>${abs(var_total):,.2f} COP/Huevo</b> respecto a {p_base}.<br>'
            f'<i>Causa Raíz:</i> El rubro que más impulsó este ahorro fue <b>{rubro_exito}</b>. '
            f'Mantener controles actuales.</div>'
        )
    rubro_critico = df_impactos.iloc[-1]["Rubro"]
    impacto_critico = df_impactos.iloc[-1]["Impacto ($/HF)"]
    pct_explicado = (impacto_critico / var_total) * 100 if var_total else 0
    recomendacion = RECOMENDACIONES_RUBRO.get(
        rubro_critico,
        f"⚠️ <b>Factor Crítico - {rubro_critico}:</b> Explicó el <b>{pct_explicado:.1f}%</b> de la "
        f"desviación desfavorable. Requiere revisión inmediata.",
    )
    return f'<div class="alert-box">🚨 <b>Alerta de Desviación Financiera (+${var_total:,.2f} COP/Huevo):</b><br>{recomendacion}</div>'

# =============================================================================
# 5. PÁGINA 1 — PRODUCCIÓN
# =============================================================================
if menu == "1. Producción":
    st.markdown('<p class="main-title">1. PRODUCCIÓN CONSOLIDADA DE HUEVO FÉRTIL (PROPIOS Y EXTERNOS)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    tot_propios = df_bd_graficas["Producción HF"].sum()
    tot_externos = df_bd_graficas["Externos HF"].sum() if "Externos HF" in df_bd_graficas.columns else 0
    tot_global = tot_propios + tot_externos
    n_meses_rango = df_filtrado_graficas["Periodo"].nunique()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Volumen Propios", f"{tot_propios:,.0f} HF")
    c2.metric("Volumen Externos", f"{tot_externos:,.0f} HF")
    c3.metric("Volumen Total Período", f"{tot_global:,.0f} HF")
    c4.metric("Promedio Mensual Total", f"{tot_global / max(n_meses_rango, 1):,.0f} HF/mes")

    st.markdown("---")
    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("📈 Evolución Histórica de Producción")
        df_ts = df_bd_graficas.groupby("Periodo")[["Producción HF", "Externos HF"]].sum().reset_index()
        fig_ts = px.line(df_ts, x="Periodo", y=["Producción HF", "Externos HF"], markers=True, title="Comportamiento del Volumen (Propios vs Externos)")
        fig_ts.update_layout(yaxis_title="Unidades de HF", legend_title="Origen")
        st.plotly_chart(fig_ts, use_container_width=True)

    with col_r:
        st.subheader("🧬 Participación por Origen")
        df_pie_data = pd.DataFrame({
            "Origen": ["Propios", "Externos"],
            "Volumen": [tot_propios, tot_externos]
        })
        fig_pie = px.pie(df_pie_data, values="Volumen", names="Origen", hole=0.4, title="Mix de Abastecimiento")
        st.plotly_chart(fig_pie, use_container_width=True)

# =============================================================================
# 6. PÁGINA 2 — PRODUCCIÓN MES A MES POR LÍNEA
# =============================================================================
elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÓN MENSUAL POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    if not df_raza.empty:
        df_raza.columns = df_raza.columns.astype(str).str.strip()
        df_raza["FECHA"] = pd.to_datetime(df_raza["FECHA"], errors="coerce")
        df_raza["Periodo"] = df_raza["FECHA"].dt.strftime("%Y-%m")
        df_raza_f = df_raza[(df_raza["Periodo"] >= rango_inicio) & (df_raza["Periodo"] <= rango_fin)]
        
        piv_raza = df_raza_f.groupby(["Periodo", "RAZA"])[["PROPIOS", "EXTERNOS"]].sum().sum(axis=1).unstack(fill_value=0)
        if not piv_raza.empty:
            fig_bar = px.bar(piv_raza.reset_index(), x="Periodo", y=piv_raza.columns, title="Producción Mensual por Raza", text_auto=".2s")
            st.plotly_chart(fig_bar, use_container_width=True)
            st.dataframe(piv_raza.style.format("{:,.0f}"), use_container_width=True)
        else:
            st.info("Sin datos suficientes en BD PN HF RAZA para el rango seleccionado.")
    elif not df_cto.empty and "linea" in df_cto.columns:
        piv_linea = df_cto[(df_cto["Periodo"] >= rango_inicio) & (df_cto["Periodo"] <= rango_fin)].groupby(["Periodo", "linea"])["Huevos fértiles"].sum().unstack(fill_value=0)
        fig_bar = px.bar(piv_linea.reset_index(), x="Periodo", y=piv_linea.columns, title="Producción Mensual por Genética", text_auto=".2s")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(piv_linea.style.format("{:,.0f}"), use_container_width=True)
    else:
        st.info("No hay desglose por línea disponible en las fuentes cargadas.")

# =============================================================================
# 7. PÁGINA 3 — COSTO HUEVO FÉRTIL
# =============================================================================
elif menu == "3. Costo Huevo Fértil":
    st.markdown('<p class="main-title">3. ANÁLISIS INTEGRAL: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)

    if p_base == p_actual:
        st.warning("⚠️ Selecciona un Rango de Meses o el Modo Comparativo con dos períodos distintos.")
        st.stop()

    fila_b = df[df["Periodo"] == p_base]
    fila_a = df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty:
        st.error("No hay datos suficientes en los períodos seleccionados.")
        st.stop()

    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    var_tot = df_a["Costo Huevo Fértil"] - df_b["Costo Huevo Fértil"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {p_base}", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo {p_actual}", f"${df_a['Costo Huevo Fértil']:,.2f}")
    delta_pct = (var_tot / df_b["Costo Huevo Fértil"]) * 100 if df_b["Costo Huevo Fértil"] else 0
    c3.metric("Desviación ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")

    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos Fértiles"], df_a["Huevos Fértiles"]
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({"Rubro": r, "Costo Unit Base ($)": ub, "Costo Unit Actual ($)": ua, "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)

    if not df_vif.empty:
        st.markdown(generar_diagnostico_costos(df_vif.sort_values("Impacto ($/HF)"), var_tot, p_actual, p_base), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Evolución Histórica de Costo ($)")
        fig_line_c = px.line(df_filtrado_graficas, x="Periodo", y="Costo Huevo Fértil", markers=True)
        fig_line_c.update_traces(line_color="#1E3A8A", line_width=3)
        st.plotly_chart(fig_line_c, use_container_width=True)
    with col2:
        st.subheader("📊 Tornado de Desviación por Rubro")
        fig_tor = px.bar(df_vif.sort_values("Impacto ($/HF)"), y="Rubro", x="Impacto ($/HF)", orientation="h", color="Impacto ($/HF)", color_continuous_scale="Reds" if var_tot > 0 else "Greens")
        st.plotly_chart(fig_tor, use_container_width=True)

# =============================================================================
# 8. PÁGINA 4 — DETALLE COSTOS POR LOTE
# =============================================================================
elif menu == "4. Detalle Costos por Lote":
    st.markdown('<p class="main-title">4. DIAGNÓSTICO MICRO-OPERATIVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    if df_cto.empty:
        st.warning("No se encontró la base de detalle por lote ('BD CTO LINEA') en el archivo.")
        st.stop()

    df_l = df_cto[df_cto["Periodo"] == p_actual].copy()
    if df_l.empty:
        st.warning("No hay registros de lotes para el mes seleccionado.")
        st.stop()

    cols_sum = ["ALIMENTO", "CAMA", "DROGA", "MATERIA PRIMA", "ELEMENTOS DE ASEO Y DESINFECCION", "ARRIENDO", "DEPRECIACIÓN CONST Y EDIF", "INDIRECTOS", "DEPRECIACIÓN GALLINA", "MANO DE OBRA", "SUBPRODUCTOS"]
    existing_cols = [c for c in cols_sum if c in df_l.columns]
    
    df_l["Costo Directo"] = df_l[[c for c in existing_cols if c != "SUBPRODUCTOS"]].sum(axis=1)
    df_l["Aprovechamiento"] = df_l["SUBPRODUCTOS"] if "SUBPRODUCTOS" in df_l.columns else 0
    df_l["Costo Total Lote"] = df_l["Costo Directo"] + df_l["Aprovechamiento"]

    df_m = df_l.groupby("lote").agg({
        "Huevos fértiles": "sum",
        "Costo Total Lote": "sum",
        "Costo Directo": "sum",
        "Aprovechamiento": "sum"
    }).reset_index()

    df_m = df_m[df_m["Huevos fértiles"] > 0].copy()
    df_m["Costo Unitario"] = df_m["Costo Total Lote"] / df_m["Huevos fértiles"]
    df_m["lote"] = df_m["lote"].astype(str)
    df_m = df_m.sort_values("Costo Unitario", ascending=False)

    promedio_mes = df_m["Costo Total Lote"].sum() / df_m["Huevos fértiles"].sum()

    st.subheader("📊 Ranking de Costos por Lote")
    fig_bar_lote = px.bar(df_m, x="lote", y="Costo Unitario", text_auto=".1f", color="Costo Unitario", color_continuous_scale="RdYlGn_r")
    fig_bar_lote.add_hline(y=promedio_mes, line_dash="dot", annotation_text=f"Promedio: ${promedio_mes:,.1f}", line_color="black")
    st.plotly_chart(fig_bar_lote, use_container_width=True)
    st.dataframe(df_m.style.format({"Huevos fértiles": "{:,.0f}", "Costo Total Lote": "${:,.0f}", "Costo Unitario": "${:,.2f}"}), use_container_width=True)

# =============================================================================
# 9. PÁGINA 5 — DETALLE COSTOS POR LÍNEA
# =============================================================================
elif menu == "5. Detalle Costos por Línea":
    st.markdown('<p class="main-title">5. EVALUACIÓN FINANCIERA POR GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis del período Actual ({p_actual})</p>', unsafe_allow_html=True)

    if df_cto.empty:
        st.warning("No se encontró la base de detalle por línea ('BD CTO LINEA') en el archivo.")
        st.stop()

    df_l = df_cto[df_cto["Periodo"] == p_actual].copy()
    if df_l.empty:
        st.warning("No hay datos de línea para el período actual.")
        st.stop()

    cols_sum = ["ALIMENTO", "CAMA", "DROGA", "MATERIA PRIMA", "ELEMENTOS DE ASEO Y DESINFECCION", "ARRIENDO", "DEPRECIACIÓN CONST Y EDIF", "INDIRECTOS", "DEPRECIACIÓN GALLINA", "MANO DE OBRA", "SUBPRODUCTOS"]
    existing_cols = [c for c in cols_sum if c in df_l.columns]
    
    df_l["Costo Directo"] = df_l[[c for c in existing_cols if c != "SUBPRODUCTOS"]].sum(axis=1)
    df_l["Aprovechamiento"] = df_l["SUBPRODUCTOS"] if "SUBPRODUCTOS" in df_l.columns else 0
    df_l["Costo Total Linea"] = df_l["Costo Directo"] + df_l["Aprovechamiento"]

    df_gen = df_l.groupby("linea").agg({
        "Huevos fértiles": "sum",
        "Costo Total Linea": "sum"
    }).reset_index()

    df_gen = df_gen[df_gen["Huevos fértiles"] > 0].copy()
    df_gen["Costo Unitario ($/HF)"] = df_gen["Costo Total Linea"] / df_gen["Huevos fértiles"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Evolución Histórica por Raza")
        df_ts_g = df_cto[(df_cto["Periodo"] >= rango_inicio) & (df_cto["Periodo"] <= rango_fin)].groupby(["Periodo", "linea"])["Huevos fértiles"].sum().reset_index()
        fig_ts = px.line(df_ts_g, x="Periodo", y="Huevos fértiles", color="linea", markers=True)
        st.plotly_chart(fig_ts, use_container_width=True)
    with col2:
        st.subheader(f"📊 Costo Unitario Promedio ({p_actual})")
        fig_bar = px.bar(df_gen, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".1f")
        st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(df_gen.style.format({"Costo Total Linea": "${:,.0f}", "Huevos fértiles": "{:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"}), use_container_width=True)

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
        fig_a.update_traces(line_color="#d62728", line_width=3)
        st.plotly_chart(fig_a, use_container_width=True)
    with col2:
        st.subheader("📈 Evolución Conversión (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x="Periodo", y="Gramos Alimento/Huevo", markers=True)
        fig_g.update_traces(line_color="#2ca02c", line_width=3)
        st.plotly_chart(fig_g, use_container_width=True)
