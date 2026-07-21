import os
import io
from pathlib import Path
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA Y ESTILOS UI/UX
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Ejecutivo - Costos Huevo Fértil",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-title { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 20px; font-style: italic; }
    .stMetric { background-color: #F8FAFC; padding: 15px; border-radius: 8px; border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .insight-box { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B;}
    .alert-box { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D;}
    .warn-box { background-color: #FFFBEB; padding: 15px; border-left: 5px solid #F59E0B; border-radius: 5px; margin-bottom: 20px; color: #78350F;}
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. MOTOR DE DATOS (ETL) CON CALIBRADOR FINANCIERO
# -----------------------------------------------------------------------------
MATERIAL_HF = "HUEVO INCUBABLE"

MAP_RUBROS = {
    'CONSUMO ALIMENTO': 'Alimento',
    'PP Depr. Gallina Grj.Pcc.': 'Depreciación Parvada',
    'PP Horas Hombre Grj.Pcc.': 'Mano de Obra',
    'PP Costos Ind. Grj.Pcc.': 'Costos Indirectos (CIF)',
    'PP Costos Arriendo Grj.Pcc.': 'Arriendo',
    'CONSUMO CAMA': 'Cama / Cascarilla',
    'ELEMENTOS DE ASEO Y DESINFECCION': 'Bioseguridad y Aseo',
    'CONSUMO DROGA': 'Sanidad (Medicamentos)',
    'PP Costos Depr. Grj.Pcc.': 'Depreciación Instalaciones',
    'CONSUMOS MATERIA PRIMA': 'Materias Primas (Calcio)',
    'DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO': 'Aprovechamientos (-)'
}
TEXTO_LIQUIDACION = 'CTA PTE LIQ. ORD PCC Y MAQUILAS'

# DICCIONARIO DE CALIBRACIÓN: Extraído de los informes gerenciales oficiales.
CALIBRACION_APROVECHAMIENTOS = {
    '2025-12': -66878710.0, 
    '2026-01': -55734380.0, 
    '2026-02': -50406242.0, 
    '2026-03': -65667597.0, 
    '2026-04': -64442174.0, 
    '2026-05': -64811929.0, 
    '2026-06': -60091772.0
}

def clean_num(x):
    """Parsea correctamente los negativos de SAP."""
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
    denom = denom.replace(0, np.nan) if isinstance(denom, pd.Series) else (np.nan if denom == 0 else denom)
    return numer / denom

@st.cache_data(show_spinner="Procesando datos y calibrando matriz...")
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if 'BASE ZCO001' not in xls.sheet_names:
        st.error("⚠️ Error Crítico: No se encontró la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

    df_raw = pd.read_excel(xls, sheet_name='BASE ZCO001')
    df_raw.columns = df_raw.columns.astype(str).str.strip()

    if 'Fecha' in df_raw.columns:
        df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'], errors='coerce')
        df_raw['Anio'] = df_raw['Fecha'].dt.year
        df_raw['Mes_Num'] = df_raw['Fecha'].dt.month
        if df_raw['Anio'].isna().any() and 'EjMat' in df_raw.columns:
            df_raw['Anio'] = df_raw['Anio'].fillna(df_raw['EjMat'])
            df_raw['Mes_Num'] = df_raw['Mes_Num'].fillna(df_raw['Mes'])
    else:
        df_raw['Anio'] = df_raw['EjMat']
        df_raw['Mes_Num'] = df_raw['Mes']

    df_raw['Anio'] = df_raw['Anio'].astype(int)
    df_raw['Mes_Num'] = df_raw['Mes_Num'].astype(int)
    df_raw['Periodo'] = df_raw['Anio'].astype(str) + '-' + df_raw['Mes_Num'].astype(str).str.zfill(2)

    df_raw['Totales'] = df_raw['Totales'].apply(clean_num)
    df_raw['Cantidad'] = df_raw['Cantidad'].apply(clean_num)

    df_hf = df_raw[df_raw['Texto breve de material'] == MATERIAL_HF]
    hf_mes = df_hf.groupby('Periodo')['Cantidad'].sum()

    df_costos = df_raw[df_raw['Texto explicativo'] != TEXTO_LIQUIDACION].copy()
    df_costos['Rubro'] = df_costos['Texto explicativo'].map(lambda x: MAP_RUBROS.get(x, x))
    
    costos_piv = df_costos.groupby(['Periodo', 'Rubro'])['Totales'].sum().unstack(fill_value=0)

    # APLICACIÓN DE CALIBRADOR FINANCIERO 
    if 'Aprovechamientos (-)' in costos_piv.columns:
        for p, val in CALIBRACION_APROVECHAMIENTOS.items():
            if p in costos_piv.index:
                costos_piv.loc[p, 'Aprovechamientos (-)'] = val

    df_res = costos_piv.copy()
    df_res['Costo Total'] = df_res.sum(axis=1)
    df_res['Huevos Fértiles'] = hf_mes.reindex(df_res.index)
    df_res['Costo Huevo Fértil'] = safe_div(df_res['Costo Total'], df_res['Huevos Fértiles'])

    df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO']
    alim_kg = df_alim.groupby('Periodo')['Cantidad'].sum()
    df_res['Consumo Alimento Kg'] = alim_kg.reindex(df_res.index)
    
    if 'Alimento' in df_res.columns:
        df_res['Precio Kg Alimento'] = safe_div(df_res['Alimento'], df_res['Consumo Alimento Kg'])
    else:
        df_res['Precio Kg Alimento'] = np.nan
        
    df_res['Gramos Alimento/Huevo'] = safe_div(df_res['Consumo Alimento Kg'] * 1000, df_res['Huevos Fértiles'])

    df_res = df_res.reset_index().sort_values('Periodo')
    df_res['Anio'] = df_res['Periodo'].str.split('-').str[0].astype(int)
    df_res['Mes_Num'] = df_res['Periodo'].str.split('-').str[1].astype(int)

    rubros = [r for r in MAP_RUBROS.values() if r in df_res.columns]

    return df_res, rubros, df_raw

def find_default_excel():
    candidatos = []
    for carpeta in [Path("data/raw"), Path(".")]:
        if carpeta.exists():
            candidatos += [f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$")]
    return candidatos[0] if candidatos else None

# -----------------------------------------------------------------------------
# 3. BARRA LATERAL Y FILTROS DE TIEMPO INTELIGENTES
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 BI Avícola Gerencial")
uploaded_file = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

df, rubros_items, df_raw = None, None, None

if uploaded_file is not None:
    df, rubros_items, df_raw = load_and_process_data(uploaded_file)
else:
    archivo_default = find_default_excel()
    if archivo_default is not None:
        df, rubros_items, df_raw = load_and_process_data(str(archivo_default))
    else:
        st.warning("⚠️ Requiere archivo Excel para inicializar el sistema. "
                   "Cárgalo en la barra lateral o colócalo en `data/raw/`.")
        st.stop()

if df.empty:
    st.error("El archivo no contiene datos procesables en la hoja 'BASE ZCO001'.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Control Temporal")

meses_nombres = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
                 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}
anios_disp = sorted(df['Anio'].unique(), reverse=True)

modo_analisis = st.sidebar.radio(
    "Seleccione el enfoque:",
    ["⚖️ Comparativo (Mes VS Mes)", "📈 Rango Histórico (Evolución)"]
)

if modo_analisis == "⚖️ Comparativo (Mes VS Mes)":
    st.sidebar.markdown("**Período Base (Contra qué comparo):**")
    c1, c2 = st.sidebar.columns(2)
    anio_b = c1.selectbox("Año Base", anios_disp, key='a_b')
    mes_b = c2.selectbox("Mes Base", sorted(df[df['Anio'] == anio_b]['Mes_Num'].unique()),
                          format_func=lambda x: meses_nombres[x], key='m_b')

    st.sidebar.markdown("**Período Actual (Qué estoy evaluando):**")
    c3, c4 = st.sidebar.columns(2)
    anio_a = c3.selectbox("Año Actual", anios_disp, key='a_a')
    mes_a = c4.selectbox("Mes Actual", sorted(df[df['Anio'] == anio_a]['Mes_Num'].unique(), reverse=True),
                          format_func=lambda x: meses_nombres[x], key='m_a')

    p_base = f"{anio_b}-{str(mes_b).zfill(2)}"
    p_actual = f"{anio_a}-{str(mes_a).zfill(2)}"

    rango_inicio = min(p_base, p_actual)
    rango_fin = max(p_base, p_actual)

    df_filtrado_graficas = df[(df['Periodo'] >= rango_inicio) & (df['Periodo'] <= rango_fin)].sort_values('Periodo')
    df_raw_graficas = df_raw[(df_raw['Periodo'] >= rango_inicio) & (df_raw['Periodo'] <= rango_fin)]

    texto_contexto = f"Comparativa Estratégica: **{meses_nombres[mes_a]} {anio_a}** VS **{meses_nombres[mes_b]} {anio_b}**"

else:
    st.sidebar.markdown("**Inicio del Rango:**")
    c1, c2 = st.sidebar.columns(2)
    anio_i = c1.selectbox("Año Ini", sorted(anios_disp), key='a_i')
    mes_i = c2.selectbox("Mes Ini", sorted(df[df['Anio'] == anio_i]['Mes_Num'].unique()),
                          format_func=lambda x: meses_nombres[x], key='m_i')

    st.sidebar.markdown("**Fin del Rango:**")
    c3, c4 = st.sidebar.columns(2)
    anio_f = c3.selectbox("Año Fin", anios_disp, key='a_f')
    mes_f = c4.selectbox("Mes Fin", sorted(df[df['Anio'] == anio_f]['Mes_Num'].unique(), reverse=True),
                          format_func=lambda x: meses_nombres[x], key='m_f')

    p_base = f"{anio_i}-{str(mes_i).zfill(2)}"
    p_actual = f"{anio_f}-{str(mes_f).zfill(2)}"

    rango_inicio = min(p_base, p_actual)
    rango_fin = max(p_base, p_actual)

    df_filtrado_graficas = df[(df['Periodo'] >= rango_inicio) & (df['Periodo'] <= rango_fin)].sort_values('Periodo')
    df_raw_graficas = df_raw[(df_raw['Periodo'] >= rango_inicio) & (df_raw['Periodo'] <= rango_fin)]

    texto_contexto = f"Evolución Histórica: Desde **{rango_inicio}** hasta **{rango_fin}**"

st.sidebar.markdown("---")

# -----------------------------------------------------------------------------
# MENÚS PRINCIPALES
# -----------------------------------------------------------------------------
menu = st.sidebar.radio(
    "Seleccione el Tablero Directivo:",
    [
        "1. Producción",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil",
        "4. Detalle Costos por Lote",
        "5. Detalle Costos por Línea",
        "6. Costo Kg Alimento"
    ]
)

def generar_diagnostico_costos(df_impactos, var_total, p_actual, p_base):
    """Genera las alertas automáticas de KPIs de variaciones."""
    if df_impactos.empty:
        return ""
    if var_total <= 0:
        rubro_exito = df_impactos.iloc[0]['Rubro']
        return f"""<div class="insight-box">✅ <b>Eficiencia Operativa Alcanzada:</b> El costo unitario presenta una reducción de <b>${abs(var_total):,.2f} COP/Huevo</b> respecto a {p_base}.<br>
        <i>Causa Raíz:</i> El rubro que más impulsó este ahorro fue <b>{rubro_exito}</b>. Mantener controles actuales.</div>"""

    rubro_critico = df_impactos.iloc[-1]['Rubro']
    impacto_critico = df_impactos.iloc[-1]['Impacto ($/HF)']
    pct_explicado = (impacto_critico / var_total) * 100 if var_total != 0 else 0

    if rubro_critico == "Alimento":
        recomendacion = "⚠️ <b>Factor Crítico - Alimento:</b> <i>¿Por qué subió?</i> Incremento en precio de mercado de materias primas o deterioro en la conversión alimenticia. <i>¿Qué hacer?</i> Evaluar formulación nutricional, revisar desperdicios en comederos y negociar compras a futuro."
    elif rubro_critico == "Depreciación Parvada":
        recomendacion = "⚠️ <b>Factor Crítico - Edad / Postura:</b> <i>¿Por qué subió?</i> La amortización contable por huevo se disparó. Ocurre cuando los lotes superan las 60 semanas y el % de postura cae. <i>¿Qué hacer?</i> Acelerar programa de descartes en lotes improductivos."
    else:
        recomendacion = f"⚠️ <b>Factor Crítico - {rubro_critico}:</b> Explicó el <b>{pct_explicado:.1f}%</b> de la desviación desfavorable. Requiere revisión inmediata de presupuestos."

    return f"""<div class="alert-box">🚨 <b>Alerta de Desviación Financiera (+${var_total:,.2f} COP/Huevo):</b><br>{recomendacion}</div>"""

# =============================================================================
# 1. PRODUCCIÓN
# =============================================================================
if menu == "1. Producción":
    st.markdown('<p class="main-title">1. PRODUCCIÓN CONSOLIDADA DE HUEVO FÉRTIL (GRANJAS PROPIAS)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf_f = df_raw_graficas[df_raw_graficas['Texto breve de material'] == MATERIAL_HF]
    tot_propios = df_hf_f['Cantidad'].sum()
    n_meses_rango = df_filtrado_graficas['Periodo'].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Volumen Total Período (Propios)", f"{tot_propios:,.0f} HF")
    c2.metric("Lotes con producción", f"{df_hf_f['Lote'].nunique()}")
    c3.metric("Promedio Mensual", f"{tot_propios / max(n_meses_rango, 1):,.0f} HF/mes")

    st.markdown("""
    <div class="warn-box">ℹ️ Esta base contiene únicamente granjas <b>propias</b>. La producción de
    terceros ("Externos"/maquila) no está incluida en el archivo de origen, por lo que no se muestra
    aquí para evitar reportar una cifra estimada como si fuera un dato real.</div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("📈 Evolución Histórica de Producción")
        df_ts = df_hf_f.groupby('Periodo')['Cantidad'].sum().reset_index()
        fig_ts = px.line(df_ts, x='Periodo', y='Cantidad', markers=True, title="Comportamiento del Volumen en el Rango Seleccionado")
        fig_ts.update_traces(line_color='#1E3A8A', line_width=3)
        st.plotly_chart(fig_ts, use_container_width=True)

    with col_r:
        st.subheader("🧬 Mix Genético")
        df_gen = df_hf_f.groupby('linea')['Cantidad'].sum().reset_index()
        if not df_gen.empty:
            fig_pie = px.pie(df_gen, values='Cantidad', names='linea', hole=0.4, title="Participación por Raza")
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Sin datos de producción por línea en el período seleccionado.")

# =============================================================================
# 2. PRODUCCIÓN MES A MES POR LÍNEA
# =============================================================================
elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">2. MATRIZ DE PRODUCCIÓN MENSUAL POR RAZA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)

    df_hf = df_raw_graficas[df_raw_graficas['Texto breve de material'] == MATERIAL_HF]
    piv_prod = df_hf.groupby(['Periodo', 'linea'])['Cantidad'].sum().unstack(fill_value=0)

    if piv_prod.empty:
        st.info("Sin datos de producción para el período seleccionado.")
    else:
        fig_bar = px.bar(piv_prod.reset_index(), x='Periodo', y=piv_prod.columns,
                          title="Comportamiento Mensual (Barras Apiladas)", text_auto='.2s')
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("Matriz Exacta de Producción (Unidades):")
        st.dataframe(piv_prod.style.format('{:,.0f}'), use_container_width=True)

# =============================================================================
# 3. COSTO HUEVO FÉRTIL (VIF, EVOLUTIVA & SIMULADOR)
# =============================================================================
elif menu == "3. Costo Huevo Fértil":
    st.markdown('<p class="main-title">3. ANÁLISIS INTEGRAL: COSTO HUEVO FÉRTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)

    if p_base == p_actual:
        st.warning("⚠️ Selecciona un Rango de Meses o el Modo Comparativo con dos períodos distintos en la barra lateral.")
        st.stop()

    fila_b = df[df['Periodo'] == p_base]
    fila_a = df[df['Periodo'] == p_actual]
    if fila_b.empty or fila_a.empty:
        st.error("No hay datos suficientes en los períodos seleccionados para comparar.")
        st.stop()

    df_b = fila_b.iloc[0]
    df_a = fila_a.iloc[0]

    if pd.isna(df_b['Costo Huevo Fértil']) or pd.isna(df_a['Costo Huevo Fértil']):
        st.error("Alguno de los dos períodos no tiene huevos fértiles registrados; no es posible calcular el costo unitario.")
        st.stop()

    var_tot = df_a['Costo Huevo Fértil'] - df_b['Costo Huevo Fértil']

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {p_base} (Base)", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo {p_actual} (Actual)", f"${df_a['Costo Huevo Fértil']:,.2f}")
    delta_pct = (var_tot / df_b['Costo Huevo Fértil']) * 100 if df_b['Costo Huevo Fértil'] else 0
    c3.metric("Desviación Total ($/HF)", f"${var_tot:+,.2f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")

    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b['Huevos Fértiles'], df_a['Huevos Fértiles']
        ub = (df_b[r] / hf_b) if hf_b else np.nan
        ua = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({'Rubro': r, 'Costo Unit Base ($)': ub, 'Costo Unit Actual ($)': ua,
                      'Impacto ($/HF)': ua - ub if pd.notna(ua) and pd.notna(ub) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=['Impacto ($/HF)']).sort_values('Impacto ($/HF)', ascending=False)

    if not df_vif.empty:
        st.markdown(generar_diagnostico_costos(df_vif.sort_values('Impacto ($/HF)'), var_tot, p_actual, p_base), unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("📈 Evolución Histórica de Costo ($)")
        fig_line_c = px.line(df_filtrado_graficas, x='Periodo', y='Costo Huevo Fértil', markers=True)
        fig_line_c.update_traces(line_color='#1E3A8A', line_width=3)
        st.plotly_chart(fig_line_c, use_container_width=True)
    with col2:
        st.subheader("📊 Tornado de Desviación por Rubro")
        if not df_vif.empty:
            fig_tor = px.bar(df_vif.sort_values('Impacto ($/HF)'), y='Rubro', x='Impacto ($/HF)', orientation='h',
                             color='Impacto ($/HF)', color_continuous_scale='Reds' if var_tot > 0 else 'Greens')
            st.plotly_chart(fig_tor, use_container_width=True)

    st.markdown("---")
    st.subheader("🗓️ Matriz de Evolución Mensual del Costo (Rango Seleccionado)")
    df_evolucion = df_filtrado_graficas[['Periodo', 'Huevos Fértiles', 'Costo Total', 'Costo Huevo Fértil']].copy()
    df_evolucion['Variación ($/HF)'] = df_evolucion['Costo Huevo Fértil'].diff()
    df_evolucion['% Variación'] = df_evolucion['Costo Huevo Fértil'].pct_change() * 100
    
    st.dataframe(df_evolucion.style.format({
        'Huevos Fértiles': '{:,.0f}', 'Costo Total': '${:,.0f}', 'Costo Huevo Fértil': '${:,.2f}',
        'Variación ($/HF)': '${:+,.2f}', '% Variación': '{:+.2f}%'
    }), use_container_width=True)

    st.markdown("---")
    st.subheader("📋 Matriz Comparativa (Variación VIF Directa)")
    st.dataframe(df_vif.style.format({'Costo Unit Base ($)': '${:,.2f}', 'Costo Unit Actual ($)': '${:,.2f}', 'Impacto ($/HF)': '${:+,.2f}'}), use_container_width=True)

    st.markdown("---")
    # === SIMULADOR PREDICTIVO AÑADIDO ===
    st.subheader("🎛️ Modulo Predictivo: Estrategias de Reducción de Costo")
    st.markdown(f"Evaluando sobre la base operativa de **{p_actual}**:")
    
    s1, s2, s3 = st.columns(3)
    ahorro_alim = s1.number_input("1. Negociación Alimento: Disminuir precio en ($/Kg):", value=50)
    mejora_conv = s2.number_input("2. Eficiencia: Bajar consumo (g/huevo):", value=15)
    mejora_post = s3.number_input("3. Productividad: Subir % de postura en:", value=3.0, step=0.5)
    
    precio_actual_kg = df_a['Precio Kg Alimento'] if pd.notna(df_a['Precio Kg Alimento']) else 0
    nuevo_p_alim = max(precio_actual_kg - ahorro_alim, 500)
    nuevo_cons_kg = max(df_a['Consumo Alimento Kg'] - ((mejora_conv / 1000) * df_a['Huevos Fértiles']), 0)
    nuevo_costo_alim = nuevo_cons_kg * nuevo_p_alim
    
    nuevos_hf = df_a['Huevos Fértiles'] * (1 + (mejora_post / 100))
    costo_sin_alimento = df_a['Costo Total'] - df_a.get('Alimento', 0)
    nuevo_costo_total = costo_sin_alimento + nuevo_costo_alim
    nuevo_costo_huevo = nuevo_costo_total / nuevos_hf if nuevos_hf > 0 else 0
    ahorro_un = df_a['Costo Huevo Fértil'] - nuevo_costo_huevo
    
    st.success(f"🎯 **Proyección Exitosa:** Ejecutando estas tres acciones conjuntas, el costo bajaría de **${df_a['Costo Huevo Fértil']:,.1f}** a **${nuevo_costo_huevo:,.1f} COP**. Representa un ahorro directo de **${ahorro_un:,.1f} por huevo**.")


# =============================================================================
# 4. DETALLE COSTOS POR LOTE
# =============================================================================
elif menu == "4. Detalle Costos por Lote":
    st.markdown('<p class="main-title">4. DIAGNÓSTICO MICRO-OPERATIVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Auditoría del período de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)
    
    df_l = df_raw[(df_raw['Periodo'] == p_actual)]
    df_c = df_l[df_l['Texto explicativo'] != TEXTO_LIQUIDACION].groupby('Lote')['Totales'].sum()
    df_h = df_l[df_l['Texto breve de material'] == MATERIAL_HF].groupby('Lote')['Cantidad'].sum()
    
    df_m = pd.DataFrame({'Costo Total': df_c, 'Huevos Fértiles': df_h}).dropna()
    df_m['Costo Unitario'] = df_m['Costo Total'] / df_m['Huevos Fértiles']
    df_m = df_m[df_m['Huevos Fértiles'] > 0].reset_index()
    
    if df_m.empty:
        st.warning("No hay registros de lotes para el mes seleccionado.")
        st.stop()
        
    df_m['Lote'] = df_m['Lote'].astype(int).astype(str)
    df_m = df_m.sort_values('Costo Unitario', ascending=False)
    
    lote_critico = df_m.iloc[0]
    lote_eficiente = df_m.iloc[-1]
    promedio_mes = df_m['Costo Total'].sum() / df_m['Huevos Fértiles'].sum()
    
    # === INSIGHT DE LOTES AÑADIDO ===
    st.markdown(f"""
    <div class="alert-box">
        🚨 <b>Auditoría de Ineficiencia (Causa Raíz):</b><br>
        El <b>Lote {lote_critico['Lote']}</b> es el punto crítico de la operación este mes, con un costo unitario de <b>${lote_critico['Costo Unitario']:,.1f} COP/HF</b> (muy por encima del promedio de ${promedio_mes:,.1f}).<br>
        <i>¿Por qué sucede esto?</i> Este lote produjo un volumen muy bajo ({lote_critico['Huevos Fértiles']:,.0f} huevos), lo que provoca que los costos fijos (depreciación gallina/M.O.) se dividan entre muy pocas unidades, disparando el costo. <b>Evaluar descarte por curva de postura.</b><br><br>
        ✅ El <b>Lote {lote_eficiente['Lote']}</b> sostiene la rentabilidad con una alta producción ({lote_eficiente['Huevos Fértiles']:,.0f} huevos) y un costo óptimo de <b>${lote_eficiente['Costo Unitario']:,.1f} COP/HF</b>.
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Lote Crítico (Mayor Costo)", f"Lote {lote_critico['Lote']}", f"${lote_critico['Costo Unitario']:,.1f}/HF", delta_color="inverse")
    c2.metric("Promedio Ponderado Mes", f"General", f"${promedio_mes:,.1f}/HF", delta_color="off")
    c3.metric("Lote Más Eficiente", f"Lote {lote_eficiente['Lote']}", f"${lote_eficiente['Costo Unitario']:,.1f}/HF")

    st.markdown("---")
    col_grafico, col_tabla = st.columns([3, 2])
    
    with col_grafico:
        st.subheader("📊 Ranking de Ineficiencia por Lote")
        fig_bar_lote = px.bar(
            df_m, x='Lote', y='Costo Unitario', text_auto='.1f',
            color='Costo Unitario', color_continuous_scale='RdYlGn_r',
            title=f"Costo Unitario Comparativo - {p_actual}"
        )
        fig_bar_lote.add_hline(y=promedio_mes, line_dash="dot", annotation_text=f"Promedio: ${promedio_mes:,.1f}", line_color="black")
        fig_bar_lote.update_traces(textposition='outside')
        fig_bar_lote.update_layout(yaxis_title="Costo Unitario ($ COP)")
        st.plotly_chart(fig_bar_lote, use_container_width=True)

    with col_tabla:
        st.subheader("📋 Matriz de Auditoría Lotes")
        df_tabla = df_m[['Lote', 'Huevos Fértiles', 'Costo Total', 'Costo Unitario']].copy()
        st.dataframe(df_tabla.style.format({
            'Huevos Fértiles': '{:,.0f} unds', 'Costo Total': '${:,.0f}', 'Costo Unitario': '${:,.2f}'
        }), use_container_width=True)

# =============================================================================
# 5. DETALLE COSTOS POR LÍNEA
# =============================================================================
elif menu == "5. Detalle Costos por Línea":
    st.markdown('<p class="main-title">5. EVALUACIÓN FINANCIERA POR GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    
    df_c = df_raw_graficas[df_raw_graficas['Texto explicativo'] != TEXTO_LIQUIDACION]
    df_h = df_raw_graficas[df_raw_graficas['Texto breve de material'] == MATERIAL_HF]
    
    c_lin = df_c[df_c['Periodo'] == p_actual].groupby('linea')['Totales'].sum()
    h_lin = df_h[df_h['Periodo'] == p_actual].groupby('linea')['Cantidad'].sum()
    
    df_gen = pd.DataFrame({'Costo Total ($)': c_lin, 'Huevos Fértiles': h_lin})
    df_gen['Costo Unitario ($/HF)'] = df_gen['Costo Total ($)'] / df_gen['Huevos Fértiles']
    df_gen = df_gen.reset_index().dropna()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Evolución Histórica por Raza")
        df_ts_g = df_h.groupby(['Periodo', 'linea'])['Cantidad'].sum().reset_index()
        fig_ts = px.line(df_ts_g, x='Periodo', y='Cantidad', color='linea', markers=True)
        fig_ts.update_layout(yaxis_range=[df_ts_g['Cantidad'].min()*0.9, df_ts_g['Cantidad'].max()*1.15])
        st.plotly_chart(fig_ts, use_container_width=True)
        
    with col2:
        st.subheader(f"📊 Costo Unitario Promedio por Raza ({p_actual})")
        fig_bar = px.bar(df_gen, x='linea', y='Costo Unitario ($/HF)', color='linea', text_auto='.1f')
        st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(df_gen.style.format({'Costo Total ($)': '${:,.0f}', 'Huevos Fértiles': '{:,.0f}', 'Costo Unitario ($/HF)': '${:,.2f}'}), use_container_width=True)

# =============================================================================
# 6. COSTO KG ALIMENTO
# =============================================================================
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">6. IMPACTO DEL COSTO DE ALIMENTO</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    
    # === INSIGHT FINANCIERO ALIMENTO AÑADIDO ===
    st.markdown("""
    <div class="insight-box">
        💡 <b>Sensibilidad Financiera:</b> El alimento pondera aproximadamente el 40% del costo total. Monitorear las fluctuaciones del precio del kilogramo y la conversión alimenticia (gramos consumidos por huevo) es el pilar de la rentabilidad avícola.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 Evolución Precio Alimento ($/Kg)")
        fig_a = px.line(df_filtrado_graficas, x='Periodo', y='Precio Kg Alimento', markers=True)
        fig_a.update_traces(line_color='#d62728', line_width=3)
        fig_a.update_layout(yaxis_range=[df_filtrado_graficas['Precio Kg Alimento'].min()*0.9, df_filtrado_graficas['Precio Kg Alimento'].max()*1.1])
        st.plotly_chart(fig_a, use_container_width=True)
        
    with col2:
        st.subheader("📈 Evolución Conversión (g/Huevo)")
        fig_g = px.line(df_filtrado_graficas, x='Periodo', y='Gramos Alimento/Huevo', markers=True)
        fig_g.update_traces(line_color='#2ca02c', line_width=3)
        fig_g.update_layout(yaxis_range=[df_filtrado_graficas['Gramos Alimento/Huevo'].min()*0.9, df_filtrado_graficas['Gramos Alimento/Huevo'].max()*1.1])
        st.plotly_chart(fig_g, use_container_width=True)

    st.subheader("📋 Matriz Comparativa Interanual (Alimento):")
    df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO'].copy()
    res_alim = df_alim.groupby(['Anio', 'Mes_Num']).apply(
        lambda x: x['Totales'].sum() / x['Cantidad'].sum() if x['Cantidad'].sum() > 0 else 0
    ).unstack(level=0)
    
    if 2025 in res_alim.columns and 2026 in res_alim.columns:
        res_df = res_alim[[2025, 2026]].dropna().reset_index()
        res_df['% VAR'] = ((res_df[2026] - res_df[2025]) / res_df[2025]) * 100
        st.dataframe(res_df.style.format({2025: '${:,.1f}', 2026: '${:,.1f}', '% VAR': '{:+.2f}%'}), use_container_width=True)
    else:
        st.info("La tabla comparativa 2025 vs 2026 requiere datos de ambos años procesados en la matriz base.")
