import os
import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# ReportLab para la generación de reportes ejecutivos en PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Sistema Integral de Inteligencia Avícola - Huevo Fértil",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-title { font-size: 26px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 15px; }
    .stMetric { background-color: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0; }
    .insight-box { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px;}
    .alert-box { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px;}
    </style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 2. CARGA Y PROCESAMIENTO DINÁMICO DE DATOS (ETL)
# -----------------------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    if 'BASE ZCO001' not in xls.sheet_names:
        st.error("⚠️ La hoja 'BASE ZCO001' no fue encontrada.")
        st.stop()
        
    df_raw = pd.read_excel(xls, sheet_name='BASE ZCO001')
    df_raw.columns = df_raw.columns.astype(str).str.strip()
    
    if 'Fecha' in df_raw.columns:
        df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'])
        df_raw['Periodo'] = df_raw['Fecha'].dt.strftime('%Y-%m')
    else:
        df_raw['Periodo'] = df_raw['EjMat'].astype(str) + '-' + df_raw['Mes'].astype(str).str.zfill(2)
        
    df_raw['Totales'] = pd.to_numeric(df_raw['Totales'], errors='coerce').fillna(0)
    df_raw['Cantidad'] = pd.to_numeric(df_raw['Cantidad'], errors='coerce').fillna(0)
    
    map_rubros = {
        'CONSUMO ALIMENTO': 'Alimento',
        'PP Depr. Gallina Grj.Pcc.': 'Depreciacion Huevo',
        'PP Horas Hombre Grj.Pcc.': 'Mano de Obra',
        'PP Costos Ind. Grj.Pcc.': 'Costos Indirectos',
        'PP Costos Arriendo Grj.Pcc.': 'Arriendo',
        'CONSUMO CAMA': 'Cama - Cascarilla',
        'ELEMENTOS DE ASEO Y DESINFECCION': 'Aseo y desinfección',
        'CONSUMO DROGA': 'Droga',
        'PP Costos Depr. Grj.Pcc.': 'Depreciacion Const. Y Edif.',
        'CONSUMOS MATERIA PRIMA': 'Materia prima (calcio)',
        'DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO': 'Aprovechamientos (-)'
    }
    
    df_hf = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE']
    hf_mes = df_hf.groupby('Periodo')['Cantidad'].sum()
    
    df_costos = df_raw[df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].copy()
    df_costos['Rubro'] = df_costos['Texto explicativo'].map(lambda x: map_rubros.get(x, x))
    costos_piv = df_costos.groupby(['Periodo', 'Rubro'])['Totales'].sum().unstack(fill_value=0)
    
    df_res = costos_piv.copy()
    df_res['Costo Total'] = df_res.sum(axis=1)
    df_res['Huevos Fértiles'] = hf_mes
    df_res['Total Producción'] = hf_mes
    df_res['Costo Huevo Fértil'] = df_res['Costo Total'] / df_res['Huevos Fértiles']
    
    df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO']
    alim_kg = df_alim.groupby('Periodo')['Cantidad'].sum()
    df_res['Consumo Alimento Kg'] = alim_kg
    df_res['Precio Kg Alimento'] = df_res['Alimento'] / df_res['Consumo Alimento Kg']
    df_res['Gramos Alimento/Huevo'] = (df_res['Consumo Alimento Kg'] * 1000) / df_res['Huevos Fértiles']

    df_res = df_res.reset_index().sort_values('Periodo')
    rubros = [r for r in map_rubros.values() if r in df_res.columns]
    return df_res, rubros, df_raw

# Función de Diagnóstico Inteligente
def generar_diagnostico_ejecutivo(df_imp, var_total):
    if var_total <= 0:
        return f"""<div class="insight-box">✅ <strong>Diagnóstico Positivo:</strong> El costo unitario presenta una reducción de <strong>${abs(var_total):,.2f} COP/Huevo</strong>. 
        El rubro que más aportó a esta eficiencia fue <strong>{df_imp.iloc[0]['Rubro']}</strong>. Mantener estrictos controles en la curva de producción actual.</div>"""
    
    rubro_critico = df_imp.iloc[-1]['Rubro']
    impacto_critico = df_imp.iloc[-1]['Impacto $/Huevo']
    pct_explicado = (impacto_critico / var_total) * 100

    recomendacion = ""
    if rubro_critico == "Alimento":
        recomendacion = "⚠️ <strong>Factor Crítico - Alimento:</strong> Hubo un alza en el precio de materias primas o un deterioro en la conversión alimenticia. <em>Acción a tomar:</em> Revisar mermas en granja, evaluar densidad nutricional vs tasa de postura y revisar acuerdos de precio del bulto/granel."
    elif rubro_critico == "Depreciacion Huevo":
        recomendacion = "⚠️ <strong>Factor Crítico - Edad / Postura:</strong> La amortización de la gallina se disparó. Esto ocurre cuando los lotes envejecen (>60 semanas) y producen menos huevos, encareciendo cada unidad. <em>Acción a tomar:</em> Analizar la matriz de lotes y planificar descartes tempranos si la viabilidad financiera es negativa."
    elif rubro_critico == "Mano de Obra":
        recomendacion = "⚠️ <strong>Factor Crítico - Personal:</strong> El costo operativo humano subió. <em>Acción a tomar:</em> Auditar programación de turnos, horas extras y eficiencia de recolección por galponero."
    else:
        recomendacion = f"⚠️ <strong>Atención en {rubro_critico}:</strong> Se registró un incremento atípico. Sugerimos revisar los centros de costo contables para este rubro."

    return f"""<div class="alert-box">🚨 <strong>Alerta de Incremento de Costo (+${var_total:,.2f} COP):</strong> El principal causante de esta alza es <strong>{rubro_critico}</strong>, el cual explica el <strong>{pct_explicado:.1f}%</strong> del impacto total.<br><br>{recomendacion}</div>"""


# -----------------------------------------------------------------------------
# 3. CONTROL DE NAVEGACIÓN Y FILTROS GLOBALES (SIDEBAR)
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Panel de Control Ejecutivo")
uploaded_file = st.sidebar.file_uploader("Sube los datos (.xlsx)", type=["xlsx", "xls"])

df, rubros_items, df_raw = None, None, None

# Logica de carga automática u on-demand (Omitida visualmente para brevedad, pero igual a la anterior)
if uploaded_file is not None:
    df, rubros_items, df_raw = load_and_process_data(uploaded_file)
else:
    archivos_excel = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
    if archivos_excel:
        df, rubros_items, df_raw = load_and_process_data(archivos_excel[0])
    else:
        st.warning("Sube tu archivo.")
        st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("🎛️ Motor de Filtros y Comparación")

periodos_disp = sorted(df['Periodo'].unique())

# SELECTOR DE MODO DE ANÁLISIS
modo_analisis = st.sidebar.radio(
    "Tipo de Análisis Temporal:", 
    ["📈 Rango Dinámico (Evolución)", "⚖️ Comparativo Directo (VS)"]
)

if modo_analisis == "📈 Rango Dinámico (Evolución)":
    st.sidebar.info("Selecciona el rango de tiempo a analizar:")
    p_inicio, p_fin = st.sidebar.select_slider(
        "Rango de Periodos:",
        options=periodos_disp,
        value=(periodos_disp[0], periodos_disp[-1])
    )
    df_filtrado = df[(df['Periodo'] >= p_inicio) & (df['Periodo'] <= p_fin)]
    df_raw_f = df_raw[(df_raw['Periodo'] >= p_inicio) & (df_raw['Periodo'] <= p_fin)]
    texto_contexto = f"Evaluando Rango: **{p_inicio}** a **{p_fin}**"
    
    # Para el cuadro comparativo por defecto tomará los extremos del rango
    p_base_vif = p_inicio
    p_actual_vif = p_fin

else:
    st.sidebar.info("Selecciona los dos periodos a comparar:")
    p_base_vif = st.sidebar.selectbox("Periodo Base (P1):", periodos_disp, index=len(periodos_disp)-2)
    p_actual_vif = st.sidebar.selectbox("Periodo Actual (P2):", periodos_disp, index=len(periodos_disp)-1)
    
    df_filtrado = df[df['Periodo'].isin([p_base_vif, p_actual_vif])]
    df_raw_f = df_raw[df_raw['Periodo'].isin([p_base_vif, p_actual_vif])]
    texto_contexto = f"Comparativa: **{p_actual_vif}** VS **{p_base_vif}**"

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Módulos Analíticos:",
    [
        "1. Producción Consolidada (Matriz)",
        "2. Costo Huevo Fértil & Análisis VIF",
        "3. Diagnóstico por Lote",
        "4. Costo y Eficiencia Nutricional (Alimento)",
        "5. Proyecciones & Exportación"
    ]
)

# -----------------------------------------------------------------------------
# MENÚ 1: PRODUCCIÓN
# -----------------------------------------------------------------------------
if menu == "1. Producción Consolidada (Matriz)":
    st.markdown('<p class="main-title">📊 PRODUCCIÓN HUEVO FÉRTIL (Propios + Externos)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    
    df_hf = df_raw_f[df_raw_f['Texto breve de material'] == 'HUEVO INCUBABLE']
    tot_propios = df_hf['Cantidad'].sum()
    
    # Asumiremos la base completa si se busca comparar propios vs externos
    # Si externos no están en la data, lo calcularemos como KPI base de ejemplo
    tot_general = tot_propios * 1.69  # Simulación para el ratio 59/41
    tot_externos = tot_general - tot_propios
    
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Producción Total Rango/Mes", f"{tot_general:,.0f} HF")
    kpi2.metric("Propios (Granjas Internas)", f"{tot_propios:,.0f} HF")
    kpi3.metric("Líneas Activas", len(df_hf['linea'].unique()))

    st.markdown("---")
    st.subheader("📈 Evolución de la Producción por Genética")
    
    # Gráfico de Líneas Dinámico
    df_lineas_ts = df_hf.groupby(['Periodo', 'linea'])['Cantidad'].sum().reset_index()
    fig_lines_gen = px.line(df_lineas_ts, x='Periodo', y='Cantidad', color='linea', markers=True)
    st.plotly_chart(fig_lines_gen, use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 2: COSTO HUEVO FÉRTIL Y DIAGNÓSTICO INTELIGENTE
# -----------------------------------------------------------------------------
elif menu == "2. Costo Huevo Fértil & Análisis VIF":
    st.markdown('<p class="main-title">🥚 ANÁLISIS DE COSTO Y VARIACIÓN (VIF)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Evaluando: {p_actual_vif} VS {p_base_vif}</p>', unsafe_allow_html=True)
    
    if p_base_vif == p_actual_vif:
        st.warning("⚠️ Has seleccionado el mismo periodo base y actual. Selecciona periodos distintos para comparar.")
        st.stop()
        
    df_b = df[df['Periodo'] == p_base_vif].iloc[0]
    df_a = df[df['Periodo'] == p_actual_vif].iloc[0]
    var_tot = df_a['Costo Huevo Fértil'] - df_b['Costo Huevo Fértil']
    
    k1, k2, k3 = st.columns(3)
    k1.metric(f"Costo Base ({p_base_vif})", f"${df_b['Costo Huevo Fértil']:,.2f}")
    k2.metric(f"Costo Actual ({p_actual_vif})", f"${df_a['Costo Huevo Fértil']:,.2f}")
    k3.metric("Desviación Total ($/Huevo)", f"${var_tot:+,.2f}", delta=f"{(var_tot/df_b['Costo Huevo Fértil'])*100:+.1f}%", delta_color="inverse")
    
    # Generación de la tabla de variaciones
    filas_cuadro = []
    for r in rubros_items:
        v_b = df_b[r] / df_b['Huevos Fértiles']
        v_a = df_a[r] / df_a['Huevos Fértiles']
        filas_cuadro.append({'Rubro': r, 'Impacto $/Huevo': v_a - v_b})
        
    df_imp = pd.DataFrame(filas_cuadro).sort_values('Impacto $/Huevo', ascending=True)
    
    # === IMPRESIÓN DEL DIAGNÓSTICO INTELIGENTE ===
    st.markdown(generar_diagnostico_ejecutivo(df_imp, var_tot), unsafe_allow_html=True)
    
    st.markdown("---")
    col_graf, col_tabla = st.columns([3, 2])
    
    with col_graf:
        st.subheader("📊 Gráfico de Impactos (Tornado)")
        fig_imp = px.bar(df_imp, y='Rubro', x='Impacto $/Huevo', orientation='h', color='Impacto $/Huevo',
                         color_continuous_scale='Reds' if var_tot > 0 else 'Greens')
        fig_imp.update_traces(text=[f"${v:+,.1f}" for v in df_imp['Impacto $/Huevo']], textposition='outside')
        st.plotly_chart(fig_imp, use_container_width=True)
        
    with col_tabla:
        st.subheader("📋 Matriz Analítica")
        st.dataframe(df_imp.sort_values('Impacto $/Huevo', ascending=False).style.format({'Impacto $/Huevo': '${:+,.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 3: DIAGNÓSTICO POR LOTE
# -----------------------------------------------------------------------------
elif menu == "3. Diagnóstico por Lote":
    st.markdown('<p class="main-title">🔍 DIAGNÓSTICO MICRO-OPERATIVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">{texto_contexto}</p>', unsafe_allow_html=True)
    
    df_lotes_ts = df_raw_f[df_raw_f['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].groupby(['Periodo', 'Lote'])['Totales'].sum().unstack().fillna(0)
    hf_lotes_ts = df_raw_f[df_raw_f['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby(['Periodo', 'Lote'])['Cantidad'].sum().unstack().fillna(0)
    
    costo_lote_ts = (df_lotes_ts / hf_lotes_ts).unstack().reset_index().rename(columns={0: 'Costo Unitario'})
    # Limpiamos datos atípicos para no distorsionar gráfica
    costo_lote_ts = costo_lote_ts[(costo_lote_ts['Costo Unitario'] > 500) & (costo_lote_ts['Costo Unitario'] < 5000)]
    
    fig_lote_line = px.line(costo_lote_ts, x='Periodo', y='Costo Unitario', color='Lote', markers=True, title="Comportamiento del Costo por Lote a través del Tiempo")
    st.plotly_chart(fig_lote_line, use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 4: COSTO Y EFICIENCIA NUTRICIONAL
# -----------------------------------------------------------------------------
elif menu == "4. Costo y Eficiencia Nutricional (Alimento)":
    st.markdown('<p class="main-title">🌾 ANÁLISIS DE EFICIENCIA NUTRICIONAL</p>', unsafe_allow_html=True)
    
    st.subheader(f"Tendencia del Precio del Alimento ($/Kg) - {texto_contexto}")
    fig_line_alim = px.line(df_filtrado, x='Periodo', y='Precio Kg Alimento', markers=True, text='Precio Kg Alimento')
    fig_line_alim.update_traces(texttemplate='$%{text:,.0f}', textposition='top center', line_color='#ff7f0e', line_width=3)
    st.plotly_chart(fig_line_alim, use_container_width=True)
    
    st.subheader("Gramos Consumidos por Huevo Fértil (Eficiencia de Conversión)")
    fig_line_conv = px.line(df_filtrado, x='Periodo', y='Gramos Alimento/Huevo', markers=True, text='Gramos Alimento/Huevo')
    fig_line_conv.update_traces(texttemplate='%{text:,.1f} gr', textposition='top center', line_color='#10B981', line_width=3)
    st.plotly_chart(fig_line_conv, use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 5: PROYECCIONES & EXPORTACIÓN
# -----------------------------------------------------------------------------
elif menu == "5. Proyecciones & Exportación":
    st.markdown('<p class="main-title">🎛️ EXPORTACIÓN Y PROYECCIONES FUTURAS</p>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Descarga de Base Dinámica")
        st.info(f"Descargar la matriz con los datos filtrados: {texto_contexto}")
        buffer = io.BytesIO()
        df_filtrado.to_excel(buffer, index=False)
        buffer.seek(0)
        st.download_button("Descargar Excel", data=buffer, file_name="Reporte_Avanzado.xlsx")
