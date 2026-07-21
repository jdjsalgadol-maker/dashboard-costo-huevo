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
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA (STREAMLIT)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Sistema Integral de Inteligencia Avícola - Huevo Fértil",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    .main-title { font-size: 26px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 15px; }
    .stMetric { background-color: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0; }
    </style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 2. CARGA Y PROCESAMIENTO DINÁMICO DE DATOS (ETL)
# -----------------------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_source):
    xls = pd.ExcelFile(file_source)
    
    if 'BASE ZCO001' in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name='BASE ZCO001')
        df_raw.columns = df_raw.columns.astype(str).str.strip()
        
        # Generar variables temporales
        if 'Fecha' in df_raw.columns:
            df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'])
            df_raw['Anio'] = df_raw['Fecha'].dt.year
            df_raw['Mes_Num'] = df_raw['Fecha'].dt.month
            df_raw['Periodo'] = df_raw['Fecha'].dt.strftime('%Y-%m')
        else:
            df_raw['Anio'] = df_raw['EjMat']
            df_raw['Mes_Num'] = df_raw['Mes']
            df_raw['Periodo'] = df_raw['EjMat'].astype(str) + '-' + df_raw['Mes'].astype(str).str.zfill(2)
            
        df_raw['Totales'] = pd.to_numeric(df_raw['Totales'], errors='coerce').fillna(0)
        df_raw['Cantidad'] = pd.to_numeric(df_raw['Cantidad'], errors='coerce').fillna(0)
        
        # Mapeo estandarizado de conceptos
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
        
        # Extracción de Huevos Fértiles
        df_hf = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE']
        hf_mes = df_hf.groupby('Periodo')['Cantidad'].sum()
        
        # Filtrado de Costos
        df_costos = df_raw[df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].copy()
        df_costos['Rubro'] = df_costos['Texto explicativo'].map(lambda x: map_rubros.get(x, x))
        
        costos_piv = df_costos.groupby(['Periodo', 'Rubro'])['Totales'].sum().unstack(fill_value=0)
        
        df_res = costos_piv.copy()
        df_res['Costo Total'] = df_res.sum(axis=1)
        df_res['Huevos Fértiles'] = hf_mes
        df_res['Total Producción'] = hf_mes
        df_res['Costo Huevo Fértil'] = df_res['Costo Total'] / df_res['Huevos Fértiles']
        
        # Consumo de Alimento
        df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO']
        alim_kg = df_alim.groupby('Periodo')['Cantidad'].sum()
        df_res['Consumo Alimento Kg'] = alim_kg
        df_res['Precio Kg Alimento'] = df_res['Alimento'] / df_res['Consumo Alimento Kg']
        df_res['Gramos Alimento/Huevo'] = (df_res['Consumo Alimento Kg'] * 1000) / df_res['Huevos Fértiles']

        df_res = df_res.reset_index().sort_values('Periodo')
        
        # Extraer variables de Año y Mes en df_res
        df_res['Anio'] = df_res['Periodo'].str.split('-').str[0].astype(int)
        df_res['Mes_Num'] = df_res['Periodo'].str.split('-').str[1].astype(int)
        
        rubros = [r for r in map_rubros.values() if r in df_res.columns]
        
        return df_res, rubros, df_raw
    else:
        st.error("⚠️ Estructura no válida: La hoja 'BASE ZCO001' no fue encontrada en el archivo Excel.")
        st.stop()


# -----------------------------------------------------------------------------
# 3. CONTROL DE NAVEGACIÓN Y FILTROS GLOBALES (SIDEBAR)
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Panel de Control Avícola")
st.sidebar.subheader("📂 Ingestión de Datos")

uploaded_file = st.sidebar.file_uploader("Sube el archivo de datos (.xlsx)", type=["xlsx", "xls"])

df, rubros_items, df_raw = None, None, None

if uploaded_file is not None:
    try:
        df, rubros_items, df_raw = load_and_process_data(uploaded_file)
        st.sidebar.success(f"¡Cargado: `{uploaded_file.name}`!")
    except Exception as e:
        st.sidebar.error(f"Error al procesar el archivo: {e}")
        st.stop()
else:
    archivos_excel = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
    if os.path.exists('data'):
        archivos_excel += [os.path.join('data', f) for f in os.listdir('data') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
        
    if archivos_excel:
        archivo_encontrado = archivos_excel[0]
        try:
            df, rubros_items, df_raw = load_and_process_data(archivo_encontrado)
            st.sidebar.info(f"Cargado automático: `{archivo_encontrado}`")
        except Exception as e:
            st.sidebar.error(f"Error cargando `{archivo_encontrado}`: {e}")
            st.stop()
    else:
        st.warning("⚠️ No se encontró ningún archivo Excel. Arrastra tu archivo a la barra lateral.")
        st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("🎛️ Filtro Temporal")

# SEGMENTADOR DE MODO: DENTRO DEL PERIODO VS HISTÓRICO
modo_temporal = st.sidebar.radio("Vista Temporal:", ["Mes Específico", "Histórico Completo (2024-2026)"])

anios_disponibles = sorted(df['Anio'].unique(), reverse=True)
meses_nombres = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

if modo_temporal == "Mes Específico":
    col_a, col_m = st.sidebar.columns(2)
    with col_a:
        anio_sel = st.selectbox("Año:", anios_disponibles, index=0)
    with col_m:
        meses_disp = sorted(df[df['Anio'] == anio_sel]['Mes_Num'].unique())
        mes_sel = st.selectbox("Mes:", meses_disp, index=len(meses_disp)-1, format_func=lambda x: meses_nombres.get(x, str(x)))
    
    periodo_sel = f"{anio_sel}-{str(mes_sel).zfill(2)}"
    df_filtrado = df[df['Periodo'] == periodo_sel]
    df_raw_filtrado = df_raw[(df_raw['Anio'] == anio_sel) & (df_raw['Mes_Num'] == mes_sel)]
else:
    periodo_sel = "HISTÓRICO COMPLETO"
    df_filtrado = df.copy()
    df_raw_filtrado = df_raw.copy()

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Selecciona el Reporte Ejecutivo:",
    [
        "1. Producción 2026 (Propios + Externos)",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil (Macro & Variaciones)",
        "4. Detalle Costos Huevo por Lote",
        "5. Detalle Costos Huevo por Línea",
        "6. Costo Kg Alimento",
        "7. Simulador What-If & Proyección",
        "8. Centro de Exportación de Informes"
    ]
)


# -----------------------------------------------------------------------------
# MENÚ 1: PRODUCCIÓN 2026 (PROPIOS + EXTERNOS)
# -----------------------------------------------------------------------------
if menu == "1. Producción 2026 (Propios + Externos)":
    st.markdown('<p class="main-title">📊 PRODUCCIÓN 2026 (Propios + Externos)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Análisis del volumen de huevo fértil - Filtro: <b>{periodo_sel}</b></p>', unsafe_allow_html=True)
    
    # Texto sin marcas LaTeX erróneas
    with st.expander("💡 **Interpretación Gerencial & Diagnóstico:**", expanded=True):
        st.write("""
        * **Autonomía de Producción:** La compañía mantiene un **59.0% de producción propia en granjas** (29.46 Millones de unidades) y recurre a **maquilas y compras externas en un 41.0%** (20.47 Millones de unidades) para cubrir el programa de incubación.
        * **Concentración de Genética:** Las razas **ROSS** (47.4%) y **ROSSAP** (35.7%) representan más del **83.1% del volumen total**, siendo las genéticas principales que sostienen la operación.
        """)

    tot_propios = 29460823
    tot_externos = 20473164
    tot_general = tot_propios + tot_externos
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Producción Total", f"{tot_general:,.0f} HF")
    kpi2.metric("Propios (Granja)", f"{tot_propios:,.0f} HF", delta="59.0% Participación")
    kpi3.metric("Externos (Maquila)", f"{tot_externos:,.0f} HF", delta="41.0% Participación")
    kpi4.metric("Línea Genética Líder", "ROSS", "47.43% del Mercado")

    st.markdown("---")
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE EVOLUCIÓN HISTÓRICA DE PRODUCCIÓN
    st.subheader("📈 Tendencia Histórica de Producción Huevo Fértil (Mensual)")
    df_hf_linea = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby('Periodo')['Cantidad'].sum().reset_index()
    fig_line_prod = px.line(df_hf_linea, x='Periodo', y='Cantidad', markers=True, text='Cantidad',
                            title="Evolución del Volumen Total Producido (2024 - 2026)")
    fig_line_prod.update_traces(texttemplate='%{text:,.0f}', textposition='top center', line_color='#1E3A8A', line_width=3)
    st.plotly_chart(fig_line_prod, use_container_width=True)
    
    st.markdown("---")
    
    col_chart, col_data = st.columns([3, 2])
    with col_chart:
        df_part = pd.DataFrame({
            'Línea': ['ROSS', 'ROSSAP', 'COBB M', 'ROSSAPFF', 'ROSSAPN'],
            'Total 2026': [23277936, 17523812, 3582415, 2361185, 2337037],
            '% Part': [47.43, 35.70, 7.30, 4.81, 4.76]
        })
        fig_gen = px.bar(df_part, x='Línea', y='Total 2026', text_auto='.2s', color='Línea', title="Distribución del Volumen por Línea Genética (2026)")
        for idx, row in df_part.iterrows():
            fig_gen.add_annotation(x=row['Línea'], y=row['Total 2026'] * 0.15, text=f"<b>{row['% Part']:.2f}%</b>", showarrow=False, bgcolor="white", bordercolor="black")
        fig_gen.update_layout(height=420, showlegend=False, yaxis_title="Huevos Fértiles")
        st.plotly_chart(fig_gen, use_container_width=True)
        
    with col_data:
        st.subheader("📌 Matriz de Participación por Genética")
        st.dataframe(df_part.style.format({'Total 2026': '{:,.0f}', '% Part': '{:.2f}%'}), use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 2: PRODUCCIÓN MES A MES POR LÍNEA
# -----------------------------------------------------------------------------
elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">📈 PRODUCCIÓN MES A MES POR LÍNEA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Comportamiento del volumen por raza - Filtro: <b>{periodo_sel}</b></p>', unsafe_allow_html=True)
    
    df_2026_hf = df_raw_filtrado[df_raw_filtrado['Texto breve de material'] == 'HUEVO INCUBABLE']
    piv_prod = df_2026_hf.groupby(['Periodo', 'linea'])['Cantidad'].sum().unstack().fillna(0)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS POR LÍNEA GENÉTICA
    st.subheader("📈 Evolución Temporal por Raza (Línea de Tiempo)")
    df_lineas_ts = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby(['Periodo', 'linea'])['Cantidad'].sum().reset_index()
    fig_lines_gen = px.line(df_lineas_ts, x='Periodo', y='Cantidad', color='linea', markers=True, title="Tendencia de Producción por Genética (2024 - 2026)")
    st.plotly_chart(fig_lines_gen, use_container_width=True)
    
    st.markdown("---")
    st.subheader("Evolución Mensual en Granjas Propias:")
    fig_prod = px.bar(piv_prod.reset_index(), x='Periodo', y=piv_prod.columns, title="Producción Huevo Fértil por Línea", barmode='group', text_auto='.2s')
    st.plotly_chart(fig_prod, use_container_width=True)
    st.dataframe(piv_prod.style.format('{:,.0f}'), use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 3: COSTO HUEVO FÉRTIL (MACRO & CUADRO DE VARIACIONES VIF)
# -----------------------------------------------------------------------------
elif menu == "3. Costo Huevo Fértil (Macro & Variaciones)":
    st.markdown('<p class="main-title">🥚 COSTO HUEVO FÉRTIL (Macro & Tabla de Variaciones)</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Análisis de variación intermensual y tabla de desviaciones en $/Huevo.</p>', unsafe_allow_html=True)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE TENDENCIA HISTÓRICA DEL COSTO
    st.subheader("📈 Tendencia Histórica del Costo por Huevo Fértil ($ COP / Huevo)")
    fig_line_costo = px.line(df, x='Periodo', y='Costo Huevo Fértil', markers=True, text='Costo Huevo Fértil', title="Evolución del Costo Unitario Histórico (2024 - 2026)")
    fig_line_costo.update_traces(texttemplate='$%{text:,.1f}', textposition='top center', line_color='#1E3A8A', line_width=3)
    st.plotly_chart(fig_line_costo, use_container_width=True)
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        mes_b = st.selectbox("Período Base (Inicial):", df['Periodo'].tolist(), index=len(df)-2)
    with col2:
        mes_a = st.selectbox("Período de Análisis (Final):", df['Periodo'].tolist(), index=len(df)-1)
        
    df_b = df[df['Periodo'] == mes_b].iloc[0]
    df_a = df[df['Periodo'] == mes_a].iloc[0]
    var_tot = df_a['Costo Huevo Fértil'] - df_b['Costo Huevo Fértil']
    
    k1, k2, k3 = st.columns(3)
    k1.metric(f"Costo {mes_b}", f"${df_b['Costo Huevo Fértil']:,.2f}")
    k2.metric(f"Costo {mes_a}", f"${df_a['Costo Huevo Fértil']:,.2f}")
    k3.metric("Variación Total Unitario", f"${var_tot:+,.2f} COP", delta=f"{(var_tot/df_b['Costo Huevo Fértil'])*100:+.1f}%", delta_color="inverse")
    
    st.markdown("---")
    
    # REINCORPORACIÓN DEL CUADRO COMPARATIVO DEL PRIMER DASHBOARD
    st.subheader("📋 Cuadro Comparativo de Variaciones por Rubro (Detalle VIF):")
    
    filas_cuadro = []
    for r in rubros_items:
        val_b = df_b[r]
        val_a = df_a[r]
        unit_b = val_b / df_b['Huevos Fértiles']
        unit_a = val_a / df_a['Huevos Fértiles']
        dif_unit = unit_a - unit_b
        pct_b = (val_b / df_b['Costo Total']) * 100
        pct_a = (val_a / df_a['Costo Total']) * 100
        
        filas_cuadro.append({
            'CONCEPTO': r,
            f'Costo Unit {mes_b} ($)': unit_b,
            f'% Part {mes_b}': pct_b,
            f'Costo Unit {mes_a} ($)': unit_a,
            f'% Part {mes_a}': pct_a,
            'VARIACIÓN ($/HUEVO)': dif_unit,
            '% VAR UNITARIO': ((unit_a - unit_b) / unit_b) * 100 if unit_b > 0 else 0
        })
        
    df_cuadro = pd.DataFrame(filas_cuadro).sort_values(by='VARIACIÓN ($/HUEVO)', ascending=False)
    
    st.dataframe(
        df_cuadro.style.format({
            f'Costo Unit {mes_b} ($)': '${:,.2f}',
            f'% Part {mes_b}': '{:.2f}%',
            f'Costo Unit {mes_a} ($)': '${:,.2f}',
            f'% Part {mes_a}': '{:.2f}%',
            'VARIACIÓN ($/HUEVO)': '${:+,.2f}',
            '% VAR UNITARIO': '{:+.2f}%'
        }),
        use_container_width=True
    )
    
    st.markdown("---")
    # Gráfico de Barras de Impacto
    df_imp = pd.DataFrame([{'Rubro': r['CONCEPTO'], 'Impacto $/Huevo': r['VARIACIÓN ($/HUEVO)']} for r in filas_cuadro]).sort_values('Impacto $/Huevo', ascending=True)
    fig_imp = px.bar(df_imp, y='Rubro', x='Impacto $/Huevo', orientation='h', color='Impacto $/Huevo',
                     color_continuous_scale='Reds' if var_tot > 0 else 'Greens',
                     title=f"Impacto Directo de cada Rubro en la Variación de ${var_tot:+,.2f} COP/Huevo")
    fig_imp.update_traces(text=[f"${v:+,.1f}" for v in df_imp['Impacto $/Huevo']], textposition='outside')
    st.plotly_chart(fig_imp, use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 4: DETALLE COSTOS HUEVO POR LOTE
# -----------------------------------------------------------------------------
elif menu == "4. Detalle Costos Huevo por Lote":
    st.markdown('<p class="main-title">🔍 DETALLE DE COSTOS HUEVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Diagnóstico micro-operativo a nivel de lote - Filtro: <b>{periodo_sel}</b></p>', unsafe_allow_html=True)
    
    df_mes_lote = df_raw_filtrado[df_raw_filtrado['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS']
    piv_lote = df_mes_lote.pivot_table(index='Texto explicativo', columns='Lote', values='Totales', aggfunc='sum').fillna(0)
    
    df_hf_lote = df_raw_filtrado[df_raw_filtrado['Texto breve de material'] == 'HUEVO INCUBABLE']
    hf_lotes = df_hf_lote.groupby('Lote')['Cantidad'].sum()
    
    piv_lote_unit = piv_lote.divide(hf_lotes, axis=1).dropna(axis=1)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE EVOLUCIÓN POR LOTE
    st.subheader("📈 Tendencia Histórica de Costo Unitario por Lote ($/Huevo)")
    df_lotes_ts = df_raw[df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].groupby(['Periodo', 'Lote'])['Totales'].sum().unstack().fillna(0)
    hf_lotes_ts = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby(['Periodo', 'Lote'])['Cantidad'].sum().unstack().fillna(0)
    costo_lote_ts = (df_lotes_ts / hf_lotes_ts).unstack().reset_index().rename(columns={0: 'Costo Unitario'})
    costo_lote_ts = costo_lote_ts[(costo_lote_ts['Costo Unitario'] > 500) & (costo_lote_ts['Costo Unitario'] < 5000)]
    
    fig_lote_line = px.line(costo_lote_ts, x='Periodo', y='Costo Unitario', color='Lote', title="Comportamiento del Costo por Lote a través del Tiempo")
    st.plotly_chart(fig_lote_line, use_container_width=True)
    
    st.markdown("---")
    st.subheader("Matriz de Costo Unitario ($/HF) por Lote:")
    st.dataframe(piv_lote_unit.style.format('${:,.1f}'), use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 5: DETALLE COSTOS HUEVO POR LÍNEA
# -----------------------------------------------------------------------------
elif menu == "5. Detalle Costos Huevo por Línea":
    st.markdown('<p class="main-title">🧬 DETALLE DE COSTOS HUEVO POR LÍNEA GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Evaluación financiera por raza - Filtro: <b>{periodo_sel}</b></p>', unsafe_allow_html=True)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE COSTO POR GENÉTICA
    st.subheader("📈 Evolución del Costo Unitario por Línea Genética ($ COP / Huevo)")
    c_gen_ts = df_raw[df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].groupby(['Periodo', 'linea'])['Totales'].sum()
    h_gen_ts = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby(['Periodo', 'linea'])['Cantidad'].sum()
    df_gen_ts = (c_gen_ts / h_gen_ts).reset_index().rename(columns={0: 'Costo Unitario'})
    
    fig_gen_line = px.line(df_gen_ts, x='Periodo', y='Costo Unitario', color='linea', markers=True, title="Tendencia del Costo por Genética (2024 - 2026)")
    st.plotly_chart(fig_gen_line, use_container_width=True)
    
    st.markdown("---")
    costo_linea = df_raw_filtrado[df_raw_filtrado['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].groupby('linea')['Totales'].sum()
    huevos_linea = df_raw_filtrado[df_raw_filtrado['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby('linea')['Cantidad'].sum()
    
    df_gen = pd.DataFrame({'Costo Total ($)': costo_linea, 'Huevos Fértiles': huevos_linea})
    df_gen['Costo / Huevo ($)'] = df_gen['Costo Total ($)'] / df_gen['Huevos Fértiles']
    df_gen = df_gen.reset_index().dropna()
    
    st.dataframe(df_gen.style.format({'Costo Total ($)': '${:,.0f}', 'Huevos Fértiles': '{:,.0f}', 'Costo / Huevo ($)': '${:,.2f}'}), use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 6: COSTO KG ALIMENTO
# -----------------------------------------------------------------------------
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">🌾 COSTO KG ALIMENTO (2025 vs 2026)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitle">Seguimiento al precio del alimento - Filtro: <b>{periodo_sel}</b></p>', unsafe_allow_html=True)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE TENDENCIA DEL ALIMENTO
    st.subheader("📈 Tendencia Histórica del Costo por Kg de Alimento ($/Kg)")
    fig_line_alim = px.line(df, x='Periodo', y='Precio Kg Alimento', markers=True, text='Precio Kg Alimento', title="Evolución del Precio del Kg de Alimento (2024 - 2026)")
    fig_line_alim.update_traces(texttemplate='$%{text:,.0f}', textposition='top center', line_color='#ff7f0e', line_width=3)
    st.plotly_chart(fig_line_alim, use_container_width=True)
    
    st.markdown("---")
    df_alim_all = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO'].copy()
    df_alim_all['Anio'] = df_alim_all['EjMat']
    
    res_alim = df_alim_all.groupby(['Anio', 'Mes']).apply(
        lambda x: x['Totales'].sum() / x['Cantidad'].sum() if x['Cantidad'].sum() > 0 else 0
    ).unstack(level=0)
    
    res_alim_df = res_alim[[2025, 2026]].dropna().reset_index()
    res_alim_df['%VAR'] = ((res_alim_df[2026] - res_alim_df[2025]) / res_alim_df[2025]) * 100
    
    fig_alim = go.Figure()
    fig_alim.add_trace(go.Bar(x=res_alim_df['Mes'], y=res_alim_df[2025], name='2025', marker_color='#8c564b', text=[f"${v:,.0f}" for v in res_alim_df[2025]], textposition='outside'))
    fig_alim.add_trace(go.Bar(x=res_alim_df['Mes'], y=res_alim_df[2026], name='2026', marker_color='#ff7f0e', text=[f"${v:,.0f}" for v in res_alim_df[2026]], textposition='outside'))
    fig_alim.add_trace(go.Scatter(x=res_alim_df['Mes'], y=res_alim_df['%VAR'], name='%VAR', yaxis='y2', mode='lines+markers+text', text=[f"{v:+.1f}%" for v in res_alim_df['%VAR']], textposition='top center', line=dict(color='black', width=3)))
    
    fig_alim.update_layout(title="Costo Ponderado por Kg de Alimento (2026 vs 2025)", yaxis_title="Costo $/Kg", yaxis2=dict(title="% Variación", overlaying='y', side='right'), barmode='group', height=450)
    st.plotly_chart(fig_alim, use_container_width=True)
    st.dataframe(res_alim_df.style.format({2025: '${:,.1f}', 2026: '${:,.1f}', '%VAR': '{:+.2f}%'}), use_container_width=True)


# -----------------------------------------------------------------------------
# MENÚ 7: SIMULADOR WHAT-IF & PROYECCIÓN
# -----------------------------------------------------------------------------
elif menu == "7. Simulador What-If & Proyección":
    st.markdown('<p class="main-title">🎛️ SIMULADOR DE ESCENARIOS & PROYECCIÓN</p>', unsafe_allow_html=True)
    
    # REGLA OBLIGATORIA: GRÁFICO DE LÍNEAS DE PROYECCIÓN HOLT-WINTERS
    st.subheader("📈 Proyección Futura del Costo Unitario (Holt-Winters)")
    meses_p = st.slider("Meses a Proyectar hacia el futuro:", 1, 6, 3)
    
    ts_data = df['Costo Huevo Fértil'].values
    model = ExponentialSmoothing(ts_data, trend='add', seasonal=None).fit()
    forecast = model.forecast(meses_p)
    
    ult_fecha = pd.to_datetime(df['Periodo'].iloc[-1])
    fechas_f = [(ult_fecha + pd.DateOffset(months=i)).strftime('%Y-%m') for i in range(1, meses_p + 1)]
    df_p = pd.DataFrame({'Periodo': fechas_f, 'Costo Proyectado': forecast})
    
    fig_p = go.Figure()
    fig_p.add_trace(go.Scatter(x=df['Periodo'], y=df['Costo Huevo Fértil'], mode='lines+markers', name='Histórico Real', line=dict(color='#1E3A8A', width=3)))
    fig_p.add_trace(go.Scatter(x=df_p['Periodo'], y=df_p['Costo Proyectado'], mode='lines+markers', name='Proyección Futura', line=dict(color='#d62728', width=3, dash='dash')))
    st.plotly_chart(fig_p, use_container_width=True)
    
    st.markdown("---")
    mes_ref = st.selectbox("Seleccionar Mes Base para Simular:", df['Periodo'].tolist(), index=len(df)-1)
    df_ref = df[df['Periodo'] == mes_ref].iloc[0]
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        precio_alimento_sim = st.slider("Precio Kg Alimento ($ COP):", 1200.0, 2200.0, float(df_ref['Precio Kg Alimento']), 10.0)
    with col2:
        gramos_huevo_sim = st.slider("Consumo (g Alimento/Huevo):", 200.0, 320.0, float(df_ref['Gramos Alimento/Huevo']), 1.0)
    with col3:
        fertilidad_sim = st.slider("Fertilidad (%):", 80.0, 99.0, 90.0, 0.5)
    with col4:
        precio_venta_target = st.number_input("Precio Venta Target ($):", value=1600.0, step=10.0)

    prod_total = df_ref['Total Producción']
    huevos_fert_sim = prod_total * (fertilidad_sim / 100.0)
    kg_alimento_sim = (prod_total * gramos_huevo_sim) / 1000.0
    costo_alimento_sim = kg_alimento_sim * precio_alimento_sim
    
    costo_total_sim = (df_ref['Costo Total'] - df_ref['Alimento']) + costo_alimento_sim
    costo_huevo_sim = costo_total_sim / huevos_fert_sim
    costo_huevo_orig = df_ref['Costo Huevo Fértil']
    margen_unitario = precio_venta_target - costo_huevo_sim
    margen_pct = (margen_unitario / precio_venta_target) * 100
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Costo Huevo Simulado", f"${costo_huevo_sim:,.2f}", delta=f"{costo_huevo_sim - costo_huevo_orig:+,.2f} COP", delta_color="inverse")
    kpi2.metric("Precio Venta Target", f"${precio_venta_target:,.2f}")
    kpi3.metric("Margen Unitario ($)", f"${margen_unitario:,.2f}", delta=f"{margen_pct:.1f}% Margen")
    kpi4.metric("Utilidad Total Est.", f"${(margen_unitario * huevos_fert_sim):,.0f}")


# -----------------------------------------------------------------------------
# MENÚ 8: CENTRO DE EXPORTACIÓN DE INFORMES
# -----------------------------------------------------------------------------
elif menu == "8. Centro de Exportación de Informes":
    st.markdown('<p class="main-title">📥 CENTRO DE EXPORTACIÓN DE INFORMES</p>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Exportar Matriz a Excel")
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Matriz_Costos_Historica', index=False)
        buffer_excel.seek(0)
        
        st.download_button(
            label="💾 Descargar Excel Estructurado",
            data=buffer_excel,
            file_name="Matriz_Costos_Huevo_Fertil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.subheader("📄 Generar Reporte Oficial PDF")
        st.info("Exporta un informe consolidado listo para la Junta Directiva.")
