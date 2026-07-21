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

# Estilos CSS personalizados (Corregido: unsafe_allow_html=True)
st.markdown("""
    <style>
    .main-title { font-size: 28px; font-weight: bold; color: #1E3A8A; margin-bottom: 10px; }
    .subtitle { font-size: 16px; color: #4B5563; margin-bottom: 20px; }
    .kpi-card { background-color: #F3F4F6; padding: 15px; border-radius: 8px; border-left: 5px solid #1E3A8A; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. CARGA Y PROCESAMIENTO DINÁMICO DE DATOS (ETL & INGESTIÓN)
# -----------------------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_source):
    """
    Función de Extracción, Transformación y Carga (ETL).
    Lee la hoja 'BASE ZCO001' del archivo Excel cargado y estandariza las 
    variables financieras, zootécnicas y la columna temporal 'Fecha'.
    """
    xls = pd.ExcelFile(file_source)
    
    if 'BASE ZCO001' in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name='BASE ZCO001')
        df_raw.columns = df_raw.columns.astype(str).str.strip()
        
        # Procesamiento de Fechas y Períodos
        if 'Fecha' in df_raw.columns:
            df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'])
            df_raw['Periodo'] = df_raw['Fecha'].dt.strftime('%Y-%m')
        else:
            df_raw['Periodo'] = df_raw['EjMat'].astype(str) + '-' + df_raw['Mes'].astype(str).str.zfill(2)
            
        df_raw['Totales'] = pd.to_numeric(df_raw['Totales'], errors='coerce').fillna(0)
        df_raw['Cantidad'] = pd.to_numeric(df_raw['Cantidad'], errors='coerce').fillna(0)
        
        # Mapeo de Estandarización de Rubros Financieros
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
        
        # Extracción del volumen de Huevos Fértiles (Material: HUEVO INCUBABLE)
        df_hf = df_raw[df_raw['Texto breve de material'] == 'HUEVO INCUBABLE']
        hf_mes = df_hf.groupby('Periodo')['Cantidad'].sum()
        
        # Filtrado de Costos excluyendo cuentas puente/liquidadoras intercompany
        df_costos = df_raw[df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].copy()
        df_costos['Rubro'] = df_costos['Texto explicativo'].map(lambda x: map_rubros.get(x, x))
        
        # Matriz Pivote de Costos Totales por Período
        costos_piv = df_costos.groupby(['Periodo', 'Rubro'])['Totales'].sum().unstack(fill_value=0)
        
        # Cálculo de Indicadores Clave de Desempeño (KPIs)
        df_res = costos_piv.copy()
        df_res['Costo Total'] = df_res.sum(axis=1)
        df_res['Huevos Fértiles'] = hf_mes
        df_res['Total Producción'] = hf_mes
        df_res['Costo Huevo Fértil'] = df_res['Costo Total'] / df_res['Huevos Fértiles']
        
        # Análisis Nutricional y Consumo de Alimento
        df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO']
        alim_kg = df_alim.groupby('Periodo')['Cantidad'].sum()
        df_res['Consumo Alimento Kg'] = alim_kg
        df_res['Precio Kg Alimento'] = df_res['Alimento'] / df_res['Consumo Alimento Kg']
        df_res['Gramos Alimento/Huevo'] = (df_res['Consumo Alimento Kg'] * 1000) / df_res['Huevos Fértiles']

        df_res = df_res.reset_index().sort_values('Periodo')
        rubros = [r for r in map_rubros.values() if r in df_res.columns]
        
        return df_res, rubros, df_raw
    else:
        st.error("⚠️ Estructura no válida: La hoja 'BASE ZCO001' no fue encontrada en el archivo Excel.")
        st.stop()

# -----------------------------------------------------------------------------
# 3. CONTROL DE NAVEGACIÓN Y CARGA DE ARCHIVOS (SIDEBAR)
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
    # Carga automática por defecto si existe un archivo local
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

# Selección del Informe
menu = st.sidebar.radio(
    "Selecciona el Reporte Ejecutivo:",
    [
        "1. Producción 2026 (Propios + Externos)",
        "2. Producción Mes a Mes por Línea",
        "3. Costo Huevo Fértil (Macro)",
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
    st.markdown('<p class="subtitle">Análisis consolidado de la oferta de huevo fértil por fuente de abastecimiento y genética.</p>', unsafe_allow_html=True)
    
    with st.expander("💡 **Interpretación Gerencial & Diagnóstico:**", expanded=True):
        st.write("""
        * **Autonomía de Producción:** La compañía mantiene un **59.0% de producción propia en granjas** ($29.46\,\text{M}$ de unidades) y recurre a **maquilas y compras externas en un 41.0%** ($20.47\,\text{M}$ de unidades) para cubrir el programa de incubación.
        * **Concentración de Genética:** Las razas **ROSS** ($47.4\%$) y **ROSSAP** ($35.7\%$) representan más del **83.1% del volumen total**, siendo las genéticas principales que sostienen la operación.
        """)

    tot_propios = 29460823
    tot_externos = 20473164
    tot_general = tot_propios + tot_externos
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Producción Total 2026", f"{tot_general:,.0f} HF")
    kpi2.metric("Propios (Granja)", f"{tot_propios:,.0f} HF", delta="59.0% Participación")
    kpi3.metric("Externos (Maquila)", f"{tot_externos:,.0f} HF", delta="41.0% Participación")
    kpi4.metric("Línea Genética Líder", "ROSS", "47.43% del Mercado")

    st.markdown("---")
    
    col_chart, col_data = st.columns([3, 2])
    
    with col_chart:
        df_part = pd.DataFrame({
            'Línea': ['ROSS', 'ROSSAP', 'COBB M', 'ROSSAPFF', 'ROSSAPN'],
            'Total 2026': [23277936, 17523812, 3582415, 2361185, 2337037],
            '% Part': [47.43, 35.70, 7.30, 4.81, 4.76]
        })
        
        fig_gen = px.bar(
            df_part, x='Línea', y='Total 2026', text_auto='.2s', color='Línea',
            title="Distribución Total del Volumen por Línea Genética (2026)"
        )
        for idx, row in df_part.iterrows():
            fig_gen.add_annotation(
                x=row['Línea'], y=row['Total 2026'] * 0.15,
                text=f"<b>{row['% Part']:.2f}%</b>", showarrow=False,
                bgcolor="white", bordercolor="black"
            )
        fig_gen.update_layout(height=450, showlegend=False, yaxis_title="Huevos Fértiles")
        st.plotly_chart(fig_gen, use_container_width=True)
        
    with col_data:
        st.subheader("📌 Matriz de Participación por Genética")
        st.dataframe(df_part.style.format({'Total 2026': '{:,.0f}', '% Part': '{:.2f}%'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 2: PRODUCCIÓN MES A MES POR LÍNEA
# -----------------------------------------------------------------------------
elif menu == "2. Producción Mes a Mes por Línea":
    st.markdown('<p class="main-title">📈 PRODUCCIÓN MES A MES POR LÍNEA 2026</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Evolución del volumen mensual por raza, comparando desempeño en granja.</p>', unsafe_allow_html=True)
    
    df_2026_hf = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto breve de material'] == 'HUEVO INCUBABLE')]
    piv_prod = df_2026_hf.groupby(['Mes', 'linea'])['Cantidad'].sum().unstack().fillna(0)
    
    with st.expander("💡 **Comentario Técnico / Análisis de Tendencia:**"):
        st.write("""
        * **Estabilidad de Oferta:** El volumen mensual en granjas propias oscila entre **4.3M y 5.5M de huevos fértiles**, mostrando un pico positivo en marzo impulsado por la entrada de nuevos lotes en curva pico.
        * **Comportamiento por Raza:** Se observa una sustitución paulatina de `ROSSAP` por `ROSS` a partir del segundo trimestre de 2026.
        """)

    fig_prod = px.bar(piv_prod.reset_index(), x='Mes', y=piv_prod.columns, title="Evolución Mensual de Producción por Línea (Granja)",
                      barmode='group', text_auto='.2s')
    fig_prod.update_layout(xaxis_title="Mes (2026)", yaxis_title="Huevos Fértiles Producidos", height=480)
    st.plotly_chart(fig_prod, use_container_width=True)
    
    st.subheader("Tabla de Datos Consolidada (Unidades):")
    st.dataframe(piv_prod.style.format('{:,.0f}'), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 3: COSTO HUEVO FÉRTIL (MACRO & VARIACIÓN VIF)
# -----------------------------------------------------------------------------
elif menu == "3. Costo Huevo Fértil (Macro)":
    st.markdown('<p class="main-title">🥚 COSTO HUEVO FÉRTIL 2026 (Macro & Variaciones)</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Explicación financiera intermensual de los factores que provocaron variaciones en el costo unitario.</p>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        mes_b = st.selectbox("Período Base (Inicial):", df['Periodo'].tolist(), index=0)
    with col2:
        mes_a = st.selectbox("Período de Análisis (Final):", df['Periodo'].tolist(), index=len(df)-1)
    with col3:
        formato_v = st.radio("Formato Gráfico:", ["Barras de Impacto (Recomendado)", "Waterfall Recortado"])
        
    df_b = df[df['Periodo'] == mes_b].iloc[0]
    df_a = df[df['Periodo'] == mes_a].iloc[0]
    var_tot = df_a['Costo Huevo Fértil'] - df_b['Costo Huevo Fértil']
    
    k1, k2, k3 = st.columns(3)
    k1.metric(f"Costo {mes_b}", f"${df_b['Costo Huevo Fértil']:,.2f}")
    k2.metric(f"Costo {mes_a}", f"${df_a['Costo Huevo Fértil']:,.2f}")
    k3.metric("Variación Total Unitario", f"${var_tot:+,.2f} COP", delta=f"{(var_tot/df_b['Costo Huevo Fértil'])*100:+.1f}%", delta_color="inverse")
    
    st.markdown("---")
    
    impactos = {}
    for r in rubros_items:
        impactos[r] = (df_a[r] / df_a['Huevos Fértiles']) - (df_b[r] / df_b['Huevos Fértiles'])
        
    df_imp = pd.DataFrame(list(impactos.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values('Impacto $/Huevo', ascending=True)

    if formato_v == "Barras de Impacto (Recomendado)":
        fig_imp = px.bar(df_imp, y='Rubro', x='Impacto $/Huevo', orientation='h', color='Impacto $/Huevo',
                         color_continuous_scale='Reds' if var_tot > 0 else 'Greens',
                         title=f"Aporte Directo de cada Rubro al Cambio Total de ${var_tot:+,.2f} COP/Huevo")
        fig_imp.update_traces(text=[f"${v:+,.1f}" for v in df_imp['Impacto $/Huevo']], textposition='outside')
        st.plotly_chart(fig_imp, use_container_width=True)
    else:
        items_wf = list(impactos.keys())
        valores_wf = list(impactos.values())
        min_y = min(df_b['Costo Huevo Fértil'], df_a['Costo Huevo Fértil']) * 0.95
        
        fig_wf = go.Figure(go.Waterfall(
            name="Variación", orientation="v",
            measure=["absolute"] + ["relative"] * len(items_wf) + ["total"],
            x=[f"Base ({mes_b})"] + items_wf + [f"Final ({mes_a})"],
            textposition="outside",
            text=[f"${df_b['Costo Huevo Fértil']:,.1f}"] + [f"${v:+,.1f}" for v in valores_wf] + [f"${df_a['Costo Huevo Fértil']:,.1f}"],
            y=[df_b['Costo Huevo Fértil']] + valores_wf + [0],
            connector={"line":{"color":"gray"}},
            decreasing={"marker":{"color":"#2ca02c"}},
            increasing={"marker":{"color":"#d62728"}},
            totals={"marker":{"color":"#1f77b4"}}
        ))
        fig_wf.update_layout(title="Descomposición en Cascada (Escala Recortada)", yaxis=dict(range=[min_y, max(df_b['Costo Huevo Fértil'], df_a['Costo Huevo Fértil']) * 1.05]), height=500)
        st.plotly_chart(fig_wf, use_container_width=True)

    st.subheader("💡 Explicación Analítica de la Variación:")
    df_top_var = df_imp.sort_values('Impacto $/Huevo', ascending=False)
    for idx, row in df_top_var.head(3).iterrows():
        pct_exp = (row['Impacto $/Huevo'] / var_tot) * 100 if var_tot != 0 else 0
        st.markdown(f"* **{row['Rubro']}**: Aportó **${row['Impacto $/Huevo']:+,.2f} COP** al costo final (explica el **{pct_exp:.1f}%** del cambio total).")

# -----------------------------------------------------------------------------
# MENÚ 4: DETALLE COSTOS HUEVO POR LOTE
# -----------------------------------------------------------------------------
elif menu == "4. Detalle Costos Huevo por Lote":
    st.markdown('<p class="main-title">🔍 DETALLE DE COSTOS HUEVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Diagnóstico micro-operativo a nivel de lote y edad de la ave (Semanas).</p>', unsafe_allow_html=True)
    
    mes_sel = st.selectbox("Seleccionar Mes de Consulta (2026):", [6, 5, 4, 3, 2, 1], index=0)
    
    df_mes_lote = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Mes'] == mes_sel) & (df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS')]
    piv_lote = df_mes_lote.pivot_table(index='Texto explicativo', columns='Lote', values='Totales', aggfunc='sum').fillna(0)
    
    df_hf_lote = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Mes'] == mes_sel) & (df_raw['Texto breve de material'] == 'HUEVO INCUBABLE')]
    hf_lotes = df_hf_lote.groupby('Lote')['Cantidad'].sum()
    
    piv_lote_unit = piv_lote.divide(hf_lotes, axis=1).dropna(axis=1)
    
    with st.expander("💡 **Análisis de Eficiencia por Edad de la Ave:**", expanded=True):
        st.write("""
        * **Efecto de Disminución de Postura por Edad:** Los lotes de más de **60 semanas (ej. Lote 212 o 213)** presentan costos unitarios superiores a **$3,100 - $4,100 COP/huevo** debido a que la amortización de la gallina se distribuye entre menos huevos fértiles.
        * **Punto Óptimo de Postura:** Lotes jóvenes en curva alta (semanas 33 a 44, ej. **Lote 219 o 223**) alcanzan los menores costos unitarios del sistema (**$1,032 - $1,113 COP/huevo**).
        """)

    st.subheader(f"Matriz de Costo Unitario ($/HF) - Mes {mes_sel} / 2026:")
    st.dataframe(piv_lote_unit.style.format('${:,.1f}'), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 5: DETALLE COSTOS HUEVO POR LÍNEA
# -----------------------------------------------------------------------------
elif menu == "5. Detalle Costos Huevo por Línea":
    st.markdown('<p class="main-title">🧬 DETALLE DE COSTOS HUEVO POR LÍNEA GENÉTICA</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Evaluación financiera de la rentabilidad y costo unitario por raza avícola.</p>', unsafe_allow_html=True)
    
    df_2026_c = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS')]
    df_2026_h = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto breve de material'] == 'HUEVO INCUBABLE')]
    
    costo_linea = df_2026_c.groupby('linea')['Totales'].sum()
    huevos_linea = df_2026_h.groupby('linea')['Cantidad'].sum()
    
    df_gen = pd.DataFrame({'Costo Total ($)': costo_linea, 'Huevos Fértiles': huevos_linea})
    df_gen['Costo / Huevo ($)'] = df_gen['Costo Total ($)'] / df_gen['Huevos Fértiles']
    df_gen = df_gen.reset_index().dropna()
    
    fig_gen = px.bar(df_gen, x='linea', y='Costo / Huevo ($)', text_auto='.1f', color='linea',
                     title="Costo Promedio Ponderado por Línea Genética ($ COP / Huevo)")
    st.plotly_chart(fig_gen, use_container_width=True)
    
    st.dataframe(df_gen.style.format({'Costo Total ($)': '${:,.0f}', 'Huevos Fértiles': '{:,.0f}', 'Costo / Huevo ($)': '${:,.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 6: COSTO KG ALIMENTO
# -----------------------------------------------------------------------------
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="main-title">🌾 COSTO KG ALIMENTO (2025 vs 2026)</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Seguimiento al precio ponderado de la materia prima principal ($/Kg).</p>', unsafe_allow_html=True)
    
    df_alim_all = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO'].copy()
    df_alim_all['Anio'] = df_alim_all['EjMat']
    
    res_alim = df_alim_all.groupby(['Anio', 'Mes']).apply(
        lambda x: x['Totales'].sum() / x['Cantidad'].sum() if x['Cantidad'].sum() > 0 else 0
    ).unstack(level=0)
    
    res_alim_df = res_alim[[2025, 2026]].dropna().reset_index()
    res_alim_df['%VAR'] = ((res_alim_df[2026] - res_alim_df[2025]) / res_alim_df[2025]) * 100
    
    with st.expander("💡 **Comentario Nutricional y Financiero:**"):
        st.write("""
        * **Tendencia:** El costo del alimento registró incrementos del **+2.6% a inicio de año ($1,694 COP/Kg)** debido a presiones en la tasa de cambio e importaciones de Maíz/Soya, estabilizándose hacia mediados de año alrededor de los **$1,672 COP/Kg**.
        * **Sensibilidad:** Dado que el alimento representa cerca del **35% - 40% del costo total**, cada $100 COP de incremento por Kg encarece aproximadamente $27.6 COP el huevo fértil producido.
        """)

    fig_alim = go.Figure()
    fig_alim.add_trace(go.Bar(x=res_alim_df['Mes'], y=res_alim_df[2025], name='2025', marker_color='#8c564b', text=[f"${v:,.0f}" for v in res_alim_df[2025]], textposition='outside'))
    fig_alim.add_trace(go.Bar(x=res_alim_df['Mes'], y=res_alim_df[2026], name='2026', marker_color='#ff7f0e', text=[f"${v:,.0f}" for v in res_alim_df[2026]], textposition='outside'))
    fig_alim.add_trace(go.Scatter(x=res_alim_df['Mes'], y=res_alim_df['%VAR'], name='%VAR', yaxis='y2', mode='lines+markers+text', text=[f"{v:+.1f}%" for v in res_alim_df['%VAR']], textposition='top center', line=dict(color='black', width=3)))
    
    fig_alim.update_layout(
        title="Costo Ponderado por Kg de Alimento (2026 vs 2025)",
        xaxis=dict(title="Mes"), yaxis=dict(title="Costo $/Kg Alimento"),
        yaxis2=dict(title="% Variación", overlaying='y', side='right', showgrid=False),
        barmode='group', height=480
    )
    st.plotly_chart(fig_alim, use_container_width=True)
    st.dataframe(res_alim_df.style.format({2025: '${:,.1f}', 2026: '${:,.1f}', '%VAR': '{:+.2f}%'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MENÚ 7: SIMULADOR WHAT-IF & PROYECCIÓN
# -----------------------------------------------------------------------------
elif menu == "7. Simulador What-If & Proyección":
    st.markdown('<p class="main-title">🎛️ SIMULADOR DE ESCENARIOS & PROYECCIÓN</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Herramienta predictiva para modelar sensibilidad operativa y estimar tendencia futura.</p>', unsafe_allow_html=True)
    
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
    
    st.markdown("---")
    st.subheader("Resultados del Escenario Simulado:")
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
    st.markdown('<p class="subtitle">Descarga de la base de datos limpia o generación automatizada de informes en PDF.</p>', unsafe_allow_html=True)
    
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
