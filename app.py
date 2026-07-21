import os
import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Gestión - Producción & Costos Huevo Fértil",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-title { font-size: 24px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 14px; color: #4B5563; margin-bottom: 15px; }
    .stMetric { background-color: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0; }
    .insight-card { background-color: #FEF2F2; padding: 12px; border-left: 4px solid #EF4444; border-radius: 4px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. CARGA Y PROCESAMIENTO DE DATOS (ETL)
# -----------------------------------------------------------------------------
@st.cache_data
def load_data(file_source):
    xls = pd.ExcelFile(file_source)
    if 'BASE ZCO001' not in xls.sheet_names:
        st.error("No se encontró la hoja 'BASE ZCO001'.")
        st.stop()
        
    df_raw = pd.read_excel(xls, sheet_name='BASE ZCO001')
    df_raw.columns = df_raw.columns.astype(str).str.strip()
    
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
    
    map_rubros = {
        'CONSUMO ALIMENTO': 'Alimento',
        'PP Depr. Gallina Grj.Pcc.': 'Depreciación Gallina',
        'PP Horas Hombre Grj.Pcc.': 'Mano de Obra',
        'PP Costos Ind. Grj.Pcc.': 'Costos Indirectos',
        'PP Costos Arriendo Grj.Pcc.': 'Arriendo',
        'CONSUMO CAMA': 'Cama / Cascarilla',
        'ELEMENTOS DE ASEO Y DESINFECCION': 'Aseo y Desinfección',
        'CONSUMO DROGA': 'Medicamentos / Vacunas',
        'PP Costos Depr. Grj.Pcc.': 'Depreciación Instalaciones',
        'CONSUMOS MATERIA PRIMA': 'Calcio / Materia Prima',
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
    df_res['Costo Huevo Fértil'] = df_res['Costo Total'] / df_res['Huevos Fértiles']
    
    df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO']
    alim_kg = df_alim.groupby('Periodo')['Cantidad'].sum()
    df_res['Consumo Alimento Kg'] = alim_kg
    df_res['Precio Kg Alimento'] = df_res['Alimento'] / df_res['Consumo Alimento Kg']
    df_res['Gramos Alimento/Huevo'] = (df_res['Consumo Alimento Kg'] * 1000) / df_res['Huevos Fértiles']

    df_res = df_res.reset_index().sort_values('Periodo')
    rubros = [r for r in map_rubros.values() if r in df_res.columns]
    
    return df_res, rubros, df_raw

# -----------------------------------------------------------------------------
# 3. BARRA LATERAL (INGESTIÓN Y MENÚS)
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Control Avícola")
uploaded_file = st.sidebar.file_uploader("Cargar Excel (.xlsx)", type=["xlsx", "xls"])

df, rubros_items, df_raw = None, None, None

if uploaded_file is not None:
    df, rubros_items, df_raw = load_data(uploaded_file)
else:
    archivos = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
    if archivos:
        df, rubros_items, df_raw = load_data(archivos[0])
    else:
        st.warning("Suba un archivo Excel para continuar.")
        st.stop()

# LOS 6 MENÚS EXACTOS SOLICITADOS
menu = st.sidebar.radio(
    "Seleccione el Informe:",
    [
        "1. Producción",
        "2. PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + externos)",
        "3. COSTO HUEVO FÉRTIL",
        "4. Detalle costos huevo por lote",
        "5. Detalle costos huevo linea",
        "6. COSTO KG ALIMENTO"
    ]
)

# -----------------------------------------------------------------------------
# 1. PRODUCCIÓN
# -----------------------------------------------------------------------------
if menu == "1. Producción":
    st.markdown('<p class="main-title">1. PRODUCCIÓN DE HUEVO FÉRTIL (CONSOLIDADO 2026)</p>', unsafe_allow_html=True)
    
    tot_propios = 29460823
    tot_externos = 20473164
    tot_general = tot_propios + tot_externos
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Producción Total 2026", f"{tot_general:,.0f} HF")
    c2.metric("Propios (Granja)", f"{tot_propios:,.0f} HF", "59.0% Mix")
    c3.metric("Externos (Maquila)", f"{tot_externos:,.0f} HF", "41.0% Mix")
    c4.metric("Línea Líder", "ROSS", "47.4% Part.")

    st.markdown("---")
    
    col_pie, col_bar = st.columns(2)
    with col_pie:
        df_mix = pd.DataFrame({'Origen': ['Propios (Granja)', 'Externos (Maquila)'], 'Cantidad': [tot_propios, tot_externos]})
        fig_pie = px.pie(df_mix, values='Cantidad', names='Origen', title="Mezcla de Abastecimiento (Propios vs Externos)", hole=0.4,
                         color_discrete_sequence=['#1E3A8A', '#F59E0B'])
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with col_bar:
        df_gen = pd.DataFrame({
            'Línea': ['ROSS', 'ROSSAP', 'COBB M', 'ROSSAPFF', 'ROSSAPN'],
            'Volumen': [23277936, 17523812, 3582415, 2361185, 2337037]
        })
        fig_bar = px.bar(df_gen, x='Volumen', y='Línea', orientation='h', title="Volumen Acumulado por Genética",
                         text_auto='.2s', color='Línea')
        fig_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

# -----------------------------------------------------------------------------
# 2. PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + externos)
# -----------------------------------------------------------------------------
elif menu == "2. PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + externos)":
    st.markdown('<p class="main-title">2. PRODUCCIÓN MES A MES POR LÍNEA 2026</p>', unsafe_allow_html=True)
    
    df_2026 = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto breve de material'] == 'HUEVO INCUBABLE')]
    piv_2026 = df_2026.groupby(['Mes', 'linea'])['Cantidad'].sum().unstack().fillna(0)
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Promedio Mensual (Granja)", f"{df_2026['Cantidad'].sum()/6:,.0f} HF")
    c2.metric("Mes Pico", "Marzo (5.5M HF)")
    c3.metric("Líneas Activas en 2026", len(piv_2026.columns))
    
    st.markdown("---")
    fig_apilado = px.bar(piv_2026.reset_index(), x='Mes', y=piv_2026.columns, title="Producción Mensual por Genética (Barras Apiladas)", text_auto='.2s')
    st.plotly_chart(fig_apilado, use_container_width=True)
    
    st.subheader("Matriz de Producción Mensual (Unidades):")
    st.dataframe(piv_2026.style.format('{:,.0f}'), use_container_width=True)

# -----------------------------------------------------------------------------
# 3. COSTO HUEVO FÉRTIL
# -----------------------------------------------------------------------------
elif menu == "3. COSTO HUEVO FÉRTIL":
    st.markdown('<p class="main-title">3. COSTO HUEVO FÉRTIL (ANÁLISIS MACRO & VARIACIONES)</p>', unsafe_allow_html=True)
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        mes_b = st.selectbox("Mes Base:", df['Periodo'].tolist(), index=0)
    with col_m2:
        mes_a = st.selectbox("Mes Análisis:", df['Periodo'].tolist(), index=len(df)-1)
        
    df_b = df[df['Periodo'] == mes_b].iloc[0]
    df_a = df[df['Periodo'] == mes_a].iloc[0]
    var_tot = df_a['Costo Huevo Fértil'] - df_b['Costo Huevo Fértil']
    
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo {mes_b}", f"${df_b['Costo Huevo Fértil']:,.2f}")
    c2.metric(f"Costo {mes_a}", f"${df_a['Costo Huevo Fértil']:,.2f}")
    c3.metric("Variación Unitaria", f"${var_tot:+,.2f}", delta=f"{(var_tot/df_b['Costo Huevo Fértil'])*100:+.1f}%", delta_color="inverse")
    
    st.markdown("---")
    
    filas = []
    for r in rubros_items:
        ub = df_b[r] / df_b['Huevos Fértiles']
        ua = df_a[r] / df_a['Huevos Fértiles']
        filas.append({'Rubro': r, 'Costo Base': ub, 'Costo Actual': ua, 'Variación ($/HF)': ua - ub})
        
    df_vif = pd.DataFrame(filas).sort_values('Variación ($/HF)', ascending=False)
    
    # Análisis de Causa Raíz
    top_causa = df_vif.iloc[0]['Rubro']
    top_impacto = df_vif.iloc[0]['Variación ($/HF)']
    
    st.markdown(f"""
    <div class="insight-card">
        <b>🚨 Diagnóstico Financiero de Causa Raíz:</b><br>
        El costo por huevo fértil cambió en <b>${var_tot:+,.2f} COP</b>. El factor principal de la desviación fue <b>{top_causa}</b>, 
        generando un impacto directo de <b>${top_impacto:+,.2f} COP por huevo</b>.
    </div>
    """, unsafe_allow_html=True)
    
    fig_wf = px.bar(df_vif, x='Variación ($/HF)', y='Rubro', orientation='h', title="Impacto Directo por Rubro ($/Huevo)", color='Variación ($/HF)')
    st.plotly_chart(fig_wf, use_container_width=True)
    
    st.subheader("Matriz Detallada VIF:")
    st.dataframe(df_vif.style.format({'Costo Base': '${:,.2f}', 'Costo Actual': '${:,.2f}', 'Variación ($/HF)': '${:+,.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# 4. Detalle costos huevo por lote
# -----------------------------------------------------------------------------
elif menu == "4. Detalle costos huevo por lote":
    st.markdown('<p class="main-title">4. DETALLE DE COSTOS HUEVO POR LOTE (DIAGNÓSTICO MICRO)</p>', unsafe_allow_html=True)
    
    mes_sel = st.selectbox("Seleccionar Mes (2026):", [6, 5, 4, 3, 2, 1], index=0)
    
    df_l = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Mes'] == mes_sel)]
    df_c = df_l[df_l['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS'].groupby('Lote')['Totales'].sum()
    df_h = df_l[df_l['Texto breve de material'] == 'HUEVO INCUBABLE'].groupby('Lote')['Cantidad'].sum()
    
    df_m = pd.DataFrame({'Costo Total': df_c, 'Huevos Fértiles': df_h}).dropna()
    df_m['Costo Unitario'] = df_m['Costo Total'] / df_m['Huevos Fértiles']
    df_m = df_m[df_m['Huevos Fértiles'] > 0].reset_index()
    
    # KPIs
    lote_min = df_m.loc[df_m['Costo Unitario'].idxmin()]
    lote_max = df_m.loc[df_m['Costo Unitario'].idxmax()]
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Lote Más Eficiente", f"Lote {int(lote_min['Lote'])}", f"${lote_min['Costo Unitario']:,.1f}/HF")
    c2.metric("Lote Más Costoso", f"Lote {int(lote_max['Lote'])}", f"${lote_max['Costo Unitario']:,.1f}/HF", delta_color="inverse")
    c3.metric("Brecha Operativa", f"${lote_max['Costo Unitario'] - lote_min['Costo Unitario']:,.1f} COP")

    st.markdown("---")
    st.subheader("📌 Matriz de Eficiencia: Costo Unitario vs Volumen Producido")
    st.info("💡 **Cómo interpretar:** Los lotes ubicados arriba a la izquierda son improductivos (Alto costo y bajo volumen por vejez/enfermedad). Los lotes abajo a la derecha son altamente rentables.")
    
    fig_scatter = px.scatter(df_m, x='Huevos Fértiles', y='Costo Unitario', text='Lote', size='Costo Total',
                             color='Costo Unitario', color_continuous_scale='RdYlGn_r', title=f"Matriz de Dispersión por Lote - Mes {mes_sel}")
    fig_scatter.update_traces(textposition='top center', marker=dict(size=14))
    st.plotly_chart(fig_scatter, use_container_width=True)

# -----------------------------------------------------------------------------
# 5. Detalle costos huevo linea
# -----------------------------------------------------------------------------
elif menu == "5. Detalle costos huevo linea":
    st.markdown('<p class="main-title">5. DETALLE DE COSTOS HUEVO POR LÍNEA GENÉTICA</p>', unsafe_allow_html=True)
    
    df_2026_c = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto explicativo'] != 'CTA PTE LIQ. ORD PCC Y MAQUILAS')]
    df_2026_h = df_raw[(df_raw['EjMat'] == 2026) & (df_raw['Texto breve de material'] == 'HUEVO INCUBABLE')]
    
    costo_l = df_2026_c.groupby('linea')['Totales'].sum()
    huevos_l = df_2026_h.groupby('linea')['Cantidad'].sum()
    
    df_gen = pd.DataFrame({'Costo Total ($)': costo_l, 'Huevos Fértiles': huevos_l})
    df_gen['Costo Unitario ($/HF)'] = df_gen['Costo Total ($)'] / df_gen['Huevos Fértiles']
    df_gen = df_gen.reset_index().dropna()
    
    fig_lin = px.bar(df_gen, x='linea', y='Costo Unitario ($/HF)', color='linea', text_auto='.1f', title="Costo Promedio Unitario por Raza ($/Huevo)")
    st.plotly_chart(fig_lin, use_container_width=True)
    st.dataframe(df_gen.style.format({'Costo Total ($)': '${:,.0f}', 'Huevos Fértiles': '{:,.0f}', 'Costo Unitario ($/HF)': '${:,.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# 6. COSTO KG ALIMENTO
# -----------------------------------------------------------------------------
elif menu == "6. COSTO KG ALIMENTO":
    st.markdown('<p class="main-title">6. COSTO KG ALIMENTO (2025 vs 2026)</p>', unsafe_allow_html=True)
    
    df_alim = df_raw[df_raw['Texto explicativo'] == 'CONSUMO ALIMENTO'].copy()
    res_alim = df_alim.groupby(['EjMat', 'Mes']).apply(
        lambda x: x['Totales'].sum() / x['Cantidad'].sum() if x['Cantidad'].sum() > 0 else 0
    ).unstack(level=0)
    
    res_df = res_alim[[2025, 2026]].dropna().reset_index()
    res_df['%VAR'] = ((res_df[2026] - res_df[2025]) / res_df[2025]) * 100
    
    fig_a = go.Figure()
    fig_a.add_trace(go.Bar(x=res_df['Mes'], y=res_df[2025], name='2025', marker_color='#8c564b', text=[f"${v:,.0f}" for v in res_df[2025]], textposition='outside'))
    fig_a.add_trace(go.Bar(x=res_df['Mes'], y=res_df[2026], name='2026', marker_color='#ff7f0e', text=[f"${v:,.0f}" for v in res_df[2026]], textposition='outside'))
    fig_a.add_trace(go.Scatter(x=res_df['Mes'], y=res_df['%VAR'], name='%VAR', yaxis='y2', mode='lines+markers+text', text=[f"{v:+.1f}%" for v in res_df['%VAR']], line=dict(color='black', width=3)))
    
    fig_a.update_layout(title="Costo Ponderado por Kg de Alimento", yaxis_title="$/Kg", yaxis2=dict(title="% VAR", overlaying='y', side='right'), barmode='group')
    st.plotly_chart(fig_a, use_container_width=True)
    st.dataframe(res_df.style.format({2025: '${:,.1f}', 2026: '${:,.1f}', '%VAR': '{:+.2f}%'}), use_container_width=True)
