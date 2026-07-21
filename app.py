import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# Configuración de página
st.set_page_config(
    page_title="Costo Huevo Fértil",
    page_icon="🥚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 1. CARGA Y PROCESAMIENTO DE DATOS
# -----------------------------------------------------------------------------
@st.cache_data
def load_data():
    filepath = 'data/Juan_costo_del_huevo.xlsx'  # Asegúrate de poner la ruta correcta
    raw_df = pd.read_excel(filepath, sheet_name='Hoja1', header=None)
    
    # Mapeo de columnas por fecha
    fechas = []
    col_indices = []
    for col in range(2, raw_df.shape[1]):
        val = str(raw_df.iloc[0, col])
        if '202' in val:
            fechas.append(pd.to_datetime(val).strftime('%Y-%m'))
            col_indices.append(col)
            
    # Conceptos de costos (Filas exactas del Excel)
    conceptos_filas = {
        'Arriendo': 1,
        'Cama - Cascarilla': 2,
        'Depreciacion Const. Y Edif.': 4,
        'Depreciacion Huevo': 5,
        'Droga': 6,
        'Materia prima (calcio)': 7,
        'Alimento': 8,
        'Mano de Obra': 9,
        'Aseo y desinfección': 10,
        'Costos Indirectos': 11,
        'Aprovechamientos (-)': 12
    }
    
    # Extraer matriz de costos por mes
    data_costos = []
    for idx, fecha in enumerate(fechas):
        col = col_indices[idx]
        fila_mes = {'Periodo': fecha}
        for item, r_idx in conceptos_filas.items():
            val = raw_df.iloc[r_idx, col]
            fila_mes[item] = float(val) if pd.notnull(val) else 0.0
            
        fila_mes['Costo Total'] = float(raw_df.iloc[13, col])
        fila_mes['Huevos Fértiles'] = float(raw_df.iloc[14, col])
        fila_mes['Total Producción'] = float(raw_df.iloc[15, col])
        fila_mes['Costo Huevo Fértil'] = float(raw_df.iloc[16, col])
        fila_mes['Consumo Alimento Kg'] = float(raw_df.iloc[17, col])
        fila_mes['Gramos Alimento/Huevo'] = float(raw_df.iloc[18, col])
        fila_mes['Precio Kg Alimento'] = float(raw_df.iloc[20, col])
        
        data_costos.append(fila_mes)
        
    df = pd.DataFrame(data_costos)
    return df, list(conceptos_filas.keys())

try:
    df, rubros_items = load_data()
except Exception as e:
    st.error(f"Error cargando el archivo de datos: {e}")
    st.stop()

# -----------------------------------------------------------------------------
# SIDEBAR / NAVEGACIÓN
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Gestión Avícola")
menu = st.sidebar.radio(
    "Selecciona un Módulo:",
    ["1. Análisis Variación (VIF Waterfall)", "2. Impacto por Item", "3. Simulador de Escenarios", "4. Proyección Futura"]
)

# -----------------------------------------------------------------------------
# MÓDULO 1: ANÁLISIS DE VARIACIÓN DE COSTO (WHY IT WENT UP)
# -----------------------------------------------------------------------------
if menu == "1. Análisis Variación (VIF Waterfall)":
    st.header("🔍 Explicación de Variación de Costo Intermensual")
    st.markdown("Analiza la razón exacta por la cual subió o bajó el costo por huevo fértil entre dos períodos seleccionados.")
    
    col1, col2 = st.columns(2)
    with col1:
        mes_base = st.selectbox("Período Base (Comparar contra):", df['Periodo'].tolist(), index=0)
    with col2:
        mes_analisis = st.selectbox("Período de Análisis:", df['Periodo'].tolist(), index=len(df)-1)
        
    if mes_base == mes_analisis:
        st.warning("Selecciona dos meses diferentes para analizar la variación.")
    else:
        df_base = df[df['Periodo'] == mes_base].iloc[0]
        df_analisis = df[df['Periodo'] == mes_analisis].iloc[0]
        
        # Calcular impacto unitario ($/huevo) de cada ítem
        huevo_base = df_base['Huevos Fértiles']
        huevo_analisis = df_analisis['Huevos Fértiles']
        
        costo_unit_base = df_base['Costo Huevo Fértil']
        costo_unit_analisis = df_analisis['Costo Huevo Fértil']
        
        variacion_total = costo_unit_analisis - costo_unit_base
        
        # Desglose de cambios por rubro expresado en $/huevo fértil
        impactos = {}
        for item in rubros_items:
            unit_b = df_base[item] / huevo_base
            unit_a = df_analisis[item] / huevo_analisis
            impactos[item] = unit_a - unit_b
            
        # Preparar gráfico Waterfall
        items_waterfall = list(impactos.keys())
        valores_waterfall = list(impactos.values())
        
        fig = go.Figure(go.Waterfall(
            name = "Variación", orientation = "v",
            measure = ["absolute"] + ["relative"] * len(items_waterfall) + ["total"],
            x = [f"Costo {mes_base}"] + items_waterfall + [f"Costo {mes_analisis}"],
            textposition = "outside",
            text = [f"${costo_unit_base:,.1f}"] + [f"{v:+.1f}" for v in valores_waterfall] + [f"${costo_unit_analisis:,.1f}"],
            y = [costo_unit_base] + valores_waterfall + [0],
            connector = {"line":{"color":"rgb(63, 63, 63)"}},
            decreasing = {"marker":{"color":"#2ca02c"}},
            increasing = {"marker":{"color":"#d62728"}},
            totals = {"marker":{"color":"#1f77b4"}}
        ))
        
        fig.update_layout(title=f"Descomposición de Variación: ${costo_unit_base:,.2f} ➔ ${costo_unit_analisis:,.2f} COP/Huevo (Variación: ${variacion_total:+,.2f})", height=550)
        st.plotly_chart(fig, use_container_width=True)
        
        # Resumen de principales culpables
        df_imp = pd.DataFrame(list(impactos.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values(by='Impacto $/Huevo', ascending=False)
        st.subheader("Top Factores que Encarecieron el Huevo:")
        st.dataframe(df_imp.head(3).style.format({'Impacto $/Huevo': '${:+.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MÓDULO 2: OPCIÓN DE IMPACTO POR ITEM
# -----------------------------------------------------------------------------
elif menu == "2. Impacto por Item":
    st.header("📊 Estructura de Costos y Matriz de Impacto")
    mes_sel = st.selectbox("Selecciona Período:", df['Periodo'].tolist(), index=len(df)-1)
    
    df_mes = df[df['Periodo'] == mes_sel].iloc[0]
    total_huevos = df_mes['Huevos Fértiles']
    
    costos_list = []
    for item in rubros_items:
        monto = df_mes[item]
        unit = monto / total_huevos
        pct = (monto / df_mes['Costo Total']) * 100
        costos_list.append({'Rubro': item, 'Monto Total ($)': monto, 'Costo por Huevo ($)': unit, 'Participación (%)': pct})
        
    df_rubros = pd.DataFrame(costos_list).sort_values(by='Monto Total ($)', ascending=False)
    
    col1, col2 = st.columns([3, 2])
    with col1:
        fig = px.bar(df_rubros, x='Rubro', y='Costo por Huevo ($)', text_auto='.2f', color='Participación (%)',
                     title=f"Costo Unitario por Rubro - {mes_sel}", color_continuous_scale='Reds')
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        fig_pie = px.pie(df_rubros, names='Rubro', values='Monto Total ($)', title=f"Distribución del Gasto - {mes_sel}", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
        
    st.dataframe(df_rubros.style.format({'Monto Total ($)': '${:,.0f}', 'Costo por Huevo ($)': '${:,.2f}', 'Participación (%)': '{:.2f}%'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MÓDULO 3: SIMULADOR DE ESCENARIOS Y MARGEN (WHAT-IF)
# -----------------------------------------------------------------------------
elif menu == "3. Simulador de Escenarios":
    st.header("🎛️ Simulador de Sensibilidad y Nuevo Margen")
    st.markdown("Modifica los parámetros operativos para calcular en tiempo real el nuevo costo del huevo y el margen esperado.")
    
    mes_ref = st.selectbox("Seleccionar Mes de Referencia para Simulación:", df['Periodo'].tolist(), index=len(df)-1)
    df_ref = df[df['Periodo'] == mes_ref].iloc[0]
    
    st.subheader("1. Parámetros Simulados")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        precio_alimento_sim = st.slider("Precio Kg Alimento ($ COP):", 1200.0, 2200.0, float(df_ref['Precio Kg Alimento']), 10.0)
    with col2:
        gramos_huevo_sim = st.slider("Consumo (g Alimento/Huevo):", 200.0, 320.0, float(df_ref['Gramos Alimento/Huevo']), 1.0)
    with col3:
        fertilidad_sim = st.slider("Fertilidad (%):", 80.0, 99.0, float((df_ref['Huevos Fértiles']/df_ref['Total Producción'])*100), 0.5)
    with col4:
        precio_venta_target = st.number_input("Precio Venta Huevo Fértil ($):", value=1600.0, step=10.0)

    # RECÁLCULO DEL ESCENARIO
    prod_total = df_ref['Total Producción']
    huevos_fert_sim = prod_total * (fertilidad_sim / 100.0)
    
    # Recálculo de costo de alimento
    kg_alimento_sim = (prod_total * gramos_huevo_sim) / 1000.0
    costo_alimento_sim = kg_alimento_sim * precio_alimento_sim
    
    # Otros costos constantes
    otros_costos = df_ref['Costo Total'] - df_ref['Alimento']
    costo_total_sim = otros_costos + costo_alimento_sim
    
    costo_huevo_sim = costo_total_sim / huevos_fert_sim
    costo_huevo_orig = df_ref['Costo Huevo Fértil']
    
    margen_unitario = precio_venta_target - costo_huevo_sim
    margen_pct = (margen_unitario / precio_venta_target) * 100
    
    st.markdown("---")
    st.subheader("2. Resultados del Escenario Simulado")
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Costo Huevo Simulado", f"${costo_huevo_sim:,.2f}", delta=f"{costo_huevo_sim - costo_huevo_orig:+,.2f} COP", delta_color="inverse")
    kpi2.metric("Precio Venta Proyectado", f"${precio_venta_target:,.2f}")
    kpi3.metric("Margen Unitario ($)", f"${margen_unitario:,.2f}", delta=f"{margen_pct:.1f}% Margen")
    kpi4.metric("Utilidad Total Proyectada", f"${(margen_unitario * huevos_fert_sim):,.0f}")

# -----------------------------------------------------------------------------
# MÓDULO 4: SLIDE DE PROYECIÓN FUTURA
# -----------------------------------------------------------------------------
elif menu == "4. Proyección Futura":
    st.header("🔮 Proyección Futura (Siguientes Meses)")
    st.markdown("Modelado de series de tiempo para estimar la tendencia del costo por huevo fértil en los próximos períodos.")
    
    meses_proyeccion = st.slider("Meses a Proyectar hacia el futuro:", 1, 6, 3)
    
    # Modelo Holt-Winters / Exponential Smoothing
    ts_data = df['Costo Huevo Fértil'].values
    model = ExponentialSmoothing(ts_data, trend='add', seasonal=None).fit()
    forecast = model.forecast(meses_proyeccion)
    
    # Fechas futuras
    ult_fecha = pd.to_datetime(df['Periodo'].iloc[-1])
    fechas_futuras = [ (ult_fecha + pd.DateOffset(months=i)).strftime('%Y-%m') for i in range(1, meses_proyeccion + 1) ]
    
    df_proj = pd.DataFrame({'Periodo': fechas_futuras, 'Costo Proyectado': forecast})
    
    # Gráfico de Líneas con Proyección
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['Periodo'], y=df['Costo Huevo Fértil'], mode='lines+markers', name='Histórico Real', line=dict(color='#1f77b4', width=3)))
    fig.add_trace(go.Scatter(x=df_proj['Periodo'], y=df_proj['Costo Proyectado'], mode='lines+markers', name='Proyección Futura', line=dict(color='#d62728', width=3, dash='dash')))
    
    fig.update_layout(title="Proyección de Costo por Huevo Fértil ($ COP)", xaxis_title="Período", yaxis_title="Costo/Huevo ($)", height=500)
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Valores Estimados de la Proyección:")
    st.dataframe(df_proj.style.format({'Costo Proyectado': '${:,.2f}'}), use_container_width=True)
