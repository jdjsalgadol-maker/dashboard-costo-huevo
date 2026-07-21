import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control Ejecutivo - Costo Huevo Fértil",
    page_icon="🥚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 1. CARGA Y PROCESAMIENTO DINÁMICO DE DATOS
# -----------------------------------------------------------------------------
@st.cache_data
def load_data_from_file(file_source):
    """Lee y limpia la estructura del Excel independientemente de su nombre."""
    raw_df = pd.read_excel(file_source, sheet_name='Hoja1', header=None)
    
    # Mapeo de columnas por fecha en la primera fila
    fechas = []
    col_indices = []
    for col in range(2, raw_df.shape[1]):
        val = str(raw_df.iloc[0, col])
        if '202' in val:
            fechas.append(pd.to_datetime(val).strftime('%Y-%m'))
            col_indices.append(col)
            
    # Conceptos de costos (Filas exactas según la estructura de la hoja)
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

# -----------------------------------------------------------------------------
# CONTROL DE CARGA EN EL SIDEBAR
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Gestión Avícola")
st.sidebar.subheader("📂 Carga de Datos")

# Opción A: Carga interactiva Drag & Drop
uploaded_file = st.sidebar.file_uploader("Sube cualquier archivo Excel (.xlsx / .xls)", type=["xlsx", "xls"])

df = None
rubros_items = None

if uploaded_file is not None:
    try:
        df, rubros_items = load_data_from_file(uploaded_file)
        st.sidebar.success(f"¡Cargado: `{uploaded_file.name}`!")
    except Exception as e:
        st.sidebar.error(f"Error procesando el archivo subido: {e}")
        st.stop()
else:
    # Opción B: Búsqueda automática de cualquier .xlsx disponible en la raíz o /data
    archivos_excel = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
    if os.path.exists('data'):
        archivos_excel += [os.path.join('data', f) for f in os.listdir('data') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]
        
    if archivos_excel:
        archivo_encontrado = archivos_excel[0]
        try:
            df, rubros_items = load_data_from_file(archivo_encontrado)
            st.sidebar.info(f"Cargado automático: `{archivo_encontrado}`")
        except Exception as e:
            st.sidebar.error(f"Error cargando `{archivo_encontrado}`: {e}")
            st.stop()
    else:
        st.warning("⚠️ No se encontró ningún archivo Excel. Por favor, arrastra y suelta tu archivo en la barra lateral.")
        st.stop()

# -----------------------------------------------------------------------------
# NAVEGACIÓN PRINCIPAL
# -----------------------------------------------------------------------------
menu = st.sidebar.radio(
    "Selecciona un Módulo:",
    [
        "1. Análisis Variación del Costo", 
        "2. Impacto por Ítem", 
        "3. Simulador de Escenarios (What-If)", 
        "4. Proyección Futura"
    ]
)

# -----------------------------------------------------------------------------
# MÓDULO 1: ANÁLISIS DE VARIACIÓN INTERMENSUAL (CLARO Y EJECUTIVO)
# -----------------------------------------------------------------------------
if menu == "1. Análisis Variación del Costo":
    st.header("🔍 ¿Por qué subió o bajó el Costo por Huevo Fértil?")
    st.markdown("Compara dos períodos y visualiza claramente los rubros responsables de la variación en pesos COP por huevo.")
    
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        mes_base = st.selectbox("Período Base (Inicial):", df['Periodo'].tolist(), index=0)
    with col2:
        mes_analisis = st.selectbox("Período de Análisis (Final):", df['Periodo'].tolist(), index=len(df)-1)
    with col3:
        tipo_vista = st.radio("Formato de Gráfico:", ["Barras de Impacto (Recomendado)", "Waterfall Ajustado"])
        
    if mes_base == mes_analisis:
        st.warning("Selecciona dos meses diferentes para analizar la variación.")
    else:
        df_base = df[df['Periodo'] == mes_base].iloc[0]
        df_analisis = df[df['Periodo'] == mes_analisis].iloc[0]
        
        huevo_base = df_base['Huevos Fértiles']
        huevo_analisis = df_analisis['Huevos Fértiles']
        
        costo_unit_base = df_base['Costo Huevo Fértil']
        costo_unit_analisis = df_analisis['Costo Huevo Fértil']
        variacion_total = costo_unit_analisis - costo_unit_base
        
        # Tarjetas Métricas Superiores (KPIs)
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric(f"Costo {mes_base}", f"${costo_unit_base:,.2f}")
        kpi2.metric(f"Costo {mes_analisis}", f"${costo_unit_analisis:,.2f}")
        kpi3.metric("Variación Total", f"${variacion_total:+,.2f} COP", delta=f"{(variacion_total/costo_unit_base)*100:+.1f}%", delta_color="inverse")
        
        st.markdown("---")

        # Cálculo de impacto por rubro ($/huevo)
        impactos = {}
        for item in rubros_items:
            unit_b = df_base[item] / huevo_base
            unit_a = df_analisis[item] / huevo_analisis
            var_item = unit_a - unit_b
            if abs(var_item) > 0.01:  # Filtra variaciones menores a 1 centavo
                impactos[item] = var_item
                
        df_imp = pd.DataFrame(list(impactos.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values(by='Impacto $/Huevo', ascending=True)

        # OPCIÓN A: BARRAS HORIZONTALES (MUCHO MÁS CLARO QUE EL WATERFALL)
        if tipo_vista == "Barras de Impacto (Recomendado)":
            fig = px.bar(
                df_imp,
                y='Rubro',
                x='Impacto $/Huevo',
                orientation='h',
                text_auto='+$.1f',
                color='Impacto $/Huevo',
                color_continuous_scale='Reds' if variacion_total > 0 else 'Greens',
                title=f"Aporte Directo de cada Rubro al Cambio Total de ${variacion_total:+,.2f} COP/Huevo"
            )
            fig.update_layout(
                xaxis_title="Impacto en el Costo Unitario ($ COP / Huevo)",
                yaxis_title="",
                height=450,
                coloraxis_showscale=False
            )
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        # OPCIÓN B: WATERFALL REESCALADO SIN DISTORSIÓN VISUAL
        else:
            items_wf = list(impactos.keys())
            valores_wf = list(impactos.values())
            
            # Recorta el eje Y para que las barras flotantes se vean anchas y claras
            min_y = min(costo_unit_base, costo_unit_analisis) * 0.95
            
            fig = go.Figure(go.Waterfall(
                name="Variación", orientation="v",
                measure=["absolute"] + ["relative"] * len(items_wf) + ["total"],
                x=[f"Base ({mes_base})"] + items_wf + [f"Final ({mes_analisis})"],
                textposition="outside",
                text=[f"${costo_unit_base:,.1f}"] + [f"{v:+.1f}" for v in valores_wf] + [f"${costo_unit_analisis:,.1f}"],
                y=[costo_unit_base] + valores_wf + [0],
                connector={"line":{"color":"gray"}},
                decreasing={"marker":{"color":"#2ca02c"}},
                increasing={"marker":{"color":"#d62728"}},
                totals={"marker":{"color":"#1f77b4"}}
            ))
            fig.update_layout(
                title="Descomposición en Cascada (Escala Recortada)",
                yaxis=dict(range=[min_y, max(costo_unit_base, costo_unit_analisis) * 1.05]),
                height=500
            )
            st.plotly_chart(fig, use_container_width=True)

        # Resumen Ejecutivo
        df_top = df_imp.sort_values(by='Impacto $/Huevo', ascending=False)
        st.subheader("💡 Conclusión del Análisis:")
        col_left, col_right = st.columns(2)
        with col_left:
            st.write(f"El cambio de **${variacion_total:+,.2f} COP** por huevo fértil se explica principalmente por:")
            for idx, row in df_top.head(3).iterrows():
                pct_explicado = (row['Impacto $/Huevo'] / variacion_total) * 100 if variacion_total != 0 else 0
                st.markdown(f"* **{row['Rubro']}**: Aportó **${row['Impacto $/Huevo']:+,.2f}** por huevo (explica el **{pct_explicado:.1f}%** de la variación).")
        with col_right:
            st.dataframe(df_top.style.format({'Impacto $/Huevo': '${:+.2f}'}), use_container_width=True)

# -----------------------------------------------------------------------------
# MÓDULO 2: OPCIÓN DE IMPACTO POR ÍTEM
# -----------------------------------------------------------------------------
elif menu == "2. Impacto por Ítem":
    st.header("📊 Estructura de Costos y Participación por Rubro")
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
elif menu == "3. Simulador de Escenarios (What-If)":
    st.header("🎛️ Simulador de Sensibilidad y Nuevo Margen")
    st.markdown("Ajusta las variables operativas para simular un nuevo escenario y recalcular en tiempo real el costo unitario y el margen.")
    
    mes_ref = st.selectbox("Seleccionar Mes Base para la Simulación:", df['Periodo'].tolist(), index=len(df)-1)
    df_ref = df[df['Periodo'] == mes_ref].iloc[0]
    
    st.subheader("1. Modificar Parámetros Operativos")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        precio_alimento_sim = st.slider("Precio Kg Alimento ($ COP):", 1200.0, 2200.0, float(df_ref['Precio Kg Alimento']), 10.0)
    with col2:
        gramos_huevo_sim = st.slider("Consumo (g Alimento/Huevo):", 200.0, 320.0, float(df_ref['Gramos Alimento/Huevo']), 1.0)
    with col3:
        fertilidad_sim = st.slider("Fertilidad (%):", 80.0, 99.0, float((df_ref['Huevos Fértiles']/df_ref['Total Producción'])*100), 0.5)
    with col4:
        precio_venta_target = st.number_input("Precio Venta Huevo Fértil ($):", value=1600.0, step=10.0)

    # RE-CÁLCULO DEL ESCENARIO
    prod_total = df_ref['Total Producción']
    huevos_fert_sim = prod_total * (fertilidad_sim / 100.0)
    
    kg_alimento_sim = (prod_total * gramos_huevo_sim) / 1000.0
    costo_alimento_sim = kg_alimento_sim * precio_alimento_sim
    
    otros_costos = df_ref['Costo Total'] - df_ref['Alimento']
    costo_total_sim = otros_costos + costo_alimento_sim
    
    costo_huevo_sim = costo_total_sim / huevos_fert_sim
    costo_huevo_orig = df_ref['Costo Huevo Fértil']
    
    margen_unitario = precio_venta_target - costo_huevo_sim
    margen_pct = (margen_unitario / precio_venta_target) * 100
    
    st.markdown("---")
    st.subheader("2. Resultado Proyectado del Nuevo Escenario")
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Costo Huevo Simulado", f"${costo_huevo_sim:,.2f}", delta=f"{costo_huevo_sim - costo_huevo_orig:+,.2f} COP", delta_color="inverse")
    kpi2.metric("Precio Venta Huevo", f"${precio_venta_target:,.2f}")
    kpi3.metric("Margen Unitario ($)", f"${margen_unitario:,.2f}", delta=f"{margen_pct:.1f}% Margen")
    kpi4.metric("Utilidad Total Est.", f"${(margen_unitario * huevos_fert_sim):,.0f}")

# -----------------------------------------------------------------------------
# MÓDULO 4: PROYECCIÓN FUTURA
# -----------------------------------------------------------------------------
elif menu == "4. Proyección Futura":
    st.header("🔮 Proyección Futura del Costo")
    st.markdown("Estimación del comportamiento del costo por huevo fértil en los próximos meses usando suavizado exponencial.")
    
    meses_proyeccion = st.slider("Meses a Proyectar hacia el futuro:", 1, 6, 3)
    
    ts_data = df['Costo Huevo Fértil'].values
    model = ExponentialSmoothing(ts_data, trend='add', seasonal=None).fit()
    forecast = model.forecast(meses_proyeccion)
    
    ult_fecha = pd.to_datetime(df['Periodo'].iloc[-1])
    fechas_futuras = [(ult_fecha + pd.DateOffset(months=i)).strftime('%Y-%m') for i in range(1, meses_proyeccion + 1)]
    
    df_proj = pd.DataFrame({'Periodo': fechas_futuras, 'Costo Proyectado': forecast})
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['Periodo'], y=df['Costo Huevo Fértil'], mode='lines+markers', name='Histórico Real', line=dict(color='#1f77b4', width=3)))
    fig.add_trace(go.Scatter(x=df_proj['Periodo'], y=df_proj['Costo Proyectado'], mode='lines+markers', name='Proyección Futura', line=dict(color='#d62728', width=3, dash='dash')))
    
    fig.update_layout(title="Tendencia y Proyección de Costo por Huevo Fértil ($ COP)", xaxis_title="Período", yaxis_title="Costo/Huevo ($)", height=500)
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Valores Estimados Proyectados:")
    st.dataframe(df_proj.style.format({'Costo Proyectado': '${:,.2f}'}), use_container_width=True)
