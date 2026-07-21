import os
import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# Funciones de exportación a PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

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

def generar_pdf_resumen(df_resumen, mes_b, mes_a, var_tot):
    """Genera un archivo PDF ejecutivo en memoria."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#1f77b4'))
    story.append(Paragraph("<b>INFORME EJECUTIVO: VARIACIÓN COSTO HUEVO FÉRTIL</b>", title_style))
    story.append(Spacer(1, 10))
    
    body_text = f"Análisis comparativo de variación entre el período <b>{mes_b}</b> y el período <b>{mes_a}</b>.<br/>" \
                f"Variación Total Registrada: <b>${var_tot:+,.2f} COP por Huevo Fértil</b>."
    story.append(Paragraph(body_text, styles['Normal']))
    story.append(Spacer(1, 15))
    
    # Tabla de datos
    table_data = [["Rubro de Costo", "Impacto $/Huevo"]]
    for idx, row in df_resumen.iterrows():
        table_data.append([str(row['Rubro']), f"${row['Impacto $/Huevo']:+,.2f} COP"])
        
    t = Table(table_data, colWidths=[250, 150])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f7f9fa')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# CONTROL DE CARGA EN EL SIDEBAR
# -----------------------------------------------------------------------------
st.sidebar.title("🐔 Gestión Avícola")
st.sidebar.subheader("📂 Carga de Datos")

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
        "4. Proyección Futura",
        "5. Exportar Reportes"
    ]
)

# -----------------------------------------------------------------------------
# MÓDULO 1: ANÁLISIS DE VARIACIÓN INTERMENSUAL + TENDENCIA
# -----------------------------------------------------------------------------
if menu == "1. Análisis Variación del Costo":
    st.header("🔍 ¿Por qué subió o bajó el Costo por Huevo Fértil?")
    st.markdown("Compara dos períodos, analiza la tendencia histórica y visualiza los rubros responsables de la variación en pesos COP.")
    
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
        
        # KPIs
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric(f"Costo {mes_base}", f"${costo_unit_base:,.2f}")
        kpi2.metric(f"Costo {mes_analisis}", f"${costo_unit_analisis:,.2f}")
        kpi3.metric("Variación Total", f"${variacion_total:+,.2f} COP", delta=f"{(variacion_total/costo_unit_base)*100:+.1f}%", delta_color="inverse")
        
        st.markdown("---")

        # --- SECCIÓN NUEVA: GRÁFICO DE TENDENCIA HISTÓRICA ---
        st.subheader("📈 Tendencia Histórica del Costo por Huevo Fértil")
        fig_line = px.line(df, x='Periodo', y='Costo Huevo Fértil', markers=True, text='Costo Huevo Fértil',
                           title="Evolución del Costo Unitario ($ COP / Huevo)")
        fig_line.update_traces(texttemplate='$%{text:,.1f}', textposition='top center', line_color='#1f77b4', line_width=3)
        
        # Resaltar en la línea los dos puntos seleccionados
        fig_line.add_scatter(x=[mes_base, mes_analisis], 
                             y=[costo_unit_base, costo_unit_analisis],
                             mode='markers', marker=dict(size=12, color='red'), name='Períodos Seleccionados')
        
        st.plotly_chart(fig_line, use_container_width=True)
        st.markdown("---")

        # Cálculo de impacto por rubro ($/huevo)
        impactos = {}
        for item in rubros_items:
            unit_b = df_base[item] / huevo_base
            unit_a = df_analisis[item] / huevo_analisis
            var_item = unit_a - unit_b
            if abs(var_item) > 0.01:
                impactos[item] = var_item
                
        df_imp = pd.DataFrame(list(impactos.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values(by='Impacto $/Huevo', ascending=True)

        # BARRAS HORIZONTALES CON ETIQUETAS FORMATEADAS (+$.1f)
        if tipo_vista == "Barras de Impacto (Recomendado)":
            fig = px.bar(
                df_imp,
                y='Rubro',
                x='Impacto $/Huevo',
                orientation='h',
                color='Impacto $/Huevo',
                color_continuous_scale='Reds' if variacion_total > 0 else 'Greens',
                title=f"Aporte Directo de cada Rubro al Cambio Total de ${variacion_total:+,.2f} COP/Huevo"
            )
            fig.update_traces(
                text=[f"${v:+,.1f}" for v in df_imp['Impacto $/Huevo']], 
                textposition='outside'
            )
            fig.update_layout(
                xaxis_title="Impacto en el Costo Unitario ($ COP / Huevo)",
                yaxis_title="",
                height=450,
                coloraxis_showscale=False
            )
            st.plotly_chart(fig, use_container_width=True)

        else:
            items_wf = list(impactos.keys())
            valores_wf = list(impactos.values())
            min_y = min(costo_unit_base, costo_unit_analisis) * 0.95
            
            fig = go.Figure(go.Waterfall(
                name="Variación", orientation="v",
                measure=["absolute"] + ["relative"] * len(items_wf) + ["total"],
                x=[f"Base ({mes_base})"] + items_wf + [f"Final ({mes_analisis})"],
                textposition="outside",
                text=[f"${costo_unit_base:,.1f}"] + [f"${v:+,.1f}" for v in valores_wf] + [f"${costo_unit_analisis:,.1f}"],
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

        # Resumen Ejecutivo y Botón PDF Rápido
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

# -----------------------------------------------------------------------------
# MÓDULO 5: EXPORTAR INFORMES (EXCEL Y PDF)
# -----------------------------------------------------------------------------
elif menu == "5. Exportar Reportes":
    st.header("📥 Centro de Exportación de Reportes")
    st.markdown("Descarga la base de datos limpia o genera un informe oficial en formato PDF para reuniones ejecutivas.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Exportar a Excel")
        st.write("Descarga los datos procesados en formato estructurado (.xlsx).")
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Matriz_Costos_Historica', index=False)
        buffer_excel.seek(0)
        
        st.download_button(
            label="💾 Descargar Excel Completo",
            data=buffer_excel,
            file_name="Matriz_Costos_Huevo_Fertil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col2:
        st.subheader("📄 Exportar Reporte Ejecutivo a PDF")
        st.write("Genera un documento PDF oficial del análisis de variación.")
        
        mes_b_pdf = st.selectbox("Período Base PDF:", df['Periodo'].tolist(), index=0)
        mes_a_pdf = st.selectbox("Período Análisis PDF:", df['Periodo'].tolist(), index=len(df)-1)
        
        if st.button("⚙️ Generar PDF"):
            df_b_pdf = df[df['Periodo'] == mes_b_pdf].iloc[0]
            df_a_pdf = df[df['Periodo'] == mes_a_pdf].iloc[0]
            
            imp_pdf = {}
            for item in rubros_items:
                v_b = df_b_pdf[item] / df_b_pdf['Huevos Fértiles']
                v_a = df_a_pdf[item] / df_a_pdf['Huevos Fértiles']
                imp_pdf[item] = v_a - v_b
                
            df_res_pdf = pd.DataFrame(list(imp_pdf.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values(by='Impacto $/Huevo', ascending=False)
            var_tot_pdf = df_a_pdf['Costo Huevo Fértil'] - df_b_pdf['Costo Huevo Fértil']
            
            pdf_bytes = generar_pdf_resumen(df_res_pdf, mes_b_pdf, mes_a_pdf, var_tot_pdf)
            
            st.download_button(
                label="📄 Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"Informe_Costo_Huevo_{mes_b_pdf}_vs_{mes_a_pdf}.pdf",
                mime="application/pdf"
            )
