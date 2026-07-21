# -----------------------------------------------------------------------------
# MÓDULO 1: ANÁLISIS DE VARIACIÓN INTERMENSUAL (MEJORADO Y CLARO)
# -----------------------------------------------------------------------------
if menu == "1. Análisis Variación (VIF Waterfall)":
    st.header("🔍 ¿Por qué subió el Costo por Huevo Fértil?")
    st.markdown("Compara dos períodos y visualiza claramente qué rubros explicaron la variación en pesos COP por huevo.")
    
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
            if abs(var_item) > 0.01:  # Filtra variaciones irrelevantes (menores a 1 centavo)
                impactos[item] = var_item
                
        df_imp = pd.DataFrame(list(impactos.items()), columns=['Rubro', 'Impacto $/Huevo']).sort_values(by='Impacto $/Huevo', ascending=True)

        # OPCIÓN A: BARRAS HORIZONTALES (LIMPIDO Y SCANNABLE)
        if tipo_vista == "Barras de Impacto (Recomendado)":
            fig = px.bar(
                df_imp,
                y='Rubro',
                x='Impacto $/Huevo',
                orientation='h',
                text_auto='+$.1f',
                color='Impacto $/Huevo',
                color_continuous_scale='Reds' if variacion_total > 0 else 'Greens',
                title=f"Aporte de cada Rubro al Cambio Total de ${variacion_total:+,.2f} COP/Huevo"
            )
            fig.update_layout(
                xaxis_title="Impacto en el Costo Unitario ($ COP / Huevo)",
                yaxis_title="",
                height=450,
                coloraxis_showscale=False
            )
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        # OPCIÓN B: WATERFALL CON RANGO DE EJE Y OPTIMIZADO
        else:
            items_wf = list(impactos.keys())
            valores_wf = list(impactos.values())
            
            # Ajuste de escala para que no se vean diminutas las barras
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
            st.write(f"El incremento de **${variacion_total:+,.2f} COP** por huevo fértil se explica principalmente por:")
            for idx, row in df_top.head(3).iterrows():
                pct_explicado = (row['Impacto $/Huevo'] / variacion_total) * 100 if variacion_total != 0 else 0
                st.markdown(f"* **{row['Rubro']}**: Subió **${row['Impacto $/Huevo']:+,.2f}** por huevo (explica el **{pct_explicado:.1f}%** del alza).")
        with col_right:
            st.dataframe(df_top.style.format({'Impacto $/Huevo': '${:+.2f}'}), use_container_width=True)
