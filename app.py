def grafico_tornado_vif(df_vif, costo_base, costo_actual, titulo, subtitulo=""):
    """
    Construye un gráfico de Barras de Impacto (Tornado) de alta claridad visual.
    Ideal para mostrar las desviaciones financieras (VIF) de mayor a menor impacto,
    facilitando la lectura de los nombres de los rubros en el eje Y.
    """
    # Excluimos los impactos cero o nulos para limpiar el gráfico
    df_plot = df_vif[df_vif["Impacto ($/HF)"].abs() > 0.1].copy()
    
    # Ordenamos: Mayor impacto negativo (ahorro) abajo, mayor impacto positivo (sobrecosto) arriba
    df_plot = df_plot.sort_values("Impacto ($/HF)", ascending=True)

    textos = [f"+${v:,.1f}" if v > 0 else f"-${abs(v):,.1f}" for v in df_plot["Impacto ($/HF)"]]
    colores = ["#EF4444" if v > 0 else "#10B981" for v in df_plot["Impacto ($/HF)"]]

    fig = go.Figure(go.Bar(
        x=df_plot["Impacto ($/HF)"],
        y=df_plot["Rubro"],
        orientation="h",
        text=textos,
        textposition="outside",
        marker_color=colores,
        width=0.6 # Hace las barras más elegantes
    ))

    # Estilización del gráfico
    fig = estilizar_grafico(fig, titulo, subtitulo, altura=480)
    
    # Agregar una línea vertical en el 0 para separar aumentos de ahorros
    fig.add_vline(x=0, line_width=2, line_color="#1E3A8A")
    
    # Ajustes finales de los ejes
    fig.update_xaxes(title_text="Impacto en el Costo Unitario ($/Huevo)", showgrid=True, gridcolor="#E2E8F0")
    fig.update_yaxes(title_text="", showline=False)
    
    # Calculamos el margen para que los textos no se corten
    max_val = df_plot["Impacto ($/HF)"].abs().max() * 1.25
    fig.update_xaxes(range=[-max_val if df_plot["Impacto ($/HF)"].min() < 0 else 0, max_val])

    return fig
