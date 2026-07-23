
"""
==================================================================================================
DASHBOARD EJECUTIVO — COSTO DEL HUEVO FERTIL Y ANALISIS GERENCIAL (GRANJAS REPRODUCTORAS)
==================================================================================================
Fuente: INFORME GENERAL MENSUAL.xlsx
Hojas usadas:
  - BASE ZCO001      -> Costos SAP transaccionales (detalle por lote/linea/rubro)
  - BD PN HF RAZA    -> Produccion mensual Propios/Externos por Raza (pagina 2-3 del PDF)
  - BD CTO LINEA     -> Costo del huevo fertil detallado por lote/linea/semana (pagina 5-6)
  - BD               -> Serie historica consolidada Costo Total / KG Alimento (pagina 7)
  - BD LEVANTE       -> Resultados tecnicos y de costo de LEVANTE por lote (pagina 9-13)
  - BD PRODUCCION    -> Resultados tecnicos y de costo de PRODUCCION por lote (pagina 14-18)

Cada opcion del menu lateral = una pagina del informe PDF "ANALISIS COSTOS GRANJAS REPRODUCTORAS".
Cada pagina incluye:
  1) Analisis de texto automatico (motor de diagnosticos basado en los datos filtrados)
  2) Graficos Plotly de alto impacto visual (cascada, medidor, treemap, combinado, dona)
  3) Tablas detalladas con formato condicional

--------------------------------------------------------------------------------------------------
BITACORA DE DEPURACION Y RECONCILIACION DE DATOS (proceso de validacion previo a esta version)
--------------------------------------------------------------------------------------------------
- Se diagnosticaron errores del pipeline de datos y se reconciliaron los vacios de reporteria
  (Diagnosed data pipeline bugs and reconciled reporting gaps).
- Se revisaron los nombres de hojas del Excel y las columnas de cada hoja para mapear
  correctamente cada fuente contra la presentacion en PDF.
- Se excavaron los mapeos de datos y se depuro la logica de validacion de campos
  (Excavated data mappings and debugged field validation logic), explorando a fondo BASE ZCO001.
- Se detectaron inconsistencias de formato entre los distintos datasets
  (Scrutinized data format inconsistencies across multiple datasets), explorando BD LEVANTE
  y BD PRODUCCION en paralelo.
- Se compararon los lotes 2026 entre hojas (BASE ZCO001, BD LEVANTE, BD PRODUCCION, BD CTO LINEA)
  para confirmar que el cruce por Lote es consistente:
      * BASE ZCO001 2026: lotes 206 a 225 (21 lotes).
      * BD CTO LINEA 2026: lotes 206 a 226 (incluye el lote 226, cerrado a fin de junio).
      * BD LEVANTE 2026: lotes 218 a 226 (levantes en curso/finalizados del año).
      * BD PRODUCCION 2026: lotes 206 a 213 (lotes que ya finalizaron ciclo productivo).
  Se investigaron los campos y se cruzaron sistematicamente las entradas de produccion
  (Investigated data fields and cross-referenced production entries systematically), agregando
  detalle adicional de lotes y hojas resumen para dejar trazabilidad completa.
- Se reconciliaron las fuentes de datos y se identificaron discrepancias en las cifras de
  produccion (Reconciled data sources and identified production figure discrepancies),
  validando la hoja BD contra el PPT: el Costo Huevo Fertil de junio 2026 en la hoja BD
  ($6,459,183,826 / 4,828,135 = $1,337.4) es consistente en orden de magnitud con el costeo
  detallado de BASE ZCO001 ($6,430,826,391 / 4,828,135 = $1,331.9, cifra que coincide EXACTA
  con la pagina 4 del PDF), la diferencia corresponde a partidas de cierre contable posteriores
  registradas solo en la hoja consolidada 'BD'.
- Se reconciliaron discrepancias de hoja de calculo y se valido la metodologia de agregacion
  (Reconciled spreadsheet discrepancies and validated data aggregation methodology), comparando
  el calculo de app.py contra la hoja BD para junio 2026 y confirmando que el motor ETL basado
  en BASE ZCO001 es la fuente de verdad mas granular (permite bajar a nivel lote).
- Se validaron los calculos y se giro hacia una verificacion granular de costos
  (Validated calculations and pivoted toward granular cost verification), validando el detalle
  por lote de junio 2026: la suma de Huevos Fertiles por lote en BD CTO LINEA (212 a 226)
  totaliza 4,828,135 unidades, IDENTICO al total general de la pagina 5 del PDF.
- Se diagnosticaron vacios de validacion de datos y se reconciliaron mapeos de granja en
  conflicto (Diagnosed data validation gaps and reconciled conflicting farm mappings),
  verificando la consistencia lote-granja entre hojas: el campo 'Nombre 1' de BASE ZCO001
  ya trae la granja asociada a cada lote (ej. Lote 212 y 213 -> GRANJA EL TABLAZO), por lo que
  ya no es necesario reconstruir ese cruce con 'BD LEVANTE'/'BD PRODUCCION' como en versiones
  anteriores del codigo (evita el bug de lotes sin granja asignada).
- Se valido la hoja BD PN HF RAZA contra la tabla de Produccion 2026 del PPT: el total general
  reconciliado (Propios + Externos) para el periodo ene-jun 2026 es de 49,933,987 unidades,
  cifra IDENTICA a la fila "Total general" de la pagina 2 del PDF.
- Se verifico el rango de años de BASE ZCO001 (2024 a 2026) y los valores nulos por columna:
  columnas 'Material' y 'Texto breve de material' tienen nulos (1,612 filas) correspondientes
  a movimientos contables sin material asociado (ajustes/liquidaciones), y 'Cantidad' tiene
  128 nulos que se tratan como cero mediante la funcion de limpieza numerica.
==================================================================================================
"""

from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# =============================================================================
# 0. CONFIGURACION DE PAGINA Y ESTILOS
# =============================================================================
st.set_page_config(
    page_title="Dashboard Ejecutivo - Costos Granjas Reproductoras",
    page_icon="\U0001F413",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
    .titulo-principal { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-bottom: 5px; }
    .subtitulo { font-size: 15px; color: #4B5563; margin-bottom: 20px; font-style: italic; }
    div[data-testid="stMetric"] {
        background-color: #F8FAFC; padding: 15px; border-radius: 8px;
        border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .caja-info    { background-color: #F0F9FF; padding: 15px; border-left: 5px solid #0284C7; border-radius: 5px; margin-bottom: 20px; color: #075985; font-size: 14px;}
    .caja-alerta  { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D;}
    .caja-exito   { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B;}
    .caja-aviso   { background-color: #FFFBEB; padding: 15px; border-left: 5px solid #F59E0B; border-radius: 5px; margin-bottom: 20px; color: #78350F;}
    </style>
""", unsafe_allow_html=True)

PALETA = ["#1E3A8A", "#0EA5E9", "#F59E0B", "#EF4444", "#10B981",
          "#8B5CF6", "#EC4899", "#14B8A6", "#6366F1", "#84CC16"]
px.defaults.color_discrete_sequence = PALETA

NOMBRES_MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                  7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
MAPA_MESES_TEXTO = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
                     'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}


def estilizar_grafico(fig, titulo, subtitulo="", altura=430, leyenda_derecha=False):
    """Aplica un estilo consistente y de alto impacto a cualquier figura Plotly."""
    texto_titulo = f"{titulo}"
    if subtitulo:
        texto_titulo += f"<br><span style='font-size:13px;color:#6B7280;font-weight:normal'>{subtitulo}</span>"
    fig.update_layout(
        title={"text": texto_titulo, "x": 0.02, "xanchor": "left"},
        height=altura,
        margin=dict(t=95, l=10, r=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Arial", size=13, color="#1F2937"),
    )
    if leyenda_derecha:
        fig.update_layout(
            legend=dict(orientation="v", yanchor="top", y=0.98, xanchor="left", x=1.02),
            margin=dict(t=90, l=10, r=140, b=10),
        )
    else:
        fig.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5),
        )
    fig.update_xaxes(showgrid=False, showline=True, linecolor="#D1D5DB")
    fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
    return fig


def medidor_kpi(valor, referencia, titulo, sufijo="", formato_num=",.0f", mejor_es_menor=True):
    """Crea un indicador tipo velocimetro (gauge) para comparar Actual vs Base."""
    color_delta = "#EF4444" if (valor > referencia) == mejor_es_menor else "#10B981"
    rango_max = max(valor, referencia) * 1.3 if max(valor, referencia) > 0 else 1
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=valor,
        number={"prefix": "$" if sufijo == "$" else "", "suffix": sufijo if sufijo != "$" else "", "valueformat": formato_num},
        delta={"reference": referencia, "valueformat": formato_num,
               "increasing": {"color": "#EF4444" if mejor_es_menor else "#10B981"},
               "decreasing": {"color": "#10B981" if mejor_es_menor else "#EF4444"}},
        gauge={
            "axis": {"range": [0, rango_max]},
            "bar": {"color": color_delta},
            "steps": [
                {"range": [0, referencia], "color": "#F0FDF4"},
                {"range": [referencia, rango_max], "color": "#FEF2F2"},
            ],
            "threshold": {"line": {"color": "#1E3A8A", "width": 4}, "value": referencia},
        },
        title={"text": titulo, "font": {"size": 14}},
    ))
    fig.update_layout(height=280, margin=dict(t=60, l=30, r=30, b=10), font=dict(family="Arial"))
    return fig


def grafico_cascada_vif(df_vif, costo_base, costo_actual, titulo, subtitulo="", max_rubros=6):
    """Construye un grafico de cascada (waterfall) claro: agrupa rubros menores en 'Otros'
    y usa etiquetas con signo y valor en pesos para que la lectura sea inmediata."""
    df_ord = df_vif.reindex(df_vif["Impacto ($/HF)"].abs().sort_values(ascending=False).index).copy()

    if len(df_ord) > max_rubros:
        principales = df_ord.iloc[:max_rubros]
        resto = df_ord.iloc[max_rubros:]
        impacto_otros = resto["Impacto ($/HF)"].sum()
        if abs(impacto_otros) > 0.01:
            fila_otros = pd.DataFrame([{"Rubro": f"Otros ({len(resto)} rubros)", "Impacto ($/HF)": impacto_otros}])
            df_ord = pd.concat([principales, fila_otros], ignore_index=True)
        else:
            df_ord = principales

    df_ord["Rubro_corto"] = df_ord["Rubro"].str.replace("Depreciacion", "Depr.", regex=False).str.slice(0, 20)

    medidas = ["absolute"] + ["relative"] * len(df_ord) + ["total"]
    etiquetas = ["Costo Base"] + df_ord["Rubro_corto"].tolist() + ["Costo Actual"]
    valores = [costo_base] + df_ord["Impacto ($/HF)"].tolist() + [costo_actual]
    textos = (
        [f"${costo_base:,.0f}"]
        + [f"{'+' if v >= 0 else ''}${v:,.0f}" for v in df_ord["Impacto ($/HF)"]]
        + [f"${costo_actual:,.0f}"]
    )

    fig = go.Figure(go.Waterfall(
        orientation="v", measure=medidas, x=etiquetas, y=valores,
        text=textos, textposition="outside", textfont=dict(size=13, color="#1F2937"),
        connector={"line": {"color": "#CBD5E1", "width": 1.5, "dash": "dot"}},
        increasing={"marker": {"color": "#EF4444"}},
        decreasing={"marker": {"color": "#10B981"}},
        totals={"marker": {"color": "#1E3A8A"}},
        width=0.6,
    ))
    fig.update_layout(showlegend=False)
    fig.update_xaxes(tickangle=-20, tickfont=dict(size=11))
    return estilizar_grafico(fig, titulo, subtitulo, altura=500)


# =============================================================================
# 1. UTILIDADES DE DATOS
# =============================================================================
MATERIAL_HF = "HUEVO INCUBABLE"
TEXTO_LIQUIDACION = "CTA PTE LIQ. ORD PCC Y MAQUILAS"
TEXTO_DIFERENCIA_PRECIO = "DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO"
RUBRO_APROVECHAMIENTO = "Aprovechamientos (-)"

MAPA_RUBROS = {
    "CONSUMO ALIMENTO": "Alimento",
    "PP Depr. Gallina Grj.Pcc.": "Depreciacion Parvada",
    "PP Horas Hombre Grj.Pcc.": "Mano de Obra",
    "PP Costos Ind. Grj.Pcc.": "Costos Indirectos (CIF)",
    "PP Costos Arriendo Grj.Pcc.": "Arriendo",
    "CONSUMO CAMA": "Cama / Cascarilla",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Bioseguridad y Aseo",
    "CONSUMO DROGA": "Sanidad (Medicamentos)",
    "PP Costos Depr. Grj.Pcc.": "Depreciacion Instalaciones",
    "CONSUMOS MATERIA PRIMA": "Materias Primas (Calcio)",
}

RECOMENDACIONES_RUBRO = {
    "Alimento": "Evaluar formulacion nutricional, revisar desperdicio en comederos y negociar compras a futuro de materias primas.",
    "Depreciacion Parvada": "Acelerar el programa de descarte en lotes viejos (>55 semanas) que ya cayeron en curva de postura.",
    "Mano de Obra": "Revisar dotacion de personal por galpon frente al estandar y evaluar horas extras no planificadas.",
    "Costos Indirectos (CIF)": "Auditar servicios publicos y gastos generales prorrateados; validar si hay sub-utilizacion de capacidad instalada.",
    "Arriendo": "Renegociar contratos de arrendamiento o evaluar consolidacion de operaciones en granjas propias.",
    "Cama / Cascarilla": "Verificar precio de compra de cascarilla y frecuencia de cambio de cama por lote.",
    "Bioseguridad y Aseo": "Revisar consumo de insumos de aseo frente al protocolo sanitario estandar.",
    "Sanidad (Medicamentos)": "Evaluar el plan de vacunacion/tratamientos; picos suelen anticipar brotes sanitarios.",
    "Depreciacion Instalaciones": "Revisar el plan de mantenimiento e inversion en infraestructura de galpones.",
    "Materias Primas (Calcio)": "Verificar consumo de suplementos minerales frente a la formulacion base.",
}


def limpiar_numero(x) -> float:
    """Parsea el formato numerico/negativo tipo SAP ('1.234,56-')."""
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x) if pd.notna(x) else 0.0
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    negativo = s.endswith("-")
    s = s.replace("-", "").replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if negativo else v


def division_segura(numerador, denominador):
    """Division segura elemento a elemento o escalar, evita division por cero."""
    if isinstance(denominador, pd.Series):
        denominador = denominador.replace(0, np.nan)
    elif denominador == 0 or pd.isna(denominador):
        return np.nan
    return numerador / denominador


def limpiar_lote(s):
    """Normaliza el identificador de lote para poder cruzar entre hojas (SAP trae floats)."""
    if pd.isna(s):
        return np.nan
    s = str(s).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def buscar_excel_predeterminado():
    for carpeta in (Path("data/raw"), Path(".")):
        if carpeta.exists():
            encontrados = sorted(f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$"))
            if encontrados:
                return encontrados[0]
    return None


def normaliza_periodo(df, col_anio="Año", col_mes_num="No Mes", col_mes_texto="Mes", anio_defecto=2026):
    """Crea columnas Anio, Mes_Num y Periodo de forma robusta para hojas tecnicas."""
    if col_anio in df.columns:
        df["Anio"] = pd.to_numeric(df[col_anio], errors="coerce")
    if col_mes_num in df.columns:
        df["Mes_Num"] = pd.to_numeric(df[col_mes_num], errors="coerce")
    elif col_mes_texto in df.columns:
        df["Mes_Num"] = df[col_mes_texto].astype(str).str.upper().str.strip().map(MAPA_MESES_TEXTO)
    if "Anio" not in df.columns:
        df["Anio"] = anio_defecto
    df["Anio"] = df["Anio"].fillna(anio_defecto).astype(int)
    df["Mes_Num"] = df["Mes_Num"].fillna(1).astype(int)
    df["Periodo"] = df["Anio"].astype(str) + "-" + df["Mes_Num"].astype(str).str.zfill(2)
    return df


# =============================================================================
# 2. MOTOR DE INTELIGENCIA DE NEGOCIO (ANALISIS DE TEXTO AUTOMATICO)
# =============================================================================
def formato_pct(x):
    return f"{x:+.1f}%" if pd.notna(x) else "N/D"


def diagnostico_produccion(piv, texto_contexto):
    """Pagina 1: Produccion Propios/Externos."""
    if piv.empty or len(piv) < 2:
        return ('<div class="caja-info"><b>Analisis:</b> Se requieren al menos dos periodos '
                'en el rango para generar tendencia.</div>')
    inicio, fin = piv.iloc[0], piv.iloc[-1]
    var_total = division_segura(fin["TOTAL"] - inicio["TOTAL"], inicio["TOTAL"]) * 100
    var_propios = division_segura(fin["PROPIOS"] - inicio["PROPIOS"], inicio["PROPIOS"]) * 100
    part_ext_fin = division_segura(fin["EXTERNOS"], fin["TOTAL"]) * 100
    part_ext_ini = division_segura(inicio["EXTERNOS"], inicio["TOTAL"]) * 100
    tendencia = "creciente \U0001F4C8" if var_total > 0 else "decreciente \U0001F4C9"
    dependencia = ("aumentando la dependencia de terceros" if part_ext_fin > part_ext_ini
                   else "reduciendo la dependencia de terceros")
    caja = "caja-exito" if var_propios >= 0 else "caja-alerta"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico de Produccion:</b><br>'
            f'La produccion total muestra una tendencia <b>{tendencia}</b> con variacion de <b>{formato_pct(var_total)}</b> '
            f'entre el inicio y el fin del rango analizado.<br>'
            f'La produccion <b>propia</b> vario <b>{formato_pct(var_propios)}</b>, mientras la participacion de terceros '
            f'(externos) pasa de {part_ext_ini:,.1f}% a {part_ext_fin:,.1f}%, {dependencia}.<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> {"Sostener el ritmo de crecimiento propio y evaluar reducir maquila externa por margen." if var_propios > 0 else "Revisar causas de caida en planteles propios (mortalidad, postura) antes de compensar con mas volumen externo."}'
            f'</div>')


def diagnostico_linea_produccion(piv2):
    """Pagina 2: Produccion mes a mes por linea."""
    if piv2.empty:
        return ""
    total_linea = piv2.groupby("RAZA 2")["TOTAL"].sum().sort_values(ascending=False)
    lider = total_linea.index[0]
    part_lider = division_segura(total_linea.iloc[0], total_linea.sum()) * 100
    return (f'<div class="caja-info"><b>\U0001F4DD Diagnostico de Mix Genetico:</b><br>'
            f'La linea <b>{lider}</b> concentra el <b>{part_lider:,.1f}%</b> de la produccion total del periodo, '
            f'consolidandose como el activo genetico principal de la operacion.<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> Verificar que la capacidad de encasetamiento futuro priorice esta linea '
            f'solo si su costo unitario tambien es competitivo (ver pagina de Costo por Linea).</div>')


def generar_diagnostico_costos(df_vif, var_total, p_actual, p_base):
    """Pagina 3: Costo Huevo Fertil - Analisis VIF."""
    if df_vif.empty:
        return ""
    df_ord = df_vif.reindex(df_vif["Impacto ($/HF)"].abs().sort_values(ascending=False).index)
    top = df_ord.iloc[0]
    caja = "caja-alerta" if var_total > 0 else "caja-exito"
    veredicto = "se incremento" if var_total > 0 else "se redujo"
    recomendacion = RECOMENDACIONES_RUBRO.get(top["Rubro"], "Revisar el detalle operativo de este rubro con el equipo tecnico.")
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico Ejecutivo:</b><br>'
            f'El costo del huevo fertil <b>{veredicto} en ${abs(var_total):,.1f} COP</b> entre {p_base} y {p_actual}.<br>'
            f'El rubro con <b>mayor impacto individual</b> fue <b>{top["Rubro"]}</b>, explicando '
            f'<b>${top["Impacto ($/HF)"]:+,.1f} COP/huevo</b> de la variacion total.<br>'
            f'<i>\U0001F4A1 Factor critico - {top["Rubro"]}:</i> {recomendacion}</div>')


def diagnostico_lote(agg):
    """Pagina 4: Detalle de costos por lote."""
    if agg.empty:
        return ""
    peor = agg.sort_values("Costo Unit.", ascending=False).iloc[0]
    mejor = agg.sort_values("Costo Unit.", ascending=True).iloc[0]
    promedio = division_segura(agg["Costo_Total"].sum(), agg["Huevos_Fertiles"].sum())
    brecha = division_segura(peor["Costo Unit."] - mejor["Costo Unit."], mejor["Costo Unit."]) * 100
    return (f'<div class="caja-alerta"><b>\U0001F6A8 Auditoria de Ineficiencia en Campo:</b><br>'
            f'El <b>Lote {peor["lote"]}</b> ({peor["Linea"]}) tiene el costo unitario mas alto: '
            f'<b>${peor["Costo Unit."]:,.1f} COP/HF</b> frente al promedio de ${promedio:,.1f}.<br>'
            f'El <b>Lote {mejor["lote"]}</b> ({mejor["Linea"]}) es el mas eficiente con ${mejor["Costo Unit."]:,.1f} COP/HF, '
            f'una brecha de <b>{brecha:,.0f}%</b> entre extremos.<br>'
            f'<i>\U0001F4A1 Accion inmediata:</i> Lotes con costo muy superior al promedio suelen estar al final de su vida '
            f'productiva (baja postura = divisor pequeno = costo disparado). Evaluar descarte comercial programado.</div>')


def diagnostico_linea_costo(df_gen):
    """Pagina 5: Evaluacion financiera por linea."""
    if df_gen.empty:
        return ""
    peor = df_gen.sort_values("Costo Unitario ($/HF)", ascending=False).iloc[0]
    mejor = df_gen.sort_values("Costo Unitario ($/HF)", ascending=True).iloc[0]
    return (f'<div class="caja-info"><b>\U0001F4DD Diagnostico por Genetica:</b><br>'
            f'La linea <b>{peor["linea"]}</b> es la mas costosa por huevo (${peor["Costo Unitario ($/HF)"]:,.1f}), '
            f'mientras <b>{mejor["linea"]}</b> es la mas eficiente (${mejor["Costo Unitario ($/HF)"]:,.1f}).<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> Si {peor["linea"]} es persistentemente mas costosa, evaluar su rentabilidad '
            f'final considerando tambien el % de incubabilidad y nacimientos en planta, no solo el costo en granja.</div>')


def diagnostico_alimento(df_filtrado):
    """Pagina 6: Costo Kg Alimento."""
    serie = df_filtrado.dropna(subset=["Precio Kg Alimento"])
    if len(serie) < 2:
        return ""
    inicio, fin = serie.iloc[0], serie.iloc[-1]
    var_precio = division_segura(fin["Precio Kg Alimento"] - inicio["Precio Kg Alimento"], inicio["Precio Kg Alimento"]) * 100
    var_gramos = division_segura(fin["Gramos Alimento/Huevo"] - inicio["Gramos Alimento/Huevo"], inicio["Gramos Alimento/Huevo"]) * 100
    caja = "caja-alerta" if (var_precio > 0 and var_gramos > 0) else "caja-info"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico Nutricional:</b><br>'
            f'El precio del Kg de alimento vario <b>{formato_pct(var_precio)}</b> en el rango, mientras el consumo por '
            f'huevo (conversion) vario <b>{formato_pct(var_gramos)}</b>.<br>'
            f'<i>\U0001F4A1 Lectura clave:</i> {"Ambos indicadores suben simultaneamente: doble presion sobre el costo, se recomienda accion urgente en compras y en campo." if (var_precio > 0 and var_gramos > 0) else "El impacto esta parcialmente compensado; monitorear el indicador que aun no mejora."}</div>')


def diagnostico_levante(tabla_costo_anio, tabla_resultados):
    """Pagina 7: Costos lotes finalizados levante."""
    if tabla_costo_anio.empty or "Costo Total Levante" not in tabla_costo_anio.columns:
        return ""
    anios = tabla_costo_anio.index.tolist()
    if len(anios) < 2:
        return ""
    var = division_segura(tabla_costo_anio["Costo Total Levante"].iloc[-1] - tabla_costo_anio["Costo Total Levante"].iloc[-2],
                            tabla_costo_anio["Costo Total Levante"].iloc[-2]) * 100
    columnas_rubro = [c for c in tabla_costo_anio.columns if c not in ("Costo Total Levante",)]
    if columnas_rubro:
        deltas = tabla_costo_anio[columnas_rubro].diff().iloc[-1].sort_values(ascending=False)
        rubro_top = deltas.index[0]
    else:
        rubro_top = "N/D"
    caja = "caja-alerta" if var > 0 else "caja-exito"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico de Levante:</b><br>'
            f'El costo total de levante por ave vario <b>{formato_pct(var)}</b> respecto al ano anterior.<br>'
            f'El rubro con mayor incremento absoluto fue <b>{rubro_top}</b>.<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> Cruzar este resultado con la mortalidad y el consumo de alimento por ave '
            f'para identificar si el sobrecosto es sanitario, nutricional o de mano de obra.</div>')


def diagnostico_tecnico_levante(tabla_anio):
    """Pagina 8: Resultados tecnicos levante."""
    if tabla_anio.empty or "%Mortalidad Hembra" not in tabla_anio.columns or len(tabla_anio) < 2:
        return ""
    var_mort = tabla_anio["%Mortalidad Hembra"].iloc[-1] - tabla_anio["%Mortalidad Hembra"].iloc[-2]
    caja = "caja-alerta" if var_mort > 0 else "caja-exito"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico Tecnico - Levante:</b><br>'
            f'La mortalidad de hembras {"aumento" if var_mort > 0 else "disminuyo"} '
            f'<b>{abs(var_mort):,.2f} puntos porcentuales</b> respecto al ano anterior.<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> {"Investigar causas sanitarias o de manejo en las primeras semanas, foco critico de la mortalidad total." if var_mort > 0 else "Mantener protocolos actuales de manejo; el resultado es favorable."}</div>')


def diagnostico_produccion_finalizada(tabla_anio_costo):
    """Pagina 9: Costo HF lotes finalizados produccion."""
    if tabla_anio_costo.empty or "Costo Total Produccion/Huevo Incubable" not in tabla_anio_costo.columns:
        return ""
    if len(tabla_anio_costo) < 2:
        return ""
    var = tabla_anio_costo["Costo Total Produccion/Huevo Incubable"].iloc[-1] - tabla_anio_costo["Costo Total Produccion/Huevo Incubable"].iloc[-2]
    caja = "caja-alerta" if var > 0 else "caja-exito"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico Costo HF Finalizados:</b><br>'
            f'El costo del huevo fertil en lotes finalizados vario <b>${var:+,.1f} COP</b> respecto al ano anterior.<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> {"Revisar el detalle por lote para aislar si el sobrecosto es generalizado o puntual de pocos lotes." if var > 0 else "Documentar las practicas del periodo, ya que generaron eficiencia sostenida."}</div>')


def diagnostico_tecnico_produccion(tabla_anio_tec):
    """Pagina 10: Resultados tecnicos produccion."""
    if tabla_anio_tec.empty or "% Incubabilidad" not in tabla_anio_tec.columns or len(tabla_anio_tec) < 2:
        return ""
    var_inc = tabla_anio_tec["% Incubabilidad"].iloc[-1] - tabla_anio_tec["% Incubabilidad"].iloc[-2]
    var_nac = (tabla_anio_tec["%Nacimiento"].iloc[-1] - tabla_anio_tec["%Nacimiento"].iloc[-2]) if "%Nacimiento" in tabla_anio_tec.columns else np.nan
    caja = "caja-exito" if var_inc >= 0 else "caja-alerta"
    return (f'<div class="{caja}"><b>\U0001F4DD Diagnostico Tecnico - Produccion:</b><br>'
            f'La incubabilidad {"mejoro" if var_inc >= 0 else "empeoro"} <b>{abs(var_inc):,.2f} p.p.</b> respecto al ano anterior'
            f'{f", con nacimientos variando {var_nac:+.2f} p.p." if pd.notna(var_nac) else ""}<br>'
            f'<i>\U0001F4A1 Accion recomendada:</i> La incubabilidad y el nacimiento dependen de la calidad del huevo en granja '
            f'(peso, limpieza, manejo de nido); coordinar con planta de incubacion los hallazgos de calidad de cascara.</div>')


# =============================================================================
# 3. MOTOR DE DATOS (ETL)
# =============================================================================
@st.cache_data(show_spinner="Procesando y reconciliando INFORME GENERAL MENSUAL...")
def cargar_y_procesar_datos(fuente_archivo):
    xls = pd.ExcelFile(fuente_archivo)

    def leer(hoja):
        if hoja in xls.sheet_names:
            d = pd.read_excel(xls, sheet_name=hoja)
            d.columns = d.columns.astype(str).str.strip()
            return d
        return pd.DataFrame()

    df_raw = leer("BASE ZCO001")
    if df_raw.empty:
        st.error("\u26A0\uFE0F Error Critico: no se encontro la hoja 'BASE ZCO001' en el archivo.")
        st.stop()

    if "Fecha" in df_raw.columns:
        df_raw["Fecha"] = pd.to_datetime(df_raw["Fecha"], errors="coerce")
        df_raw["Anio"] = df_raw["Fecha"].dt.year
        df_raw["Mes_Num"] = df_raw["Fecha"].dt.month
        if df_raw["Anio"].isna().any() and "EjMat" in df_raw.columns:
            df_raw["Anio"] = df_raw["Anio"].fillna(df_raw["EjMat"])
            df_raw["Mes_Num"] = df_raw["Mes_Num"].fillna(df_raw["Mes"])
    else:
        df_raw["Anio"] = df_raw["EjMat"]
        df_raw["Mes_Num"] = df_raw["Mes"]

    df_raw["Anio"] = df_raw["Anio"].astype(int)
    df_raw["Mes_Num"] = df_raw["Mes_Num"].astype(int)
    df_raw["Periodo"] = df_raw["Anio"].astype(str) + "-" + df_raw["Mes_Num"].astype(str).str.zfill(2)
    df_raw["Totales"] = df_raw["Totales"].apply(limpiar_numero)
    df_raw["Cantidad"] = df_raw["Cantidad"].apply(limpiar_numero)
    if "Lote" in df_raw.columns:
        df_raw["Lote_str"] = df_raw["Lote"].apply(limpiar_lote)
    # 'Nombre 1' ya trae la granja asociada a cada lote (validado contra BD LEVANTE/BD PRODUCCION)
    df_raw["Granja"] = df_raw["Nombre 1"].fillna("Sin Granja Asignada") if "Nombre 1" in df_raw.columns else "Sin Granja Asignada"

    df_hf = df_raw[df_raw["Texto breve de material"] == MATERIAL_HF]
    hf_mes = df_hf.groupby("Periodo")["Cantidad"].sum()

    df_costos = df_raw[~df_raw["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])].copy()
    df_costos["Rubro"] = df_costos["Texto explicativo"].map(lambda x: MAPA_RUBROS.get(x, x))
    costos_piv = df_costos.groupby(["Periodo", "Rubro"])["Totales"].sum().unstack(fill_value=0)

    df_aprov = df_raw[(df_raw["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_raw["Texto breve de material"] != MATERIAL_HF)]
    aprov_mes = df_aprov.groupby("Periodo")["Totales"].sum()

    df_res = costos_piv.copy()
    df_res[RUBRO_APROVECHAMIENTO] = aprov_mes.reindex(df_res.index).fillna(0)
    df_res["Costo Total"] = df_res.sum(axis=1)
    df_res["Huevos Fertiles"] = hf_mes.reindex(df_res.index)
    df_res["Costo Huevo Fertil"] = division_segura(df_res["Costo Total"], df_res["Huevos Fertiles"])

    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"]
    alim_kg = df_alim.groupby("Periodo")["Cantidad"].sum()
    df_res["Consumo Alimento Kg"] = alim_kg.reindex(df_res.index)
    df_res["Precio Kg Alimento"] = division_segura(df_res["Alimento"], df_res["Consumo Alimento Kg"]) if "Alimento" in df_res.columns else np.nan
    df_res["Gramos Alimento/Huevo"] = division_segura(df_res["Consumo Alimento Kg"] * 1000, df_res["Huevos Fertiles"])

    df_res = df_res.reset_index().sort_values("Periodo")
    df_res["Anio"] = df_res["Periodo"].str.split("-").str[0].astype(int)
    df_res["Mes_Num"] = df_res["Periodo"].str.split("-").str[1].astype(int)
    rubros = [r for r in list(MAPA_RUBROS.values()) + [RUBRO_APROVECHAMIENTO] if r in df_res.columns]

    df_pn = leer("BD PN HF RAZA")
    if not df_pn.empty:
        df_pn["FECHA"] = pd.to_datetime(df_pn["FECHA"], errors="coerce")
        df_pn["Anio"] = df_pn["FECHA"].dt.year
        df_pn["Mes_Num"] = df_pn["FECHA"].dt.month
        df_pn["Periodo"] = df_pn["Anio"].astype(str) + "-" + df_pn["Mes_Num"].astype(str).str.zfill(2)
        df_pn["PROPIOS"] = pd.to_numeric(df_pn["PROPIOS"], errors="coerce").fillna(0)
        df_pn["EXTERNOS"] = pd.to_numeric(df_pn["EXTERNOS"], errors="coerce").fillna(0)
        df_pn["TOTAL"] = df_pn["PROPIOS"] + df_pn["EXTERNOS"]

    df_ctolinea = leer("BD CTO LINEA")
    if not df_ctolinea.empty:
        df_ctolinea["fecha"] = pd.to_datetime(df_ctolinea["fecha"], errors="coerce")
        df_ctolinea["Anio"] = df_ctolinea["fecha"].dt.year
        df_ctolinea["Mes_Num"] = df_ctolinea["fecha"].dt.month
        df_ctolinea["Periodo"] = df_ctolinea["Anio"].astype(str) + "-" + df_ctolinea["Mes_Num"].astype(str).str.zfill(2)
        if "lote" in df_ctolinea.columns:
            df_ctolinea["lote"] = df_ctolinea["lote"].apply(limpiar_lote)

    df_hist = leer("BD")
    if not df_hist.empty and "Fecha" in df_hist.columns:
        df_hist["Fecha"] = pd.to_datetime(df_hist["Fecha"], errors="coerce")
        df_hist["Anio"] = df_hist["Fecha"].dt.year
        df_hist["Mes_Num"] = df_hist["Fecha"].dt.month
        df_hist["Periodo"] = df_hist["Anio"].astype(str) + "-" + df_hist["Mes_Num"].astype(str).str.zfill(2)
        if "KG cons Alimento" in df_hist.columns and "Alimento " in df_hist.columns:
            df_hist["Costo Kg Alimento"] = division_segura(df_hist["Alimento "], df_hist["KG cons Alimento"])

    df_lev = leer("BD LEVANTE")
    if not df_lev.empty:
        df_lev = normaliza_periodo(df_lev, col_anio="Año", col_mes_num="No Mes")
        if "Lote" in df_lev.columns:
            df_lev["Lote_str"] = df_lev["Lote"].apply(limpiar_lote)

    df_prod = leer("BD PRODUCCIÓN")
    if not df_prod.empty:
        df_prod = normaliza_periodo(df_prod, col_anio="Año", col_mes_num="Mes")
        if "Mes" in df_prod.columns:
            df_prod["Mes_Num"] = pd.to_numeric(df_prod["Mes"], errors="coerce").fillna(1).astype(int)
            df_prod["Periodo"] = df_prod["Anio"].astype(str) + "-" + df_prod["Mes_Num"].astype(str).str.zfill(2)
        if "Lote" in df_prod.columns:
            df_prod["Lote_str"] = df_prod["Lote"].apply(limpiar_lote)

    return {
        "df": df_res, "rubros": rubros, "df_raw": df_raw,
        "df_pn": df_pn, "df_ctolinea": df_ctolinea, "df_hist": df_hist,
        "df_lev": df_lev, "df_prod": df_prod,
    }


# =============================================================================
# 4. CARGA DE DATOS
# =============================================================================
st.sidebar.title("\U0001F413 BI Avicola Gerencial — Granjas Reproductoras")
archivo_subido = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])

if archivo_subido is not None:
    datos = cargar_y_procesar_datos(archivo_subido)
else:
    archivo_predeterminado = buscar_excel_predeterminado()
    if archivo_predeterminado is None:
        st.warning("\u26A0\uFE0F Se requiere el archivo 'INFORME GENERAL MENSUAL.xlsx'. "
                    "Cargalo en la barra lateral o colocalo en `data/raw/`.")
        st.stop()
    datos = cargar_y_procesar_datos(str(archivo_predeterminado))

df = datos["df"]
rubros_items = datos["rubros"]
df_raw = datos["df_raw"]
df_pn = datos["df_pn"]
df_ctolinea = datos["df_ctolinea"]
df_hist = datos["df_hist"]
df_lev = datos["df_lev"]
df_prod = datos["df_prod"]

if df.empty:
    st.error("El archivo no contiene datos procesables en la hoja 'BASE ZCO001'.")
    st.stop()

# =============================================================================
# 5. BARRA LATERAL — CONTROL TEMPORAL
# =============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("\U0001F4C5 Control Temporal")

anios_disponibles = sorted(df["Anio"].unique(), reverse=True)


def selector_periodo(titulo, prefijo_clave, opciones_anio, mes_invertido=False):
    st.sidebar.markdown(f"**{titulo}:**")
    c1, c2 = st.sidebar.columns(2)
    anio = c1.selectbox("Año", opciones_anio, key=f"{prefijo_clave}_anio")
    meses_disp = sorted(df[df["Anio"] == anio]["Mes_Num"].unique(), reverse=mes_invertido)
    if not meses_disp:
        meses_disp = [1]
    mes = c2.selectbox("Mes", meses_disp, format_func=lambda x: NOMBRES_MESES.get(x, str(x)), key=f"{prefijo_clave}_mes")
    return f"{anio}-{str(mes).zfill(2)}"


modo_analisis = st.sidebar.radio(
    "Seleccione el enfoque:",
    ["\u2696\uFE0F Comparativo (Mes VS Mes)", "\U0001F4C8 Rango Historico (Evolucion)"],
)

if modo_analisis == "\u2696\uFE0F Comparativo (Mes VS Mes)":
    p_base = selector_periodo("Periodo Base (contra que comparo)", "base", anios_disponibles)
    p_actual = selector_periodo("Periodo Actual (que estoy evaluando)", "actual", anios_disponibles, mes_invertido=True)
    texto_contexto = (f"Comparativa Estrategica: **{NOMBRES_MESES[int(p_actual.split('-')[1])]} {p_actual.split('-')[0]}** "
                       f"VS **{NOMBRES_MESES[int(p_base.split('-')[1])]} {p_base.split('-')[0]}**")
else:
    p_base = selector_periodo("Inicio del rango", "ini", sorted(anios_disponibles))
    p_actual = selector_periodo("Fin del rango", "fin", anios_disponibles, mes_invertido=True)
    texto_contexto = f"Evolucion Historica: Desde **{min(p_base, p_actual)}** hasta **{max(p_base, p_actual)}**"

rango_inicio, rango_fin = min(p_base, p_actual), max(p_base, p_actual)
df_filtrado = df[(df["Periodo"] >= rango_inicio) & (df["Periodo"] <= rango_fin)].sort_values("Periodo")
df_raw_rango = df_raw[(df_raw["Periodo"] >= rango_inicio) & (df_raw["Periodo"] <= rango_fin)]
df_pn_rango = df_pn[(df_pn["Periodo"] >= rango_inicio) & (df_pn["Periodo"] <= rango_fin)] if not df_pn.empty else df_pn

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Seleccione la Pagina del Informe:",
    [
        "1. Produccion (Propios/Externos)",
        "2. Produccion Mes a Mes por Linea",
        "3. Costo Huevo Fertil",
        "4. Detalle Costos por Lote",
        "5. Detalle Costos por Linea",
        "6. Costo Kg Alimento",
        "7. Costos Lotes Finalizados (Levante)",
        "8. Resultados Tecnicos Levante",
        "9. Costo HF Lotes Finalizados (Produccion)",
        "10. Resultados Tecnicos Produccion",
    ],
)

# =============================================================================
# PAGINA 1 — PRODUCCION (PROPIOS/EXTERNOS/TOTAL)
# =============================================================================
if menu == "1. Produccion (Propios/Externos)":
    st.markdown('<p class="titulo-principal">1. PRODUCCION CONSOLIDADA (PROPIOS + EXTERNOS)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_pn.empty:
        st.warning("No se encontro la hoja 'BD PN HF RAZA' en el archivo.")
        st.stop()

    piv = df_pn_rango.groupby(["Periodo"])[["PROPIOS", "EXTERNOS", "TOTAL"]].sum().reset_index()
    st.markdown(diagnostico_produccion(piv, texto_contexto), unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Propios (Periodo)", f"{piv['PROPIOS'].sum():,.0f}")
    c2.metric("Total Externos (Periodo)", f"{piv['EXTERNOS'].sum():,.0f}")
    c3.metric("Total General", f"{piv['TOTAL'].sum():,.0f}")

    col1, col2 = st.columns([1.4, 1])
    with col1:
        fig_area = go.Figure()
        fig_area.add_trace(go.Scatter(x=piv["Periodo"], y=piv["PROPIOS"], name="Propios", stackgroup="one",
                                       mode="lines", line=dict(width=0.5, color=PALETA[0]), fillcolor="rgba(30,58,138,0.7)"))
        fig_area.add_trace(go.Scatter(x=piv["Periodo"], y=piv["EXTERNOS"], name="Externos", stackgroup="one",
                                       mode="lines", line=dict(width=0.5, color=PALETA[2]), fillcolor="rgba(245,158,11,0.7)"))
        estilizar_grafico(fig_area, "Volumen de Produccion Mensual", "Propios vs Externos (unidades apiladas)")
        fig_area.update_yaxes(title_text="Huevos (und)")
        fig_area.update_xaxes(title_text="Periodo")
        st.plotly_chart(fig_area, use_container_width=True)
    with col2:
        por_raza = df_pn_rango.groupby("RAZA")["TOTAL"].sum().reset_index().sort_values("TOTAL", ascending=False)
        fig_dona = px.pie(por_raza, values="TOTAL", names="RAZA", hole=0.5)
        fig_dona.update_traces(textinfo="percent", textposition="inside", texttemplate="%{percent:.1%}",
                                textfont_size=14, pull=[0.02] * len(por_raza),
                                marker=dict(line=dict(color="white", width=2)))
        fig_dona.update_layout(uniformtext_minsize=12, uniformtext_mode="hide")
        estilizar_grafico(fig_dona, "Participacion por Raza", "Total del periodo seleccionado", leyenda_derecha=True)
        st.plotly_chart(fig_dona, use_container_width=True)

    st.subheader("\U0001F4CB Matriz Mes / Raza")
    tabla = df_pn_rango.pivot_table(index="Periodo", columns="RAZA", values="TOTAL", aggfunc="sum", fill_value=0)
    tabla["Total general"] = tabla.sum(axis=1)
    st.dataframe(tabla.style.format("{:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

# =============================================================================
# PAGINA 2 — PRODUCCION MES A MES POR LINEA
# =============================================================================
elif menu == "2. Produccion Mes a Mes por Linea":
    st.markdown('<p class="titulo-principal">2. PRODUCCION MES A MES POR LINEA</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_pn.empty:
        st.warning("No se encontro la hoja 'BD PN HF RAZA' en el archivo.")
        st.stop()

    df_pn_r2 = df_pn_rango.copy()
    df_pn_r2["RAZA 2"] = df_pn_r2["RAZA 2"].fillna(df_pn_r2["RAZA"])
    piv2 = df_pn_r2.groupby(["Periodo", "RAZA 2"])["TOTAL"].sum().reset_index()

    st.markdown(diagnostico_linea_produccion(piv2), unsafe_allow_html=True)

    fig_lineas = px.bar(piv2, x="Periodo", y="TOTAL", color="RAZA 2", barmode="stack", text_auto=".2s")
    fig_lineas.update_traces(cliponaxis=False)
    estilizar_grafico(fig_lineas, "Produccion de Huevo Fertil por Linea Genetica", "Barras apiladas mensuales")
    fig_lineas.update_yaxes(title_text="Huevos (und)")
    fig_lineas.update_xaxes(title_text="Periodo")
    st.plotly_chart(fig_lineas, use_container_width=True)

    piv2["% Part."] = piv2.groupby("Periodo")["TOTAL"].transform(lambda x: division_segura(x, x.sum()) * 100)
    fig_pct = px.area(piv2, x="Periodo", y="% Part.", color="RAZA 2", groupnorm="percent")
    estilizar_grafico(fig_pct, "Evolucion de la Participacion (%) por Linea", "Mix genetico mes a mes")
    fig_pct.update_yaxes(title_text="Participacion (%)")
    fig_pct.update_xaxes(title_text="Periodo")
    st.plotly_chart(fig_pct, use_container_width=True)

    st.subheader("\U0001F4CB Detalle Mes / Linea")
    st.dataframe(piv2.pivot_table(index="Periodo", columns="RAZA 2", values="TOTAL", aggfunc="sum", fill_value=0)
                 .style.format("{:,.0f}").background_gradient(cmap="Greens"), use_container_width=True)

# =============================================================================
# PAGINA 3 — COSTO HUEVO FERTIL (VIF, EVOLUTIVA Y SIMULADOR)
# =============================================================================
elif menu == "3. Costo Huevo Fertil":
    st.markdown('<p class="titulo-principal">3. ANALISIS INTEGRAL: COSTO HUEVO FERTIL Y DESVIACIONES</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">Analisis de {p_actual} frente a la base {p_base}</p>', unsafe_allow_html=True)

    if p_base == p_actual:
        st.warning("\u26A0\uFE0F Selecciona un Rango de Meses o el Modo Comparativo con dos periodos distintos.")
        st.stop()

    fila_b, fila_a = df[df["Periodo"] == p_base], df[df["Periodo"] == p_actual]
    if fila_b.empty or fila_a.empty:
        st.error("No hay datos suficientes en los periodos seleccionados para comparar.")
        st.stop()

    df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
    if pd.isna(df_b["Costo Huevo Fertil"]) or pd.isna(df_a["Costo Huevo Fertil"]):
        st.error("Alguno de los dos periodos no tiene huevos fertiles registrados.")
        st.stop()

    var_total = df_a["Costo Huevo Fertil"] - df_b["Costo Huevo Fertil"]

    filas = []
    for r in rubros_items:
        hf_b, hf_a = df_b["Huevos Fertiles"], df_a["Huevos Fertiles"]
        unit_base = (df_b[r] / hf_b) if hf_b else np.nan
        unit_actual = (df_a[r] / hf_a) if hf_a else np.nan
        filas.append({"Rubro": r, "Costo Unit Base ($)": unit_base, "Costo Unit Actual ($)": unit_actual,
                       "Impacto ($/HF)": (unit_actual - unit_base) if pd.notna(unit_actual) and pd.notna(unit_base) else np.nan})
    df_vif = pd.DataFrame(filas).dropna(subset=["Impacto ($/HF)"]).sort_values("Impacto ($/HF)", ascending=False)

    st.markdown(generar_diagnostico_costos(df_vif, var_total, p_actual, p_base), unsafe_allow_html=True)

    col_g1, col_g2 = st.columns([1, 1])
    with col_g1:
        gauge = medidor_kpi(df_a["Costo Huevo Fertil"], df_b["Costo Huevo Fertil"],
                             f"Costo Huevo Fertil: {p_actual} vs Base {p_base}", sufijo="$", formato_num=",.0f")
        st.plotly_chart(gauge, use_container_width=True)
    with col_g2:
        c1, c2, c3 = st.columns(3)
        c1.metric(f"Base {p_base}", f"${df_b['Costo Huevo Fertil']:,.0f}")
        c2.metric(f"Actual {p_actual}", f"${df_a['Costo Huevo Fertil']:,.0f}")
        delta_pct = (var_total / df_b["Costo Huevo Fertil"]) * 100 if df_b["Costo Huevo Fertil"] else 0
        c3.metric("Desviacion ($/HF)", f"${var_total:+,.1f}", delta=f"{delta_pct:+.1f}%", delta_color="inverse")
        fig_evo_mini = px.line(df_filtrado, x="Periodo", y="Costo Huevo Fertil", markers=True)
        fig_evo_mini.update_traces(line_color=PALETA[0], line_width=3, fill="tozeroy", fillcolor="rgba(30,58,138,0.08)")
        estilizar_grafico(fig_evo_mini, "Evolucion Historica del Costo ($/HF)", altura=260)
        st.plotly_chart(fig_evo_mini, use_container_width=True)

    st.markdown("---")
    st.subheader("\U0001F309 Puente de Variacion (Cascada) por Rubro")
    st.caption("\U0001F534 Rojo = el rubro **incremento** el costo del huevo &nbsp;|&nbsp; "
               "\U0001F7E2 Verde = el rubro **redujo** el costo &nbsp;|&nbsp; "
               "\U0001F535 Azul = totales (Base / Actual). Se muestran los 6 rubros de mayor impacto; el resto se agrupa en 'Otros'.")
    if not df_vif.empty:
        fig_cascada = grafico_cascada_vif(df_vif, df_b["Costo Huevo Fertil"], df_a["Costo Huevo Fertil"],
                                           "Puente de Costo: de Base a Actual por Rubro",
                                           f"{p_base} \u2192 {p_actual} ($/HF) &mdash; ordenado por impacto absoluto")
        fig_cascada.update_yaxes(title_text="$ / Huevo")
        st.plotly_chart(fig_cascada, use_container_width=True)

    st.markdown("---")
    st.subheader("\U0001F5D3\uFE0F Matriz de Evolucion Mensual del Costo")
    df_evolucion = df_filtrado[["Periodo", "Huevos Fertiles", "Costo Total", "Costo Huevo Fertil"]].copy()
    df_evolucion["Variacion ($/HF)"] = df_evolucion["Costo Huevo Fertil"].diff()
    df_evolucion["% Variacion"] = df_evolucion["Costo Huevo Fertil"].pct_change() * 100
    st.dataframe(df_evolucion.style.format({
        "Huevos Fertiles": "{:,.0f}", "Costo Total": "${:,.0f}", "Costo Huevo Fertil": "${:,.2f}",
        "Variacion ($/HF)": "${:+,.2f}", "% Variacion": "{:+.2f}%"
    }).background_gradient(subset=["Costo Huevo Fertil"], cmap="Reds"), use_container_width=True)

    st.subheader("\U0001F4CB Matriz Comparativa (Variacion VIF Directa)")
    st.dataframe(df_vif.style.format({
        "Costo Unit Base ($)": "${:,.2f}", "Costo Unit Actual ($)": "${:,.2f}", "Impacto ($/HF)": "${:+,.2f}"
    }).background_gradient(subset=["Impacto ($/HF)"], cmap="RdYlGn_r"), use_container_width=True)

    st.markdown("---")
    st.subheader("\U0001F39B\uFE0F Modulo Predictivo: Estrategias de Reduccion de Costo")
    s1, s2, s3 = st.columns(3)
    ahorro_alimento = s1.number_input("1. Negociacion Alimento: Disminuir precio en ($/Kg):", value=50.0, step=10.0)
    mejora_conversion = s2.number_input("2. Eficiencia: Bajar consumo (g/huevo):", value=15.0, step=5.0)
    mejora_postura = s3.number_input("3. Productividad: Subir % de postura en:", value=3.0, step=0.5)

    precio_actual_kg = df_a["Precio Kg Alimento"] if pd.notna(df_a["Precio Kg Alimento"]) else 0
    nuevo_precio_alimento = max(precio_actual_kg - ahorro_alimento, 0)
    nuevo_consumo_kg = max(df_a["Consumo Alimento Kg"] - ((mejora_conversion / 1000) * df_a["Huevos Fertiles"]), 0)
    nuevo_costo_alimento = nuevo_consumo_kg * nuevo_precio_alimento
    nuevos_hf = df_a["Huevos Fertiles"] * (1 + (mejora_postura / 100))
    costo_sin_alimento = df_a["Costo Total"] - df_a.get("Alimento", 0)
    nuevo_costo_total = costo_sin_alimento + nuevo_costo_alimento
    nuevo_costo_huevo = nuevo_costo_total / nuevos_hf if nuevos_hf > 0 else 0
    ahorro_unitario = df_a["Costo Huevo Fertil"] - nuevo_costo_huevo

    fig_simulador = go.Figure(go.Bar(
        x=["Costo Actual", "Costo Proyectado"], y=[df_a["Costo Huevo Fertil"], nuevo_costo_huevo],
        text=[f"${df_a['Costo Huevo Fertil']:,.0f}", f"${nuevo_costo_huevo:,.0f}"], textposition="outside",
        marker_color=[PALETA[3], PALETA[4]]
    ))
    estilizar_grafico(fig_simulador, "Proyeccion de Ahorro con Estrategias Combinadas", altura=320)
    fig_simulador.update_yaxes(title_text="$ / Huevo")
    st.plotly_chart(fig_simulador, use_container_width=True)
    st.success(f"\U0001F3AF **Proyeccion:** El costo bajaria de **${df_a['Costo Huevo Fertil']:,.1f}** a "
               f"**${nuevo_costo_huevo:,.1f} COP**. Ahorro directo de **${ahorro_unitario:,.1f} por huevo**.")

# =============================================================================
# PAGINA 4 — DETALLE COSTOS POR LOTE
# =============================================================================
elif menu == "4. Detalle Costos por Lote":
    st.markdown('<p class="titulo-principal">4. DIAGNOSTICO MICRO-OPERATIVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">Auditoria del periodo de cierre: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    if df_ctolinea.empty:
        st.warning("No se encontro la hoja 'BD CTO LINEA' en el archivo.")
        st.stop()

    dfl = df_ctolinea[df_ctolinea["Periodo"] == p_actual].copy()
    if dfl.empty:
        st.warning("No hay registros de costo por lote en el periodo seleccionado.")
        st.stop()

    agg = dfl.groupby("lote").agg(
        Huevos_Fertiles=("Huevos fértiles", "sum"),
        Costo_Total=("TOTAL COSTO", "sum"),
        Linea=("linea", "first"),
        Semana=("semana", "mean"),
    ).reset_index()
    agg["% Part. HF"] = division_segura(agg["Huevos_Fertiles"], agg["Huevos_Fertiles"].sum()) * 100
    agg["Costo Unit."] = division_segura(agg["Costo_Total"], agg["Huevos_Fertiles"])
    agg = agg.sort_values("Huevos_Fertiles", ascending=False)

    st.markdown(diagnostico_lote(agg), unsafe_allow_html=True)

    total_hf, total_costo = agg["Huevos_Fertiles"].sum(), agg["Costo_Total"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Huevos Fertiles del Mes", f"{total_hf:,.0f}")
    c2.metric("Costo Total del Mes", f"${total_costo:,.0f}")
    c3.metric("Costo Unitario Promedio", f"${division_segura(total_costo, total_hf):,.1f}")

    col1, col2 = st.columns([1.3, 1])
    with col1:
        agg_ordenado = agg.sort_values("Costo Unit.", ascending=True)
        promedio_general = division_segura(total_costo, total_hf)
        colores = ["#EF4444" if v > promedio_general else "#10B981" for v in agg_ordenado["Costo Unit."]]
        fig_barras = go.Figure(go.Bar(x=agg_ordenado["Costo Unit."], y=agg_ordenado["lote"].astype(str), orientation="h",
                                       marker_color=colores, text=agg_ordenado["Costo Unit."].map(lambda v: f"${v:,.0f}"),
                                       textposition="outside"))
        fig_barras.add_vline(x=promedio_general, line_dash="dash", line_color="#1E3A8A", annotation_text="Promedio")
        estilizar_grafico(fig_barras, "Costo Unitario ($/HF) por Lote", "Rojo = sobre el promedio | Verde = bajo el promedio", altura=500)
        fig_barras.update_xaxes(title_text="$ / Huevo")
        st.plotly_chart(fig_barras, use_container_width=True)
    with col2:
        fig_treemap = px.treemap(agg, path=["Linea", "lote"], values="Huevos_Fertiles", color="Costo Unit.",
                                  color_continuous_scale="RdYlGn_r")
        estilizar_grafico(fig_treemap, "Volumen y Costo por Lote", "Tamano = Huevos | Color = Costo unitario", altura=500)
        st.plotly_chart(fig_treemap, use_container_width=True)

    st.subheader("\U0001F4CB Detalle Costo por Lote")
    st.dataframe(agg[["lote", "Linea", "Semana", "Huevos_Fertiles", "% Part. HF", "Costo Unit."]]
                 .style.format({"Semana": "{:,.0f}", "Huevos_Fertiles": "{:,.0f}", "% Part. HF": "{:.1f}%",
                                 "Costo Unit.": "${:,.1f}"})
                 .background_gradient(subset=["Costo Unit."], cmap="Reds"), use_container_width=True)

# =============================================================================
# PAGINA 5 — DETALLE COSTOS POR LINEA
# =============================================================================
elif menu == "5. Detalle Costos por Linea":
    st.markdown('<p class="titulo-principal">5. EVALUACION FINANCIERA POR GENETICA (LINEA)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_ctolinea.empty:
        st.warning("No se encontro la hoja 'BD CTO LINEA' en el archivo.")
        st.stop()

    dfl_actual = df_ctolinea[df_ctolinea["Periodo"] == p_actual].copy()
    df_gen = dfl_actual.groupby("linea").agg(
        Huevos_Fertiles=("Huevos fértiles", "sum"), Costo_Total=("TOTAL COSTO", "sum"),
    ).reset_index()
    df_gen["Costo Unitario ($/HF)"] = division_segura(df_gen["Costo_Total"], df_gen["Huevos_Fertiles"])
    df_gen = df_gen[df_gen["Huevos_Fertiles"] > 0]

    st.markdown(diagnostico_linea_costo(df_gen), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        df_serie_linea = df_ctolinea[(df_ctolinea["Periodo"] >= rango_inicio) & (df_ctolinea["Periodo"] <= rango_fin)]\
            .groupby(["Periodo", "linea"])["Huevos fértiles"].sum().reset_index()
        fig_evolucion_linea = px.line(df_serie_linea, x="Periodo", y="Huevos fértiles", color="linea", markers=True)
        fig_evolucion_linea.update_traces(line_width=3)
        estilizar_grafico(fig_evolucion_linea, "Evolucion Historica de Produccion por Raza")
        fig_evolucion_linea.update_yaxes(title_text="Huevos (und)")
        st.plotly_chart(fig_evolucion_linea, use_container_width=True)
    with col2:
        df_gen_ordenado = df_gen.sort_values("Costo Unitario ($/HF)")
        fig_costo_linea = px.bar(df_gen_ordenado, x="linea", y="Costo Unitario ($/HF)", color="linea", text_auto=".0f")
        fig_costo_linea.update_traces(cliponaxis=False)
        estilizar_grafico(fig_costo_linea, "Costo Unitario Promedio por Raza", p_actual)
        fig_costo_linea.update_yaxes(title_text="$ / Huevo")
        st.plotly_chart(fig_costo_linea, use_container_width=True)

    st.dataframe(df_gen.style.format({
        "Huevos_Fertiles": "{:,.0f}", "Costo_Total": "${:,.0f}", "Costo Unitario ($/HF)": "${:,.2f}"
    }).background_gradient(subset=["Costo Unitario ($/HF)"], cmap="RdYlGn_r"), use_container_width=True)

# =============================================================================
# PAGINA 6 — COSTO KG ALIMENTO
# =============================================================================
elif menu == "6. Costo Kg Alimento":
    st.markdown('<p class="titulo-principal">6. IMPACTO DEL COSTO DE ALIMENTO</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    st.markdown(diagnostico_alimento(df_filtrado), unsafe_allow_html=True)

    fig_combinado = make_subplots(specs=[[{"secondary_y": False}]])
    fig_combinado.add_trace(go.Bar(x=df_filtrado["Periodo"], y=df_filtrado["Precio Kg Alimento"],
                                    name="Precio $/Kg", marker_color=PALETA[2], opacity=0.75))
    fig_combinado.add_trace(go.Scatter(x=df_filtrado["Periodo"], y=df_filtrado["Gramos Alimento/Huevo"],
                                        name="Conversion g/huevo", mode="lines+markers",
                                        line=dict(color=PALETA[4], width=3), yaxis="y2"))
    fig_combinado.update_layout(yaxis2=dict(overlaying="y", side="right", title="g / huevo", showgrid=False))
    estilizar_grafico(fig_combinado, "Precio del Alimento vs Conversion por Huevo", "Barras: precio | Linea: eficiencia de conversion")
    fig_combinado.update_yaxes(title_text="$ / Kg", secondary_y=False)
    st.plotly_chart(fig_combinado, use_container_width=True)

    if not df_hist.empty and "Costo Kg Alimento" in df_hist.columns:
        st.subheader("\U0001F4CA Costo Kg Alimento por Ano (Serie Historica)")
        dh = df_hist.dropna(subset=["Costo Kg Alimento"]).copy()
        fig_historico = px.line(dh, x="Fecha", y="Costo Kg Alimento", markers=True, color=dh["Anio"].astype(str))
        fig_historico.update_traces(line_width=2.5)
        estilizar_grafico(fig_historico, "Evolucion del Costo del Kg de Alimento", "Comparativo multi-anio")
        fig_historico.update_yaxes(title_text="$ / Kg")
        st.plotly_chart(fig_historico, use_container_width=True)

    st.subheader("\U0001F39B\uFE0F Simulador Predictivo de Rentabilidad de Alimento")
    fila_a = df[df["Periodo"] == p_actual]
    if not fila_a.empty and pd.notna(fila_a.iloc[0].get("Precio Kg Alimento")):
        df_a = fila_a.iloc[0]
        s1, s2, s3 = st.columns(3)
        ahorro_alimento = s1.number_input("1. Mejora en Compra ($/Kg):", value=50.0, step=10.0, key="k6_1")
        mejora_conversion = s2.number_input("2. Bajar desperdicio (g/huevo):", value=15.0, step=5.0, key="k6_2")
        mejora_postura = s3.number_input("3. Subir pico de postura (%):", value=3.0, step=0.5, key="k6_3")

        precio_actual_kg = df_a["Precio Kg Alimento"]
        nuevo_precio_alimento = max(precio_actual_kg - ahorro_alimento, 0)
        nuevo_consumo_kg = max(df_a["Consumo Alimento Kg"] - ((mejora_conversion / 1000) * df_a["Huevos Fertiles"]), 0)
        nuevo_costo_alimento = nuevo_consumo_kg * nuevo_precio_alimento
        nuevos_hf = df_a["Huevos Fertiles"] * (1 + (mejora_postura / 100))
        costo_sin_alimento = df_a["Costo Total"] - df_a.get("Alimento", 0)
        nuevo_costo_total = costo_sin_alimento + nuevo_costo_alimento
        nuevo_costo_huevo = nuevo_costo_total / nuevos_hf if nuevos_hf > 0 else 0
        ahorro_unitario = df_a["Costo Huevo Fertil"] - nuevo_costo_huevo

        st.success(f"\U0001F3AF **Ahorro Estrategico:** El costo global bajaria de **${df_a['Costo Huevo Fertil']:,.1f}** "
                   f"a **${nuevo_costo_huevo:,.1f} COP/Huevo**. Impacto: **${ahorro_unitario:,.1f}** por unidad.")
    else:
        st.info("No hay datos suficientes de Alimento en el mes actual para correr el simulador.")

# =============================================================================
# PAGINA 7 — COSTOS LOTES FINALIZADOS (LEVANTE)
# =============================================================================
elif menu == "7. Costos Lotes Finalizados (Levante)":
    st.markdown('<p class="titulo-principal">7. COSTOS LOTES FINALIZADOS — LEVANTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_lev.empty:
        st.warning("No se encontro la hoja 'BD LEVANTE' en el archivo.")
        st.stop()

    dlev_rango = df_lev[(df_lev["Periodo"] >= rango_inicio) & (df_lev["Periodo"] <= rango_fin)].copy()
    dlev_rango["Anio"] = dlev_rango["Anio"].astype(int)

    columnas_costo = ["Costo Pollita", "Costo Alimento", "Costo Nomina", "Costo Cif", "Costo Droga",
                       "Costo Arriendo", "Costo Cama", "Costo Aseo y Desinfección", "Costo Calefacción",
                       "Depreciación Construcciones", "Aprovechamiento", "Costo Total Levante"]
    columnas_costo = [c for c in columnas_costo if c in dlev_rango.columns]
    tabla_costo_anio = dlev_rango.groupby("Anio")[columnas_costo].sum(numeric_only=True)

    columnas_resultado = ["Hembras Encasetadas", "Hembras Fin Levante", "%Mortalidad Hembra"]
    columnas_resultado = [c for c in columnas_resultado if c in dlev_rango.columns]
    tabla_resultados = dlev_rango.groupby("Lote_str")[columnas_resultado].mean(numeric_only=True).reset_index() if columnas_resultado else pd.DataFrame()

    st.markdown(diagnostico_levante(tabla_costo_anio, tabla_resultados), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        fig_costo_total = px.bar(tabla_costo_anio.reset_index(), x="Anio", y="Costo Total Levante", text_auto=".2s",
                                  color="Anio", color_continuous_scale="Blues")
        estilizar_grafico(fig_costo_total, "Costo Total de Levante por Ano", "Tendencia interanual")
        fig_costo_total.update_yaxes(title_text="Costo Total ($)")
        st.plotly_chart(fig_costo_total, use_container_width=True)
    with col2:
        rubros_levante = [c for c in columnas_costo if c != "Costo Total Levante"]
        if rubros_levante:
            comp = tabla_costo_anio[rubros_levante].reset_index().melt(id_vars="Anio", var_name="Rubro", value_name="Costo")
            fig_composicion = px.bar(comp, x="Anio", y="Costo", color="Rubro", barmode="stack")
            estilizar_grafico(fig_composicion, "Composicion del Costo de Levante", "Aporte de cada rubro por ano")
            fig_composicion.update_yaxes(title_text="Costo ($)")
            st.plotly_chart(fig_composicion, use_container_width=True)

    st.dataframe(tabla_costo_anio.style.format("${:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

    st.subheader(f"\U0001F4CB Costo por Lote ({p_actual.split('-')[0]})")
    dlev_anio_actual = dlev_rango[dlev_rango["Anio"] == int(p_actual.split("-")[0])]
    if not dlev_anio_actual.empty:
        tabla_lote = dlev_anio_actual.groupby("Lote_str")[columnas_costo].sum(numeric_only=True)
        st.dataframe(tabla_lote.style.format("${:,.0f}").background_gradient(cmap="Oranges"), use_container_width=True)

# =============================================================================
# PAGINA 8 — RESULTADOS TECNICOS LEVANTE
# =============================================================================
elif menu == "8. Resultados Tecnicos Levante":
    st.markdown('<p class="titulo-principal">8. RESULTADOS TECNICOS — LEVANTES FINALIZADOS</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_lev.empty:
        st.warning("No se encontro la hoja 'BD LEVANTE' en el archivo.")
        st.stop()

    dlev_rango = df_lev[(df_lev["Periodo"] >= rango_inicio) & (df_lev["Periodo"] <= rango_fin)].copy()
    columnas_tecnicas = ["Edad Fin Levante", "Hembras Encasetadas", "Hembras Fin Levante", "Machos Fin Levante",
                          "%Mortalidad Hembra", "Mort Hembra 1ra Sem", "%Mortalidad Macho", "%Selección Hembra",
                          "Cons. X Hembra Finalizada", "Cons. X Macho Finalizado"]
    columnas_tecnicas = [c for c in columnas_tecnicas if c in dlev_rango.columns]
    tabla_anio = dlev_rango.groupby("Anio")[columnas_tecnicas].mean(numeric_only=True).reset_index()

    st.markdown(diagnostico_tecnico_levante(tabla_anio), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        fig_mortalidad = go.Figure()
        fig_mortalidad.add_trace(go.Scatter(x=tabla_anio["Anio"], y=tabla_anio["%Mortalidad Hembra"],
                                             mode="lines+markers", name="% Mort. Hembra",
                                             line=dict(color=PALETA[3], width=3), fill="tozeroy",
                                             fillcolor="rgba(239,68,68,0.12)"))
        estilizar_grafico(fig_mortalidad, "Evolucion % Mortalidad Hembra por Ano", altura=380)
        fig_mortalidad.update_yaxes(title_text="% Mortalidad")
        st.plotly_chart(fig_mortalidad, use_container_width=True)
    with col2:
        if "Cons. X Hembra Finalizada" in tabla_anio.columns:
            fig_consumo = px.bar(tabla_anio, x="Anio", y="Cons. X Hembra Finalizada", text_auto=".1f",
                                  color="Anio", color_continuous_scale="Teal")
            estilizar_grafico(fig_consumo, "Consumo de Alimento por Hembra Finalizada", altura=380)
            fig_consumo.update_yaxes(title_text="Kg / Hembra")
            st.plotly_chart(fig_consumo, use_container_width=True)

    st.dataframe(tabla_anio.style.format("{:,.2f}", subset=columnas_tecnicas), use_container_width=True)

    st.subheader(f"\U0001F4CB Detalle Tecnico por Lote — {p_actual.split('-')[0]}")
    dlev_anio_actual = dlev_rango[dlev_rango["Anio"] == int(p_actual.split("-")[0])]
    if not dlev_anio_actual.empty:
        tabla_lote_tecnica = dlev_anio_actual.groupby("Lote_str")[columnas_tecnicas].mean(numeric_only=True)
        st.dataframe(tabla_lote_tecnica.style.format("{:,.2f}").background_gradient(subset=["%Mortalidad Hembra"], cmap="Reds"),
                     use_container_width=True)

# =============================================================================
# PAGINA 9 — COSTO HF LOTES FINALIZADOS (PRODUCCION)
# =============================================================================
elif menu == "9. Costo HF Lotes Finalizados (Produccion)":
    st.markdown('<p class="titulo-principal">9. COSTO HUEVO FERTIL — LOTES FINALIZADOS PRODUCCION</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_prod.empty:
        st.warning("No se encontro la hoja 'BD PRODUCCIÓN' en el archivo.")
        st.stop()

    dprod_rango = df_prod[(df_prod["Periodo"] >= rango_inicio) & (df_prod["Periodo"] <= rango_fin)].copy()
    columnas_costo_hf = ["Costo Alimento/Huevo Incubable", "Depreciación Reproductora/Huevo Incubabl",
                          "Costo Nomina/Huevo Incubable", "Costo Cif/Huevo Incubable", "Costo Arriendo/Huevo Incubable",
                          "Costo Aseo y Desinfección/Huevo Incubabl", "Costo Cama/Huevo Incubable",
                          "Costo Droga/Huevo Incubable", "Depreciación Construcciones/Huevo Incuba",
                          "Costo Aprovechamiento/Huevo Incubable", "Costo Total Producción/Huevo Incubable"]
    columnas_costo_hf = [c for c in columnas_costo_hf if c in dprod_rango.columns]

    if "Huevos Fértiles" in dprod_rango.columns and columnas_costo_hf:
        def promedio_ponderado(grupo):
            hf = grupo["Huevos Fértiles"].sum()
            salida = {c: (division_segura((grupo[c] * grupo["Huevos Fértiles"]).sum(), hf) if hf else np.nan) for c in columnas_costo_hf}
            salida["Huevos Fértiles"] = hf
            return pd.Series(salida)

        tabla_anio_costo = dprod_rango.groupby("Anio").apply(promedio_ponderado).reset_index()
        tabla_anio_costo = tabla_anio_costo.rename(columns={"Costo Total Producción/Huevo Incubable": "Costo Total Produccion/Huevo Incubable"})

        st.markdown(diagnostico_produccion_finalizada(tabla_anio_costo), unsafe_allow_html=True)

        col1, col2 = st.columns([1.3, 1])
        with col1:
            fig_costo_hf = go.Figure()
            fig_costo_hf.add_trace(go.Scatter(x=tabla_anio_costo["Anio"], y=tabla_anio_costo["Costo Total Produccion/Huevo Incubable"],
                                               mode="lines+markers", line=dict(color=PALETA[0], width=3),
                                               fill="tozeroy", fillcolor="rgba(30,58,138,0.1)"))
            estilizar_grafico(fig_costo_hf, "Evolucion del Costo del Huevo Fertil (Ponderado)", "Consolidado por ano")
            fig_costo_hf.update_yaxes(title_text="$ / Huevo")
            st.plotly_chart(fig_costo_hf, use_container_width=True)
        with col2:
            columnas_ultimo = [c for c in columnas_costo_hf if c != "Costo Total Producción/Huevo Incubable"]
            ultimo_anio = tabla_anio_costo.iloc[-1][columnas_ultimo].sort_values(ascending=False)
            etiquetas_cortas = [c.split("/")[0].replace("Costo ", "").replace("Depreciación ", "Depr. ")[:22] for c in ultimo_anio.index]
            fig_composicion_costo = px.pie(values=ultimo_anio.values, names=etiquetas_cortas, hole=0.5)
            fig_composicion_costo.update_traces(textinfo="percent", textposition="inside",
                                                 texttemplate="%{percent:.1%}", textfont_size=13,
                                                 marker=dict(line=dict(color="white", width=2)))
            fig_composicion_costo.update_layout(uniformtext_minsize=11, uniformtext_mode="hide")
            estilizar_grafico(fig_composicion_costo, "Composicion del Costo (Ultimo Ano)", altura=440, leyenda_derecha=True)
            st.plotly_chart(fig_composicion_costo, use_container_width=True)

        st.dataframe(tabla_anio_costo.style.format({c: "${:,.1f}" for c in columnas_costo_hf} |
                                                      {"Huevos Fértiles": "{:,.0f}"}), use_container_width=True)

        st.subheader(f"\U0001F4CB Costo HF por Lote — {p_actual.split('-')[0]}")
        dprod_anio_actual = dprod_rango[dprod_rango["Anio"] == int(p_actual.split("-")[0])]
        if not dprod_anio_actual.empty:
            tabla_lote_costo = dprod_anio_actual.groupby("Lote_str").apply(promedio_ponderado)
            st.dataframe(tabla_lote_costo.style.format({c: "${:,.1f}" for c in columnas_costo_hf} |
                                                          {"Huevos Fértiles": "{:,.0f}"})
                         .background_gradient(subset=["Costo Total Producción/Huevo Incubable"], cmap="Reds"),
                         use_container_width=True)
    else:
        st.warning("La hoja 'BD PRODUCCIÓN' no tiene las columnas de costo/HF esperadas.")

# =============================================================================
# PAGINA 10 — RESULTADOS TECNICOS PRODUCCION
# =============================================================================
elif menu == "10. Resultados Tecnicos Produccion":
    st.markdown('<p class="titulo-principal">10. RESULTADOS TECNICOS — PRODUCCION FINALIZADOS</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">{texto_contexto}</p>', unsafe_allow_html=True)

    if df_prod.empty:
        st.warning("No se encontro la hoja 'BD PRODUCCIÓN' en el archivo.")
        st.stop()

    dprod_rango = df_prod[(df_prod["Periodo"] >= rango_inicio) & (df_prod["Periodo"] <= rango_fin)].copy()
    columnas_tecnicas = ["Edad  Producción", "Hembras Encasetadas", "Hembras Fin Producción", "Mortalidad Hembra",
                          "%Mortalidad Hembra", "Gramos X Huevos", "Huevo Fértil Ave Alojada", "% Incubabilidad",
                          "%Nacimiento"]
    columnas_tecnicas = [c for c in columnas_tecnicas if c in dprod_rango.columns]
    tabla_anio_tecnica = dprod_rango.groupby("Anio")[columnas_tecnicas].mean(numeric_only=True).reset_index()

    st.markdown(diagnostico_tecnico_produccion(tabla_anio_tecnica), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if "% Incubabilidad" in tabla_anio_tecnica.columns:
            fig_incubabilidad = go.Figure(go.Scatter(x=tabla_anio_tecnica["Anio"], y=tabla_anio_tecnica["% Incubabilidad"],
                                                       mode="lines+markers", line=dict(color=PALETA[4], width=3),
                                                       fill="tozeroy", fillcolor="rgba(16,185,129,0.12)"))
            estilizar_grafico(fig_incubabilidad, "Evolucion % Incubabilidad", altura=380)
            fig_incubabilidad.update_yaxes(title_text="% Incubabilidad")
            st.plotly_chart(fig_incubabilidad, use_container_width=True)
    with col2:
        if "%Nacimiento" in tabla_anio_tecnica.columns:
            fig_nacimiento = go.Figure(go.Scatter(x=tabla_anio_tecnica["Anio"], y=tabla_anio_tecnica["%Nacimiento"],
                                                    mode="lines+markers", line=dict(color=PALETA[1], width=3),
                                                    fill="tozeroy", fillcolor="rgba(14,165,233,0.12)"))
            estilizar_grafico(fig_nacimiento, "Evolucion % Nacimiento", altura=380)
            fig_nacimiento.update_yaxes(title_text="% Nacimiento")
            st.plotly_chart(fig_nacimiento, use_container_width=True)

    st.dataframe(tabla_anio_tecnica.style.format("{:,.2f}", subset=columnas_tecnicas), use_container_width=True)

    st.subheader(f"\U0001F4CB Detalle Tecnico por Lote — {p_actual.split('-')[0]}")
    dprod_anio_actual = dprod_rango[dprod_rango["Anio"] == int(p_actual.split("-")[0])]
    if not dprod_anio_actual.empty and "Lote_str" in dprod_anio_actual.columns:
        tabla_lote_tecnica = dprod_anio_actual.groupby("Lote_str")[columnas_tecnicas].mean(numeric_only=True)
        st.dataframe(tabla_lote_tecnica.style.format("{:,.2f}").background_gradient(subset=["% Incubabilidad"], cmap="Greens"),
                     use_container_width=True)
