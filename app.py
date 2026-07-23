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
  2) Graficos Plotly de alto impacto visual (Tornado, medidor, treemap, combinado, dona)
  3) Tablas detalladas con formato condicional
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
# 0. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# =============================================================================
st.set_page_config(
    page_title="Dashboard Ejecutivo - Costos Huevo Fértil",
    page_icon="🐔",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
    .titulo-principal { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-bottom: 5px; text-transform: uppercase;}
    .subtitulo { font-size: 15px; color: #4B5563; margin-bottom: 20px; font-style: italic; }
    div[data-testid="stMetric"] {
        background-color: #F8FAFC; padding: 15px; border-radius: 8px;
        border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .caja-info    { background-color: #F0F9FF; padding: 15px; border-left: 5px solid #0284C7; border-radius: 5px; margin-bottom: 20px; color: #075985; font-size: 14px;}
    .caja-alerta  { background-color: #FEF2F2; padding: 15px; border-left: 5px solid #EF4444; border-radius: 5px; margin-bottom: 20px; color: #7F1D1D; font-size: 14px;}
    .caja-exito   { background-color: #ECFDF5; padding: 15px; border-left: 5px solid #10B981; border-radius: 5px; margin-bottom: 20px; color: #064E3B; font-size: 14px;}
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    .dataframe th { background-color: #1E3A8A !important; color: white !important; font-weight: bold !important; text-align: center !important;}
    </style>
""", unsafe_allow_html=True)

PALETA = ["#1E3A8A", "#0EA5E9", "#F59E0B", "#EF4444", "#10B981",
          "#8B5CF6", "#EC4899", "#14B8A6", "#6366F1", "#84CC16"]
px.defaults.color_discrete_sequence = PALETA

NOMBRES_MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                  7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
MAPA_MESES_TEXTO = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
                     'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}

def estilizar_grafico(fig, titulo, subtitulo="", altura=430):
    """Aplica un estilo consistente y de alto impacto a cualquier figura Plotly."""
    texto_titulo = f"{titulo}"
    if subtitulo:
        texto_titulo += f"<br><span style='font-size:13px;color:#6B7280;font-weight:normal'>{subtitulo}</span>"
    fig.update_layout(
        title={"text": texto_titulo, "x": 0.02, "xanchor": "left"},
        height=altura,
        # Reubicamos la leyenda abajo para evitar que pise los títulos largos
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
        margin=dict(t=90, l=10, r=10, b=60), 
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Arial", size=13, color="#1F2937"),
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

def grafico_tornado_vif(df_vif, costo_base, costo_actual, titulo, subtitulo=""):
    """Construye un gráfico de Barras de Impacto (Tornado) de alta claridad visual."""
    df_plot = df_vif[df_vif["Impacto ($/HF)"].abs() > 0.1].copy()
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
        width=0.6
    ))

    fig = estilizar_grafico(fig, titulo, subtitulo, altura=480)
    fig.add_vline(x=0, line_width=2, line_color="#1E3A8A")
    fig.update_xaxes(title_text="Impacto en el Costo Unitario ($/Huevo)", showgrid=True, gridcolor="#E2E8F0")
    fig.update_yaxes(title_text="", showline=False)
    
    if not df_plot.empty:
        max_val = df_plot["Impacto ($/HF)"].abs().max() * 1.25
        fig.update_xaxes(range=[-max_val if df_plot["Impacto ($/HF)"].min() < 0 else 0, max_val])

    return fig

# =============================================================================
# 1. UTILIDADES DE DATOS
# =============================================================================
MATERIAL_HF = "HUEVO INCUBABLE"
TEXTO_LIQUIDACION = "CTA PTE LIQ. ORD PCC Y MAQUILAS"
TEXTO_DIFERENCIA_PRECIO = "DIFERENCIA EN PRECIO PRODUCTOS SEMIELABO"
RUBRO_APROVECHAMIENTO = "Aprovechamiento"

MAPA_RUBROS = {
    "CONSUMO ALIMENTO": "Alimento",
    "PP Depr. Gallina Grj.Pcc.": "Depreciación Huevo",
    "PP Horas Hombre Grj.Pcc.": "Mano de obra",
    "PP Costos Ind. Grj.Pcc.": "Indirectos",
    "PP Costos Arriendo Grj.Pcc.": "Arriendo",
    "CONSUMO CAMA": "Cama",
    "ELEMENTOS DE ASEO Y DESINFECCION": "Aseo y desinfección",
    "CONSUMO DROGA": "Droga",
    "PP Costos Depr. Grj.Pcc.": "Depreciacion Const. Y Edif.",
    "CONSUMOS MATERIA PRIMA": "Materia prima",
}

def limpiar_numero(x):
    if isinstance(x, (int, float, np.integer, np.floating)): return float(x)
    if pd.isna(x): return 0.0
    s = str(x).strip().replace("-", "").replace(",", "")
    try: v = float(s)
    except ValueError: return 0.0
    return -v if str(x).strip().endswith("-") else v

def division_segura(numerador, denominador):
    if isinstance(denominador, pd.Series): denominador = denominador.replace(0, np.nan)
    elif denominador == 0 or pd.isna(denominador): return np.nan
    return numerador / denominador

def limpiar_lote(s):
    if pd.isna(s): return "S/N"
    s = str(s).strip().upper()
    if s.endswith(".0"): s = s[:-2]
    return s

def buscar_excel_predeterminado():
    for carpeta in (Path("data/raw"), Path(".")):
        if carpeta.exists():
            encontrados = sorted(f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$"))
            if encontrados: return encontrados[0]
    return None

def normaliza_periodo(df, col_anio="Año", col_mes_num="No Mes", col_mes_texto="Mes", anio_defecto=2026):
    if col_anio in df.columns: df["Anio"] = pd.to_numeric(df[col_anio], errors="coerce")
    if col_mes_num in df.columns: df["Mes_Num"] = pd.to_numeric(df[col_mes_num], errors="coerce")
    elif col_mes_texto in df.columns: df["Mes_Num"] = df[col_mes_texto].astype(str).str.upper().str.strip().map(MAPA_MESES_TEXTO)
    if "Anio" not in df.columns: df["Anio"] = anio_defecto
    df["Anio"] = df["Anio"].fillna(anio_defecto).astype(int)
    df["Mes_Num"] = df["Mes_Num"].fillna(1).astype(int)
    df["Periodo"] = df["Anio"].astype(str) + "-" + df["Mes_Num"].astype(str).str.zfill(2)
    return df

# =============================================================================
# MOTOR DE DATOS (ETL)
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
        df_raw["Anio"] = df_raw.get("EjMat", 2026)
        df_raw["Mes_Num"] = df_raw.get("Mes", 6)

    df_raw["Anio"] = df_raw["Anio"].astype(int)
    df_raw["Mes_Num"] = df_raw["Mes_Num"].astype(int)
    df_raw["Periodo"] = df_raw["Anio"].astype(str) + "-" + df_raw["Mes_Num"].astype(str).str.zfill(2)
    df_raw["Totales"] = df_raw["Totales"].apply(limpiar_numero)
    df_raw["Cantidad"] = df_raw["Cantidad"].apply(limpiar_numero)
    if "Lote" in df_raw.columns: df_raw["Lote_str"] = df_raw["Lote"].apply(limpiar_lote)
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
        if "lote" in df_ctolinea.columns: df_ctolinea["lote"] = df_ctolinea["lote"].apply(limpiar_lote)

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
        if "Lote" in df_lev.columns: df_lev["Lote_str"] = df_lev["Lote"].apply(limpiar_lote)

    df_prod = leer("BD PRODUCCIÓN")
    if not df_prod.empty:
        df_prod = normaliza_periodo(df_prod, col_anio="Año", col_mes_num="Mes")
        if "Mes" in df_prod.columns:
            df_prod["Mes_Num"] = pd.to_numeric(df_prod["Mes"], errors="coerce").fillna(1).astype(int)
            df_prod["Periodo"] = df_prod["Anio"].astype(str) + "-" + df_prod["Mes_Num"].astype(str).str.zfill(2)
        if "Lote" in df_prod.columns: df_prod["Lote_str"] = df_prod["Lote"].apply(limpiar_lote)

    return {
        "df": df_res, "rubros": rubros, "df_raw": df_raw,
        "df_pn": df_pn, "df_ctolinea": df_ctolinea, "df_hist": df_hist,
        "df_lev": df_lev, "df_prod": df_prod,
    }


# =============================================================================
# 4. CARGA DE DATOS Y BARRA LATERAL
# =============================================================================
st.sidebar.title("🐔 BI Avícola Gerencial")

archivo_subido = st.sidebar.file_uploader("Actualizar Data (.xlsx)", type=["xlsx", "xls"])
file_to_load = archivo_subido if archivo_subido else buscar_excel_predeterminado()

if file_to_load is None:
    st.warning("⚠️ Se requiere el archivo 'INFORME GENERAL MENSUAL.xlsx'.")
    st.stop()

try:
    datos = cargar_y_procesar_datos(file_to_load)
    df_res, rubros_items, df_raw = datos["df"], datos["rubros"], datos["df_raw"]
    df_pn, df_ctolinea, df_hist = datos["df_pn"], datos["df_ctolinea"], datos["df_hist"]
    df_lev, df_prod = datos["df_lev"], datos["df_prod"]
    if df_res.empty: st.stop()
except Exception as e:
    st.error(f"⚠️ Error al procesar el archivo: {e}")
    st.stop()

st.sidebar.markdown("---")
anios_disp = sorted(df_res["Anio"].unique(), reverse=True)
st.sidebar.markdown("**📅 Control Temporal**")
c1, c2 = st.sidebar.columns(2)
anio_act = c1.selectbox("Año", anios_disp, key="anio_act")
meses_disp = sorted(df_res[df_res["Anio"] == anio_act]["Mes_Num"].unique(), reverse=True)
mes_act = c2.selectbox("Mes", meses_disp, format_func=lambda x: NOMBRES_MESES[x], key="mes_act")
p_actual = f"{anio_act}-{str(mes_act).zfill(2)}"

# Control base para el reporte de EVOLUCIÓN (Página 2)
st.sidebar.markdown("**🔍 Base Comparativa (Solo para Inf. 2)**")
anio_base = st.sidebar.selectbox("Año Base", anios_disp, key="anio_base")
mes_base = st.sidebar.selectbox("Mes Base", sorted(df_res[df_res["Anio"] == anio_base]["Mes_Num"].unique()), format_func=lambda x: NOMBRES_MESES[x], key="mes_base")
p_base = f"{anio_base}-{str(mes_base).zfill(2)}"

texto_contexto = f"Cierre Operativo: **{NOMBRES_MESES[mes_act]} {anio_act}**"

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "📊 Informes Directivos:",
    [
        "1. PRODUCCIÓN GLOBAL Y MIX",
        "2. EVOLUCIÓN COSTO HUEVO FÉRTIL",
        "3. MICRO-AUDITORÍA COSTOS POR LOTE",
        "4. COSTO Y EFICIENCIA ALIMENTO",
        "5. RESULTADOS LEVANTE FINALIZADOS",
        "6. RESULTADOS PRODUCCIÓN FINALIZADOS"
    ]
)

# =============================================================================
# MÓDULOS DEL DASHBOARD (REPLICANDO IMÁGENES EXACTAS)
# =============================================================================

if menu == "1. PRODUCCIÓN GLOBAL Y MIX":
    st.markdown('<p class="titulo-principal">PRODUCCIÓN 2026</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="caja-info">
        <b>📝 Análisis Directivo - Apalancamiento Externo:</b><br>
        La raza <b>ROSS</b> lidera el mix genético operativo. Destaca la fuerte integración de producciones "Externas" (Maquilas) para sostener el volumen hacia incubación.
    </div>
    """, unsafe_allow_html=True)

    if not df_pn.empty and 'Periodo' in df_pn.columns:
        df_r = df_pn[df_pn['Anio'] == anio_act].copy()
        
        col_tab, col_graf = st.columns([1.2, 2])
        
        with col_tab:
            st.markdown("### Producción Mes a Mes")
            tabla_prod = df_r.groupby(["Mes_Num", "RAZA"])[["PROPIOS", "EXTERNOS"]].sum().reset_index()
            tabla_prod["TOTAL"] = tabla_prod["PROPIOS"] + tabla_prod["EXTERNOS"]
            tabla_prod["MES"] = tabla_prod["Mes_Num"].map(NOMBRES_MESES)
            
            st.dataframe(tabla_prod[["MES", "RAZA", "PROPIOS", "EXTERNOS", "TOTAL"]].style.format({
                "PROPIOS": "{:,.0f}", "EXTERNOS": "{:,.0f}", "TOTAL": "{:,.0f}"
            }), use_container_width=True, hide_index=True)

        with col_graf:
            st.markdown("### Total 2026 / % Participación")
            tot_raza = df_r.groupby("RAZA")[["PROPIOS", "EXTERNOS"]].sum().sum(axis=1).reset_index(name="TOTAL 2026")
            tot_raza["% PART"] = division_segura(tot_raza["TOTAL 2026"], tot_raza["TOTAL 2026"].sum())
            
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=tot_raza["RAZA"], y=tot_raza["TOTAL 2026"], name="TOTAL 2026", text=tot_raza["TOTAL 2026"].apply(lambda x: f"{x:,.0f}"), textposition="auto"), secondary_y=False)
            fig.add_trace(go.Scatter(x=tot_raza["RAZA"], y=tot_raza["% PART"], name="% PART", mode="markers+text", text=tot_raza["% PART"].apply(lambda x: f"{x:.2%}" if pd.notna(x) else ""), textposition="top center", marker=dict(color="black", size=10)), secondary_y=True)
            
            fig = estilizar_grafico(fig, "")
            fig.update_yaxes(visible=False, secondary_y=True)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("### PRODUCCIÓN MES A MES POR LÍNEA 2026 (Propios + Externos)")
        df_r["TOTAL"] = df_r["PROPIOS"] + df_r["EXTERNOS"]
        df_linea = df_r.groupby(["Mes_Num", "RAZA"])["TOTAL"].sum().reset_index()
        df_linea["MES"] = df_linea["Mes_Num"].map(NOMBRES_MESES)
        
        # Agrupar las razas principales para el gráfico de barras (ROSS vs COBB)
        df_linea["GRUPO_RAZA"] = df_linea["RAZA"].apply(lambda x: "ROSS" if "ROSS" in str(x).upper() else "COBB")
        df_graf2 = df_linea.groupby(["MES", "GRUPO_RAZA", "Mes_Num"])["TOTAL"].sum().reset_index().sort_values("Mes_Num")
        
        fig2 = px.bar(df_graf2, x="MES", y="TOTAL", color="GRUPO_RAZA", barmode="group", text_auto=".2s")
        fig2 = estilizar_grafico(fig2, "")
        fig2.update_layout(xaxis_title="", yaxis_title="Huevo Fértil")
        st.plotly_chart(fig2, use_container_width=True)

elif menu == "2. EVOLUCIÓN COSTO HUEVO FÉRTIL":
    st.markdown('<p class="titulo-principal">COSTO HUEVO FÉRTIL 2026 E HISTÓRICO</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="caja-info">
        💡 <b>Diagnóstico Financiero:</b> Identificación de los inductores de costo. Un incremento persistente en <b>Depreciación Huevo</b> o <b>Mano de Obra</b> explica la pérdida de margen operativo a pesar de posibles mejoras en el precio de la materia prima.
    </div>
    """, unsafe_allow_html=True)

    # 1. Tabla de Evolución Mensual
    df_y = df_res[df_res['Anio'] == anio_act].copy()
    if not df_y.empty:
        st.subheader(f"Evolución de Variables - {anio_act}")
        rubros_ord = ["Alimento", "Depreciación Huevo", "Mano de obra", "Indirectos", "Arriendo", "Cama", "Aseo y desinfección", "Droga", "Depreciacion Const. Y Edif.", "Materia prima", RUBRO_APROVECHAMIENTO]
        valid_rubros = [r for r in rubros_ord if r in df_y.columns]
        
        df_y_rubros = df_y.set_index('Mes_Num')[valid_rubros].T
        for col in df_y_rubros.columns:
            hf_mes = df_y.loc[df_y['Mes_Num'] == col, 'Huevos Fertiles'].values[0]
            df_y_rubros[col] = df_y_rubros[col] / hf_mes if hf_mes > 0 else 0
            
        df_y_rubros.columns = [NOMBRES_MESES.get(c, str(c)) for c in df_y_rubros.columns]
        df_y_rubros.loc["Costo HF"] = df_y_rubros.sum()
        
        cols = list(df_y_rubros.columns)
        if len(cols) >= 2:
            df_y_rubros["$ Var"] = df_y_rubros[cols[-1]] - df_y_rubros[cols[-2]]
            df_y_rubros["% Var"] = division_segura(df_y_rubros["$ Var"], df_y_rubros[cols[-2]])
        else:
            df_y_rubros["$ Var"] = 0.0; df_y_rubros["% Var"] = 0.0

        st.dataframe(df_y_rubros.style.format(formatter="{:,.1f}", subset=cols + ["$ Var"]).format(formatter="{:.1%}", subset=["% Var"]).background_gradient(subset=["$ Var"], cmap="RdYlGn_r"), use_container_width=True)

    # 2. Análisis VIF con TORNADO
    st.markdown("---")
    st.subheader("🌪️ Tornado de Variaciones por Rubro (Factores de Desviación)")
    
    if p_base == p_actual:
        st.warning(f"⚠️ Selecciona un mes base diferente en la barra lateral para ver la variación (actualmente comparando {p_base} vs {p_actual}).")
    else:
        fila_b, fila_a = df_res[df_res["Periodo"] == p_base], df_res[df_res["Periodo"] == p_actual]
        if not fila_b.empty and not fila_a.empty:
            df_b, df_a = fila_b.iloc[0], fila_a.iloc[0]
            
            filas_vif = []
            for r in rubros_items:
                hf_b, hf_a = df_b["Huevos Fertiles"], df_a["Huevos Fertiles"]
                ub = (df_b[r] / hf_b) if hf_b else np.nan
                ua = (df_a[r] / hf_a) if hf_a else np.nan
                filas_vif.append({
                    "Rubro": r, 
                    "Impacto ($/HF)": (ua - ub) if pd.notna(ua) and pd.notna(ub) else np.nan
                })
            
            df_vif = pd.DataFrame(filas_vif).dropna(subset=["Impacto ($/HF)"])
            
            if not df_vif.empty:
                fig_tornado = grafico_tornado_vif(
                    df_vif, 
                    df_b["Costo Huevo Fertil"], 
                    df_a["Costo Huevo Fertil"],
                    f"Explicación de la Variación de Costo (VIF)",
                    f"Factores que explican el paso de ${df_b['Costo Huevo Fertil']:,.1f} a ${df_a['Costo Huevo Fertil']:,.1f} por huevo ({p_base} -> {p_actual})."
                )
                st.plotly_chart(fig_tornado, use_container_width=True)

    # 3. Tabla Histórica
    st.markdown("---")
    st.subheader("Comparativa Histórica (Anual)")
    df_hist_g = df_res.groupby("Anio")[valid_rubros].sum()
    hf_hist_g = df_res.groupby("Anio")["Huevos Fertiles"].sum()
    
    for c in df_hist_g.columns:
        df_hist_g[c] = division_segura(df_hist_g[c], hf_hist_g)
    
    df_hist_g["Costo Huevo fértil"] = df_hist_g.sum(axis=1)
    df_hist_g = df_hist_g.T
    
    anios_cols = sorted(list(df_hist_g.columns))
    if len(anios_cols) >= 2:
        ult = anios_cols[-1]
        ant = anios_cols[-2]
        df_hist_g["$ Var"] = df_hist_g[ult] - df_hist_g[ant]
        df_hist_g["% Var"] = division_segura(df_hist_g["$ Var"], df_hist_g[ant])
        
        st.dataframe(df_hist_g.style.format(formatter="{:,.1f}", subset=anios_cols + ["$ Var"]).format(formatter="{:.1%}", subset=["% Var"]), use_container_width=True)

elif menu == "3. MICRO-AUDITORÍA COSTOS POR LOTE":
    st.markdown('<p class="titulo-principal">DETALLE COSTOS HUEVO POR LOTE</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="subtitulo">Cierre Operativo: <b>{p_actual}</b></p>', unsafe_allow_html=True)

    df_l = df_raw[(df_raw["Periodo"] == p_actual)].copy()
    if df_l.empty: st.stop()

    df_costo = df_l[~df_l["Texto explicativo"].isin([TEXTO_LIQUIDACION, TEXTO_DIFERENCIA_PRECIO])]
    df_aprov = df_l[(df_l["Texto explicativo"] == TEXTO_LIQUIDACION) & (df_l["Texto breve de material"] != MATERIAL_HF)]
    
    alimento_lote = df_costo[df_costo["Texto explicativo"] == "CONSUMO ALIMENTO"].groupby("Lote_str")["Totales"].sum()
    depre_lote = df_costo[df_costo["Texto explicativo"] == "PP Depr. Gallina Grj.Pcc."].groupby("Lote_str")["Totales"].sum()
    ind_lote = df_costo[df_costo["Texto explicativo"] == "PP Costos Ind. Grj.Pcc."].groupby("Lote_str")["Totales"].sum()
    
    c_lote = df_costo.groupby("Lote_str")["Totales"].sum()
    a_lote = df_aprov.groupby("Lote_str")["Totales"].sum()
    h_lote = df_l[df_l["Texto breve de material"] == MATERIAL_HF].groupby("Lote_str")["Cantidad"].sum()

    df_m = pd.DataFrame({
        "Huevos fértiles": h_lote, "Costo Base": c_lote, "Aprovechamientos": a_lote,
        "Total Alimento": alimento_lote, "Total Depreciación": depre_lote, "Total Indirectos": ind_lote
    }).fillna(0)
    
    df_m["Costo Total"] = df_m["Costo Base"] + df_m["Aprovechamientos"]
    df_m["Costo Unit."] = division_segura(df_m["Costo Total"], df_m["Huevos fértiles"])
    df_m = df_m[df_m["Huevos fértiles"] > 0].reset_index().rename(columns={"index": "Lote"})
    
    if not df_prod.empty:
        df_p_act = df_prod[df_prod["Periodo"] == p_actual].copy()
        df_p_act = df_p_act.groupby("Lote_str").last().reset_index()
        df_m = df_m.merge(df_p_act[["Lote_str", "Edad", "Gramos x Huevos", "RAZA"]], left_on="Lote", right_on="Lote_str", how="left")
    else:
        df_m["Edad"] = np.nan; df_m["Gramos x Huevos"] = np.nan; df_m["RAZA"] = "S/N"

    df_m["% Part. HF"] = division_segura(df_m["Huevos fértiles"], df_m["Huevos fértiles"].sum())
    df_m["$ Alimento"] = division_segura(df_m["Total Alimento"], df_m["Huevos fértiles"])
    df_m["% Alimento"] = division_segura(df_m["$ Alimento"], df_m["Costo Unit."])
    df_m["$ Depreciación"] = division_segura(df_m["Total Depreciación"], df_m["Huevos fértiles"])
    df_m["% Depreciación"] = division_segura(df_m["$ Depreciación"], df_m["Costo Unit."])
    df_m["$ Indirectos"] = division_segura(df_m["Total Indirectos"], df_m["Huevos fértiles"])
    df_m["% Indirectos"] = division_segura(df_m["$ Indirectos"], df_m["Costo Unit."])
    df_m["$ Otros"] = df_m["Costo Unit."] - df_m["$ Alimento"] - df_m["$ Depreciación"] - df_m["$ Indirectos"]
    df_m["% Otros"] = division_segura(df_m["$ Otros"], df_m["Costo Unit."])

    df_m = df_m.sort_values("Costo Unit.", ascending=False)
    
    st.markdown("""
    <div class="caja-alerta">
        🚨 <b>Auditoría de Ineficiencia en Campo:</b> Evaluar el descarte inmediato de los lotes ubicados en la parte superior de la matriz. Su baja productividad eleva exponencialmente el <b>% Depreciación</b>, mermando el margen de la compañía.
    </div>
    """, unsafe_allow_html=True)
    
    cols_mostrar = ["Lote", "Huevos fértiles", "% Part. HF", "Costo Unit.", "Edad", "Gramos x Huevos", "$ Alimento", "% Alimento", "$ Depreciación", "% Depreciación", "$ Indirectos", "% Indirectos", "$ Otros", "% Otros"]
    st.dataframe(df_m[[c for c in cols_mostrar if c in df_m.columns]].rename(columns={"Gramos x Huevos": "Consumo/HF"}).style.format({
        "Huevos fértiles": "{:,.0f}", "% Part. HF": "{:.1%}", "Costo Unit.": "${:,.1f}", "Edad": "{:.0f}", "Consumo/HF": "{:.1f}",
        "$ Alimento": "${:,.0f}", "% Alimento": "{:.1%}", "$ Depreciación": "${:,.0f}", "% Depreciación": "{:.1%}",
        "$ Indirectos": "${:,.0f}", "% Indirectos": "{:.1%}", "$ Otros": "${:,.0f}", "% Otros": "{:.1%}"
    }).background_gradient(subset=["Costo Unit."], cmap="Reds"), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### Matriz de Rubros por Línea y Lote")
    df_costo["Rubro"] = df_costo["Texto explicativo"].map(lambda x: MAPA_RUBROS.get(x, x))
    piv_raza = df_costo.groupby(["Lote_str", "Rubro"])["Totales"].sum().unstack(fill_value=0)
    piv_aprov = df_aprov.groupby("Lote_str")["Totales"].sum()
    piv_raza["SUBPRODUCTOS HF"] = piv_aprov
    piv_raza = piv_raza.merge(df_m[["Lote", "RAZA", "Huevos fértiles", "Edad", "Costo Unit."]], left_index=True, right_on="Lote")
    
    for c in piv_raza.columns:
        if c not in ["Lote", "RAZA", "Huevos fértiles", "Edad", "Costo Unit."]:
            piv_raza[c] = division_segura(piv_raza[c], piv_raza["Huevos fértiles"])
            
    piv_raza = piv_raza.set_index(["RAZA", "Lote"]).T
    st.dataframe(piv_raza.style.format("{:,.2f}"), use_container_width=True)

elif menu == "4. COSTO Y EFICIENCIA ALIMENTO":
    st.markdown('<p class="titulo-principal">COSTO KG ALIMENTO 2026-2025</p>', unsafe_allow_html=True)
    
    df_alim = df_raw[df_raw["Texto explicativo"] == "CONSUMO ALIMENTO"].copy()
    res_alim = df_alim.groupby(["Anio", "Mes_Num"]).apply(
        lambda x: x["Totales"].sum() / x["Cantidad"].sum() if x["Cantidad"].sum() > 0 else 0
    ).unstack(level=0)

    if 2025 in res_alim.columns and 2026 in res_alim.columns:
        res_df = res_alim[[2025, 2026]].dropna(how="all").reset_index()
        res_df["MES"] = res_df["Mes_Num"].map(NOMBRES_MESES)
        res_df["%VAR"] = division_segura(res_df[2026] - res_df[2025], res_df[2025])
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=res_df["MES"], y=res_df[2025], name="2025", marker_color="#A16207"), secondary_y=False)
        fig.add_trace(go.Bar(x=res_df["MES"], y=res_df[2026], name="2026", marker_color="#FACC15"), secondary_y=False)
        fig.add_trace(go.Scatter(x=res_df["MES"], y=res_df["%VAR"], name="%VAR", mode="lines+markers+text", line=dict(color="black", width=3), text=res_df["%VAR"].apply(lambda x: f"{x:.1%}" if pd.notna(x) else ""), textposition="top right"), secondary_y=True)
        
        fig = estilizar_grafico(fig, "Evolución Precio del Alimento y Variación")
        fig.update_layout(barmode='group', hovermode="x unified")
        fig.update_yaxes(title_text="Precio ($/Kg)", secondary_y=False, range=[1450, 1800])
        fig.update_yaxes(title_text="% Variación", secondary_y=True, tickformat=".1%")
        
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(res_df[["MES", 2025, 2026, "%VAR"]].style.format({
            2025: "${:,.1f}", 2026: "${:,.1f}", "%VAR": "{:+.1%}"
        }), use_container_width=True, hide_index=True)
    else:
        st.info("La comparativa requiere que el archivo contenga datos tanto del 2025 como del 2026.")

elif menu == "5. RESULTADOS LEVANTE FINALIZADOS":
    st.markdown('<p class="titulo-principal">RESULTADOS TÉCNICOS POR LOTE LEVANTES FINALIZADOS</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="caja-info">
        💡 <b>Mega-Informe de Capitalización (Crianza):</b> Esta sección consolida la viabilidad técnica y el peso financiero de la pollita que ingresará a producción. 
    </div>
    """, unsafe_allow_html=True)

    if not df_lev.empty:
        df_l = df_lev[df_lev["Anio"] == anio_act].copy()
        if not df_l.empty:
            st.subheader(f"📊 Detalle Técnico por Lote ({anio_act})")
            cols_tech = ["Lote_str", "Edad", "Hembras Encasetadas", "Total Aves Fin Levante", "Mortalidad Hembra", "% Mort Hembra", "% Mort Hembra 1ra Sem", "% Bajas H", "% Selección Hembra", "Consumo x Ave H", "Encasetados M", "Finalizados M", "Mortalidad M", "%Mort M", "%Mort Macho 1ra Sem", "% Bajas M", "Consumo x Ave M"]
            existentes = [c for c in cols_tech if c in df_l.columns]
            st.dataframe(df_l[existentes].set_index("Lote_str").T.style.format("{:,.2f}"), use_container_width=True)

            st.markdown("---")
            st.subheader("📈 Comparativa Técnica Histórica (2023-2026)")
            df_hist_tech = df_lev.groupby("Anio").agg(
                Edad_Fin_Levante=("Edad", "mean"),
                Encasetadas_H=("Hembras Encasetadas", "sum"),
                Finalizadas_H=("Total Aves Fin Levante", "sum") if "Total Aves Fin Levante" in df_lev.columns else ("Hembras Encasetadas", "sum"),
                Mortalidad_H=("Mortalidad Hembra", "sum")
            ).T
            st.dataframe(df_hist_tech.style.format("{:,.1f}"), use_container_width=True)

            st.markdown("---")
            st.subheader("💰 COSTO POR LOTE LEVANTES FINALIZADOS")
            rubros_costo = ["Costo Pollita", "Costo Alimento", "Costo Nomina", "Costo Cif", "Costo Droga", "Costo Arriendo", "Costo Cama", "Costo Aseo y Desinfección", "Costo Calefacción", "Depreciación Construcciones", "Aprovechamiento", "Costo Total Levante"]
            exist_c = [c for c in rubros_costo if c in df_l.columns]
            if exist_c:
                st.dataframe(df_l[["Lote_str"] + exist_c].set_index("Lote_str").T.style.format("${:,.0f}"), use_container_width=True)

        else:
            st.info(f"No hay lotes de Levante reportados para el año {anio_act}.")

elif menu == "6. RESULTADOS PRODUCCIÓN FINALIZADOS":
    st.markdown('<p class="titulo-principal">RESULTADOS LOTES FINALIZADOS PRODUCCIÓN</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="caja-info">
        💡 <b>Mega-Informe de Producción:</b> Control estricto sobre el ciclo de vida, la mortalidad acumulada y los costos de los galpones que completaron su ciclo.
    </div>
    """, unsafe_allow_html=True)
    
    if not df_prod.empty:
        df_p = df_prod[df_prod["Anio"] == anio_act].copy()
        if not df_p.empty:
            st.subheader(f"📊 Detalle Técnico por Lote ({anio_act})")
            cols_tech = ["Lote_str", "Mes_Num", "Edad", "Hembras Encasetadas", "Total Aves Fin Producción", "%Mortalidad Hembra", "%Mortalidad Macho", "%Sel Hembra", "HT/AA", "HF /AA", "% APROV", "%Nacimiento", "%Nacimiento con 2das", "Pollitos Hembra Alojada", "Gramos x Huevo Fértil", "%Incubabilidad", "Costo total Producción"]
            existentes = [c for c in cols_tech if c in df_p.columns]
            st.dataframe(df_p[existentes].rename(columns={"Lote_str":"Lote", "Costo total Producción": "Costo HF"}).style.format(na_rep="—"), use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("📈 RESULTADOS TÉCNICOS PRODUCCIÓN FINALIZADOS (Histórico)")
            df_hist_p = df_prod.groupby("Anio").agg(
                Edad=("Edad", "mean"),
                Encasetadas_H=("Hembras Encasetadas", "sum"),
                Finalizadas_H=("Total Aves Fin Producción", "sum") if "Total Aves Fin Producción" in df_prod.columns else ("Hembras Encasetadas", "sum"),
                Mortalidad_H=("Mortalidad Hembra", "sum"),
                Gramos_x_HF=("Gramos x Huevo Fértil", "mean"),
                HF_Ave=("HF /AA", "mean") if "HF /AA" in df_prod.columns else ("Gramos x Huevo Fértil", "mean")
            ).T
            st.dataframe(df_hist_p.style.format("{:,.1f}"), use_container_width=True)
            
            st.markdown("---")
            st.subheader("💰 COSTO HF POR LOTE PRODUCCIÓN FINALIZADOS")
            rubros_costo = ["Costo Alimento/Huevo Incubable", "Depreciación Reproductora/Huevo Incubabl", "Costo Nomina/Huevo Incubable", "Costo Cif/Huevo Incubable", "Costo Arriendo/Huevo Incubable", "Costo Aseo y Desinfección/Huevo Incubabl", "Costo Cama/Huevo Incubable", "Costo Droga/Huevo Incubable", "Depreciación Construcciones/Huevo Incuba", "Costo Aprovechamiento/Huevo Incubable", "Costo Total Producción/Huevo Incubable"]
            exist_c = [c for c in rubros_costo if c in df_p.columns]
            if exist_c:
                st.dataframe(df_p[["Lote_str"] + exist_c].set_index("Lote_str").T.style.format("${:,.1f}"), use_container_width=True)

        else:
            st.info(f"No hay lotes de Producción reportados para el año {anio_act}.")
