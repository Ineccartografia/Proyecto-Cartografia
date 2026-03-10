# =============================================================================
# PLANIFICACIÓN CARTOGRÁFICA ENDI 2025 - INTERFAZ STREAMLIT
# INEC · Zonal Litoral
# Autores: Franklin López, Carlos Quinto
#
# CAMBIOS PRINCIPALES:
#
# 1. FLUJO LÓGICO CORREGIDO:
#    Antes: los gráficos de carga aparecían ANTES de generar rutas (sin sentido).
#    Ahora: el flujo es estrictamente secuencial:
#      Paso 1 → Cargar datos (.gpkg)
#      Paso 2 → Cargar grafo vial (.graphml) - embebido en sidebar
#      Paso 3 → Configurar equipos
#      Paso 4 → Generar clusters + rutas (un solo botón)
#      Paso 5 → Ver resultados: mapa, análisis estadístico, reporte
#
# 2. CLUSTERING POR CONGLOMERADOS:
#    K-Means con k = n_equipos * 2 (dos jornadas).
#    Cada equipo trabaja en sus clusters asignados, no en todos los puntos.
#
# 3. EL MAPA NO DESAPARECE:
#    El problema anterior era que el botón "Generar Rutas" disparaba un
#    re-render completo que borraba el mapa. Ahora los resultados se guardan
#    en st.session_state y se muestran en una pestaña separada que solo
#    aparece después de generar.
#
# 4. ANÁLISIS ESTADÍSTICO MOVIDO:
#    Solo aparece DESPUÉS de generar los clusters/rutas, con los equipos
#    ya asignados. Los gráficos muestran carga por equipo con drilldown
#    horizontal (todos los equipos en una sola fila de tarjetas).
#
# 5. EQUIPO BOMBERO MEJORADO:
#    Toggle visible. Se muestra en una sección separada del reporte.
#    El usuario puede ver exactamente qué UPMs van al bombero y por qué.
#
# 6. CONFIGURACIÓN DE EQUIPOS MEJORADA:
#    Formulario dinámico más limpio: añadir/quitar equipos con nombre
#    personalizable y número de encuestadores independiente por equipo.
# =============================================================================

import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
import pyogrio
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import tempfile
import os
import osmnx as ox
import networkx as nx
from networkx.algorithms import approximation
from pyproj import Transformer
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
#  CONFIG DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ENDI · Planificación Cartográfica",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
#  CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

[data-testid="stSidebar"] { background-color: #0c0f1a; border-right: 1px solid #1e2540; }
[data-testid="stSidebar"] * { color: #d0d8e8 !important; }

.header-banner {
    background: linear-gradient(135deg, #071e3d 0%, #0d3b6e 60%, #0a2a52 100%);
    border-radius: 12px; padding: 28px 36px; margin-bottom: 24px;
    border-left: 5px solid #2e86de; position: relative; overflow: hidden;
}
.header-banner::after {
    content: "ENDI"; position: absolute; right: 28px; top: 50%;
    transform: translateY(-50%); font-family: 'IBM Plex Mono', monospace;
    font-size: 80px; font-weight: 600; color: rgba(255,255,255,0.04); letter-spacing: 6px;
}
.header-banner h1 {
    color: #fff !important; font-size: 20px !important; font-weight: 600 !important;
    margin: 0 0 4px 0 !important; font-family: 'IBM Plex Mono', monospace !important;
}
.header-banner p { color: #7eb3d8 !important; font-size: 12px !important; margin: 0 !important; }

.metric-card {
    background: #111827; border: 1px solid #1f2d45; border-radius: 10px;
    padding: 18px 20px; text-align: center; transition: border-color 0.2s;
}
.metric-card:hover { border-color: #2e86de; }
.metric-card .val {
    font-family: 'IBM Plex Mono', monospace; font-size: 28px;
    font-weight: 600; color: #2e86de; line-height: 1;
}
.metric-card .lbl { font-size: 11px; color: #7a8fa6; margin-top: 5px; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-card .sub { font-size: 10px; color: #4a6070; margin-top: 2px; }

.step-badge {
    display: inline-block; background: #0d2035; color: #2e86de;
    border: 1px solid #1a4060; border-radius: 4px; padding: 2px 8px;
    font-family: 'IBM Plex Mono', monospace; font-size: 10px; font-weight: 600;
    letter-spacing: 1px; margin-bottom: 8px;
}
.section-title {
    font-family: 'IBM Plex Mono', monospace; font-size: 12px; font-weight: 600;
    color: #2e86de; text-transform: uppercase; letter-spacing: 1px;
    border-bottom: 1px solid #1f2d45; padding-bottom: 8px; margin: 20px 0 14px 0;
}
.info-box {
    background: #0a1f35; border: 1px solid #143050; border-left: 3px solid #2e86de;
    border-radius: 8px; padding: 12px 16px; margin: 10px 0; font-size: 13px; color: #7eb3d8;
}
.warn-box {
    background: #1a1400; border: 1px solid #3a2800; border-left: 3px solid #f39c12;
    border-radius: 8px; padding: 12px 16px; margin: 10px 0; font-size: 13px; color: #c9a227;
}
.team-card {
    background: #0d1520; border: 1px solid #1f2d45; border-radius: 8px;
    padding: 12px 16px; margin-bottom: 8px;
}
.pill-ok {
    display: inline-block; background: #0a2e1a; color: #27ae60; border: 1px solid #1a5e35;
    border-radius: 20px; padding: 2px 10px; font-size: 11px;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
}
.pill-wait {
    display: inline-block; background: #1a1500; color: #e67e22; border: 1px solid #5a3c00;
    border-radius: 20px; padding: 2px 10px; font-size: 11px;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
}
.bombero-card {
    background: #1a0d2e; border: 1px solid #3d1a6e; border-left: 3px solid #9b59b6;
    border-radius: 8px; padding: 14px 18px; margin: 10px 0;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FUNCIONES AUXILIARES
# ─────────────────────────────────────────────

def cv(series):
    """Coeficiente de variación en porcentaje. Retorna 0 si la media es 0."""
    m = series.mean()
    return (series.std() / m * 100) if m > 0 else 0.0


def utm_to_wgs84(df):
    """Convierte columnas x,y de UTM zona 17S (EPSG:32717) a lat/lon WGS84."""
    t = Transformer.from_crs("epsg:32717", "epsg:4326", always_xy=True)
    lons, lats = t.transform(df["x"].values, df["y"].values)
    df = df.copy()
    df["lon"] = lons
    df["lat"] = lats
    return df


def cargar_gpkg(path, dissolve_by_upm=True):
    """
    Carga el GeoPackage, filtra Zonal Litoral y retorna un DataFrame
    con coordenadas UTM y WGS84 por UPM o por manzana/sector.
    """
    capas = pyogrio.list_layers(path)
    man  = gpd.read_file(path, layer=capas[0][0])
    disp = gpd.read_file(path, layer=capas[1][0])

    man  = man[man['zonal'] == 'LITORAL']
    disp = disp[disp['zonal'] == 'LITORAL']

    man_utm  = man.to_crs(epsg=32717)
    disp_utm = disp.to_crs(epsg=32717)

    if dissolve_by_upm:
        # Disolvemos por UPM: cada UPM queda como un punto representativo
        # y las viviendas se suman dentro de la UPM
        man_d  = man_utm.dissolve(by='upm', aggfunc={'mes':'first','viv':'sum'})
        disp_d = disp_utm.dissolve(by='upm', aggfunc={'mes':'first','viv':'sum'})

        for df_d, tipo in [(man_d, 'man_upm'), (disp_d, 'sec_upm')]:
            df_d['geometry'] = df_d.geometry.representative_point()

        def _fmt(df_d, tipo):
            out = df_d[['mes','viv']].copy()
            out['id_entidad'] = df_d.index
            out['upm']        = df_d.index
            out['tipo_entidad']= tipo
            out['x'] = df_d.geometry.x
            out['y'] = df_d.geometry.y
            return out[['id_entidad','upm','mes','viv','x','y','tipo_entidad']]

        man_sel  = _fmt(man_d,  'man_upm')
        disp_sel = _fmt(disp_d, 'sec_upm')
    else:
        # Nivel manzana/sector: un punto por fila original
        for gdf in [man_utm, disp_utm]:
            gdf['geometry'] = gdf.geometry.representative_point()

        man_utm['x']  = man_utm.geometry.x
        man_utm['y']  = man_utm.geometry.y
        disp_utm['x'] = disp_utm.geometry.x
        disp_utm['y'] = disp_utm.geometry.y

        man_sel  = man_utm[['man','upm','mes','viv','x','y']].copy().rename(columns={'man':'id_entidad'})
        man_sel['tipo_entidad'] = 'man'
        disp_sel = disp_utm[['sec','upm','mes','viv','x','y']].copy().rename(columns={'sec':'id_entidad'})
        disp_sel['tipo_entidad'] = 'sec'

    data = pd.concat([man_sel, disp_sel], ignore_index=True)
    if not dissolve_by_upm:
        data = data.drop_duplicates(subset=['id_entidad','upm'], keep='first')

    return utm_to_wgs84(data)


# Constantes
BASE_LAT = -2.145825935522539
BASE_LON = -79.89383956329586
MESES_NOMBRES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}
COLORES_EQUIPOS = ['#e74c3c','#2e86de','#27ae60','#f39c12','#9b59b6','#1abc9c','#e67e22','#e91e63']

# ─────────────────────────────────────────────
#  SESSION STATE — inicialización única
# ─────────────────────────────────────────────
_defaults = {
    "data_raw": None,            # DataFrame completo procesado del .gpkg
    "data_mes": None,            # DataFrame filtrado por mes seleccionado
    "graph_G": None,             # Grafo vial OSMnx
    "resultados_generados": False,  # Flag: ¿se generaron rutas?
    "df_planificado": None,      # DataFrame con equipo/jornada/cluster asignados
    "tsp_results": {},           # Resultados TSP por equipo/jornada
    "road_paths": {},            # Rutas geométricas para el mapa
    "equipos_config": [          # Configuración inicial: 3 equipos de 3 enc. c/u
        {"id": 1, "nombre": "Equipo 1", "encuestadores": 3},
        {"id": 2, "nombre": "Equipo 2", "encuestadores": 3},
        {"id": 3, "nombre": "Equipo 3", "encuestadores": 3},
    ],
    "silhouette_score": None,
    "resumen_balance": None,
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗺️ ENDI 2025")
    st.markdown("<p style='font-size:10px;color:#445566;margin-top:-8px'>INEC · Zonal Litoral · Cartografía</p>", unsafe_allow_html=True)
    st.divider()

    # ── PASO 1: Cargar .gpkg ──────────────────
    st.markdown("<div class='step-badge'>PASO 1</div>", unsafe_allow_html=True)
    st.markdown("**Cargar muestra (.gpkg)**")

    gpkg_file = st.file_uploader("Archivo GeoPackage", type=["gpkg"], key="gpkg_uploader",
                                  help="GeoPackage de la muestra ENDI con capas de manzanas y dispersos")

    if gpkg_file:
        dissolve = st.radio("Nivel de análisis", ["Por UPM (recomendado)", "Por manzana/sector"],
                            index=0, help="UPM agrupa manzanas contiguas. Reduce el número de puntos y mejora el clustering.")
        dissolve_upm = dissolve.startswith("Por UPM")

        if st.button("⚡ Procesar GeoPackage", use_container_width=True, type="primary"):
            with st.spinner("Leyendo geometrías y convirtiendo coordenadas..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".gpkg") as tmp:
                        tmp.write(gpkg_file.read())
                        tmp_path = tmp.name
                    data = cargar_gpkg(tmp_path, dissolve_by_upm=dissolve_upm)
                    os.unlink(tmp_path)
                    st.session_state.data_raw = data
                    # Reseteamos resultados al recargar datos
                    st.session_state.resultados_generados = False
                    st.session_state.df_planificado = None
                    st.success(f"✓ {len(data):,} entidades cargadas")
                except Exception as e:
                    st.error(f"Error al leer GeoPackage: {e}")

        if st.session_state.data_raw is not None:
            st.markdown("<span class='pill-ok'>✓ Datos cargados</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-wait'>⏳ Sin archivo</span>", unsafe_allow_html=True)

    st.divider()

    # ── PASO 2: Cargar grafo vial ─────────────
    st.markdown("<div class='step-badge'>PASO 2</div>", unsafe_allow_html=True)
    st.markdown("**Cargar red vial (.graphml)**")
    st.markdown("<p style='font-size:11px;color:#445566'>Archivo generado con OSMnx (carreteras principales de la Costa)</p>", unsafe_allow_html=True)

    graphml_file = st.file_uploader("Archivo GraphML", type=["graphml"], key="graphml_uploader",
                                     help="Grafo vial guardado como zonal.graphml. Solo carreteras primarias y secundarias.")

    if graphml_file:
        if st.button("⚡ Cargar grafo vial", use_container_width=True):
            with st.spinner("Cargando red vial... (puede tardar 1-2 minutos)"):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".graphml") as tmp:
                        tmp.write(graphml_file.read())
                        tmp_path_g = tmp.name
                    G = ox.load_graphml(tmp_path_g)
                    os.unlink(tmp_path_g)
                    st.session_state.graph_G = G
                    st.success(f"✓ Grafo cargado: {len(G.nodes):,} nodos")
                except Exception as e:
                    st.error(f"Error al cargar grafo: {e}")

        if st.session_state.graph_G is not None:
            st.markdown("<span class='pill-ok'>✓ Red vial lista</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-wait'>⏳ Sin grafo vial</span>", unsafe_allow_html=True)

    st.divider()

    # ── Filtro de mes ─────────────────────────
    if st.session_state.data_raw is not None:
        st.markdown("<div class='step-badge'>PASO 3</div>", unsafe_allow_html=True)
        st.markdown("**Seleccionar mes operativo**")

        data = st.session_state.data_raw
        meses_disp = sorted(data["mes"].dropna().unique().tolist())
        mes_sel = st.selectbox(
            "Mes",
            options=meses_disp,
            format_func=lambda x: f"{MESES_NOMBRES.get(int(x), str(int(x)))} (mes {int(x)})"
        )
        df_mes = data[data["mes"] == mes_sel].copy()
        st.session_state.data_mes = df_mes

        st.divider()

        # ── Configuración de equipos ──────────
        st.markdown("<div class='step-badge'>PASO 4</div>", unsafe_allow_html=True)
        st.markdown("**Configurar equipos de campo**")

        col_a, col_b = st.columns(2)
        with col_a:
            # Añadir equipo: crea un nuevo equipo con ID autoincremental
            if st.button("＋ Equipo", use_container_width=True):
                next_id = max(t["id"] for t in st.session_state.equipos_config) + 1
                st.session_state.equipos_config.append({
                    "id": next_id,
                    "nombre": f"Equipo {next_id}",
                    "encuestadores": 3
                })
                st.session_state.resultados_generados = False
        with col_b:
            # Eliminar el último equipo (mínimo 1)
            if st.button("－ Equipo", use_container_width=True,
                         disabled=len(st.session_state.equipos_config) <= 1):
                st.session_state.equipos_config.pop()
                st.session_state.resultados_generados = False

        st.markdown("<br>", unsafe_allow_html=True)

        # Renderizamos una tarjeta por equipo para configurar nombre y encuestadores
        for i, eq in enumerate(st.session_state.equipos_config):
            with st.container():
                c1, c2 = st.columns([2, 1])
                with c1:
                    nuevo_nombre = st.text_input(
                        f"Nombre equipo {eq['id']}",
                        value=eq["nombre"],
                        key=f"nombre_eq_{eq['id']}",
                        label_visibility="collapsed"
                    )
                    st.session_state.equipos_config[i]["nombre"] = nuevo_nombre
                with c2:
                    nuevo_n = st.number_input(
                        f"Enc.",
                        min_value=1, max_value=6,
                        value=eq["encuestadores"],
                        key=f"enc_eq_{eq['id']}",
                        label_visibility="collapsed"
                    )
                    st.session_state.equipos_config[i]["encuestadores"] = nuevo_n

        st.divider()

        # ── Opciones avanzadas ────────────────
        st.markdown("**Opciones avanzadas**")
        usar_bombero = st.toggle(
            "Equipo Bombero activo",
            value=True,
            help="Detecta UPMs muy alejadas del resto (outliers por IQR) y las asigna a un equipo especial con ruta libre."
        )
        factor_rural = st.slider(
            "Factor de carga rural (×)",
            min_value=1.0, max_value=2.5, value=1.5, step=0.1,
            help="Las UPMs dispersas/rurales se ponderan con este factor por ser más difíciles de encuestar."
        )

        n_vehiculos = st.number_input("Vehículos disponibles", min_value=1, max_value=30, value=4)

        # Resumen rápido de la configuración
        n_eq = len(st.session_state.equipos_config)
        tot_enc = sum(e["encuestadores"] for e in st.session_state.equipos_config)
        tot_viv = int(df_mes["viv"].sum()) if len(df_mes) > 0 else 0
        viv_enc = tot_viv // tot_enc if tot_enc > 0 else 0
        st.markdown(f"""
        <div style='font-size:11px;color:#445566;line-height:2;margin-top:8px'>
        📍 <b style='color:#7eb3d8'>{len(df_mes):,}</b> UPMs en mes {int(mes_sel)}<br>
        🏠 <b style='color:#7eb3d8'>{tot_viv:,}</b> viviendas estimadas<br>
        👥 <b style='color:#7eb3d8'>{n_eq}</b> equipos · <b style='color:#7eb3d8'>{tot_enc}</b> encuestadores<br>
        👤 <b style='color:#7eb3d8'>{viv_enc:,}</b> viv/encuestador (ideal)<br>
        🚗 <b style='color:#7eb3d8'>{n_vehiculos}</b> vehículos
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  CONTENIDO PRINCIPAL
# ─────────────────────────────────────────────
st.markdown("""
<div class='header-banner'>
    <h1>Planificación Automática · Actualización Cartográfica</h1>
    <p>ENDI 2025 &nbsp;·&nbsp; Zonal Litoral &nbsp;·&nbsp; INEC Ecuador</p>
</div>
""", unsafe_allow_html=True)

# ── Pantalla de bienvenida si no hay datos ──
if st.session_state.data_raw is None:
    st.markdown("""
    <div class='info-box'>
    👈 &nbsp; Comienza cargando el archivo <code>.gpkg</code> de la muestra desde el panel lateral.
    Luego carga la red vial <code>.graphml</code>, configura los equipos y genera la planificación.
    </div>
    """, unsafe_allow_html=True)

    cols = st.columns(4)
    pasos = [
        ("📂", "1. Cargar .gpkg", "GeoPackage con la muestra ENDI de la Zonal Litoral"),
        ("🛣️", "2. Cargar .graphml", "Red vial de la costa ecuatoriana (OSMnx)"),
        ("🏢", "3. Configurar equipos", "Número de equipos, encuestadores y opciones"),
        ("⚡", "4. Generar planificación", "Clustering + rutas optimizadas por jornada"),
    ]
    for col, (icon, title, desc) in zip(cols, pasos):
        with col:
            st.markdown(f"""
            <div style='background:#0d1520;border:1px solid #1f2d45;border-radius:10px;
                        padding:28px 20px;text-align:center;'>
                <div style='font-size:32px;margin-bottom:10px'>{icon}</div>
                <div style='font-family:"IBM Plex Mono",monospace;font-size:12px;
                            color:#2e86de;font-weight:600;margin-bottom:6px'>{title}</div>
                <div style='font-size:11px;color:#445566'>{desc}</div>
            </div>
            """, unsafe_allow_html=True)
    st.stop()

# ── Verificamos que hay mes seleccionado ──
df = st.session_state.data_mes
if df is None or len(df) == 0:
    st.warning("No hay datos para el mes seleccionado. Elige otro mes en el panel lateral.")
    st.stop()

# ── KPIs de la muestra ──
k1, k2, k3, k4, k5 = st.columns(5)
mes_actual = int(df["mes"].iloc[0])
n_man  = len(df[df["tipo_entidad"].isin(["man","man_upm"])])
n_disp = len(df[df["tipo_entidad"].isin(["sec","sec_upm"])])
cv_viv = cv(df["viv"])
cv_color = "#27ae60" if cv_viv < 50 else "#e74c3c"

for col, (val, lbl, sub) in zip(
    [k1,k2,k3,k4,k5],
    [
        (f"{len(df):,}", "UPMs", f"mes {mes_actual}"),
        (f"{int(df['viv'].sum()):,}", "Viviendas est.", "precenso 2020"),
        (f"{n_man:,}", "Amanzanadas", "man / man_upm"),
        (f"{n_disp:,}", "Dispersas", "sec / sec_upm"),
        (f"{cv_viv:.1f}%", "CV viviendas", "dispersión carga"),
    ]
):
    color = cv_color if lbl == "CV viviendas" else "#2e86de"
    with col:
        st.markdown(f"""<div class='metric-card'>
            <div class='val' style='color:{color}'>{val}</div>
            <div class='lbl'>{lbl}</div>
            <div class='sub'>{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  BOTÓN PRINCIPAL: GENERAR PLANIFICACIÓN
# ─────────────────────────────────────────────
# CAMBIO: el botón está en la zona principal, no en el sidebar.
# Esto evita que un re-render del sidebar limpie los resultados.
# Los resultados se guardan en session_state y persisten entre re-renders.

col_btn, col_info = st.columns([1, 3])
with col_btn:
    btn_generar = st.button(
        "⚡ Generar Planificación",
        use_container_width=True,
        type="primary",
        disabled=(st.session_state.graph_G is None),
        help="Requiere el grafo vial cargado (Paso 2)"
    )

with col_info:
    if st.session_state.graph_G is None:
        st.markdown("<div class='warn-box'>⚠️ Carga el archivo <code>.graphml</code> de la red vial (Paso 2) para habilitar la generación.</div>", unsafe_allow_html=True)
    elif st.session_state.resultados_generados:
        st.markdown("<div class='info-box'>✓ Planificación generada. Puedes regenerar si cambias la configuración de equipos.</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ALGORITMO PRINCIPAL
# Todos los resultados van a session_state.
# El mapa y los análisis se muestran DEBAJO, en pestañas.
# ─────────────────────────────────────────────
if btn_generar:
    G = st.session_state.graph_G
    equipos_cfg = st.session_state.equipos_config
    n_equipos = len(equipos_cfg)
    nombres_equipos = [e["nombre"] for e in equipos_cfg]

    # n_clusters = n_equipos * 2 (un cluster por equipo por jornada)
    n_clusters = n_equipos * 2

    df_trabajo = df.copy()

    progress_bar = st.progress(0, text="Iniciando...")

    # ── 1. Outliers → equipo_bombero ──────────
    progress_bar.progress(10, text="Detectando outliers geográficos...")

    t_wgs_utm = Transformer.from_crs("EPSG:4326","EPSG:32717",always_xy=True)
    base_x, base_y = t_wgs_utm.transform(BASE_LON, BASE_LAT)

    df_trabajo['dist_base_m'] = np.sqrt(
        (df_trabajo['x'] - base_x)**2 + (df_trabajo['y'] - base_y)**2
    )
    Q1d = df_trabajo['dist_base_m'].quantile(0.25)
    Q3d = df_trabajo['dist_base_m'].quantile(0.75)
    umbral = Q3d + 1.5*(Q3d - Q1d)

    mask_bombero = df_trabajo['dist_base_m'] > umbral
    df_trabajo['equipo']  = 'sin_asignar'
    df_trabajo['jornada'] = 'sin_asignar'
    df_trabajo['cluster_geo'] = -1

    if usar_bombero:
        df_trabajo.loc[mask_bombero, 'equipo']  = 'Equipo Bombero'
        df_trabajo.loc[mask_bombero, 'jornada'] = 'Jornada Especial'
    else:
        # Si no usamos bombero, esos puntos entran al clustering normal
        mask_bombero[:] = False

    # ── 2. K-Means clustering ─────────────────
    progress_bar.progress(25, text=f"Generando {n_clusters} conglomerados geográficos...")

    df_clust = df_trabajo[~mask_bombero].copy()

    if len(df_clust) >= n_clusters:
        coords = df_clust[['x','y']].values
        km = KMeans(n_clusters=n_clusters, init='k-means++', n_init=20,
                    max_iter=500, random_state=42)
        df_clust['cluster_geo'] = km.fit_predict(coords)

        # Silhouette score para evaluar la calidad del clustering
        if len(df_clust) > n_clusters:
            try:
                sil = silhouette_score(coords, df_clust['cluster_geo'])
                st.session_state.silhouette_score = sil
            except:
                st.session_state.silhouette_score = None

        # Asignamos clusters a equipos y jornadas por distancia a la base:
        # Los N clusters más lejanos → Jornada 1
        # Los N clusters más cercanos → Jornada 2
        centroides = km.cluster_centers_
        dist_c_base = np.sqrt(
            (centroides[:,0] - base_x)**2 + (centroides[:,1] - base_y)**2
        )
        orden = np.argsort(dist_c_base)[::-1]  # De más lejano a más cercano
        clusters_j1 = orden[:n_equipos]         # Más lejanos → Jornada 1
        clusters_j2 = orden[n_equipos:]         # Más cercanos → Jornada 2

        # Mapeamos cada cluster_id → (nombre_equipo, jornada)
        mapa_asignacion = {}
        for i, (cj1, cj2) in enumerate(zip(clusters_j1, clusters_j2)):
            nombre = nombres_equipos[i]
            mapa_asignacion[cj1] = (nombre, 'Jornada 1')
            mapa_asignacion[cj2] = (nombre, 'Jornada 2')

        df_clust['equipo']  = df_clust['cluster_geo'].map(lambda c: mapa_asignacion[c][0])
        df_clust['jornada'] = df_clust['cluster_geo'].map(lambda c: mapa_asignacion[c][1])

        # Actualizamos el dataframe principal con los resultados del clustering
        df_trabajo.loc[df_clust.index, 'equipo']      = df_clust['equipo']
        df_trabajo.loc[df_clust.index, 'jornada']     = df_clust['jornada']
        df_trabajo.loc[df_clust.index, 'cluster_geo'] = df_clust['cluster_geo']
    else:
        st.warning(f"Pocos puntos ({len(df_clust)}) para {n_clusters} clusters. Reduce el número de equipos.")

    # ── 3. Carga ponderada y asignación de encuestadores ──
    progress_bar.progress(40, text="Calculando carga ponderada y asignando encuestadores...")

    df_trabajo['carga_pond'] = df_trabajo.apply(
        lambda r: r['viv'] * factor_rural
        if str(r.get('tipo_entidad','')).startswith('sec') else r['viv'],
        axis=1
    )

    # Asignación de encuestadores dentro de cada equipo (greedy por carga)
    # Cada encuestador recibe UPMs hasta igualar la carga del equipo
    df_trabajo['encuestador'] = 0
    for eq in equipos_cfg:
        n_enc = eq["encuestadores"]
        mask_eq = df_trabajo['equipo'] == eq["nombre"]
        idx_eq  = df_trabajo[mask_eq].sort_values('carga_pond', ascending=False).index
        cargas  = np.zeros(n_enc)
        asigs   = []
        for idx in idx_eq:
            enc_min = int(np.argmin(cargas))
            asigs.append(enc_min + 1)
            cargas[enc_min] += df_trabajo.loc[idx, 'carga_pond']
        df_trabajo.loc[idx_eq, 'encuestador'] = asigs

    # ── 4. TSP por equipo y jornada ───────────
    progress_bar.progress(50, text="Optimizando rutas TSP (esto puede tardar)...")

    tsp_results = {}
    road_paths  = {}
    base_node   = ox.nearest_nodes(G, BASE_LON, BASE_LAT)
    G_undir     = G.to_undirected()
    componente_base = nx.node_connected_component(G_undir, base_node)

    total_rutas = n_equipos * 2
    ruta_actual = 0

    for eq_cfg in equipos_cfg:
        nombre_eq = eq_cfg["nombre"]
        for jornada in ['Jornada 1', 'Jornada 2']:
            ruta_actual += 1
            pct = 50 + int(ruta_actual / total_rutas * 45)
            progress_bar.progress(pct, text=f"Ruta TSP: {nombre_eq} | {jornada}...")

            mask_g = (df_trabajo['equipo'] == nombre_eq) & (df_trabajo['jornada'] == jornada)
            grp    = df_trabajo[mask_g]
            if len(grp) == 0:
                continue

            # Encontramos nodos en el grafo y filtramos alcanzables
            nodos_raw = ox.nearest_nodes(G, grp['lon'].values, grp['lat'].values)
            nodos_ok  = [n for n in nodos_raw if n in componente_base]

            if len(nodos_ok) == 0:
                continue

            nodos_unicos = [base_node] + list(dict.fromkeys(nodos_ok))
            n = len(nodos_unicos)
            if n <= 2:
                continue

            # Matriz de distancias O(n²)
            D = np.zeros((n, n))
            for i in range(n):
                for j in range(i+1, n):
                    try:
                        d = nx.shortest_path_length(G, nodos_unicos[i], nodos_unicos[j], weight='length')
                        D[i,j] = D[j,i] = d
                    except:
                        D[i,j] = D[j,i] = 1e9

            # Grafo para TSP
            G_tsp = nx.Graph()
            for i in range(n):
                for j in range(i+1, n):
                    if D[i,j] < 1e8:
                        G_tsp.add_edge(i, j, weight=D[i,j])

            if not nx.is_connected(G_tsp):
                continue

            try:
                ciclo = approximation.traveling_salesman_problem(G_tsp, weight='weight', cycle=True)
            except:
                continue

            # Rotamos para empezar en la base
            if 0 in ciclo:
                idx0 = ciclo.index(0)
                ciclo = ciclo[idx0:] + ciclo[1:idx0+1]

            dist_total = sum(D[ciclo[i], ciclo[i+1]] for i in range(len(ciclo)-1))

            # Reconstruimos ruta geométrica completa
            ruta_coords = []
            nodos_grafo = [nodos_unicos[idx] for idx in ciclo]
            for k in range(len(nodos_grafo)-1):
                u, v = nodos_grafo[k], nodos_grafo[k+1]
                try:
                    seg = nx.shortest_path(G, u, v, weight='length')
                    ruta_coords.extend((G.nodes[nd]['y'], G.nodes[nd]['x']) for nd in seg[:-1])
                except:
                    continue
            if nodos_grafo:
                last = nodos_grafo[-1]
                ruta_coords.append((G.nodes[last]['y'], G.nodes[last]['x']))

            clave = f"{nombre_eq}||{jornada}"
            tsp_results[clave] = {
                'equipo': nombre_eq, 'jornada': jornada,
                'n_puntos': len(grp), 'dist_km': dist_total / 1000
            }
            road_paths[clave] = ruta_coords

    progress_bar.progress(100, text="¡Planificación completa!")
    progress_bar.empty()

    # Guardamos todo en session_state
    # ESTO ES LA CLAVE: al guardar en session_state, los resultados
    # no se pierden cuando Streamlit hace re-render de la página.
    st.session_state.df_planificado      = df_trabajo
    st.session_state.tsp_results         = tsp_results
    st.session_state.road_paths          = road_paths
    st.session_state.resultados_generados = True

    # Cálculo del resumen de balance para el análisis estadístico
    resumen = df_trabajo.groupby(['equipo','jornada']).agg(
        n_upms          = ('id_entidad','count'),
        viv_reales      = ('viv','sum'),
        carga_ponderada = ('carga_pond','sum')
    ).reset_index()
    dist_df = pd.DataFrame([
        {'equipo': v['equipo'], 'jornada': v['jornada'], 'dist_km': round(v['dist_km'],1)}
        for v in tsp_results.values()
    ])
    if len(dist_df) > 0:
        st.session_state.resumen_balance = pd.merge(resumen, dist_df, on=['equipo','jornada'], how='left')
    else:
        st.session_state.resumen_balance = resumen

    st.success("✓ Planificación generada exitosamente. Explora los resultados en las pestañas.")

# ─────────────────────────────────────────────
#  RESULTADOS — solo se muestran si existen
# ─────────────────────────────────────────────
if not st.session_state.resultados_generados:
    st.markdown("<div class='info-box'>👆 Configura los equipos y presiona <b>Generar Planificación</b> para ver los resultados.</div>", unsafe_allow_html=True)
    st.stop()

# A partir de aquí tenemos resultados válidos en session_state
df_plan     = st.session_state.df_planificado
tsp_results = st.session_state.tsp_results
road_paths  = st.session_state.road_paths
resumen_bal = st.session_state.resumen_balance

tab_mapa, tab_analisis, tab_reporte = st.tabs([
    "🗺️  Mapa de Rutas",
    "📊  Análisis de Carga",
    "📋  Reporte Mensual"
])

# ══════════════════════════════════════════════
#  TAB 1 — MAPA DE RUTAS
# ══════════════════════════════════════════════
with tab_mapa:
    st.markdown("<div class='section-title'>Mapa del Operativo de Campo</div>", unsafe_allow_html=True)

    col_ctrl, col_m = st.columns([1, 3])

    with col_ctrl:
        # Construimos paleta de colores para equipos
        nombres_eq_plan = [e["nombre"] for e in st.session_state.equipos_config]
        color_map = {nombre: COLORES_EQUIPOS[i % len(COLORES_EQUIPOS)]
                     for i, nombre in enumerate(nombres_eq_plan)}
        color_map['Equipo Bombero'] = '#9b59b6'

        st.markdown("**Filtros del mapa**")
        mostrar_j1  = st.checkbox("Mostrar Jornada 1", value=True)
        mostrar_j2  = st.checkbox("Mostrar Jornada 2", value=True)
        mostrar_bomb = st.checkbox("Mostrar Equipo Bombero", value=True)
        mostrar_rutas = st.checkbox("Mostrar rutas TSP", value=True)
        fondo = st.selectbox("Fondo", ["CartoDB dark_matter","CartoDB positron","OpenStreetMap"])

        st.divider()
        # Leyenda de colores
        st.markdown("**Leyenda de equipos:**")
        for nombre, color in color_map.items():
            if nombre in nombres_eq_plan or nombre == 'Equipo Bombero':
                st.markdown(f"<span style='color:{color};font-size:18px'>●</span> {nombre}", unsafe_allow_html=True)

    with col_m:
        m = folium.Map(location=[BASE_LAT, BASE_LON], zoom_start=8, tiles=fondo)

        # Marcador de la base
        folium.Marker(
            location=[BASE_LAT, BASE_LON],
            popup="<b>Base INEC Guayaquil</b>",
            tooltip="Base de operaciones",
            icon=folium.Icon(color="white", icon="home", prefix="fa")
        ).add_to(m)

        # Añadimos puntos
        for _, row in df_plan.iterrows():
            eq  = row.get('equipo','')
            jor = row.get('jornada','')
            if jor == 'Jornada 1'       and not mostrar_j1:   continue
            if jor == 'Jornada 2'       and not mostrar_j2:   continue
            if jor == 'Jornada Especial' and not mostrar_bomb: continue

            color = color_map.get(eq, '#888888')
            folium.CircleMarker(
                location=[row['lat'], row['lon']],
                radius=5, color=color, fill=True,
                fill_color=color, fill_opacity=0.85,
                popup=folium.Popup(
                    f"<b>UPM:</b> {row['id_entidad']}<br>"
                    f"<b>Viviendas:</b> {int(row['viv'])}<br>"
                    f"<b>Equipo:</b> {eq}<br>"
                    f"<b>Jornada:</b> {jor}<br>"
                    f"<b>Encuestador:</b> {int(row.get('encuestador',0))}",
                    max_width=200
                ),
                tooltip=f"{eq} · {jor} · {int(row['viv'])} viv"
            ).add_to(m)

        # Añadimos rutas TSP
        if mostrar_rutas:
            for clave, coords in road_paths.items():
                eq, jor = clave.split('||')
                if jor == 'Jornada 1'        and not mostrar_j1:   continue
                if jor == 'Jornada 2'        and not mostrar_j2:   continue
                if jor == 'Jornada Especial' and not mostrar_bomb: continue
                if len(coords) > 1:
                    folium.PolyLine(
                        locations=coords, weight=3,
                        color=color_map.get(eq,'#888888'),
                        opacity=0.75, tooltip=f"Ruta: {eq} | {jor}"
                    ).add_to(m)

        # Usamos st_folium con key fija para que no se borre al interactuar
        # CAMBIO: key fija evita que Streamlit recree el componente en cada render
        st_folium(m, width=None, height=540, returned_objects=[], key="mapa_rutas")

# ══════════════════════════════════════════════
#  TAB 2 — ANÁLISIS DE CARGA
# ══════════════════════════════════════════════
with tab_analisis:
    st.markdown("<div class='section-title'>Análisis Estadístico de Carga</div>", unsafe_allow_html=True)

    # Métricas de calidad del clustering
    sil = st.session_state.silhouette_score
    col_sil1, col_sil2, col_sil3 = st.columns(3)
    with col_sil1:
        sil_text = f"{sil:.3f}" if sil is not None else "N/A"
        sil_color = "#27ae60" if (sil or 0) > 0.5 else ("#f39c12" if (sil or 0) > 0.3 else "#e74c3c")
        st.markdown(f"""<div class='metric-card'>
            <div class='val' style='color:{sil_color}'>{sil_text}</div>
            <div class='lbl'>Índice de Silueta</div>
            <div class='sub'>&gt;0.5 = clusters coherentes</div>
        </div>""", unsafe_allow_html=True)
    with col_sil2:
        n_clusters_generados = len(st.session_state.equipos_config) * 2
        st.markdown(f"""<div class='metric-card'>
            <div class='val'>{n_clusters_generados}</div>
            <div class='lbl'>Clusters generados</div>
            <div class='sub'>{len(st.session_state.equipos_config)} equipos × 2 jornadas</div>
        </div>""", unsafe_allow_html=True)
    with col_sil3:
        n_bombero = len(df_plan[df_plan['equipo'] == 'Equipo Bombero'])
        st.markdown(f"""<div class='metric-card'>
            <div class='val'>{n_bombero}</div>
            <div class='lbl'>UPMs equipo bombero</div>
            <div class='sub'>outliers IQR</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── CARGA POR EQUIPO — vista horizontal con drilldown ──
    # CAMBIO: en lugar de gráficos antes de generar rutas, ahora mostramos
    # una tarjeta por equipo con su CV y drilldown por encuestador.
    st.markdown("<div class='section-title'>Carga por equipo (click para detalle de encuestadores)</div>", unsafe_allow_html=True)

    # Mostramos todos los equipos en una fila horizontal
    eq_nombres_activos = df_plan[df_plan['equipo'] != 'Equipo Bombero']['equipo'].unique().tolist()
    eq_nombres_activos = sorted(eq_nombres_activos)

    if len(eq_nombres_activos) > 0:
        # Fila de métricas por equipo (horizontal)
        cols_eq = st.columns(len(eq_nombres_activos))
        for col_eq, nombre_eq in zip(cols_eq, eq_nombres_activos):
            viv_eq = df_plan[df_plan['equipo'] == nombre_eq]['viv'].sum()
            cv_eq  = cv(df_plan[df_plan['equipo'] == nombre_eq]['carga_pond'])
            color_cv = "#27ae60" if cv_eq < 20 else ("#f39c12" if cv_eq < 40 else "#e74c3c")
            color_eq = color_map.get(nombre_eq, '#2e86de')
            with col_eq:
                st.markdown(f"""<div class='metric-card' style='border-color:{color_eq}44'>
                    <div style='width:10px;height:10px;background:{color_eq};border-radius:50%;
                                margin:0 auto 8px auto'></div>
                    <div class='val' style='font-size:20px;color:{color_eq}'>{nombre_eq}</div>
                    <div class='lbl' style='margin-top:6px'>{int(viv_eq):,} viviendas</div>
                    <div class='sub'>CV carga: <span style='color:{color_cv}'>{cv_eq:.1f}%</span></div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Drilldown: selector de equipo para ver encuestadores
        eq_sel = st.selectbox(
            "Ver detalle de encuestadores del equipo:",
            options=eq_nombres_activos,
            format_func=lambda x: x
        )

        df_eq_sel = df_plan[df_plan['equipo'] == eq_sel].copy()
        df_enc = df_eq_sel.groupby(['jornada','encuestador']).agg(
            n_upms          = ('id_entidad','count'),
            viv_total       = ('viv','sum'),
            carga_ponderada = ('carga_pond','sum')
        ).reset_index()

        # Gráfico de barras agrupado: jornada x encuestador
        col_drill1, col_drill2 = st.columns(2)
        with col_drill1:
            fig_drill = px.bar(
                df_enc, x='encuestador', y='carga_ponderada', color='jornada',
                barmode='group', title=f'Carga ponderada por encuestador — {eq_sel}',
                labels={'carga_ponderada':'Carga ponderada','encuestador':'Encuestador','jornada':'Jornada'},
                template='plotly_dark',
                color_discrete_sequence=['#2e86de','#27ae60']
            )
            fig_drill.update_layout(paper_bgcolor="#111827", plot_bgcolor="#0a1020",
                                     title_font_size=13)
            st.plotly_chart(fig_drill, use_container_width=True)

        with col_drill2:
            fig_upms = px.bar(
                df_enc, x='encuestador', y='n_upms', color='jornada',
                barmode='group', title=f'UPMs asignadas por encuestador — {eq_sel}',
                labels={'n_upms':'UPMs','encuestador':'Encuestador','jornada':'Jornada'},
                template='plotly_dark',
                color_discrete_sequence=['#2e86de','#27ae60']
            )
            fig_upms.update_layout(paper_bgcolor="#111827", plot_bgcolor="#0a1020",
                                    title_font_size=13)
            st.plotly_chart(fig_upms, use_container_width=True)

        st.dataframe(df_enc.rename(columns={
            'jornada':'Jornada','encuestador':'Encuestador',
            'n_upms':'UPMs','viv_total':'Viviendas','carga_ponderada':'Carga pond.'
        }), use_container_width=True)

    # ── CV entre equipos por jornada ──
    st.markdown("<div class='section-title'>Equidad entre equipos</div>", unsafe_allow_html=True)

    for jornada in ['Jornada 1', 'Jornada 2']:
        sub = resumen_bal[resumen_bal['jornada'] == jornada]
        if len(sub) > 1:
            cv_j = cv(sub['carga_ponderada'])
            color_j = "#27ae60" if cv_j < 20 else ("#f39c12" if cv_j < 40 else "#e74c3c")
            st.markdown(f"""
            <div class='info-box'>
            <b>{jornada}:</b> CV de carga ponderada entre equipos =
            <span style='color:{color_j};font-weight:600;font-family:monospace'>{cv_j:.1f}%</span>
            {"&nbsp;✓ Muy bueno" if cv_j < 20 else ("&nbsp;⚠ Aceptable" if cv_j < 40 else "&nbsp;✗ Revisar")}
            </div>
            """, unsafe_allow_html=True)

    # ── Equipo Bombero ──
    df_bomb = df_plan[df_plan['equipo'] == 'Equipo Bombero']
    if len(df_bomb) > 0:
        st.markdown("<div class='section-title'>Equipo Bombero</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='bombero-card'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — {len(df_bomb)} UPMs asignadas<br>
        <span style='font-size:12px;color:#7a5a9a'>
        Estas UPMs superan el umbral de distancia (Q3+1.5×IQR) y tienen rutas no optimizadas.
        El equipo bombero cubre estos puntos aislados con flexibilidad operativa.
        Viviendas estimadas: {int(df_bomb['viv'].sum()):,}
        </span>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(
            df_bomb[['id_entidad','upm','tipo_entidad','viv','lat','lon','dist_base_m']]
            .rename(columns={'dist_base_m':'Dist. base (m)'})
            .sort_values('Dist. base (m)', ascending=False)
            .reset_index(drop=True),
            use_container_width=True, height=220
        )

# ══════════════════════════════════════════════
#  TAB 3 — REPORTE MENSUAL
# ══════════════════════════════════════════════
with tab_reporte:
    st.markdown("<div class='section-title'>Reporte Logístico del Mes</div>", unsafe_allow_html=True)

    # Tabla de resumen
    if resumen_bal is not None:
        st.markdown("**Resumen por equipo y jornada:**")
        total_row = pd.DataFrame([{
            'equipo':'TOTAL','jornada':'—',
            'n_upms': resumen_bal['n_upms'].sum(),
            'viv_reales': resumen_bal['viv_reales'].sum(),
            'carga_ponderada': resumen_bal['carga_ponderada'].sum(),
            'dist_km': resumen_bal['dist_km'].sum() if 'dist_km' in resumen_bal.columns else 0
        }])
        reporte_completo = pd.concat([resumen_bal, total_row], ignore_index=True)
        st.dataframe(
            reporte_completo.rename(columns={
                'equipo':'Equipo','jornada':'Jornada','n_upms':'UPMs',
                'viv_reales':'Viviendas','carga_ponderada':'Carga pond.','dist_km':'Dist. (km)'
            }),
            use_container_width=True
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Detalle completo de asignación (descargable):**")

    cols_reporte = ['id_entidad','upm','tipo_entidad','mes','viv','carga_pond',
                    'equipo','jornada','encuestador','lat','lon']
    cols_disponibles = [c for c in cols_reporte if c in df_plan.columns]
    df_export = df_plan[cols_disponibles].copy()
    df_export = df_export.sort_values(['equipo','jornada','encuestador']).reset_index(drop=True)

    st.dataframe(df_export, use_container_width=True, height=350)

    # Botón de descarga como CSV
    csv_data = df_export.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Descargar planificación (CSV)",
        data=csv_data,
        file_name=f"planificacion_mes{int(df['mes'].iloc[0])}.csv",
        mime="text/csv",
        use_container_width=True
    )
