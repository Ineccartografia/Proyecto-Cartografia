import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import tempfile
import os

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ENDI · Planificación Cartográfica",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #0f1117;
    border-right: 1px solid #1e2130;
}
[data-testid="stSidebar"] * {
    color: #e0e0e0 !important;
}

/* Header banner */
.header-banner {
    background: linear-gradient(135deg, #0a3d62 0%, #1a5276 50%, #0e4d92 100%);
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 24px;
    border-left: 5px solid #3498db;
    position: relative;
    overflow: hidden;
}
.header-banner::after {
    content: "ENDI";
    position: absolute;
    right: 24px;
    top: 50%;
    transform: translateY(-50%);
    font-family: 'IBM Plex Mono', monospace;
    font-size: 72px;
    font-weight: 600;
    color: rgba(255,255,255,0.06);
    letter-spacing: 4px;
}
.header-banner h1 {
    color: #ffffff !important;
    font-size: 22px !important;
    font-weight: 600 !important;
    margin: 0 0 4px 0 !important;
    font-family: 'IBM Plex Mono', monospace !important;
}
.header-banner p {
    color: #89b4d4 !important;
    font-size: 13px !important;
    margin: 0 !important;
}

/* Metric cards */
.metric-card {
    background: #1a1f2e;
    border: 1px solid #2a3145;
    border-radius: 10px;
    padding: 20px 24px;
    text-align: center;
    transition: border-color 0.2s;
}
.metric-card:hover { border-color: #3498db; }
.metric-card .value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 32px;
    font-weight: 600;
    color: #3498db;
    line-height: 1;
}
.metric-card .label {
    font-size: 12px;
    color: #8899aa;
    margin-top: 6px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.metric-card .sublabel {
    font-size: 11px;
    color: #556677;
    margin-top: 3px;
}

/* Section headers */
.section-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 28px 0 16px 0;
    padding-bottom: 10px;
    border-bottom: 1px solid #2a3145;
}
.section-header span {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 13px;
    font-weight: 600;
    color: #3498db;
    text-transform: uppercase;
    letter-spacing: 1px;
}
.section-header .line {
    flex: 1;
    height: 1px;
    background: #2a3145;
}

/* Tab override */
[data-testid="stTabs"] [role="tab"] {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.5px;
    color: #667788;
    text-transform: uppercase;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #3498db;
}

/* Upload zone */
.upload-zone {
    border: 2px dashed #2a3145;
    border-radius: 12px;
    padding: 40px;
    text-align: center;
    background: #0d1117;
    margin-bottom: 20px;
}
.upload-zone p {
    color: #556677;
    font-size: 13px;
    margin: 0;
}

/* Status pill */
.status-ok {
    display: inline-block;
    background: #0d3b27;
    color: #2ecc71;
    border: 1px solid #1a6b40;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}
.status-wait {
    display: inline-block;
    background: #1a1a0d;
    color: #f39c12;
    border: 1px solid #6b5a1a;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}

/* Coming soon badge */
.coming-soon {
    background: #1a1f2e;
    border: 1px dashed #2a3145;
    border-radius: 10px;
    padding: 48px 24px;
    text-align: center;
    color: #445566;
}
.coming-soon .icon { font-size: 36px; margin-bottom: 12px; }
.coming-soon .title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 14px;
    color: #556677;
    font-weight: 600;
    margin-bottom: 6px;
}
.coming-soon .desc { font-size: 12px; color: #334455; }

/* Dark theme for charts */
.js-plotly-plot { border-radius: 10px; overflow: hidden; }

/* Info box */
.info-box {
    background: #0d1f35;
    border: 1px solid #1a3a5c;
    border-left: 3px solid #3498db;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 12px 0;
    font-size: 13px;
    color: #89b4d4;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FUNCIÓN PRINCIPAL (la tuya)
# ─────────────────────────────────────────────
def muestra_coordenada(archivo_gpkg, dissolve_by_upm=False):
    capas = fiona.listlayers(archivo_gpkg)
    man = gpd.read_file(archivo_gpkg, layer=capas[0])
    disp = gpd.read_file(archivo_gpkg, layer=capas[1])

    litoral_man = man[man['zonal'] == 'LITORAL']
    litoral_disp = disp[disp['zonal'] == 'LITORAL']

    litoral_man_utm = litoral_man.to_crs(epsg=32717)
    litoral_disp_utm = litoral_disp.to_crs(epsg=32717)

    if dissolve_by_upm:
        man_dissolved = litoral_man_utm.dissolve(by='upm', aggfunc={'mes': 'first', 'viv': 'sum'})
        man_dissolved['geometry'] = man_dissolved.geometry.representative_point()
        man_selected = man_dissolved[['mes', 'viv']].copy()
        man_selected['id_entidad'] = man_dissolved.index
        man_selected['upm'] = man_dissolved.index
        man_selected['tipo_entidad'] = 'man_upm'
        man_selected['x'] = man_dissolved.geometry.x
        man_selected['y'] = man_dissolved.geometry.y
        man_selected = man_selected[['id_entidad', 'upm', 'mes', 'viv', 'x', 'y', 'tipo_entidad']]

        disp_dissolved = litoral_disp_utm.dissolve(by='upm', aggfunc={'mes': 'first', 'viv': 'sum'})
        disp_dissolved['geometry'] = disp_dissolved.geometry.representative_point()
        disp_selected = disp_dissolved[['mes', 'viv']].copy()
        disp_selected['id_entidad'] = disp_dissolved.index
        disp_selected['upm'] = disp_dissolved.index
        disp_selected['tipo_entidad'] = 'sec_upm'
        disp_selected['x'] = disp_dissolved.geometry.x
        disp_selected['y'] = disp_dissolved.geometry.y
        disp_selected = disp_selected[['id_entidad', 'upm', 'mes', 'viv', 'x', 'y', 'tipo_entidad']]
    else:
        litoral_man_points_utm = litoral_man_utm.copy()
        litoral_man_points_utm['geometry'] = litoral_man_points_utm.geometry.representative_point()
        litoral_disp_points_utm = litoral_disp_utm.copy()
        litoral_disp_points_utm['geometry'] = litoral_disp_points_utm.geometry.representative_point()

        litoral_man_points_utm['x'] = litoral_man_points_utm.geometry.x
        litoral_man_points_utm['y'] = litoral_man_points_utm.geometry.y
        litoral_disp_points_utm['x'] = litoral_disp_points_utm.geometry.x
        litoral_disp_points_utm['y'] = litoral_disp_points_utm.geometry.y

        litoral_man_points_utm = litoral_man_points_utm.drop(columns=['geometry'])
        litoral_disp_points_utm = litoral_disp_points_utm.drop(columns=['geometry'])

        man_selected = litoral_man_points_utm[['man', 'upm', 'mes', 'viv', 'x', 'y']].copy()
        man_selected = man_selected.rename(columns={'man': 'id_entidad'})
        man_selected['tipo_entidad'] = 'man'

        disp_selected = litoral_disp_points_utm[['sec', 'upm', 'mes', 'viv', 'x', 'y']].copy()
        disp_selected = disp_selected.rename(columns={'sec': 'id_entidad'})
        disp_selected['tipo_entidad'] = 'sec'

    data = pd.concat([man_selected, disp_selected], ignore_index=True)

    if not dissolve_by_upm:
        data = data.drop_duplicates(subset=['id_entidad', 'upm'], keep='first')

    return data

# ─────────────────────────────────────────────
#  UTM → WGS84
# ─────────────────────────────────────────────
def utm_to_wgs84(df):
    from pyproj import Transformer
    transformer = Transformer.from_crs("epsg:32717", "epsg:4326", always_xy=True)
    lons, lats = transformer.transform(df["x"].values, df["y"].values)
    df = df.copy()
    df["lon"] = lons
    df["lat"] = lats
    return df

# ─────────────────────────────────────────────
#  CALCULATE COEFFICIENT OF VARIATION (CV)
# ─────────────────────────────────────────────
def calculate_cv(series):
    if series.mean() == 0: # Avoid division by zero
        return 0.0
    return (series.std() / series.mean()) * 100

# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
if "data_raw" not in st.session_state:
    st.session_state.data_raw = None
if "data_filtered" not in st.session_state:
    st.session_state.data_filtered = None
if "dissolve" not in st.session_state:
    st.session_state.dissolve = True

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗺️ ENDI 2025")
    st.markdown("<p style='font-size:11px;color:#556677;margin-top:-8px'>Zonal Litoral · INEC</p>", unsafe_allow_html=True)
    st.divider()

    st.markdown("**📂 Cargar muestra**")
    uploaded_file = st.file_uploader(
        "Archivo .gpkg",
        type=["gpkg"],
        help="GeoPackage con la muestra seleccionada de la ENDI"
    )

    if uploaded_file:
        st.markdown("<span class='status-ok'>✓ Archivo cargado</span>", unsafe_allow_html=True)
        st.divider()

        st.markdown("**⚙️ Parámetros de procesamiento**")
        dissolve_option = st.radio(
            "Nivel de análisis",
            options=["Por UPM (disuelto)", "Por manzana/sector"],
            index=0,
            help="UPM agrupa manzanas contiguas en un solo punto representativo"
        )
        st.session_state.dissolve = dissolve_option == "Por UPM (disuelto)"

        if st.button("⚡ Procesar muestra", use_container_width=True, type="primary"):
            with st.spinner("Procesando geometrías..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".gpkg") as tmp:
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name
                    data = muestra_coordenada(tmp_path, dissolve_by_upm=st.session_state.dissolve)
                    data = utm_to_wgs84(data)
                    st.session_state.data_raw = data
                    os.unlink(tmp_path)
                    st.success(f"✓ {len(data):,} entidades procesadas")
                except Exception as e:
                    st.error(f"Error: {e}")
    else:
        st.markdown("<span class='status-wait'>⏳ Sin archivo</span>", unsafe_allow_html=True)

    st.divider()

    # Filtros — solo si hay datos
    if st.session_state.data_raw is not None:
        data = st.session_state.data_raw
        meses_disponibles = sorted(data["mes"].dropna().unique().tolist())
        meses_nombres = {
            1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
            5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
            9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
        }

        st.markdown("**🗓️ Filtro de mes**")
        mes_sel = st.selectbox(
            "Mes operativo",
            options=meses_disponibles,
            format_func=lambda x: f"{meses_nombres.get(int(x), x)} (mes {int(x)})"
        )

        df_mes = data[data["mes"] == mes_sel].copy()
        
        if st.session_state.get('n_equipos') is not None and st.session_state.get('n_enc') is not None:
            n_equipos_val = st.session_state.n_equipos
            n_enc_val = st.session_state.n_enc
            
            if len(df_mes) > 0:
                # Assign teams cyclically
                df_mes['equipo'] = (np.arange(len(df_mes)) % n_equipos_val) + 1
                
                # Assign surveyors cyclically within each team
                # Sort first to ensure consistent assignment within teams
                df_mes = df_mes.sort_values(by=['equipo']).reset_index(drop=True)
                df_mes['encuestador'] = (df_mes.groupby('equipo').cumcount() % n_enc_val) + 1
                
        st.session_state.data_filtered = df_mes

        st.markdown("**🏢 Equipos de campo**")
        n_equipos = st.number_input("Número de equipos", min_value=1, max_value=20, value=4, key='n_equipos')
        n_enc = st.number_input("Encuestadores por equipo", min_value=1, max_value=5, value=3, key='n_enc')
        n_vehiculos = st.number_input("Número de vehículos", min_value=1, max_value=20, value=4, key='n_vehiculos')

        st.session_state.n_equipos = n_equipos
        st.session_state.n_enc = n_enc
        st.session_state.n_vehiculos = n_vehiculos

        st.divider()
        total_enc = n_equipos * n_enc
        total_viv = int(df_mes["viv"].sum()) if len(df_mes) > 0 else 0
        viv_x_enc = total_viv // total_enc if total_enc > 0 else 0
        st.markdown(f"""
        <div style='font-size:11px;color:#556677;line-height:2'>
        📍 <b style='color:#89b4d4'>{len(df_mes):,}</b> entidades en mes {int(mes_sel)}<br>
        🏠 <b style='color:#89b4d4'>{total_viv:,}</b> viviendas estimadas<br>
        👤 <b style='color:#89b4d4'>{viv_x_enc:,}</b> viv/encuestador (ideal)<br>
        🚗 <b style='color:#89b4d4'>{n_vehiculos}</b> vehículos disponibles
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MAIN CONTENT
# ─────────────────────────────────────────────

# Header
st.markdown("""
<div class='header-banner'>
    <h1>Planificación Automática · Actualización Cartográfica</h1>
    <p>ENDI 2025 &nbsp;·&nbsp; Zonal Litoral &nbsp;·&nbsp; INEC Ecuador</p>
</div>
""", unsafe_allow_html=True)

# ── Sin datos: pantalla de bienvenida ──
if st.session_state.data_raw is None:
    st.markdown("""
    <div class='info-box'>
    👈 &nbsp; Carga un archivo <code>.gpkg</code> desde el panel lateral para comenzar.
    Una vez procesado, podrás explorar el mapa, analizar la muestra y generar la planificación del operativo.
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    modules = [
        ("🗺️", "Mapa de UPMs", "Visualización geográfica por mes con filtros de tipo de zona"),
        ("📊", "Análisis Estadístico", "Distribución de viviendas, outliers y carga por equipo"),
        ("🚗", "Generador de Rutas", "Algoritmo de clustering + rutas óptimas por carretera real"),
        ("📋", "Reporte Mensual", "Exportar programa mensual con asignación por equipo y día"),
    ]
    for col, (icon, title, desc) in zip([col1, col2, col3, col4], modules):
        with col:
            st.markdown(f"""
            <div class='coming-soon'>
                <div class='icon'>{icon}</div>
                <div class='title'>{title}</div>
                <div class='desc'>{desc}</div>
            </div>
            """, unsafe_allow_html=True)
    st.stop()

# ── Con datos: tabs ──
df = st.session_state.data_filtered
data_all = st.session_state.data_raw

# KPIs rápidos arriba
c1, c2, c3, c4, c5 = st.columns(5)
mes_actual = df["mes"].iloc[0] if df is not None and len(df) > 0 else "—"
meses_nombres2 = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
                  7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

with c1:
    st.markdown(f"""<div class='metric-card'>
        <div class='value'>{len(df):,}</div>
        <div class='label'>Entidades</div>
        <div class='sublabel'>mes {int(mes_actual)}</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class='metric-card'>
        <div class='value'>{int(df['viv'].sum()):,}</div>
        <div class='label'>Viviendas est.</div>
        <div class='sublabel'>precenso 2020</div>
    </div>""", unsafe_allow_html=True)
with c3:
    n_man = len(df[df["tipo_entidad"].isin(["man","man_upm"])])
    st.markdown(f"""<div class='metric-card'>
        <div class='value'>{n_man:,}</div>
        <div class='label'>Amanzanadas</div>
        <div class='sublabel'>UPMs / manzanas</div>
    </div>""", unsafe_allow_html=True)
with c4:
    n_disp = len(df[df["tipo_entidad"].isin(["sec","sec_upm"])])
    st.markdown(f"""<div class='metric-card'>
        <div class='value'>{n_disp:,}</div>
        <div class='label'>Dispersas</div>
        <div class='sublabel'>sectores</div>
    </div>""", unsafe_allow_html=True)
with c5:
    cv = calculate_cv(df["viv"]) # Use the new function
    color = "#2ecc71" if cv < 50 else "#e74c3c"
    st.markdown(f"""<div class='metric-card'>
        <div class='value' style='color:{color}'>{cv:.1f}%</div>
        <div class='label'>CV viviendas</div>
        <div class='sublabel'>dispersión carga</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3, tab4 = st.tabs([
    "🗺️  Mapa",
    "📊  Análisis Estadístico",
    "🚗  Generador de Rutas",
    "📋  Reporte Mensual"
])

# ══════════════════════════════════════════════
#  TAB 1 — MAPA
# ══════════════════════════════════════════════
with tab1:
    if len(df) == 0:
        st.warning("No hay datos para el mes seleccionado.")
    else:
        col_ctrl, col_map = st.columns([1, 3])

        with col_ctrl:
            st.markdown("**Opciones de visualización**")
            color_map_by = st.radio("Colorear puntos por", ["Tipo de zona", "Viviendas (intensidad)", "Equipo"]) # Added Equipo option
            show_inec = st.checkbox("Mostrar base INEC Guayaquil", value=True)
            map_tiles = st.selectbox("Fondo del mapa", ["CartoDB dark_matter", "CartoDB positron", "OpenStreetMap"])

            st.divider()
            st.markdown(f"""
            <div style='font-size:12px;color:#556677'>
            🔵 Amanzanadas: <b style='color:#3498db'>{len(df[df['tipo_entidad'].isin(['man','man_upm'])]):,}</b><br>
            🟠 Dispersas: <b style='color:#e67e22'>{len(df[df['tipo_entidad'].isin(['sec','sec_upm'])]):,}</b>
            </div>
            """, unsafe_allow_html=True)

        with col_map:
            centro_lat = df["lat"].mean()
            centro_lon = df["lon"].mean()
            m = folium.Map(location=[centro_lat, centro_lon], zoom_start=8, tiles=map_tiles)

            # Base INEC
            if show_inec:
                folium.Marker(
                    location=[-2.145825935522539, -79.89383956329586],
                    popup="<b>Base INEC Guayaquil</b>",
                    tooltip="INEC Guayaquil",
                    icon=folium.Icon(color="white", icon="home", prefix="fa")
                ).add_to(m)

            # Puntos
            color_palette_teams = px.colors.sequential.Viridis # Use a sequential palette for teams
            n_equipos_displayed = df['equipo'].nunique() if 'equipo' in df.columns else 1
            team_colors = {team_id: color_palette_teams[i % len(color_palette_teams)] for i, team_id in enumerate(sorted(df['equipo'].unique()))} if 'equipo' in df.columns else {}

            for _, row in df.iterrows():
                es_man = row["tipo_entidad"] in ["man", "man_upm"]
                if color_map_by == "Tipo de zona":
                    color = "#3498db" if es_man else "#e67e22"
                elif color_map_by == "Viviendas (intensidad)":
                    max_viv = df["viv"].max()
                    intensidad = int(row["viv"] / max_viv * 200) + 55
                    color = f"#{intensidad:02x}{'88'}{'ff'}" if es_man else f"#ff{intensidad:02x}22"
                elif color_map_by == "Equipo" and 'equipo' in row:
                    color = team_colors.get(row['equipo'], '#888888') # Default to grey if team not in palette
                else:
                    color = '#888888' # Default color

                popup_html = f"<b>ID:</b> {row['id_entidad']}<br>"\
                             f"<b>UPM:</b> {row['upm']}<br>"\
                             f"<b>Tipo:</b> {row['tipo_entidad']}<br>"\
                             f"<b>Viviendas:</b> {int(row['viv'])}"
                tooltip_text = f"{row['tipo_entidad']} · {int(row['viv'])} viv"

                if 'equipo' in row:
                    popup_html += f"<br><b>Equipo:</b> {int(row['equipo'])}"
                    tooltip_text += f" · Equipo {int(row['equipo'])}"
                if 'encuestador' in row:
                    popup_html += f"<br><b>Encuestador:</b> {int(row['encuestador'])}"
                    tooltip_text += f" · Enc. {int(row['encuestador'])}"

                folium.CircleMarker(
                    location=[row["lat"], row["lon"]],
                    radius=5,
                    color=color,
                    fill=True,
                    fill_color=color,
                    fill_opacity=0.8,
                    popup=folium.Popup(popup_html, max_width=200),
                    tooltip=tooltip_text
                ).add_to(m)

            st_folium(m, width=None, height=520, returned_objects=[])

# ══════════════════════════════════════════════
#  TAB 2 — ANÁLISIS ESTADÍSTICO
# ══════════════════════════════════════════════
with tab2:
    if len(df) == 0:
        st.warning("No hay datos para el mes seleccionado.")
    else:
        st.markdown("<div class='section-header'><span>Distribución de viviendas</span><div class='line'></div></div>", unsafe_allow_html=True)

        col_a, col_b = st.columns(2)

        with col_a:
            fig_hist = px.histogram(
                df, x="viv", color="tipo_entidad",
                nbins=40,
                color_discrete_map={"man":"#3498db","man_upm":"#2980b9","sec":"#e67e22","sec_upm":"#d35400"},
                template="plotly_dark",
                title="Distribución de viviendas por entidad",
                labels={"viv": "Viviendas estimadas", "tipo_entidad": "Tipo"}
            )
            fig_hist.update_layout(
                paper_bgcolor="#1a1f2e", plot_bgcolor="#0d1117",
                title_font_size=13, legend_title_text="Tipo"
            )
            st.plotly_chart(fig_hist, use_container_width=True)

        with col_b:
            fig_box = px.box(
                df, x="tipo_entidad", y="viv", color="tipo_entidad",
                color_discrete_map={"man":"#3498db","man_upm":"#2980b9","sec":"#e67e22","sec_upm":"#d35400"},
                template="plotly_dark",
                title="Dispersión por tipo de zona",
                labels={"viv": "Viviendas", "tipo_entidad": "Tipo"}
            )
            fig_box.update_layout(paper_bgcolor="#1a1f2e", plot_bgcolor="#0d1117",
                                   title_font_size=13, showlegend=False)
            st.plotly_chart(fig_box, use_container_width=True)

        st.markdown("<div class='section-header'><span>Detección de valores atípicos (IQR)</span><div class='line'></div></div>", unsafe_allow_html=True)

        Q1 = df["viv"].quantile(0.25)
        Q3 = df["viv"].quantile(0.75)
        IQR = Q3 - Q1
        limite_sup = Q3 + 1.5 * IQR
        outliers = df[df["viv"] > limite_sup]

        col_o1, col_o2, col_o3 = st.columns(3)
        with col_o1:
            st.markdown(f"""<div class='metric-card'>
                <div class='value'>{len(outliers)}</div>
                <div class='label'>Outliers detectados</div>
                <div class='sublabel'>viv > {limite_sup:.0f} (Q3+1.5×IQR)</div>
            </div>""", unsafe_allow_html=True)
        with col_o2:
            st.markdown(f"""<div class='metric-card'>
                <div class='value'>{Q1:.0f} – {Q3:.0f}</div>
                <div class='label'>Rango IQR</div>
                <div class='sublabel'>50% central de datos</div>
            </div>""", unsafe_allow_html=True)
        with col_o3:
            mediana = df["viv"].median()
            st.markdown(f"""<div class='metric-card'>
                <div class='value'>{mediana:.0f}</div>
                <div class='label'>Mediana viviendas</div>
                <div class='sublabel'>vs media {df['viv'].mean():.0f}</div>
            </div>""", unsafe_allow_html=True)

        if len(outliers) > 0:
            st.markdown("**Entidades con viviendas atípicas:**")
            st.dataframe(
                outliers[["id_entidad","upm","tipo_entidad","viv","lat","lon"]]
                .sort_values("viv", ascending=False)
                .reset_index(drop=True),
                use_container_width=True,
                height=220
            )

        st.markdown("<div class='section-header'><span>Resumen por mes (muestra completa)</span><div class='line'></div></div>", unsafe_allow_html=True)

        resumen_mes = data_all.groupby("mes").agg(
            entidades=("id_entidad","count"),
            viviendas_total=("viv","sum"),
            viviendas_media=("viv","mean"),
            viviendas_cv=("viv", calculate_cv) # Use the new function
        ).reset_index()
        resumen_mes.columns = ["Mes","Entidades","Viviendas totales","Media viv","CV (%)"]
        resumen_mes["CV (%)"] = resumen_mes["CV (%)"].round(1)
        resumen_mes["Media viv"] = resumen_mes["Media viv"].round(1)
        st.dataframe(resumen_mes, use_container_width=True, height=280)
        
        # Display load distribution per team and surveyor (New charts)
        if 'equipo' in df.columns and 'encuestador' in df.columns:
            st.markdown("<div class='section-header'><span>Análisis de carga de trabajo por equipo y encuestador</span><div class='line'></div></div>", unsafe_allow_html=True)

            # CV for teams
            team_cv = df.groupby('equipo')['viv'].apply(calculate_cv).reset_index(name='CV (%)')
            team_cv = team_cv.sort_values(by='equipo')
            
            st.markdown("**Coeficiente de Variación (CV) de viviendas por Equipo**")
            st.info("Un CV bajo (idealmente <50%) indica una carga de trabajo más equilibrada entre los equipos.")
            st.dataframe(team_cv, use_container_width=True)

            # Drill-down for surveyors within a selected team
            st.markdown("**CV de viviendas por Encuestador (drill-down por Equipo)**")
            selected_team = st.selectbox(
                "Seleccione un equipo para ver el CV de sus encuestadores:",
                options=sorted(df['equipo'].unique()),
                format_func=lambda x: f"Equipo {int(x)}"
            )

            if selected_team:
                df_selected_team = df[df['equipo'] == selected_team]
                surveyor_cv = df_selected_team.groupby('encuestador')['viv'].apply(calculate_cv).reset_index(name='CV (%)')
                surveyor_cv = surveyor_cv.sort_values(by='encuestador')

                st.info(f"CV de viviendas para los encuestadores del Equipo {int(selected_team)}. De manera similar, un CV bajo es deseable.")
                st.dataframe(surveyor_cv, use_container_width=True)

            st.markdown("<br>")

            col_load1, col_load2 = st.columns(2)

            with col_load1:
                team_load = df.groupby('equipo')['viv'].sum().reset_index()
                fig_team_load = px.bar(
                    team_load, x='equipo', y='viv',
                    title='Viviendas asignadas por Equipo (Suma)',
                    labels={'equipo': 'Equipo', 'viv': 'Total Viviendas'},
                    template='plotly_dark',
                    color_discrete_sequence=px.colors.sequential.Viridis
                )
                fig_team_load.update_layout(paper_bgcolor="#1a1f2e", plot_bgcolor="#0d1117", title_font_size=13)
                st.plotly_chart(fig_team_load, use_container_width=True)

            with col_load2:
                surveyor_load = df.groupby(['equipo', 'encuestador'])['viv'].sum().reset_index()
                fig_surveyor_load = px.bar(
                    surveyor_load, x='encuestador', y='viv', color='equipo',
                    title='Viviendas asignadas por Encuestador (por Equipo)',
                    labels={'encuestador': 'Encuestador', 'viv': 'Total Viviendas', 'equipo': 'Equipo'},
                    template='plotly_dark',
                    barmode='group',
                    color_discrete_sequence=px.colors.sequential.Viridis
                )
                fig_surveyor_load.update_layout(paper_bgcolor="#1a1f2e", plot_bgcolor="#0d1117", title_font_size=13)
                st.plotly_chart(fig_surveyor_load, use_container_width=True)

# ══════════════════════════════════════════════
#  TAB 3 — GENERADOR DE RUTAS
# ══════════════════════════════════════════════
with tab3:
    st.markdown("""
    <div class='info-box'>
    🔧 &nbsp; Esta sección integrará el algoritmo de clustering con restricción de capacidad
    y la generación de rutas óptimas por red vial real (OSMnx + NetworkX).
    El módulo está en desarrollo por <b>Franklin López</b>.
    </div>
    """, unsafe_allow_html=True)

    col_r1, col_r2 = st.columns(2)

    with col_r1:
        st.markdown("""
        <div class='coming-soon'>
            <div class='icon'>🔵</div>
            <div class='title'>Clustering con restricción de carga</div>
            <div class='desc'>K-Means + balanceo greedy por viviendas.<br>
            Número de clusters = número de equipos configurado en el panel lateral.</div>
        </div>
        """, unsafe_allow_html=True)

    with col_r2:
        st.markdown("""
        <div class='coming-soon'>
            <div class='icon'>🛣️</div>
            <div class='title'>Ruta óptima por carretera</div>
            <div class='desc'>Nearest Neighbor sobre grafo OSMnx.<br>
            Requiere archivo <code>zonal.graphml</code> con la red vial de la costa.</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**📂 Cargar grafo vial (cuando esté disponible)**")
    graphml_file = st.file_uploader("Archivo zonal.graphml", type=["graphml"],
                                     help="Generado con OSMnx por Franklin. Necesario para rutas reales.")
    if graphml_file:
        st.success("✓ Grafo cargado. Listo para integrar con el algoritmo de rutas.")
    else:
        st.markdown("""
        <div class='coming-soon' style='padding:24px'>
            <div class='icon'>⏳</div>
            <div class='title'>Esperando grafo vial</div>
            <div class='desc'>Carga el archivo <code>zonal.graphml</code> para habilitar esta sección.</div>
        </div>
        """, unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 4 — REPORTE MENSUAL
# ══════════════════════════════════════════════
with tab4:
    st.markdown("""
    <div class='info-box'>
    📋 &nbsp; Una vez generadas las rutas, esta sección permitirá exportar el programa mensual
    en formato Excel: equipo, encuestador, día, UPM asignada, viviendas estimadas y distancia al siguiente punto.
    También incluirá el resumen estadístico de CV de carga y comparación entre jornadas.
    </div>
    """, unsafe_allow_html=True)

    col_rep1, col_rep2, col_rep3 = st.columns(3)
    with col_rep1:
        st.markdown("""<div class='coming-soon'>
            <div class='icon'>📅</div>
            <div class='title'>Programa mensual</div>
            <div class='desc'>Asignación día a día por equipo y encuestador para las dos rondas de 12 días.</div>
        </div>""", unsafe_allow_html=True)
    with col_rep2:
        st.markdown("""<div class='coming-soon'>
            <div class='icon'>📈</div>
            <div class='title'>Resumen estadístico</div>
            <div class='desc'>CV de carga, distancia total por equipo, balance entre jornadas.</div>
        </div>""", unsafe_allow_html=True)
    with col_rep3:
        st.markdown("""<div class='coming-soon'>
            <div class='icon'>⬇️</div>
            <div class='title'>Exportar Excel</div>
            <div class='desc'>Reporte descargable listo para entregar a los equipos de campo.</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Vista previa de estructura del reporte:**")
    preview = pd.DataFrame({
        "Equipo": ["Equipo 1"]*3 + ["Equipo 2"]*3,
        "Encuestador": ["Enc. A","Enc. B","Enc. C"]*2,
        "Jornada": [1]*3 + [1]*3,
        "Día": [1,1,1,1,1,1],
        "UPM": ["—"]*6,
        "id_entidad": ["—"]*6,
        "Viviendas est.": ["—"]*6,
        "Dist. siguiente (km)": ["—"]*6
    })
    st.dataframe(preview, use_container_width=True, height=200)
    st.caption("Esta estructura se completará automáticamente al generar las rutas desde la pestaña anterior.")
