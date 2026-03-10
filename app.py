# =============================================================================
# PLANIFICACIÓN CARTOGRÁFICA ENDI 2025 — STREAMLIT v2
# INEC · Zonal Litoral
# Autores: Franklin López, Carlos Quinto
#
# CAMBIOS EN ESTA VERSIÓN:
#
# 1. DISTRIBUCIÓN EN 12 DÍAS:
#    Cada encuestador tiene sus UPMs distribuidas en los 12 días de la jornada
#    (50-80 viviendas/día). El día 12 se planifica igual aunque sea de regreso.
#
# 2. FACTOR RURAL EXPLICADO EN INTERFAZ:
#    La diferencia entre "carga real" y "carga ponderada" se explica visualmente
#    con un tooltip y una sección dedicada en el análisis estadístico.
#
# 3. EQUIPO BOMBERO SIEMPRE VISIBLE:
#    Aunque no haya outliers, el Equipo Bombero aparece en el mapa y el reporte
#    con 0 UPMs. Se muestra un mensaje informativo claro.
#
# 4. RESTRICCIÓN DE GUAYAQUIL:
#    Toggle configurable. Si está activo y hay suficientes UPMs en Guayaquil,
#    los primeros N días todos los equipos trabajan en Guayaquil.
#
# 5. CV CORREGIDO:
#    El CV entre equipos se calcula correctamente DESPUÉS del balance greedy.
#    Se muestra tanto el CV de viviendas reales como el de carga ponderada,
#    con explicación de la diferencia entre ambos.
#
# 6. VISTA HORIZONTAL DE EQUIPOS CON DRILLDOWN:
#    Una tarjeta por equipo en fila horizontal. El drilldown de encuestadores
#    muestra viviendas reales Y carga ponderada por día.
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import geopandas as gpd
import folium
import pyogrio
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
import tempfile, os, warnings
import osmnx as ox
import networkx as nx
from networkx.algorithms import approximation
from pyproj import Transformer
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
warnings.filterwarnings('ignore')

# ─── CONFIG ───────────────────────────────────
st.set_page_config(
    page_title="ENDI · Planificación Cartográfica",
    page_icon="🗺️", layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS ──────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif}
[data-testid="stSidebar"]{background:#0c0f1a;border-right:1px solid #1e2540}
[data-testid="stSidebar"] *{color:#d0d8e8 !important}
.hdr{background:linear-gradient(135deg,#071e3d,#0d3b6e 60%,#0a2a52);border-radius:12px;
     padding:26px 34px;margin-bottom:22px;border-left:5px solid #2e86de;
     position:relative;overflow:hidden}
.hdr::after{content:"INEC";position:absolute;right:24px;top:50%;transform:translateY(-50%);
            font-family:'IBM Plex Mono',monospace;font-size:76px;font-weight:600;
            color:rgba(255,255,255,.04);letter-spacing:6px}
.hdr h1{color:#fff!important;font-size:19px!important;font-weight:600!important;
        margin:0 0 3px!important;font-family:'IBM Plex Mono',monospace!important}
.hdr p{color:#7eb3d8!important;font-size:12px!important;margin:0!important}
.kcard{background:#111827;border:1px solid #1f2d45;border-radius:10px;
       padding:16px 18px;text-align:center;transition:border-color .2s}
.kcard:hover{border-color:#2e86de}
.kcard .v{font-family:'IBM Plex Mono',monospace;font-size:26px;font-weight:600;
          color:#2e86de;line-height:1}
.kcard .l{font-size:10px;color:#7a8fa6;margin-top:4px;text-transform:uppercase;letter-spacing:.5px}
.kcard .s{font-size:10px;color:#4a6070;margin-top:2px}
.step{display:inline-block;background:#0d2035;color:#2e86de;border:1px solid #1a4060;
      border-radius:4px;padding:2px 7px;font-family:'IBM Plex Mono',monospace;
      font-size:10px;font-weight:600;letter-spacing:1px;margin-bottom:6px}
.stitle{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;color:#2e86de;
        text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid #1f2d45;
        padding-bottom:7px;margin:18px 0 12px}
.ibox{background:#0a1f35;border:1px solid #143050;border-left:3px solid #2e86de;
      border-radius:7px;padding:11px 15px;margin:9px 0;font-size:13px;color:#7eb3d8}
.wbox{background:#1a1400;border:1px solid #3a2800;border-left:3px solid #f39c12;
      border-radius:7px;padding:11px 15px;margin:9px 0;font-size:13px;color:#c9a227}
.bcard{background:#1a0d2e;border:1px solid #3d1a6e;border-left:3px solid #9b59b6;
       border-radius:7px;padding:13px 16px;margin:9px 0}
.pill-ok{display:inline-block;background:#0a2e1a;color:#27ae60;border:1px solid #1a5e35;
         border-radius:20px;padding:2px 9px;font-size:11px;font-family:'IBM Plex Mono',monospace;font-weight:600}
.pill-w{display:inline-block;background:#1a1500;color:#e67e22;border:1px solid #5a3c00;
        border-radius:20px;padding:2px 9px;font-size:11px;font-family:'IBM Plex Mono',monospace;font-weight:600}
.eq-card{background:#0d1520;border:1px solid #1f2d45;border-radius:9px;
         padding:14px 16px;text-align:center;transition:border-color .2s}
.eq-card:hover{border-color:#2e86de}
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────
BASE_LAT = -2.145825935522539
BASE_LON = -79.89383956329586
PRO_GYE  = "09"
CAN_GYE  = "01"
MESES_NOMBRES = {
    1:"Julio",2:"Agosto",3:"Septiembre",4:"Octubre",5:"Noviembre",6:"Diciembre",
    7:"Enero",8:"Febrero",9:"Marzo",10:"Abril",11:"Mayo",12:"Junio"
}
COLORES_EQ = ['#e74c3c','#2e86de','#27ae60','#f39c12','#9b59b6',
              '#1abc9c','#e67e22','#e91e63']

# ─── FUNCIONES AUXILIARES ─────────────────────

def cv_pct(s):
    """Coeficiente de Variación en %. Retorna 0 si la media es 0."""
    m = s.mean()
    return float(s.std() / m * 100) if m > 0 else 0.0


def utm_to_wgs84(df):
    t = Transformer.from_crs("epsg:32717","epsg:4326",always_xy=True)
    lons, lats = t.transform(df["x"].values, df["y"].values)
    df = df.copy()
    df["lon"] = lons; df["lat"] = lats
    return df


def cargar_gpkg(path, dissolve_upm=True):
    capas = pyogrio.list_layers(path)
    man  = gpd.read_file(path, layer=capas[0][0])
    disp = gpd.read_file(path, layer=capas[1][0])
    man  = man[man['zonal']=='LITORAL']
    disp = disp[disp['zonal']=='LITORAL']
    man_utm  = man.to_crs(epsg=32717)
    disp_utm = disp.to_crs(epsg=32717)

    if dissolve_upm:
        def _diss(gdf, tipo):
            d = gdf.dissolve(by='upm', aggfunc={'mes':'first','viv':'sum'})
            d['geometry'] = d.geometry.representative_point()
            o = d[['mes','viv']].copy()
            o['id_entidad']=d.index; o['upm']=d.index
            o['tipo_entidad']=tipo
            o['x']=d.geometry.x; o['y']=d.geometry.y
            return o[['id_entidad','upm','mes','viv','x','y','tipo_entidad']]
        man_sel  = _diss(man_utm,  'man_upm')
        disp_sel = _diss(disp_utm, 'sec_upm')
    else:
        for g in [man_utm, disp_utm]:
            g['geometry'] = g.geometry.representative_point()
            g['x'] = g.geometry.x; g['y'] = g.geometry.y
        man_sel  = man_utm[['man','upm','mes','viv','x','y']].rename(columns={'man':'id_entidad'})
        man_sel['tipo_entidad'] = 'man'
        disp_sel = disp_utm[['sec','upm','mes','viv','x','y']].rename(columns={'sec':'id_entidad'})
        disp_sel['tipo_entidad'] = 'sec'
        
        man_sel["pro_x"] = man_sel["id_entidad"].str[:2]
        disp_sel["pro_x"] = disp_sel["id_entidad"].str[:2]
        
        man_sel["can_x"] = man_sel["id_entidad"].str[2:4]
        disp_sel["can_x"] = disp_sel["id_entidad"].str[2:4]

    data = pd.concat([man_sel, disp_sel], ignore_index=True)
    if not dissolve_upm:
        data = data.drop_duplicates(subset=['id_entidad','upm'], keep='first')
    return utm_to_wgs84(data)


def asignar_encuestadores_y_dias(df_grp, n_enc, dias_tot, viv_min, viv_max, inicio_dia=1):
    """
    Asignación greedy + distribución en días operativos.
    Retorna df_grp con columnas 'encuestador' y 'dia_operativo'.
    """
    df_g = df_grp.sort_values('carga_pond', ascending=False).copy()
    cargas = np.zeros(n_enc)
    enc_asig = []
    for _, row in df_g.iterrows():
        em = int(np.argmin(cargas))
        enc_asig.append(em + 1)
        cargas[em] += row['carga_pond']
    df_g['encuestador'] = enc_asig

    # Distribuimos en días: acumulamos viv_y reales hasta llegar a viv_max
    dias_col = [0]*len(df_g)
    for enc_id in range(1, n_enc+1):
        idx_enc = df_g[df_g['encuestador']==enc_id].index.tolist()
        viv_acum = 0
        dia = inicio_dia
        for pos, idx in enumerate(idx_enc):
            loc = df_g.index.get_loc(idx)
            viv_acum += df_g.iloc[loc]['viv']
            dias_col[loc] = dia
            if viv_acum >= viv_max and dia < inicio_dia + dias_tot - 1:
                dia += 1; viv_acum = 0
    df_g['dia_operativo'] = dias_col
    return df_g





# ─── SESSION STATE ────────────────────────────
_defs = {
    "data_raw": None, "data_mes": None, "graph_G": None,
    "resultados_generados": False, "df_plan": None,
    "tsp_results": {}, "road_paths": {}, "resumen_bal": None,
    "sil_score": None, "n_bombero": 0,
    "equipos_cfg": [
        {"id":1,"nombre":"Equipo 1","enc":3},
        {"id":2,"nombre":"Equipo 2","enc":3},
        {"id":3,"nombre":"Equipo 3","enc":3},
    ],
}
for k,v in _defs.items():
    if k not in st.session_state: st.session_state[k] = v

# ─── SIDEBAR ──────────────────────────────────
with st.sidebar:
    st.markdown("### 🗺️ Encuesta Nacional")
    st.markdown("<p style='font-size:10px;color:#445566;margin-top:-8px'>INEC · Zonal Litoral · Cartografía</p>", unsafe_allow_html=True)
    st.divider()

    # PASO 1 — GeoPackage
    st.markdown("<div class='step'>PASO 1</div>", unsafe_allow_html=True)
    st.markdown("**Cargar muestra (.gpkg)**")
    gpkg_f = st.file_uploader("GeoPackage", type=["gpkg"], key="gpkg_up")
    if gpkg_f:
        dissolve = st.radio("Nivel de análisis",
                            ["Por UPM","Por manzana"], index=0)
        if st.button("⚡ Procesar GeoPackage", use_container_width=True, type="primary"):
            with st.spinner("Leyendo geometrías..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".gpkg") as tmp:
                        tmp.write(gpkg_f.read()); p=tmp.name
                    data = cargar_gpkg(p, dissolve_upm=dissolve.startswith("Por UPM"))
                    os.unlink(p)
                    st.session_state.data_raw = data
                    st.session_state.resultados_generados = False
                    st.success(f"✓ {len(data):,} entidades")
                except Exception as e:
                    st.error(str(e))
        if st.session_state.data_raw is not None:
            st.markdown("<span class='pill-ok'>✓ Datos listos</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin archivo</span>", unsafe_allow_html=True)

    st.divider()

    # PASO 2 — GraphML
    st.markdown("<div class='step'>PASO 2</div>", unsafe_allow_html=True)
    st.markdown("**Cargar red vial (.graphml)**")
    gml_f = st.file_uploader("GraphML", type=["graphml"], key="gml_up")
    if gml_f:
        if st.button("⚡ Cargar grafo", use_container_width=True):
            with st.spinner("Cargando grafo vial..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".graphml") as tmp:
                        tmp.write(gml_f.read()); pg=tmp.name
                    G = ox.load_graphml(pg); os.unlink(pg)
                    st.session_state.graph_G = G
                    st.success(f"✓ {len(G.nodes):,} nodos")
                except Exception as e:
                    st.error(str(e))
        if st.session_state.graph_G is not None:
            st.markdown("<span class='pill-ok'>✓ Red vial lista</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin grafo</span>", unsafe_allow_html=True)

    st.divider()

    # Filtros (solo si hay datos)
    if st.session_state.data_raw is not None:
        st.markdown("<div class='step'>PASO 3</div>", unsafe_allow_html=True)
        st.markdown("**Mes operativo**")
        meses_disp = sorted(st.session_state.data_raw["mes"].dropna().unique().tolist())
        mes_sel = st.selectbox("Mes", meses_disp,
            format_func=lambda x: f"{MESES_NOMBRES.get(int(x), str(int(x)))} (mes {int(x)})")
        df_mes = st.session_state.data_raw[st.session_state.data_raw["mes"]==mes_sel].copy()
        st.session_state.data_mes = df_mes

        st.divider()
        st.markdown("<div class='step'>PASO 4</div>", unsafe_allow_html=True)
        st.markdown("**Equipos de campo**")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("＋", use_container_width=True):
                nid = max(t["id"] for t in st.session_state.equipos_cfg)+1
                st.session_state.equipos_cfg.append({"id":nid,"nombre":f"Equipo {nid}","enc":3})
                st.session_state.resultados_generados = False
        with c2:
            if st.button("－", use_container_width=True,
                         disabled=len(st.session_state.equipos_cfg)<=1):
                st.session_state.equipos_cfg.pop()
                st.session_state.resultados_generados = False

        for i, eq in enumerate(st.session_state.equipos_cfg):
            cc1, cc2 = st.columns([2,1])
            with cc1:
                nn = st.text_input(f"n{eq['id']}", value=eq["nombre"],
                                   key=f"n_{eq['id']}", label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["nombre"] = nn
            with cc2:
                ne = st.number_input("e", min_value=1, max_value=6, value=eq["enc"],
                                     key=f"e_{eq['id']}", label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["enc"] = ne

        st.divider()
        st.markdown("**Parámetros avanzados**")

        dias_op   = st.slider("Días operativos", 10, 14, 12)
        viv_min   = st.slider("Mín viv/día", 30, 60, 50)
        viv_max   = st.slider("Máx viv/día", 60, 120, 80)
        factor_r  = st.slider("Factor rural (×)", 1.0, 2.5, 1.5, 0.1,
            help="Multiplicador de carga para zonas dispersas. Ver explicación en 'Análisis de Carga'.")
        usar_bomb = st.toggle("Equipo Bombero", value=True,
            help="Asigna UPMs muy alejadas (outliers IQR) a un equipo especial.")
        usar_gye  = st.toggle("Restricción Guayaquil", value=True,
            help="Si hay suficientes UPMs en Guayaquil, los primeros días todos trabajan ahí.")
        dias_gye  = st.slider("Días en Guayaquil", 1, 5, 3,
            disabled=not usar_gye)
        umbral_gye = st.slider("Umbral Guayaquil (%)", 5, 30, 10,
            help="% mínimo de UPMs en Guayaquil para activar la restricción.",
            disabled=not usar_gye)

        # Resumen
        n_eq      = len(st.session_state.equipos_cfg)
        tot_enc   = sum(e["enc"] for e in st.session_state.equipos_cfg)
        tot_viv   = int(df_mes["viv"].sum()) if len(df_mes)>0 else 0
        st.markdown(f"""
        <div style='font-size:11px;color:#445566;line-height:2;margin-top:8px'>
        📍 <b style='color:#7eb3d8'>{len(df_mes):,}</b> UPMs · mes {int(mes_sel)}<br>
        🏠 <b style='color:#7eb3d8'>{tot_viv:,}</b> viviendas<br>
        👥 <b style='color:#7eb3d8'>{n_eq}</b> equipos · <b style='color:#7eb3d8'>{tot_enc}</b> enc.
        </div>
        """, unsafe_allow_html=True)

# ─── HEADER PRINCIPAL ─────────────────────────
st.markdown("""
<div class='hdr'>
  <h1>Planificación Automática · Actualización Cartográfica</h1>
  <p>Encuesta Nacional &nbsp;·&nbsp; Zonal Litoral &nbsp;·&nbsp; INEC Ecuador</p>
</div>
""", unsafe_allow_html=True)

if st.session_state.data_raw is None:
    st.markdown("<div class='ibox'>👈 Carga el archivo <code>.gpkg</code> desde el panel lateral para comenzar.</div>", unsafe_allow_html=True)
    st.stop()

df = st.session_state.data_mes
if df is None or len(df)==0:
    st.warning("Sin datos para el mes seleccionado.")
    st.stop()

# KPIs
k1,k2,k3,k4,k5 = st.columns(5)
cv_v = cv_pct(df["viv"])
cv_c = "#27ae60" if cv_v<50 else "#e74c3c"
for col,(val,lbl,sub,c) in zip([k1,k2,k3,k4,k5],[
    (f"{len(df):,}","UPMs",f"mes {int(df['mes'].iloc[0])}","#2e86de"),
    (f"{int(df['viv'].sum()):,}","Viviendas","precenso 2020","#2e86de"),
    (f"{len(df[df['tipo_entidad'].isin(['man','man_upm'])]):,}","Amanzanadas","man/man_upm","#2e86de"),
    (f"{len(df[df['tipo_entidad'].isin(['sec','sec_upm'])]):,}","Dispersas","sec/sec_upm","#2e86de"),
    (f"{cv_v:.1f}%","CV viviendas","dispersión",cv_c),
]):
    with col:
        st.markdown(f"<div class='kcard'><div class='v' style='color:{c}'>{val}</div>"
                    f"<div class='l'>{lbl}</div><div class='s'>{sub}</div></div>",
                    unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── BOTÓN GENERAR ────────────────────────────
cb1, cb2 = st.columns([1,3])
with cb1:
    btn = st.button("⚡ Generar Planificación", use_container_width=True,
                    type="primary", disabled=(st.session_state.graph_G is None))
with cb2:
    if st.session_state.graph_G is None:
        st.markdown("<div class='wbox'>⚠️ Carga el <code>.graphml</code> de la red vial (Paso 2) para habilitar.</div>", unsafe_allow_html=True)
    elif st.session_state.resultados_generados:
        st.markdown("<div class='ibox'>✓ Planificación lista. Puedes regenerar si cambias la configuración.</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ALGORITMO PRINCIPAL
# ─────────────────────────────────────────────
if btn:
    G          = st.session_state.graph_G
    eq_cfg     = st.session_state.equipos_cfg
    n_eq       = len(eq_cfg)
    nombres_eq = [e["nombre"] for e in eq_cfg]
    n_clusters = n_eq * 2

    df_w = df.copy()
    prog = st.progress(0, text="Iniciando...")

    # 1. Outliers → Equipo Bombero
    prog.progress(8, "Detectando outliers...")
    t = Transformer.from_crs("EPSG:4326","EPSG:32717",always_xy=True)
    bx, by = t.transform(BASE_LON, BASE_LAT)
    df_w['dist_base_m'] = np.sqrt((df_w['x']-bx)**2+(df_w['y']-by)**2) # calcula distancia euclidiana de base a manzanas.
    Q1d = df_w['dist_base_m'].quantile(.25)
    Q3d = df_w['dist_base_m'].quantile(.75)
    umb = Q3d + 1.5*(Q3d-Q1d)
    # la distancia está o no dentro del rango intercuantílico: si no, se considera punto outlier para el bombero. ¿Es esto adecuado?
    mask_bomb = df_w['dist_base_m'] > umb if usar_bomb else pd.Series(False, index=df_w.index)

    df_w['equipo']       = 'sin_asignar'
    df_w['jornada']      = 'sin_asignar'
    df_w['cluster_geo']  = -1
    df_w['carga_pond']   = df_w.apply(
        lambda r: r['viv']*factor_r if str(r.get('tipo_entidad','')).startswith('sec') else r['viv'], axis=1)
    df_w['encuestador']  = 0
    df_w['dia_operativo']= 0

    # EQUIPO BOMBERO: siempre se inicializa aunque esté vacío
    df_w.loc[mask_bomb, 'equipo']  = 'Equipo Bombero'
    df_w.loc[mask_bomb, 'jornada'] = 'Jornada Especial'
    st.session_state.n_bombero = int(mask_bomb.sum())

    # 2. Restricción Guayaquil
    prog.progress(15, "Verificando restricción Guayaquil...")
    upms_gye_mask = pd.Series(False, index=df_w.index)
    if usar_gye and 'pro_x' in df_w.columns and 'can_x' in df_w.columns:
        upms_gye_mask = (df_w['pro_x']==PRO_GYE)&(df_w['can_x']==CAN_GYE)
    pct_gye = upms_gye_mask.sum()/len(df_w) if len(df_w)>0 else 0
    activar_gye = usar_gye and (pct_gye >= umbral_gye/100) and upms_gye_mask.sum()>0

    if activar_gye:
        df_gye    = df_w[upms_gye_mask & ~mask_bomb].copy()
        df_no_gye = df_w[~upms_gye_mask & ~mask_bomb].copy()
    else:
        df_gye    = pd.DataFrame()
        df_no_gye = df_w[~mask_bomb].copy()

    # 3. Clustering
    prog.progress(25, f"Generando {n_clusters} conglomerados...")
    if len(df_no_gye) >= n_clusters:
        coords = df_no_gye[['x','y']].values
        km = KMeans(n_clusters=n_clusters, init='k-means++', n_init=20, max_iter=500, random_state=42)
        df_no_gye = df_no_gye.copy()
        df_no_gye['cluster_geo'] = km.fit_predict(coords)
        if len(df_no_gye) > n_clusters:
            try: st.session_state.sil_score = silhouette_score(coords, df_no_gye['cluster_geo'])
            except: st.session_state.sil_score = None

        centroides = km.cluster_centers_
        dist_c = np.sqrt((centroides[:,0]-bx)**2+(centroides[:,1]-by)**2)
        orden  = np.argsort(dist_c)[::-1]
        asig   = {}
        for i,(cj1,cj2) in enumerate(zip(orden[:n_eq],orden[n_eq:])):
            asig[cj1] = (nombres_eq[i],'Jornada 1')
            asig[cj2] = (nombres_eq[i],'Jornada 2')

        df_no_gye['equipo']  = df_no_gye['cluster_geo'].map(lambda c: asig[c][0])
        df_no_gye['jornada'] = df_no_gye['cluster_geo'].map(lambda c: asig[c][1])
        df_w.update(df_no_gye[['equipo','jornada','cluster_geo']])

    # 4. Asignación de encuestadores y días
    prog.progress(40, "Asignando encuestadores y distribuyendo en días...")
    enc_dict = {e["nombre"]:e["enc"] for e in eq_cfg}

    for nombre_eq in nombres_eq:
        for jornada in ['Jornada 1','Jornada 2']:
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp    = df_w[mask_g].copy()
            if len(grp)==0: continue
            n_enc      = enc_dict.get(nombre_eq, 3)
            inicio_dia = (dias_gye+1) if activar_gye else 1
            dias_disp  = dias_op - (dias_gye if activar_gye else 0)
            grp_a = asignar_encuestadores_y_dias(grp, n_enc, dias_disp, viv_min, viv_max, inicio_dia)
            df_w.update(grp_a[['encuestador','dia_operativo']])

    # Fase Guayaquil
    if activar_gye and len(df_gye)>0:
        for i,(idx,row) in enumerate(df_gye.sort_values('carga_pond',ascending=False).iterrows()):
            eq_a  = nombres_eq[i % n_eq]
            enc_a = (i // n_eq) % enc_dict.get(eq_a,3) + 1
            dia_a = min((i // (n_eq * enc_dict.get(eq_a,3)))+1, dias_gye)
            df_w.loc[idx,['equipo','jornada','encuestador','dia_operativo']] = [eq_a,'Jornada 1',enc_a,dia_a]

    # 5. TSP
    prog.progress(50, "Optimizando rutas TSP...")
    base_node = ox.nearest_nodes(G, BASE_LON, BASE_LAT)
    G_u       = G.to_undirected()
    comp_base = nx.node_connected_component(G_u, base_node)
    tsp_r     = {}
    road_p    = {}

    total_rutas = n_eq * 2
    for ri,(nombre_eq) in enumerate(nombres_eq):
        for jornada in ['Jornada 1','Jornada 2']:
            pct = 50 + int((ri*2 + ['Jornada 1','Jornada 2'].index(jornada)+1)/total_rutas*44)
            prog.progress(pct, f"TSP: {nombre_eq} | {jornada}...")
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp    = df_w[mask_g]
            if len(grp)==0: continue

            nodos_r = ox.nearest_nodes(G, grp['lon'].values, grp['lat'].values)
            nodos_k = [n for n in nodos_r if n in comp_base]
            if not nodos_k: continue

            nu = [base_node]+list(dict.fromkeys(nodos_k))
            n  = len(nu)
            if n<=2: continue

            D = np.zeros((n,n))
            for i in range(n):
                for j in range(i+1,n):
                    try: d=nx.shortest_path_length(G,nu[i],nu[j],weight='length'); D[i,j]=D[j,i]=d
                    except: D[i,j]=D[j,i]=1e9

            Gt = nx.Graph()
            for i in range(n):
                for j in range(i+1,n):
                    if D[i,j]<1e8: Gt.add_edge(i,j,weight=D[i,j])
            if not nx.is_connected(Gt): continue

            try: ciclo = approximation.traveling_salesman_problem(Gt, weight='weight', cycle=True)
            except: continue

            if 0 in ciclo:
                i0=ciclo.index(0); ciclo=ciclo[i0:]+ciclo[1:i0+1]

            dist = sum(D[ciclo[i],ciclo[i+1]] for i in range(len(ciclo)-1))
            ruta = []
            ng   = [nu[idx] for idx in ciclo]
            for k in range(len(ng)-1):
                try:
                    seg=nx.shortest_path(G,ng[k],ng[k+1],weight='length')
                    ruta.extend((G.nodes[nd]['y'],G.nodes[nd]['x']) for nd in seg[:-1])
                except: continue
            if ng: ruta.append((G.nodes[ng[-1]]['y'],G.nodes[ng[-1]]['x']))

            clave = f"{nombre_eq}||{jornada}"
            tsp_r[clave] = {'equipo':nombre_eq,'jornada':jornada,
                            'n_puntos':len(grp),'dist_km':dist/1000}
            road_p[clave] = ruta

    prog.progress(98, "Calculando métricas finales...")

    # Resumen de balance — calculado DESPUÉS del greedy (corrección del CV)
    resumen = df_w[df_w['equipo']!='Equipo Bombero'].groupby(['equipo','jornada']).agg(
        n_upms          = ('id_entidad','count'),
        viv_reales      = ('viv','sum'),
        carga_ponderada = ('carga_pond','sum')
    ).reset_index()
    dist_df = pd.DataFrame([
        {'equipo':v['equipo'],'jornada':v['jornada'],'dist_km':round(v['dist_km'],1)}
        for v in tsp_r.values()
    ]) if tsp_r else pd.DataFrame(columns=['equipo','jornada','dist_km'])
    if len(dist_df)>0:
        resumen_bal = pd.merge(resumen, dist_df, on=['equipo','jornada'], how='left')
    else:
        resumen_bal = resumen.copy()
        resumen_bal['dist_km'] = 0.0

    prog.progress(100, "¡Listo!")
    prog.empty()

    # Guardamos en session_state (persiste entre re-renders)
    st.session_state.df_plan             = df_w
    st.session_state.tsp_results         = tsp_r
    st.session_state.road_paths          = road_p
    st.session_state.resumen_bal         = resumen_bal
    st.session_state.resultados_generados = True

    st.success("✓ Planificación generada. Explora los resultados abajo.")

# ─────────────────────────────────────────────
#  RESULTADOS
# ─────────────────────────────────────────────
if not st.session_state.resultados_generados:
    st.markdown("<div class='ibox'>👆 Configura los equipos y presiona <b>Generar Planificación</b>.</div>", unsafe_allow_html=True)
    st.stop()

df_plan     = st.session_state.df_plan
tsp_results = st.session_state.tsp_results
road_paths  = st.session_state.road_paths
resumen_bal = st.session_state.resumen_bal
eq_cfg      = st.session_state.equipos_cfg
nombres_eq  = [e["nombre"] for e in eq_cfg]

color_map = {n: COLORES_EQ[i%len(COLORES_EQ)] for i,n in enumerate(nombres_eq)}
color_map['Equipo Bombero'] = '#9b59b6'

tab_mapa, tab_analisis, tab_reporte = st.tabs([
    "🗺️  Mapa de Rutas", "📊  Análisis de Carga", "📋  Reporte Mensual"
])

# ══ TAB 1 — MAPA ══════════════════════════════
with tab_mapa:
    st.markdown("<div class='stitle'>Mapa del Operativo de Campo</div>", unsafe_allow_html=True)
    cc1, cc2 = st.columns([1,3])
    with cc1:
        mj1   = st.checkbox("Jornada 1",     value=True)
        mj2   = st.checkbox("Jornada 2",     value=True)
        # Equipo Bombero siempre visible aunque tenga 0 UPMs
        n_bomb_mapa = int((df_plan['equipo']=='Equipo Bombero').sum())
        mbomb = st.checkbox(
            f"Equipo Bombero ({n_bomb_mapa} UPMs)",
            value=True,
            help="Muestra/oculta los puntos del Equipo Bombero. Si hay 0 UPMs no cambia nada en el mapa."
        )
        mrutas = st.checkbox("Mostrar rutas", value=True)
        fondo  = st.selectbox("Fondo", ["CartoDB dark_matter","CartoDB positron","OpenStreetMap"])
        st.divider()
        st.markdown("**Leyenda:**")
        for n,c in color_map.items():
            if n in nombres_eq:
                st.markdown(f"<span style='color:{c};font-size:17px'>●</span> {n}", unsafe_allow_html=True)
        # Bombero siempre en leyenda
        bomb_label = "Equipo Bombero" + (" (sin UPMs este mes)" if n_bomb_mapa==0 else f" — {n_bomb_mapa} UPMs")
        st.markdown(f"<span style='color:#9b59b6;font-size:17px'>●</span> {bomb_label}", unsafe_allow_html=True)

    with cc2:
        m = folium.Map(location=[BASE_LAT,BASE_LON], zoom_start=8, tiles=fondo)
        folium.Marker([BASE_LAT,BASE_LON], popup="<b>Base INEC Guayaquil</b>",
            icon=folium.Icon(color='white',icon='home',prefix='fa')).add_to(m)

        for _,row in df_plan.iterrows():
            eq,jor = row.get('equipo',''), row.get('jornada','')
            if jor=='Jornada 1' and not mj1: continue
            if jor=='Jornada 2' and not mj2: continue
            if jor=='Jornada Especial' and not mbomb: continue
            color = color_map.get(eq,'#888')
            folium.CircleMarker(
                location=[row['lat'],row['lon']], radius=5, color=color,
                fill=True, fill_color=color, fill_opacity=.85,
                popup=folium.Popup(
                    f"<b>ID:</b> {row['id_entidad']}<br>"
                    f"<b>Viv. reales:</b> {int(row['viv'])}<br>"
                    f"<b>Carga pond.:</b> {row.get('carga_pond',0):.0f}<br>"
                    f"<b>Equipo:</b> {eq}<br><b>Jornada:</b> {jor}<br>"
                    f"<b>Encuestador:</b> {int(row.get('encuestador',0))}<br>"
                    f"<b>Día:</b> {int(row.get('dia_operativo',0))}",
                    max_width=210),
                tooltip=f"{eq}·Día {int(row.get('dia_operativo',0))}·{int(row['viv'])}viv"
            ).add_to(m)

        if mrutas:
            for clave,coords in road_paths.items():
                eq,jor = clave.split('||')
                if jor=='Jornada 1' and not mj1: continue
                if jor=='Jornada 2' and not mj2: continue
                if len(coords)>1:
                    folium.PolyLine(coords,weight=3,color=color_map.get(eq,'#888'),
                                    opacity=.75,tooltip=f"Ruta:{eq}|{jor}").add_to(m)

        # KEY FIJA: evita que el mapa desaparezca al interactuar con controles
        st_folium(m, width=None, height=540, returned_objects=[], key="mapa_principal")

# ══ TAB 2 — ANÁLISIS ══════════════════════════
with tab_analisis:
    st.markdown("<div class='stitle'>Análisis Estadístico de Carga</div>", unsafe_allow_html=True)

    # Explicación del factor rural
    with st.expander("ℹ️ ¿Qué es la carga ponderada y por qué difiere de las viviendas reales?"):
        st.markdown(f"""
        **Carga real** = número de viviendas del precenso 2020 que se visitarán.
        
        **Carga ponderada** = criterio *interno* que usa el algoritmo para balancear.
        No es el número de visitas, es un peso ajustado.
        
        **¿Por qué ajustar?**  
        Visitar 30 casas en Guayaquil (urbano) toma ~3 horas.  
        Visitar 30 casas en zona dispersa rural puede tomar 6-8 horas por los desplazamientos.
        
        Si balanceamos solo por viviendas reales, el encuestador rural trabajará el doble.
        Por eso multiplicamos las zonas dispersas (`sec`, `sec_upm`) por el factor **{factor_r}×**.
        
        **Resultado:** al encuestador rural se le asignan *menos* viviendas reales,
        compensando el mayor tiempo de desplazamiento. El CV que mide equidad real
        es el de **carga ponderada**, no el de viviendas reales.
        """)

    # Métricas de clustering
    sil = st.session_state.sil_score
    n_bomb = st.session_state.n_bombero
    mc1,mc2,mc3 = st.columns(3)
    sil_v = f"{sil:.3f}" if sil else "N/A"
    sil_c = "#27ae60" if (sil or 0)>.5 else ("#f39c12" if (sil or 0)>.3 else "#e74c3c")
    with mc1:
        st.markdown(f"<div class='kcard'><div class='v' style='color:{sil_c}'>{sil_v}</div>"
                    f"<div class='l'>Índice Silueta</div><div class='s'>>0.5 = clusters coherentes</div></div>",
                    unsafe_allow_html=True)
    with mc2:
        st.markdown(f"<div class='kcard'><div class='v'>{len(eq_cfg)*2}</div>"
                    f"<div class='l'>Clusters</div><div class='s'>{len(eq_cfg)} eq × 2 jornadas</div></div>",
                    unsafe_allow_html=True)
    with mc3:
        bc = "#9b59b6" if n_bomb>0 else "#445566"
        st.markdown(f"<div class='kcard'><div class='v' style='color:{bc}'>{n_bomb}</div>"
                    f"<div class='l'>UPMs Bombero</div>"
                    f"<div class='s'>{'outliers IQR' if n_bomb>0 else 'ninguno detectado'}</div></div>",
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Equidad entre equipos (CV corregido) ──
    st.markdown("<div class='stitle'>Equidad entre equipos</div>", unsafe_allow_html=True)
    st.markdown("""
    <div class='ibox'>
    <b>CV viviendas reales</b> = dispersión de lo que el encuestador nota en campo.<br>
    <b>CV carga ponderada</b> = criterio interno de equidad (incluye factor rural).<br>
    Un CV &lt;20% indica balance muy bueno; 20-40% aceptable; &gt;40% revisar configuración.
    </div>
    """, unsafe_allow_html=True)

    # ── CV entre equipos — calculado post-greedy ──────────────────────────────
    # Recalculamos aquí directamente desde df_plan para garantizar que los
    # valores estén presentes, independientemente del estado de resumen_bal.
    # Esto soluciona el problema de que el CV "no aparecía" — la causa era que
    # resumen_bal podía no tener filas para cada jornada si el TSP fallaba.
    df_plan_main = df_plan[~df_plan['equipo'].isin(['Equipo Bombero','sin_asignar'])].copy()
    resumen_cv   = df_plan_main.groupby(['equipo','jornada']).agg(
        viv_reales      = ('viv','sum'),
        carga_ponderada = ('carga_pond','sum')
    ).reset_index()

    jornadas_presentes = resumen_cv['jornada'].unique().tolist()

    if len(resumen_cv) == 0:
        st.warning("No hay datos de equipos asignados para calcular equidad.")
    else:
        for jornada in ['Jornada 1','Jornada 2']:
            sub = resumen_cv[resumen_cv['jornada']==jornada]
            if len(sub) == 0:
                # Puede ocurrir si todos los puntos del mes son de Guayaquil
                st.markdown(f"<div class='ibox'><b>{jornada}:</b> sin UPMs asignadas.</div>",
                            unsafe_allow_html=True)
                continue
            if len(sub) == 1:
                st.markdown(f"<div class='ibox'><b>{jornada}:</b> solo 1 equipo — CV no aplicable.</div>",
                            unsafe_allow_html=True)
                continue
            cv_real = cv_pct(sub['viv_reales'])
            cv_pond = cv_pct(sub['carga_ponderada'])
            c_real  = "#27ae60" if cv_real<20 else ("#f39c12" if cv_real<40 else "#e74c3c")
            c_pond  = "#27ae60" if cv_pond<20 else ("#f39c12" if cv_pond<40 else "#e74c3c")
            emoji   = "✓" if cv_pond<20 else ("⚠" if cv_pond<40 else "✗")
            st.markdown(f"""
            <div class='ibox'>
            <b>{jornada}:</b><br>
            &nbsp;&nbsp;CV viviendas reales:&nbsp;&nbsp;
            <span style='color:{c_real};font-family:monospace;font-weight:600'>{cv_real:.1f}%</span>
            &nbsp;<i style='font-size:11px;color:#556677'>(lo que el encuestador ve en campo)</i><br>
            &nbsp;&nbsp;CV carga ponderada:&nbsp;&nbsp;
            <span style='color:{c_pond};font-family:monospace;font-weight:600'>{cv_pond:.1f}%</span>
            &nbsp;{emoji} <b>{'Muy bueno' if cv_pond<20 else ('Aceptable' if cv_pond<40 else 'Revisar')}</b>
            &nbsp;<i style='font-size:11px;color:#556677'>(criterio interno de equidad)</i>
            </div>
            """, unsafe_allow_html=True)

        # Tabla de detalle para inspección
        with st.expander("Ver tabla de balance completa"):
            st.dataframe(resumen_cv.rename(columns={
                'equipo':'Equipo','jornada':'Jornada',
                'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.'
            }), use_container_width=True)

    # ── Vista horizontal de equipos con drilldown ──
    st.markdown("<div class='stitle'>Carga por equipo</div>", unsafe_allow_html=True)

    eq_activos = [n for n in nombres_eq if n in df_plan['equipo'].values]
    cols_eq    = st.columns(len(eq_activos))

    for col_eq, nombre_eq in zip(cols_eq, eq_activos):
        sub_eq  = df_plan[df_plan['equipo']==nombre_eq]
        viv_tot = int(sub_eq['viv'].sum())
        cv_eq   = cv_pct(sub_eq['carga_pond'])
        c_eq    = color_map.get(nombre_eq,'#2e86de')
        c_cv    = "#27ae60" if cv_eq<20 else ("#f39c12" if cv_eq<40 else "#e74c3c")
        with col_eq:
            st.markdown(f"""
            <div class='eq-card' style='border-color:{c_eq}55'>
              <div style='width:10px;height:10px;background:{c_eq};border-radius:50%;margin:0 auto 7px'></div>
              <div style='font-family:"IBM Plex Mono",monospace;font-size:12px;
                          color:{c_eq};font-weight:600'>{nombre_eq}</div>
              <div style='font-size:18px;font-weight:600;color:#d0d8e8;margin:5px 0'>{viv_tot:,}</div>
              <div style='font-size:10px;color:#7a8fa6'>viviendas reales</div>
              <div style='font-size:11px;color:{c_cv};margin-top:5px'>CV {cv_eq:.1f}%</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Drilldown por encuestador
    eq_sel = st.selectbox("Ver detalle de encuestadores:", eq_activos)
    df_eq  = df_plan[df_plan['equipo']==eq_sel].copy()
    df_enc = df_eq.groupby(['jornada','encuestador','dia_operativo']).agg(
        upms        = ('id_entidad','count'),
        viv_reales  = ('viv','sum'),
        carga_pond  = ('carga_pond','sum')
    ).reset_index()

    cd1, cd2 = st.columns(2)
    with cd1:
        fig = px.bar(df_enc, x='encuestador', y='viv_reales', color='jornada',
                     barmode='group', title=f'Viviendas reales — {eq_sel}',
                     labels={'viv_reales':'Viv. reales','encuestador':'Encuestador','jornada':'Jornada'},
                     template='plotly_dark', color_discrete_sequence=['#2e86de','#27ae60'])
        fig.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",title_font_size=12)
        st.plotly_chart(fig, use_container_width=True)
    with cd2:
        fig2 = px.bar(df_enc, x='encuestador', y='carga_pond', color='jornada',
                      barmode='group', title=f'Carga ponderada — {eq_sel}',
                      labels={'carga_pond':'Carga pond.','encuestador':'Encuestador','jornada':'Jornada'},
                      template='plotly_dark', color_discrete_sequence=['#e74c3c','#f39c12'])
        fig2.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",title_font_size=12)
        st.plotly_chart(fig2, use_container_width=True)

    # Distribución por días
    st.markdown("<div class='stitle'>Distribución diaria de viviendas</div>", unsafe_allow_html=True)
    pivot = df_plan[df_plan['equipo'].isin(eq_activos)].groupby(
        ['equipo','dia_operativo'])['viv'].sum().reset_index()
    fig_dias = px.bar(pivot, x='dia_operativo', y='viv', color='equipo',
                      barmode='group', title='Viviendas por día operativo por equipo',
                      labels={'dia_operativo':'Día','viv':'Viviendas','equipo':'Equipo'},
                      template='plotly_dark',
                      color_discrete_map=color_map)
    fig_dias.add_hline(y=viv_min*sum(e["enc"] for e in eq_cfg)/len(eq_cfg),
                       line_dash="dot", line_color="#f39c12",
                       annotation_text=f"Mín esperado ({viv_min} viv/enc)")
    fig_dias.add_hline(y=viv_max*sum(e["enc"] for e in eq_cfg)/len(eq_cfg),
                       line_dash="dot", line_color="#e74c3c",
                       annotation_text=f"Máx esperado ({viv_max} viv/enc)")
    fig_dias.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020")
    st.plotly_chart(fig_dias, use_container_width=True)

    # Equipo Bombero
    df_bomb = df_plan[df_plan['equipo']=='Equipo Bombero']
    n_bomb  = st.session_state.n_bombero
    st.markdown("<div class='stitle'>Equipo Bombero</div>", unsafe_allow_html=True)
    if n_bomb == 0:
        st.markdown("""
        <div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — 0 UPMs asignadas<br>
        <span style='font-size:12px;color:#7a5a9a'>
        No se detectaron outliers geográficos en este mes.
        Todas las UPMs están dentro del rango IQR de distancia a la base.
        El equipo bombero está disponible como contingencia.
        </span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — {n_bomb} UPMs asignadas<br>
        <span style='font-size:12px;color:#7a5a9a'>
        Estas UPMs superan el umbral IQR de distancia y tienen ruta flexible (no optimizada TSP).
        Viviendas: {int(df_bomb['viv'].sum()):,}
        </span>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(
            df_bomb[['id_entidad','tipo_entidad','viv','lat','lon','dist_base_m']]
            .rename(columns={'dist_base_m':'Dist. base (m)'})
            .sort_values('Dist. base (m)', ascending=False).reset_index(drop=True),
            use_container_width=True, height=200
        )

# ══ TAB 3 — REPORTE ═══════════════════════════
with tab_reporte:
    st.markdown("<div class='stitle'>Reporte Mensual del Operativo</div>", unsafe_allow_html=True)
    st.markdown("<div class='ibox'>Estructura similar al formato INEC (Jornada, Grupo, Supervisor, Encuestador, Jurisdicción, Fechas).</div>", unsafe_allow_html=True)

    if resumen_bal is not None:
        total_row = pd.DataFrame([{
            'equipo':'TOTAL','jornada':'—',
            'n_upms':resumen_bal['n_upms'].sum(),
            'viv_reales':resumen_bal['viv_reales'].sum(),
            'carga_ponderada':resumen_bal['carga_ponderada'].sum(),
            'dist_km':resumen_bal['dist_km'].sum() if 'dist_km' in resumen_bal.columns else 0
        }])
        rep = pd.concat([resumen_bal, total_row], ignore_index=True)
        st.dataframe(rep.rename(columns={
            'equipo':'Equipo','jornada':'Jornada','n_upms':'UPMs',
            'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.','dist_km':'Dist. (km)'
        }), use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Detalle completo de asignación:**")
    cols_ok = [c for c in ['id_entidad','upm','tipo_entidad','mes','viv','carga_pond',
                            'equipo','jornada','encuestador','dia_operativo','lat','lon']
               if c in df_plan.columns]
    df_exp  = df_plan[cols_ok].sort_values(['equipo','jornada','encuestador','dia_operativo']).reset_index(drop=True)
    st.dataframe(df_exp, use_container_width=True, height=360)

    csv = df_exp.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Descargar planificación (CSV)", data=csv,
                       file_name=f"planificacion_mes{int(df['mes'].iloc[0])}.csv",
                       mime="text/csv", use_container_width=True)
