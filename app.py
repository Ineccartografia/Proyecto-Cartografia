# =============================================================================
# PLANIFICACIÓN CARTOGRÁFICA ENDI 2025 — STREAMLIT v4
# INEC · Zonal Litoral · Autores: Franklin López, Carlos Quinto
#
# CAMBIOS v4 (sobre v3):
#
# ── A. CLUSTERING BALANCEADO POR VIVIENDAS ──────────────────────────────────
#   Problema v3: KMeans puro balancea por distancia geométrica, no por viv.
#   Resultado: clusters con 8 000 vs 18 000 viviendas → inequidad grave.
#
#   Solución v4: clustering en DOS FASES:
#   1. KMeans(n_clusters) genera semillas geográficas coherentes.
#   2. Post-balance iterativo: se identifican clusters "pesados" y "livianos"
#      (por suma de viv ponderadas). En cada iteración se mueve la UPM
#      fronteriza del cluster más pesado al vecino más liviano, siempre que:
#        a) La UPM sea fronteriza (tiene vecinos en otro cluster).
#        b) El movimiento reduzca el CV global de cargas.
#      Se detiene cuando CV < umbral (default 10%) o max_iter alcanzado.
#
#   Ventaja: mantiene cohesión geográfica (solo frontera) y equilibra viv.
#
# ── B. ASIGNACIÓN POR ENCUESTADOR CON CONTIGÜIDAD GEOGRÁFICA ────────────────
#   Problema v3: greedy puro asigna UPMs en orden de carga DESC → encuestador
#   puede tener manzanas dispersas por todo el cluster sin continuidad.
#
#   Solución v4:
#   1. División balanceada por bin-packing (igual que v3 para equidad).
#   2. NUEVO: dentro de la porción asignada a cada encuestador, las UPMs
#      se ordenan por nearest-neighbor greedy (partiendo del centroide del
#      sub-grupo) → el encuestador recorre manzanas adyacentes, no salta.
#
# ── C. DISTRIBUCIÓN DIARIA CON CALENDARIO DE BLOQUEO ───────────────────────
#   Problema v3: manzana grande de 4 días avanza el cursor, pero días
#   intermedios quedaban con carga irregular (días 1-2 sobrecargados,
#   días finales casi vacíos).
#
#   Solución v4:
#   1. Cada encuestador tiene un "calendario" dict {dia: viv_acum}.
#   2. Para manzana normal: se busca el primer día con espacio disponible
#      (viv_acum < target) respetando el orden geográfico.
#   3. Para manzana grande: se reservan D días consecutivos desde el primer
#      día libre con D slots disponibles seguidos.
#   4. Las manzanas se procesan en orden geográfico (nearest-neighbor),
#      por lo que días consecutivos = manzanas físicamente contiguas.
#
# ── D. MÉTRICAS DE BALANCE VISIBLES ─────────────────────────────────────────
#   Se añade al tab Análisis:
#   - Gráfico de barras: viv por cluster antes y después del rebalanceo.
#   - Tabla: iteraciones del post-balance, CV inicial vs final.
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
import tempfile, os, warnings, io
from datetime import date, timedelta
import osmnx as ox
import networkx as nx
from networkx.algorithms import approximation
from pyproj import Transformer
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.neighbors import BallTree
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

# ── PAGE CONFIG ───────────────────────────────
st.set_page_config(page_title="ENDI · Planificación",
                   page_icon="🗺️", layout="wide",
                   initial_sidebar_state="expanded")

# ── CSS ───────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif}
[data-testid="stSidebar"]{background:#0c0f1a;border-right:1px solid #1e2540}
[data-testid="stSidebar"] *{color:#d0d8e8 !important}
.hdr{background:linear-gradient(135deg,#071e3d,#0d3b6e 60%,#0a2a52);
     border-radius:12px;padding:24px 32px;margin-bottom:20px;
     border-left:5px solid #2e86de;position:relative;overflow:hidden}
.hdr::after{content:"INEC";position:absolute;right:24px;top:50%;
            transform:translateY(-50%);font-family:'IBM Plex Mono',monospace;
            font-size:76px;font-weight:600;color:rgba(255,255,255,.04);letter-spacing:6px}
.hdr h1{color:#fff!important;font-size:18px!important;font-weight:600!important;
        margin:0 0 3px!important;font-family:'IBM Plex Mono',monospace!important}
.hdr p{color:#7eb3d8!important;font-size:12px!important;margin:0!important}
.kcard{background:#111827;border:1px solid #1f2d45;border-radius:10px;
       padding:14px 16px;text-align:center;transition:border-color .2s}
.kcard:hover{border-color:#2e86de}
.kcard .v{font-family:'IBM Plex Mono',monospace;font-size:24px;font-weight:600;
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
         border-radius:20px;padding:2px 9px;font-size:11px;
         font-family:'IBM Plex Mono',monospace;font-weight:600}
.pill-w{display:inline-block;background:#1a1500;color:#e67e22;border:1px solid #5a3c00;
        border-radius:20px;padding:2px 9px;font-size:11px;
        font-family:'IBM Plex Mono',monospace;font-weight:600}
.eq-card{background:#0d1520;border:1px solid #1f2d45;border-radius:9px;
         padding:14px 16px;text-align:center;transition:border-color .2s}
.eq-card:hover{border-color:#2e86de}
.pi-form{background:#0d1520;border:1px solid #1f2d45;border-radius:8px;
         padding:16px;margin-bottom:12px}
.balance-box{background:#071a10;border:1px solid #0d4020;border-left:3px solid #27ae60;
             border-radius:7px;padding:11px 15px;margin:9px 0;font-size:12px;color:#5dca8a}
</style>
""", unsafe_allow_html=True)

# ── CONSTANTES ────────────────────────────────
BASE_LAT = -2.145825935522539
BASE_LON = -79.89383956329586
PRO_GYE  = "09"
CAN_GYE  = "01"
MESES_N  = {1:"Julio",2:"Agosto",3:"Septiembre",4:"Octubre",5:"Noviembre",
             6:"Diciembre",7:"Enero",8:"Febrero",9:"Marzo",10:"Abril",
             11:"Mayo",12:"Junio"}
COLORES  = ['#e74c3c','#2e86de','#27ae60','#f39c12','#9b59b6',
            '#1abc9c','#e67e22','#e91e63']

# ── HELPERS ───────────────────────────────────

def cv_pct(s):
    m = s.mean()
    return float(s.std()/m*100) if m > 0 else 0.0

def utm_to_wgs84(df):
    t = Transformer.from_crs("epsg:32717","epsg:4326",always_xy=True)
    lons,lats = t.transform(df["x"].values,df["y"].values)
    df=df.copy(); df["lon"]=lons; df["lat"]=lats
    return df

def parse_codigo(codigo):
    c = str(codigo).strip()
    r = {'prov':'','canton':'','ciudad_parroq':'','zona':'','sector':'','man':''}
    if len(c)>=6:  r['prov']=c[:2]; r['canton']=c[2:4]; r['ciudad_parroq']=c[4:6]
    if len(c)>=9:  r['zona']=c[6:9]
    if len(c)>=12: r['sector']=c[9:12]
    if len(c)>=15: r['man']=c[12:15]
    return r

def cargar_gpkg(path, dissolve_upm=True):
    capas = pyogrio.list_layers(path)
    if len(capas) == 1:
        gdf = gpd.read_file(path, layer=capas[0][0])
        col_map = {
            '1_mes_cart': 'mes', 'viv_total': 'viv', '1_zonal': 'zonal',
            '1_id_upm': 'upm', 'ManSec': 'id_entidad'
        }
        gdf = gdf.rename(columns={k: v for k, v in col_map.items() if k in gdf.columns})
        if 'id_entidad' in gdf.columns:
            gdf['tipo_entidad'] = gdf['id_entidad'].astype(str).apply(
                lambda x: 'sec' if '999' in x else 'man')
        gdf_u = gdf.to_crs(epsg=32717)
        gdf_u['geometry'] = gdf_u.geometry.representative_point()
        gdf_u['x'] = gdf_u.geometry.x
        gdf_u['y'] = gdf_u.geometry.y
        if 'pro' in gdf_u.columns: gdf_u['pro_x'] = gdf_u['pro']
        if 'can' in gdf_u.columns: gdf_u['can_x'] = gdf_u['can']
        if 'mes' in gdf_u.columns: gdf_u['mes'] = pd.to_numeric(gdf_u['mes'], errors='coerce')
        if dissolve_upm and 'upm' in gdf_u.columns:
            agg_dict = {'viv':'sum','mes':'first','x':'first','y':'first','tipo_entidad':'first'}
            if 'pro_x' in gdf_u.columns: agg_dict['pro_x'] = 'first'
            if 'can_x' in gdf_u.columns: agg_dict['can_x'] = 'first'
            gdf_final = gdf_u.groupby('upm').agg(agg_dict).reset_index()
            gdf_final['id_entidad'] = gdf_final['upm']
            gdf_final['tipo_entidad'] = gdf_final['tipo_entidad'].apply(lambda t: f"{t}_upm")
        else:
            gdf_final = gdf_u
        return utm_to_wgs84(gdf_final)
    else:
        man   = gpd.read_file(path, layer=capas[0][0])
        disp  = gpd.read_file(path, layer=capas[1][0])
        man   = man[man['zonal']=='LITORAL']
        disp  = disp[disp['zonal']=='LITORAL']
        man_u = man.to_crs(epsg=32717)
        dis_u = disp.to_crs(epsg=32717)
        if dissolve_upm:
            def _d(gdf, tipo):
                d = gdf.dissolve(by='upm', aggfunc={'mes':'first','viv':'sum'})
                d['geometry'] = d.geometry.representative_point()
                o = d[['mes','viv']].copy()
                o['id_entidad'] = d.index; o['upm'] = d.index
                o['tipo_entidad'] = tipo
                o['x'] = d.geometry.x; o['y'] = d.geometry.y
                if 'mes' in o.columns: o['mes'] = pd.to_numeric(o['mes'], errors='coerce')
                return o[['id_entidad','upm','mes','viv','x','y','tipo_entidad']]
            ms = _d(man_u,'man_upm'); ds = _d(dis_u,'sec_upm')
        else:
            for g in [man_u, dis_u]:
                g['geometry'] = g.geometry.representative_point()
                g['x'] = g.geometry.x; g['y'] = g.geometry.y
            ms = man_u[['man','upm','mes','viv','x','y']].rename(columns={'man':'id_entidad'})
            ms['tipo_entidad'] = 'man'
            ds = dis_u[['sec','upm','mes','viv','x','y']].rename(columns={'sec':'id_entidad'})
            ds['tipo_entidad'] = 'sec'
            for df_t in [ms, ds]:
                ms['pro_x'] = ms['id_entidad'].astype(str).str[:2]
                ms['can_x'] = ms['id_entidad'].astype(str).str[2:4]
                if 'mes' in df_t.columns: df_t['mes'] = pd.to_numeric(df_t['mes'], errors='coerce')
        data = pd.concat([ms, ds], ignore_index=True)
        if not dissolve_upm: data = data.drop_duplicates(subset=['id_entidad','upm'], keep='first')
        return utm_to_wgs84(data)


# ══════════════════════════════════════════════════════════════════════════════
#  NUEVO — A. CLUSTERING BALANCEADO POR VIVIENDAS
# ══════════════════════════════════════════════════════════════════════════════

def nearest_neighbor_order(points_xy, start_xy=None):
    """
    Ordena un array de puntos (N×2) por el algoritmo nearest-neighbor greedy.
    Retorna los índices en orden de visita.
    Complejidad O(N²) — aceptable para N < 2000 UPMs por cluster.
    """
    n = len(points_xy)
    if n == 0:
        return []
    if n == 1:
        return [0]

    visited = [False] * n
    if start_xy is not None:
        # Empezar por el punto más cercano al centroide del sub-grupo
        dists = np.linalg.norm(points_xy - start_xy, axis=1)
        cur = int(np.argmin(dists))
    else:
        cur = 0

    order = [cur]
    visited[cur] = True

    for _ in range(n - 1):
        best_d, best_j = np.inf, -1
        px, py = points_xy[cur]
        for j in range(n):
            if not visited[j]:
                d = (points_xy[j,0]-px)**2 + (points_xy[j,1]-py)**2
                if d < best_d:
                    best_d, best_j = d, j
        cur = best_j
        order.append(cur)
        visited[cur] = True

    return order


def clustering_balanceado(df, n_clusters, factor_r=1.5,
                          cv_objetivo=0.10, max_iter=300,
                          k_vecinos=8):
    """
    Clustering en dos fases para balancear suma de viviendas ponderadas.

    FASE 1 — KMeans geográfico
    --------------------------
    Genera n_clusters semillas geográficamente coherentes.

    FASE 2 — Post-balance iterativo con DOS modos de swap
    ------------------------------------------------------
    MODO FRONTERA (preferido — preserva cohesión geográfica):
      Busca UPMs en el cluster pesado que tengan al menos un vecino
      en el cluster liviano (detectado por BallTree k-NN). Mueve la
      UPM que maximiza la reducción del CV.

    MODO GLOBAL (fallback — se activa cuando frontera se atasca):
      Cuando el modo frontera no encuentra mejoras en 3 iteraciones
      consecutivas (por ejemplo, clusters geográficamente separados
      sin frontera directa), mueve las UPMs del cluster pesado que
      estén geográficamente más cerca del centroide del cluster
      liviano. Esto permite balancear zonas densas vs. dispersas
      que no son vecinas geográficas.

      Ejemplo práctico: zona urbana densa (500 viv/UPM) y zona rural
      dispersa (20 viv/UPM) a 50 km de distancia. El modo frontera no
      puede hacer swaps porque no hay vecinos en común. El modo global
      toma las UPMs de la zona densa más cercanas a la zona rural y
      las reasigna, logrando balance aunque sacrifique algo de compacidad.

    Criterio de parada: CV < cv_objetivo O 5 iteraciones globales
    consecutivas sin mejora.

    Retorna
    -------
    labels : np.ndarray shape (N,)  — cluster id para cada UPM
    log    : list[dict]             — historial por iteración
    cv_ini : float                  — CV antes del rebalanceo
    cv_fin : float                  — CV después del rebalanceo
    """
    coords = df[['x','y']].values.astype(float)
    cargas = df['carga_pond'].values.astype(float)
    n      = len(df)

    # ── Fase 1: KMeans ──────────────────────────────────────────────────────
    km = KMeans(n_clusters=n_clusters, init='k-means++', n_init=20,
                max_iter=500, random_state=42)
    labels = km.fit_predict(coords).copy()

    def cluster_sums():
        return np.array([cargas[labels == c].sum() for c in range(n_clusters)])

    def centroides_actuales():
        return np.array([
            coords[labels == c].mean(axis=0) if (labels == c).sum() > 0
            else np.zeros(2)
            for c in range(n_clusters)
        ])

    cv_ini = cv_pct(pd.Series(cluster_sums()))
    log    = [{'iter': 0, 'cv': cv_ini, 'modo': 'inicial'}]

    # ── BallTree para detección de vecinos de frontera ───────────────────────
    tree       = BallTree(coords, leaf_size=40)
    _, nbr_idx = tree.query(coords, k=min(k_vecinos + 1, n))

    no_mejora_frontera = 0   # iteraciones consecutivas sin swap de frontera

    for it in range(1, max_iter + 1):
        sums = cluster_sums()
        cv   = cv_pct(pd.Series(sums))
        if cv <= cv_objetivo * 100:
            log.append({'iter': it, 'cv': cv, 'modo': 'CV objetivo alcanzado'})
            break

        orden_pesados  = np.argsort(sums)[::-1]
        orden_livianos = np.argsort(sums)
        mejora_encontrada = False

        # ── MODO 1: Frontera ─────────────────────────────────────────────────
        if no_mejora_frontera < 3:
            for ci_p in orden_pesados[:3]:
                for ci_l in orden_livianos[:3]:
                    if ci_p == ci_l:
                        continue
                    mask_p = np.where(labels == ci_p)[0]
                    if len(mask_p) == 0:
                        continue
                    mejor_cv, mejor_idx = cv, -1
                    for idx in mask_p:
                        vecinos = nbr_idx[idx]
                        if not any(labels[v] == ci_l for v in vecinos if v != idx):
                            continue
                        labels[idx] = ci_l
                        cv_nuevo = cv_pct(pd.Series(cluster_sums()))
                        labels[idx] = ci_p
                        if cv_nuevo < mejor_cv:
                            mejor_cv, mejor_idx = cv_nuevo, idx
                    if mejor_idx >= 0:
                        labels[mejor_idx] = ci_l
                        log.append({'iter': it, 'cv': mejor_cv, 'modo': 'frontera',
                                    'movimiento': f'UPM {mejor_idx}: {ci_p}→{ci_l}'})
                        mejora_encontrada    = True
                        no_mejora_frontera   = 0
                        break
                if mejora_encontrada:
                    break
            if not mejora_encontrada:
                no_mejora_frontera += 1

        # ── MODO 2: Global (fallback) ─────────────────────────────────────────
        # Se activa cuando el modo frontera lleva 3+ iteraciones sin mejorar.
        # Busca la mejor UPM del cluster pesado para mover al cluster liviano
        # más cercano geográficamente, priorizando las UPMs del pesado que
        # están más cerca del centroide del liviano.
        if not mejora_encontrada:
            cents = centroides_actuales()
            ci_p  = orden_pesados[0]
            mask_p = np.where(labels == ci_p)[0]
            mejor_cv_g, mejor_idx_g, mejor_dest = cv, -1, -1

            for ci_l in orden_livianos[:2]:
                if ci_p == ci_l or len(mask_p) == 0:
                    continue
                # Candidatas: las 10 UPMs de ci_p más cercanas al centroide de ci_l
                dists_al_dest = np.linalg.norm(coords[mask_p] - cents[ci_l], axis=1)
                top_cands     = mask_p[np.argsort(dists_al_dest)[:10]]
                for idx in top_cands:
                    labels[idx] = ci_l
                    cv_nuevo = cv_pct(pd.Series(cluster_sums()))
                    labels[idx] = ci_p
                    if cv_nuevo < mejor_cv_g:
                        mejor_cv_g, mejor_idx_g, mejor_dest = cv_nuevo, idx, ci_l

            if mejor_idx_g >= 0 and mejor_cv_g < cv:
                labels[mejor_idx_g] = mejor_dest
                log.append({'iter': it, 'cv': mejor_cv_g, 'modo': 'global',
                            'movimiento': f'UPM {mejor_idx_g}: {ci_p}→{mejor_dest}'})
                mejora_encontrada  = True
                no_mejora_frontera = 0   # volver al modo frontera
            else:
                log.append({'iter': it, 'cv': cv, 'modo': 'sin mejora'})
                consec = sum(1 for l in log[-6:] if l.get('modo') == 'sin mejora')
                if consec >= 5:
                    break   # óptimo local estable

    cv_fin = cv_pct(pd.Series(cluster_sums()))
    log.append({'iter': len(log), 'cv': cv_fin, 'modo': 'final'})
    return labels, log, cv_ini, cv_fin


# ══════════════════════════════════════════════════════════════════════════════
#  NUEVO — B+C. ASIGNACIÓN ENCUESTADORES + DÍAS (con contigüidad geográfica)
# ══════════════════════════════════════════════════════════════════════════════

def asignar_encuestadores_y_dias(df_grp, n_enc, dias_tot, viv_min, viv_max,
                                  inicio_dia=1):
    """
    Asignación de encuestadores con contigüidad geográfica y distribución
    diaria balanceada con calendario de bloqueo.

    PASO 1 — Bin-packing balanceado
    ─────────────────────────────────
    Divide las UPMs del cluster entre n_enc encuestadores intentando igualar
    la suma de carga_pond. Usa el mismo greedy que v3 (argmin de carga acum.)
    pero sobre UPMs ordenadas por carga DESC para mejor balance.

    PASO 2 — Reordenamiento geográfico por encuestador
    ───────────────────────────────────────────────────
    Una vez que cada encuestador tiene su lista de UPMs, las reordena por
    nearest-neighbor greedy partiendo del centroide de su sub-grupo.
    Resultado: días consecutivos = manzanas físicamente adyacentes.

    PASO 3 — Calendario de bloqueo por día
    ───────────────────────────────────────
    Cada encuestador tiene un dict: calendario = {día: viv_acum}.
    target_dia = promedio(viv_min, viv_max).

    Para cada manzana (en orden geográfico):
    · Normal (viv ≤ viv_max): buscar el primer día d ∈ [cursor, último]
      donde calendario[d] < target. Asignar allí. Actualizar calendario.
    · Grande (viv > viv_max): necesita D = ceil(viv/target) días consecutivos.
      Buscar el primer bloque de D días seguidos donde todos tengan espacio.
      Reservar esos D días. Avanzar cursor al día D+1.

    Ventaja sobre v3: días intermedios no quedan vacíos porque las manzanas
    normales "rellenan" los huecos que deja una manzana grande.
    """
    target    = (viv_min + viv_max) / 2.0
    ultimo    = inicio_dia + dias_tot - 1

    df_g      = df_grp.copy().reset_index(drop=True)
    n_rows    = len(df_g)
    cargas    = df_g['carga_pond'].values.astype(float)
    coords    = df_g[['x','y']].values.astype(float)

    # ── PASO 1: bin-packing (igual que v3) ──────────────────────────────────
    orden_bp  = np.argsort(cargas)[::-1]          # mayor carga primero
    enc_acum  = np.zeros(n_enc)
    enc_asig  = np.zeros(n_rows, dtype=int)

    for loc in orden_bp:
        em = int(np.argmin(enc_acum))
        enc_asig[loc] = em + 1
        enc_acum[em] += cargas[loc]

    df_g['encuestador'] = enc_asig

    # ── PASO 2: reordenamiento geográfico por encuestador ───────────────────
    # Almacenamos el orden geográfico como lista de índices originales
    geo_order_global = []   # lista de (idx_original, enc_id) en orden de visita

    for enc_id in range(1, n_enc + 1):
        idx_enc = df_g[df_g['encuestador'] == enc_id].index.tolist()
        if not idx_enc:
            continue
        sub_coords = coords[idx_enc]
        centroide  = sub_coords.mean(axis=0)
        nn_order   = nearest_neighbor_order(sub_coords, start_xy=centroide)
        # nn_order[i] = posición dentro de idx_enc
        for pos in nn_order:
            geo_order_global.append((idx_enc[pos], enc_id))

    # ── PASO 3: calendario de bloqueo por día ───────────────────────────────
    dias_range    = list(range(inicio_dia, ultimo + 1))
    # calendario[enc_id][dia] = viviendas ya asignadas ese día
    calendario    = {e: {d: 0.0 for d in dias_range} for e in range(1, n_enc + 1)}
    # cursor[enc_id] = día mínimo desde el que buscar (evita retroceder)
    cursor        = {e: inicio_dia for e in range(1, n_enc + 1)}

    dia_ini_col   = [inicio_dia] * n_rows
    dia_fin_col   = [inicio_dia] * n_rows

    for idx_orig, enc_id in geo_order_global:
        viv_m = max(0.0, float(df_g.at[idx_orig, 'viv']))
        cal   = calendario[enc_id]
        cur   = cursor[enc_id]

        if viv_m > viv_max:
            # ── MANZANA GRANDE ──────────────────────────────────────────────
            dias_m = max(1, int(np.ceil(viv_m / target)))
            dias_m = min(dias_m, dias_tot)   # no más que el total disponible

            # Buscar primer bloque de días_m días consecutivos desde cursor
            # donde TODOS tengan espacio (cal[d] < target)
            bloque_inicio = None
            for d_start in range(cur, ultimo - dias_m + 2):
                if all(cal.get(d, target) < target
                       for d in range(d_start, d_start + dias_m)):
                    bloque_inicio = d_start
                    break

            if bloque_inicio is None:
                # No hay bloque limpio → usar desde cursor hasta el límite
                bloque_inicio = min(cur, ultimo - dias_m + 1)
                bloque_inicio = max(bloque_inicio, inicio_dia)

            d_ini = bloque_inicio
            d_fin = min(d_ini + dias_m - 1, ultimo)

            viv_por_dia = viv_m / max(1, d_fin - d_ini + 1)
            for dd in range(d_ini, d_fin + 1):
                cal[dd] = cal.get(dd, 0.0) + viv_por_dia

            cursor[enc_id] = d_fin + 1   # siguiente manzana empieza después

        else:
            # ── MANZANA NORMAL ──────────────────────────────────────────────
            # Buscar primer día desde cursor con espacio disponible
            dia_asig = None
            for d in range(cur, ultimo + 1):
                if cal.get(d, 0.0) < target:
                    dia_asig = d
                    break
            if dia_asig is None:
                dia_asig = ultimo    # overflow: ir al último día

            cal[dia_asig] = cal.get(dia_asig, 0.0) + viv_m
            d_ini = dia_asig
            d_fin = dia_asig

            # Avanzar cursor solo si el día actual ya está "lleno"
            if cal[dia_asig] >= target and cursor[enc_id] == dia_asig:
                cursor[enc_id] = min(dia_asig + 1, ultimo)

        # Clamp de seguridad
        d_ini = max(inicio_dia, min(d_ini, ultimo))
        d_fin = max(d_ini,      min(d_fin, ultimo))

        dia_ini_col[idx_orig] = d_ini
        dia_fin_col[idx_orig] = d_fin

    df_g['dia_inicio']    = dia_ini_col
    df_g['dia_fin']       = dia_fin_col
    df_g['dia_operativo'] = dia_ini_col   # retrocompatibilidad
    return df_g


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL (sin cambios funcionales respecto a v3)
# ══════════════════════════════════════════════════════════════════════════════

def generar_excel(df_plan, eq_cfg, personal_info,
                  fecha_j1, fecha_j2, dias_op, j1_num, j2_num, mes_nombre):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    AZ_OSCURO = "0D3B6E"; AZ_MEDIO="1A5276"; AZ_CLARO="D6EAF8"
    VRD_CHECK  = "D5F5E3"; BLANCO="FFFFFF"
    ENC_PALETAS = [
        {"par":"DBEAFE","impar":"EFF6FF","subtot":"BFDBFE","hdr":"1D4ED8"},
        {"par":"D1FAE5","impar":"ECFDF5","subtot":"A7F3D0","hdr":"065F46"},
        {"par":"FEF9C3","impar":"FEFCE8","subtot":"FDE68A","hdr":"854D0E"},
        {"par":"FCE7F3","impar":"FDF4FF","subtot":"F9A8D4","hdr":"831843"},
        {"par":"FFE4E6","impar":"FFF1F2","subtot":"FECACA","hdr":"9F1239"},
        {"par":"E0E7FF","impar":"EEF2FF","subtot":"C7D2FE","hdr":"3730A3"},
    ]

    def sc(cell, bold=False, bg=None, fg="000000",
           ha="left", sz=9, brd=False, wrap=False, italic=False):
        cell.font      = Font(bold=bold, size=sz, color=fg, italic=italic)
        cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        if brd:
            t = Side(style='thin')
            cell.border = Border(left=t, right=t, top=t, bottom=t)

    ct_counter = [700]

    for jornada_nombre, fecha_inicio, n_jornada_hoja in [
        ("Jornada 1", fecha_j1, j1_num),
        ("Jornada 2", fecha_j2, j2_num)
    ]:
        df_jor = df_plan[df_plan['jornada'] == jornada_nombre].copy()
        if len(df_jor) == 0:
            continue

        ws = wb.create_sheet(title=jornada_nombre)
        ws.sheet_view.showGridLines = False

        anchos = {'A':14,'B':14,'C':8,'D':5,'E':5,'F':6,
                  'G':6,'H':6,'I':5,'J':18,'K':13,'L':10,'M':14,'N':5}
        for col_l, w in anchos.items():
            ws.column_dimensions[col_l].width = w
        for i in range(dias_op):
            ws.column_dimensions[get_column_letter(15+i)].width = 7
        ws.column_dimensions[get_column_letter(15+dias_op)].width = 6

        cur = 1
        equipos_jor = [e['nombre'] for e in eq_cfg
                       if e['nombre'] in df_jor['equipo'].values]

        for grupo_num, nombre_eq in enumerate(equipos_jor, 1):
            df_eq = df_jor[df_jor['equipo'] == nombre_eq].copy()
            if len(df_eq) == 0: continue

            pi    = personal_info.get(nombre_eq, {})
            n_enc = next((e['enc'] for e in eq_cfg if e['nombre']==nombre_eq), 3)

            if fecha_inicio:
                fechas  = [fecha_inicio + timedelta(days=i) for i in range(dias_op)]
                fi_str  = fecha_inicio.strftime("%d-%b-%y").upper()
                ff_str  = fechas[-1].strftime("%d-%b-%y").upper()
            else:
                fechas  = None; fi_str = "____"; ff_str = "____"

            last_col_idx = 15 + dias_op

            def merge_row(row, c1, c2, val, **kw):
                ws.merge_cells(f'{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}')
                c = ws.cell(row, c1, val)
                sc(c, **kw)
                return c

            for txt in ["INSTITUTO NACIONAL DE ESTADÍSTICA Y CENSOS",
                        "COORDINACIÓN ZONAL LITORAL CZ8L",
                        "ACTUALIZACIÓN CARTOGRÁFICA - ENDI ENLISTAMIENTO",
                        "PROGRAMACIÓN OPERATIVO DE CAMPO"]:
                merge_row(cur,1,last_col_idx,txt,bold=True,bg=AZ_OSCURO,
                          fg=BLANCO,ha="center",sz=9)
                cur += 1
            cur += 1

            ws.cell(cur,1,"JORNADA")
            jorn_cell = ws.cell(cur,2,str(n_jornada_hoja))
            jorn_cell.font = Font(bold=True,size=11)
            ws.cell(cur,7,"GRUPO")
            ws.cell(cur,9,str(grupo_num)).font = Font(bold=True,size=11)
            sc(ws.cell(cur,1),bold=True,sz=10)
            sc(ws.cell(cur,7),bold=True,sz=10)
            cur += 2

            ws.cell(cur,1,"PERÍODO DE ACTUALIZACIÓN:")
            sc(ws.cell(cur,1),bold=True,sz=9)
            ws.cell(cur,5,"DEL"); ws.cell(cur,6,fi_str)
            ws.cell(cur,9,"AL"); ws.cell(cur,10,ff_str)
            cur += 2

            for col,txt in [(3,"COD."),(4,"NOMBRE"),(8,"No. CÉDULA"),(11,"No. CELULAR")]:
                sc(ws.cell(cur,col,txt),bold=True,sz=8)
            cur += 1

            ws.cell(cur,1,"SUPERVISOR:")
            ws.cell(cur,3,pi.get('supervisor_cod',''))
            ws.cell(cur,4,pi.get('supervisor_nombre',''))
            ws.cell(cur,8,pi.get('supervisor_cedula',''))
            ws.cell(cur,11,pi.get('supervisor_celular',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 2

            enc_list = pi.get('encuestadores', [])
            for j in range(n_enc):
                info = enc_list[j] if j < len(enc_list) else {}
                ws.cell(cur,1,"ENCUESTADOR")
                ws.cell(cur,3,info.get('cod',''))
                ws.cell(cur,4,info.get('nombre',''))
                ws.cell(cur,8,info.get('cedula',''))
                ws.cell(cur,11,info.get('celular',''))
                sc(ws.cell(cur,1),bold=True,sz=9)
                cur += 1
            cur += 1

            ws.cell(cur,1,"VEHÍCULO: CHOFER")
            ws.cell(cur,4,pi.get('chofer_nombre',''))
            ws.cell(cur,8,pi.get('chofer_cedula',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 1
            ws.cell(cur,1,"PLACA:")
            ws.cell(cur,4,pi.get('placa',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 2

            merge_row(cur,1,4,"EQUIPO",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            merge_row(cur,5,14,"IDENTIFICACIÓN",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            merge_row(cur,15,14+dias_op,"RECORRIDO DE LOS SECTORES EN LA JORNADA — FECHA",
                      bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=7,brd=True)
            sc(ws.cell(cur,last_col_idx,"# VIV"),bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            ws.row_dimensions[cur].height = 24
            cur += 1

            sub_hdrs = ["SUPERVISOR","ENCUESTADOR","CARGA DE TRABAJO",
                        "PROV","CANTON","CIUDAD O PARROQ","ZONA","SECTOR","MAN",
                        "CÓDIGO DE LA JURISDICCIÓN","PROVINCIA","CANTÓN",
                        "CIUDAD, PARROQ. O LOC AMAZ.","NRO EDIF"]
            for ci,h in enumerate(sub_hdrs,1):
                sc(ws.cell(cur,ci,h),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True,wrap=True)
            for i in range(dias_op):
                lbl = fechas[i].strftime("%d/%m") if fechas else f"Día {i+1}"
                sc(ws.cell(cur,15+i,lbl),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)
            sc(ws.cell(cur,last_col_idx,"# VIV"),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)
            ws.row_dimensions[cur].height = 32
            cur += 1

            df_sorted = df_eq.sort_values(['encuestador','dia_inicio']).copy()
            enc_actual = None; fila_enc = 0; viv_enc_acum = 0; enc_color_idx = -1

            for ri, (_, rd) in enumerate(df_sorted.iterrows()):
                enc_id = int(rd.get('encuestador', 0))
                if enc_id != enc_actual and enc_actual is not None:
                    pal_sub = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                    enc_info_prev = enc_list[enc_actual-1] if 0 < enc_actual <= len(enc_list) else {}
                    merge_row(cur,1,9,
                              f"SUBTOTAL {enc_info_prev.get('nombre',f'Encuestador {enc_actual}')}",
                              bold=True,bg=pal_sub["subtot"],fg=pal_sub["hdr"],ha="right",sz=8)
                    for ci in range(10,last_col_idx):
                        sc(ws.cell(cur,ci,""),bg=pal_sub["subtot"],brd=True)
                    sc(ws.cell(cur,last_col_idx,viv_enc_acum),
                       bold=True,ha="center",sz=9,bg=pal_sub["subtot"],fg=pal_sub["hdr"],brd=True)
                    ws.row_dimensions[cur].height = 14
                    cur += 1; viv_enc_acum = 0

                if enc_id != enc_actual:
                    enc_actual = enc_id; fila_enc = 0
                    enc_color_idx = (enc_color_idx + 1) % len(ENC_PALETAS)

                pal    = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                bg_row = pal["par"] if fila_enc % 2 == 0 else pal["impar"]
                fila_enc += 1
                viv_enc_acum += int(rd.get('viv',0))

                p_cod   = parse_codigo(str(rd['id_entidad']))
                enc_i   = enc_list[enc_id-1] if 0 < enc_id <= len(enc_list) else {}
                ct_str  = f"CT{ct_counter[0]:03d}"
                ct_counter[0] += 1

                row_vals = [
                    pi.get('supervisor_cedula',''), enc_i.get('cedula',''), ct_str,
                    p_cod['prov'],p_cod['canton'],p_cod['ciudad_parroq'],
                    p_cod['zona'],p_cod['sector'],p_cod['man'],
                    str(rd['id_entidad']),'','','','',
                ]
                for ci,val in enumerate(row_vals,1):
                    c = ws.cell(cur,ci,val)
                    sc(c,bg=AZ_CLARO if ci==10 else bg_row,ha="center",sz=8,brd=True)

                d_ini = int(rd.get('dia_inicio',rd.get('dia_operativo',1)))
                d_fin = int(rd.get('dia_fin',d_ini))
                for i in range(dias_op):
                    dia_num = i + 1
                    if d_ini <= dia_num <= d_fin:
                        c = ws.cell(cur,15+i,"✓")
                        sc(c,bold=True,bg=VRD_CHECK,ha="center",sz=11,brd=True)
                    else:
                        sc(ws.cell(cur,15+i,""),bg=bg_row,ha="center",brd=True)

                sc(ws.cell(cur,last_col_idx,int(rd.get('viv',0))),
                   ha="center",sz=8,brd=True,bg=bg_row)
                cur += 1

            if enc_actual is not None:
                pal_sub  = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                enc_info = enc_list[enc_actual-1] if 0 < enc_actual <= len(enc_list) else {}
                merge_row(cur,1,9,
                          f"SUBTOTAL {enc_info.get('nombre',f'Encuestador {enc_actual}')}",
                          bold=True,bg=pal_sub["subtot"],fg=pal_sub["hdr"],ha="right",sz=8)
                for ci in range(10,last_col_idx):
                    sc(ws.cell(cur,ci,""),bg=pal_sub["subtot"],brd=True)
                sc(ws.cell(cur,last_col_idx,viv_enc_acum),
                   bold=True,ha="center",sz=9,
                   bg=pal_sub["subtot"],fg=pal_sub["hdr"],brd=True)
                ws.row_dimensions[cur].height = 14
                cur += 1

            tot_viv_eq = int(df_eq['viv'].sum())
            sc(ws.cell(cur,last_col_idx-1,"TOTAL"),bold=True,ha="right",sz=8,bg=AZ_CLARO,brd=True)
            sc(ws.cell(cur,last_col_idx,tot_viv_eq),bold=True,ha="center",sz=8,bg=AZ_CLARO,brd=True)
            cur += 4

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── SESSION STATE ─────────────────────────────
_defs = {
    "data_raw": None, "data_mes": None, "graph_G": None,
    "resultados_generados": False, "df_plan": None,
    "tsp_results": {}, "road_paths": {}, "resumen_bal": None,
    "sil_score": None, "n_bombero": 0,
    "personal_info": {},
    "fecha_j1": None, "fecha_j2": None,
    "j1_num": 1, "j2_num": 2,
    "balance_log": [],          # NUEVO: historial del rebalanceo
    "cv_ini_bal": None,         # NUEVO: CV antes de rebalanceo
    "cv_fin_bal": None,         # NUEVO: CV después de rebalanceo
    "viv_por_cluster_antes": None,   # NUEVO
    "viv_por_cluster_despues": None, # NUEVO
    "params": {
        "dias_op":12,"viv_min":50,"viv_max":80,"factor_r":1.5,
        "usar_bomb":True,"usar_gye":True,"dias_gye":3,"umbral_gye":10,
        # NUEVO: parámetros de rebalanceo
        "cv_objetivo": 10,    # % — detener cuando CV < esto
        "max_iter_bal": 200,  # iteraciones máximas de swap
        "k_vecinos": 8,       # vecinos para detección de frontera
    },
    "equipos_cfg": [
        {"id":1,"nombre":"Equipo 1","enc":3},
        {"id":2,"nombre":"Equipo 2","enc":3},
        {"id":3,"nombre":"Equipo 3","enc":3},
    ],
}
for k,v in _defs.items():
    if k not in st.session_state: st.session_state[k] = v

# ── SIDEBAR ───────────────────────────────────
with st.sidebar:
    st.markdown("### 🗺️ Encuesta Nacional")
    st.markdown("<p style='font-size:10px;color:#445566;margin-top:-8px'>INEC · Zonal Litoral</p>",
                unsafe_allow_html=True)
    st.divider()

    st.markdown("<div class='step'>PASO 1</div>", unsafe_allow_html=True)
    st.markdown("**Muestra (.gpkg)**")
    gpkg_f = st.file_uploader("GeoPackage", type=["gpkg"], key="gpkg_up")
    if gpkg_f:
        dissolve = st.radio("Nivel",["Por UPM","Por manzana"],index=0)
        if st.button("⚡ Procesar", use_container_width=True, type="primary"):
            with st.spinner("Leyendo geometrías..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False,suffix=".gpkg") as tmp:
                        tmp.write(gpkg_f.read()); p=tmp.name
                    data=cargar_gpkg(p,dissolve_upm=dissolve.startswith("Por UPM"))
                    os.unlink(p)
                    st.session_state.data_raw=data
                    st.session_state.resultados_generados=False
                    st.success(f"✓ {len(data):,} entidades")
                except Exception as e: st.error(str(e))
        if st.session_state.data_raw is not None:
            st.markdown("<span class='pill-ok'>✓ Listo</span>",unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin archivo</span>",unsafe_allow_html=True)

    st.divider()

    st.markdown("<div class='step'>PASO 2</div>", unsafe_allow_html=True)
    st.markdown("**Red vial (.graphml)**")
    gml_f = st.file_uploader("GraphML", type=["graphml"], key="gml_up")
    if gml_f:
        if st.button("⚡ Cargar grafo", use_container_width=True):
            with st.spinner("Cargando..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False,suffix=".graphml") as tmp:
                        tmp.write(gml_f.read()); pg=tmp.name
                    G=ox.load_graphml(pg); os.unlink(pg)
                    st.session_state.graph_G=G
                    st.success(f"✓ {len(G.nodes):,} nodos")
                except Exception as e: st.error(str(e))
        if st.session_state.graph_G is not None:
            st.markdown("<span class='pill-ok'>✓ Red lista</span>",unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin grafo</span>",unsafe_allow_html=True)

    st.divider()

    if st.session_state.data_raw is not None:
        st.markdown("<div class='step'>PASO 3</div>", unsafe_allow_html=True)
        st.markdown("**Mes operativo**")
        meses_disp = sorted(st.session_state.data_raw["mes"].dropna().unique().tolist())
        mes_sel = st.selectbox("Mes", meses_disp,
            format_func=lambda x: f"{MESES_N.get(int(x),str(int(x)))} (mes {int(x)})")
        df_mes = st.session_state.data_raw[st.session_state.data_raw["mes"]==mes_sel].copy()
        st.session_state.data_mes = df_mes

        st.divider()

        st.markdown("<div class='step'>PASO 4</div>", unsafe_allow_html=True)
        st.markdown("**Equipos**")
        c1,c2 = st.columns(2)
        with c1:
            if st.button("＋",use_container_width=True):
                nid=max(t["id"] for t in st.session_state.equipos_cfg)+1
                st.session_state.equipos_cfg.append({"id":nid,"nombre":f"Equipo {nid}","enc":3})
                st.session_state.resultados_generados=False
        with c2:
            if st.button("－",use_container_width=True,
                         disabled=len(st.session_state.equipos_cfg)<=1):
                st.session_state.equipos_cfg.pop()
                st.session_state.resultados_generados=False

        for i,eq in enumerate(st.session_state.equipos_cfg):
            cc1,cc2=st.columns([2,1])
            with cc1:
                nn=st.text_input(f"n{eq['id']}",value=eq["nombre"],
                                 key=f"n_{eq['id']}",label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["nombre"]=nn
            with cc2:
                ne=st.number_input("e",min_value=1,max_value=6,value=eq["enc"],
                                   key=f"e_{eq['id']}",label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["enc"]=ne

        st.divider()
        st.markdown("**Parámetros operativos**")
        p = st.session_state.params
        p["dias_op"]  = st.slider("Días operativos",10,14,p["dias_op"])
        p["viv_min"]  = st.slider("Mín viv/día",30,60,p["viv_min"])
        p["viv_max"]  = st.slider("Máx viv/día",60,120,p["viv_max"])
        p["factor_r"] = st.slider("Factor rural (×)",1.0,2.5,p["factor_r"],0.1,
            help="Viviendas dispersas pesan X veces más para el balance.")

        st.markdown("**Parámetros de rebalanceo** *(nuevo v4)*")
        p["cv_objetivo"]  = st.slider("CV objetivo clusters (%)", 3, 25,
                                       p.get("cv_objetivo",10),
            help="El rebalanceo se detiene cuando CV de viviendas entre clusters cae por debajo de este valor.")
        p["max_iter_bal"] = st.slider("Iter. máx. rebalanceo", 50, 500,
                                       p.get("max_iter_bal",200), step=50,
            help="Máximo de intercambios de UPMs entre clusters.")
        p["k_vecinos"]    = st.slider("Vecinos frontera (k)", 4, 20,
                                       p.get("k_vecinos",8),
            help="Cuántos vecinos cercanos se consideran para detectar UPMs en la frontera entre clusters.")

        p["usar_bomb"] = st.toggle("Equipo Bombero",value=p["usar_bomb"])
        if p["usar_bomb"]:
            p["min_dist_bomb_m"] = st.slider(
                "Distancia mín. Bombero (km)", 10, 150,
                p.get("min_dist_bomb_m", 40000) // 1000) * 1000
        p["usar_gye"]   = st.toggle("Restricción Guayaquil",value=p["usar_gye"])
        p["dias_gye"]   = st.slider("Días GYE",1,5,p["dias_gye"],disabled=not p["usar_gye"])
        p["umbral_gye"] = st.slider("Umbral GYE (%)",5,30,p["umbral_gye"],disabled=not p["usar_gye"])

        j1_num = (int(mes_sel) - 1) * 2 + 1
        j2_num = j1_num + 1
        st.session_state.j1_num = j1_num
        st.session_state.j2_num = j2_num
        st.divider()
        st.markdown(f"""
        <div style='font-size:11px;background:#0d2035;border-radius:6px;
                    padding:8px 12px;border-left:3px solid #2e86de'>
        📅 Mes {int(mes_sel)} →
        <b style='color:#2e86de'>Jornada {j1_num}</b> +
        <b style='color:#27ae60'>Jornada {j2_num}</b>
        </div>""", unsafe_allow_html=True)

        tot_enc = sum(e["enc"] for e in st.session_state.equipos_cfg)
        tot_viv = int(df_mes["viv"].sum()) if len(df_mes)>0 else 0
        st.markdown(f"""
        <div style='font-size:11px;color:#445566;line-height:2;margin-top:8px'>
        📍 <b style='color:#7eb3d8'>{len(df_mes):,}</b> UPMs · mes {int(mes_sel)}<br>
        🏠 <b style='color:#7eb3d8'>{tot_viv:,}</b> viviendas<br>
        👥 <b style='color:#7eb3d8'>{len(st.session_state.equipos_cfg)}</b> equipos ·
           <b style='color:#7eb3d8'>{tot_enc}</b> enc.
        </div>""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────
st.markdown("""
<div class='hdr'>
  <h1>Planificación Automática · Actualización Cartográfica · v4</h1>
  <p>Encuesta Nacional &nbsp;·&nbsp; Zonal Litoral &nbsp;·&nbsp; INEC Ecuador</p>
</div>""", unsafe_allow_html=True)

if st.session_state.data_raw is None:
    st.markdown("<div class='ibox'>👈 Carga el <code>.gpkg</code> desde el panel lateral.</div>",
                unsafe_allow_html=True)
    st.stop()

df = st.session_state.data_mes
if df is None or len(df)==0:
    st.warning("Sin datos para el mes seleccionado."); st.stop()

p = st.session_state.params
k1,k2,k3,k4,k5 = st.columns(5)
cv_v=cv_pct(df["viv"]); cv_c="#27ae60" if cv_v<50 else "#e74c3c"
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

st.markdown("<br>",unsafe_allow_html=True)

cb1,cb2=st.columns([1,3])
with cb1:
    btn=st.button("⚡ Generar Planificación",use_container_width=True,
                  type="primary",disabled=(st.session_state.graph_G is None))
with cb2:
    if st.session_state.graph_G is None:
        st.markdown("<div class='wbox'>⚠️ Carga el <code>.graphml</code> (Paso 2).</div>",
                    unsafe_allow_html=True)
    elif st.session_state.resultados_generados:
        st.markdown("<div class='ibox'>✓ Planificación lista. Puedes regenerar.</div>",
                    unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  ALGORITMO PRINCIPAL v4
# ═══════════════════════════════════════════════════════════════════════════════
if btn:
    G       = st.session_state.graph_G
    eq_cfg  = st.session_state.equipos_cfg
    n_eq    = len(eq_cfg)
    nombres = [e["nombre"] for e in eq_cfg]
    n_clust = n_eq * 2
    p       = st.session_state.params

    df_w = df.copy()
    df_w['equipo']      = 'sin_asignar'
    df_w['jornada']     = 'sin_asignar'
    df_w['cluster_geo'] = -1
    df_w['carga_pond']  = df_w.apply(
        lambda r: r['viv']*p["factor_r"]
        if str(r.get('tipo_entidad','')).startswith('sec') else r['viv'], axis=1)
    df_w['encuestador']   = 0
    df_w['dia_operativo'] = 0
    df_w['dia_inicio']    = 0
    df_w['dia_fin']       = 0
    df_w['dist_base_m']   = 0.0

    prog = st.progress(0,"Iniciando...")

    # 1. Distancias a la base
    prog.progress(5,"Calculando distancias a la base...")
    t_utm = Transformer.from_crs("EPSG:4326","EPSG:32717",always_xy=True)
    bx,by = t_utm.transform(BASE_LON,BASE_LAT)
    df_w['dist_base_m'] = np.sqrt((df_w['x']-bx)**2+(df_w['y']-by)**2)

    # 2. Restricción Guayaquil
    prog.progress(8,"Verificando restricción Guayaquil...")
    upms_gye = pd.Series(False,index=df_w.index)
    if p["usar_gye"] and 'pro_x' in df_w.columns and 'can_x' in df_w.columns:
        upms_gye = (df_w['pro_x']==PRO_GYE)&(df_w['can_x']==CAN_GYE)
    pct_gye  = upms_gye.sum()/len(df_w) if len(df_w)>0 else 0
    act_gye  = p["usar_gye"] and (pct_gye >= p["umbral_gye"]/100) and upms_gye.sum()>0

    df_gye    = df_w[upms_gye].copy()  if act_gye else pd.DataFrame()
    df_no_gye = df_w[~upms_gye].copy()

    # ── 3. CLUSTERING BALANCEADO (NUEVO v4) ──────────────────────────────────
    prog.progress(12,f"Fase 1: KMeans geográfico ({n_clust} clusters)...")
    mask_bomb_global = pd.Series(False, index=df_w.index)

    if len(df_no_gye) >= n_clust:
        df_no_gye = df_no_gye.copy()

        # Guardar viv por cluster ANTES del rebalanceo (para el gráfico comparativo)
        km_init = KMeans(n_clusters=n_clust, init='k-means++', n_init=20,
                         max_iter=500, random_state=42)
        labels_init = km_init.fit_predict(df_no_gye[['x','y']].values.astype(float))
        viv_antes = {c: float(df_no_gye['carga_pond'].values[labels_init==c].sum())
                     for c in range(n_clust)}
        st.session_state.viv_por_cluster_antes = viv_antes

        prog.progress(22,f"Fase 2: rebalanceo por viviendas (CV objetivo {p['cv_objetivo']}%)...")
        labels_bal, bal_log, cv_ini, cv_fin = clustering_balanceado(
            df_no_gye,
            n_clusters  = n_clust,
            factor_r    = p["factor_r"],
            cv_objetivo = p["cv_objetivo"] / 100.0,
            max_iter    = p["max_iter_bal"],
            k_vecinos   = p["k_vecinos"],
        )
        df_no_gye['cluster_geo'] = labels_bal
        st.session_state.balance_log      = bal_log
        st.session_state.cv_ini_bal       = cv_ini
        st.session_state.cv_fin_bal       = cv_fin
        st.session_state.sil_score        = None
        try:
            from sklearn.metrics import silhouette_score as sil_fn
            st.session_state.sil_score = sil_fn(
                df_no_gye[['x','y']].values, labels_bal)
        except: pass

        viv_despues = {c: float(df_no_gye['carga_pond'].values[labels_bal==c].sum())
                       for c in range(n_clust)}
        st.session_state.viv_por_cluster_despues = vib_despues = viv_despues

        # Asignación cluster → (equipo, jornada)
        # Ordenar clusters por distancia al centroide de la base (desc = más lejanos primero)
        centroides = np.array([
            df_no_gye[['x','y']].values[labels_bal==c].mean(axis=0)
            if (labels_bal==c).sum() > 0 else np.array([bx,by])
            for c in range(n_clust)
        ])
        dist_c = np.sqrt((centroides[:,0]-bx)**2+(centroides[:,1]-by)**2)
        orden  = np.argsort(dist_c)[::-1]
        asig   = {}
        for i,(cj1,cj2) in enumerate(zip(orden[:n_eq],orden[n_eq:])):
            asig[cj1]=(nombres[i],'Jornada 1')
            asig[cj2]=(nombres[i],'Jornada 2')

        df_no_gye['equipo']  = df_no_gye['cluster_geo'].map(lambda c: asig[c][0])
        df_no_gye['jornada'] = df_no_gye['cluster_geo'].map(lambda c: asig[c][1])

        # Equipo Bombero (igual que v3 pero sobre clusters ya balanceados)
        if p["usar_bomb"]:
            prog.progress(32,"Detectando outliers por cluster (Equipo Bombero)...")
            MIN_DIST_BOMBERO_M = p.get("min_dist_bomb_m", 40000)
            for c_id in range(n_clust):
                if c_id not in asig: continue
                mask_c = df_no_gye['cluster_geo'] == c_id
                pts    = df_no_gye[mask_c]
                if len(pts) < 8: continue
                cx, cy = centroides[c_id]
                dists  = np.sqrt((pts['x']-cx)**2+(pts['y']-cy)**2)
                Q1c, Q3c = dists.quantile(.25), dists.quantile(.75)
                iqrc   = Q3c - Q1c
                if iqrc == 0: continue
                umbral_iqr = Q3c + 3.0 * iqrc
                bomb_cands = dists[(dists > umbral_iqr) & (dists > MIN_DIST_BOMBERO_M)]
                bomb_idx   = bomb_cands.index
                if len(bomb_idx) > 0:
                    df_no_gye.loc[bomb_idx,'equipo']  = 'Equipo Bombero'
                    df_no_gye.loc[bomb_idx,'jornada'] = 'Jornada Especial'
                    mask_bomb_global.loc[bomb_idx]    = True

        df_w.update(df_no_gye[['equipo','jornada','cluster_geo']])

    n_bomb = int((df_w['equipo']=='Equipo Bombero').sum())
    st.session_state.n_bombero = n_bomb

    # ── 4. ENCUESTADORES + DÍAS (NUEVO v4) ───────────────────────────────────
    prog.progress(42,"Asignando encuestadores con contigüidad geográfica...")
    enc_dict = {e["nombre"]:e["enc"] for e in eq_cfg}

    for nombre_eq in nombres:
        for jornada in ['Jornada 1','Jornada 2']:
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp    = df_w[mask_g].copy()
            if len(grp) == 0: continue
            n_enc  = enc_dict.get(nombre_eq, 3)

            if jornada == 'Jornada 1' and act_gye:
                inicio    = p["dias_gye"] + 1
                dias_disp = p["dias_op"] - p["dias_gye"]
            else:
                inicio    = 1
                dias_disp = p["dias_op"]

            ga = asignar_encuestadores_y_dias(
                grp, n_enc, dias_disp,
                p["viv_min"], p["viv_max"], inicio)
            df_w.update(ga[['encuestador','dia_operativo','dia_inicio','dia_fin']])

    # Fase Guayaquil
    if act_gye and len(df_gye) > 0:
        for i,(idx,row) in enumerate(df_gye.sort_values('carga_pond',ascending=False).iterrows()):
            eq_a  = nombres[i % n_eq]
            enc_a = (i // n_eq) % enc_dict.get(eq_a,3) + 1
            dia_a = min((i//(n_eq*enc_dict.get(eq_a,3)))+1, p["dias_gye"])
            df_w.loc[idx,['equipo','jornada','encuestador',
                          'dia_operativo','dia_inicio','dia_fin']] = \
                [eq_a,'Jornada 1',enc_a,dia_a,dia_a,dia_a]

    # 5. TSP
    prog.progress(52,"Optimizando rutas TSP...")
    base_nd   = ox.nearest_nodes(G,BASE_LON,BASE_LAT)
    G_u       = G.to_undirected()
    comp_base = nx.node_connected_component(G_u,base_nd)
    tsp_r, road_p = {}, {}

    for ri,nombre_eq in enumerate(nombres):
        for jornada in ['Jornada 1','Jornada 2']:
            pct = 52+int((ri*2+['Jornada 1','Jornada 2'].index(jornada)+1)/(n_eq*2)*42)
            prog.progress(pct,f"TSP: {nombre_eq} | {jornada}...")
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp    = df_w[mask_g]
            if len(grp)==0: continue
            nr  = ox.nearest_nodes(G,grp['lon'].values,grp['lat'].values)
            nk  = [n for n in nr if n in comp_base]
            if not nk: continue
            nu  = [base_nd]+list(dict.fromkeys(nk))
            n   = len(nu)
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
            try: ciclo=approximation.traveling_salesman_problem(Gt,weight='weight',cycle=True)
            except: continue
            if 0 in ciclo:
                i0=ciclo.index(0); ciclo=ciclo[i0:]+ciclo[1:i0+1]
            dist=sum(D[ciclo[i],ciclo[i+1]] for i in range(len(ciclo)-1))
            ruta=[]; ng=[nu[idx] for idx in ciclo]
            for k in range(len(ng)-1):
                try:
                    seg=nx.shortest_path(G,ng[k],ng[k+1],weight='length')
                    ruta.extend((G.nodes[nd]['y'],G.nodes[nd]['x']) for nd in seg[:-1])
                except: continue
            if ng: ruta.append((G.nodes[ng[-1]]['y'],G.nodes[ng[-1]]['x']))
            clave = f"{nombre_eq}||{jornada}"
            tsp_r[clave]  = {'equipo':nombre_eq,'jornada':jornada,
                              'n_puntos':len(grp),'dist_km':dist/1000}
            road_p[clave] = ruta

    prog.progress(98,"Calculando métricas finales...")
    resumen = df_w[~df_w['equipo'].isin(['Equipo Bombero','sin_asignar'])].groupby(
        ['equipo','jornada']).agg(
        n_upms=('id_entidad','count'),
        viv_reales=('viv','sum'),
        carga_ponderada=('carga_pond','sum')).reset_index()
    dist_df = pd.DataFrame([
        {'equipo':v['equipo'],'jornada':v['jornada'],'dist_km':round(v['dist_km'],1)}
        for v in tsp_r.values()
    ]) if tsp_r else pd.DataFrame(columns=['equipo','jornada','dist_km'])
    resumen_bal = pd.merge(resumen,dist_df,on=['equipo','jornada'],how='left').fillna(0)

    prog.progress(100,"¡Listo!"); prog.empty()
    st.session_state.df_plan     = df_w
    st.session_state.tsp_results = tsp_r
    st.session_state.road_paths  = road_p
    st.session_state.resumen_bal = resumen_bal
    st.session_state.resultados_generados = True
    st.success("✓ Planificación v4 generada con clustering balanceado.")

# ── RESULTADOS ────────────────────────────────
if not st.session_state.resultados_generados:
    st.markdown("<div class='ibox'>👆 Presiona <b>Generar Planificación</b>.</div>",
                unsafe_allow_html=True)
    st.stop()

df_plan = st.session_state.df_plan
tsp_r   = st.session_state.tsp_results
road_p  = st.session_state.road_paths
res_bal = st.session_state.resumen_bal
eq_cfg  = st.session_state.equipos_cfg
nombres = [e["nombre"] for e in eq_cfg]
p       = st.session_state.params

color_map = {n:COLORES[i%len(COLORES)] for i,n in enumerate(nombres)}
color_map['Equipo Bombero'] = '#9b59b6'

tab_mapa, tab_analisis, tab_reporte = st.tabs([
    "🗺️  Mapa de Rutas", "📊  Análisis de Carga", "📋  Reporte y Descarga"
])

# ══ TAB 1 — MAPA ══════════════════════════════
with tab_mapa:
    st.markdown("<div class='stitle'>Mapa del Operativo de Campo</div>",unsafe_allow_html=True)
    cc1,cc2 = st.columns([1,3])
    with cc1:
        mj1  = st.checkbox("Jornada 1",value=True)
        mj2  = st.checkbox("Jornada 2",value=True)
        n_b  = int((df_plan['equipo']=='Equipo Bombero').sum())
        mbm  = st.checkbox(f"Equipo Bombero ({n_b} UPMs)",value=True)
        mrts = st.checkbox("Mostrar rutas",value=True)
        fnd  = st.selectbox("Fondo",["CartoDB dark_matter","CartoDB positron","OpenStreetMap"])
        st.divider()
        st.markdown("**Leyenda:**")
        for n,c in color_map.items():
            if n in nombres:
                st.markdown(f"<span style='color:{c};font-size:17px'>●</span> {n}",
                            unsafe_allow_html=True)
        bl = f"Equipo Bombero {'(sin UPMs)' if n_b==0 else f'({n_b} UPMs)'}"
        st.markdown(f"<span style='color:#9b59b6;font-size:17px'>●</span> {bl}",
                    unsafe_allow_html=True)

    with cc2:
        m=folium.Map(location=[BASE_LAT,BASE_LON],zoom_start=8,tiles=fnd)
        folium.Marker([BASE_LAT,BASE_LON],popup="<b>Base INEC Guayaquil</b>",
            icon=folium.Icon(color='white',icon='home',prefix='fa')).add_to(m)

        for _,row in df_plan.iterrows():
            eq,jor = row.get('equipo',''),row.get('jornada','')
            if jor=='Jornada 1' and not mj1: continue
            if jor=='Jornada 2' and not mj2: continue
            if jor=='Jornada Especial' and not mbm: continue
            clr=color_map.get(eq,'#888')
            d_ini=int(row.get('dia_inicio',row.get('dia_operativo',0)))
            d_fin=int(row.get('dia_fin',d_ini))
            dias_str=f"Día {d_ini}" if d_ini==d_fin else f"Días {d_ini}–{d_fin}"
            folium.CircleMarker(
                location=[row['lat'],row['lon']],radius=5,color=clr,
                fill=True,fill_color=clr,fill_opacity=.85,
                popup=folium.Popup(
                    f"<b>ID:</b> {row['id_entidad']}<br>"
                    f"<b>Viv:</b> {int(row['viv'])}<br>"
                    f"<b>Equipo:</b> {eq}<br><b>Jornada:</b> {jor}<br>"
                    f"<b>Encuestador:</b> {int(row.get('encuestador',0))}<br>"
                    f"<b>{dias_str}</b>",max_width=210),
                tooltip=f"{eq}·{dias_str}·{int(row['viv'])}viv"
            ).add_to(m)

        if mrts:
            for clave,coords in road_p.items():
                eq,jor=clave.split('||')
                if jor=='Jornada 1' and not mj1: continue
                if jor=='Jornada 2' and not mj2: continue
                if len(coords)>1:
                    folium.PolyLine(coords,weight=3,color=color_map.get(eq,'#888'),
                                    opacity=.75,tooltip=f"{eq}|{jor}").add_to(m)

        st_folium(m,width=None,height=540,returned_objects=[],key="mapa_v4")

# ══ TAB 2 — ANÁLISIS ══════════════════════════
with tab_analisis:
    st.markdown("<div class='stitle'>Análisis Estadístico de Carga</div>",unsafe_allow_html=True)

    # ── NUEVO: Panel de rebalanceo ────────────────────────────────────────
    cv_ini = st.session_state.get("cv_ini_bal")
    cv_fin = st.session_state.get("cv_fin_bal")
    bal_log= st.session_state.get("balance_log", [])
    viv_ant = st.session_state.get("viv_por_cluster_antes")
    viv_dep = st.session_state.get("vib_despues") or st.session_state.get("viv_por_cluster_despues")

    if cv_ini is not None and cv_fin is not None:
        mejora = cv_ini - cv_fin
        color_mejora = "#27ae60" if mejora > 5 else ("#f39c12" if mejora > 0 else "#e74c3c")
        st.markdown(f"""
        <div class='balance-box'>
        <b>Rebalanceo de clusters (v4)</b><br>
        CV inicial (solo KMeans): <b style='color:#e74c3c'>{cv_ini:.1f}%</b> &nbsp;→&nbsp;
        CV final (post-balance): <b style='color:#27ae60'>{cv_fin:.1f}%</b>
        &nbsp;&nbsp;<b style='color:{color_mejora}'>Δ {mejora:.1f}% de mejora</b><br>
        <span style='font-size:11px;color:#3a7a55'>
        Iteraciones realizadas: {len([l for l in bal_log if l.get('movimiento','') not in ['inicial','final']])}
        </span>
        </div>""", unsafe_allow_html=True)

    # Gráfico comparativo: viv por cluster antes vs después
    if viv_ant and viv_dep:
        n_cl = len(viv_ant)
        df_comp = pd.DataFrame({
            'Cluster': [f"C{c}" for c in range(n_cl)] * 2,
            'Viviendas (carga pond.)': list(viv_ant.values()) + list(viv_dep.values()),
            'Fase': ['Antes (KMeans puro)']*n_cl + ['Después (rebalanceo)']*n_cl
        })
        fig_comp = px.bar(df_comp, x='Cluster', y='Viviendas (carga pond.)',
                          color='Fase', barmode='group',
                          title='Carga ponderada por cluster — antes vs después del rebalanceo',
                          template='plotly_dark',
                          color_discrete_map={
                              'Antes (KMeans puro)':'#e74c3c',
                              'Después (rebalanceo)':'#27ae60'})
        fig_comp.update_layout(paper_bgcolor="#111827", plot_bgcolor="#0a1020",
                               title_font_size=12)
        st.plotly_chart(fig_comp, use_container_width=True)

        with st.expander("Ver historial de iteraciones del rebalanceo"):
            if bal_log:
                df_log = pd.DataFrame(bal_log)
                st.dataframe(df_log, use_container_width=True, height=250)

    st.divider()

    # ── Métricas de clustering ──────────────────────────────────────────
    sil  = st.session_state.sil_score
    n_bm = st.session_state.n_bombero
    mc1,mc2,mc3 = st.columns(3)
    sv   = f"{sil:.3f}" if sil else "N/A"
    sc_  = "#27ae60" if (sil or 0)>.5 else ("#f39c12" if (sil or 0)>.3 else "#e74c3c")
    with mc1:
        st.markdown(f"<div class='kcard'><div class='v' style='color:{sc_}'>{sv}</div>"
                    f"<div class='l'>Índice Silueta</div>"
                    f"<div class='s'>>0.5 = clusters coherentes</div></div>",
                    unsafe_allow_html=True)
    with mc2:
        cv_fin_disp = f"{cv_fin:.1f}%" if cv_fin is not None else "N/A"
        cc_cv = "#27ae60" if (cv_fin or 100) < 10 else ("#f39c12" if (cv_fin or 100) < 20 else "#e74c3c")
        st.markdown(f"<div class='kcard'><div class='v' style='color:{cc_cv}'>{cv_fin_disp}</div>"
                    f"<div class='l'>CV clusters (post-balance)</div>"
                    f"<div class='s'>&lt;10% = muy equilibrado</div></div>",
                    unsafe_allow_html=True)
    with mc3:
        bc = "#9b59b6" if n_bm>0 else "#445566"
        st.markdown(f"<div class='kcard'><div class='v' style='color:{bc}'>{n_bm}</div>"
                    f"<div class='l'>UPMs Bombero</div>"
                    f"<div class='s'>{'outliers por cluster' if n_bm>0 else 'ninguno'}</div></div>",
                    unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)

    # CV entre equipos
    st.markdown("<div class='stitle'>Equidad entre equipos</div>",unsafe_allow_html=True)
    df_main = df_plan[~df_plan['equipo'].isin(['Equipo Bombero','sin_asignar'])].copy()
    res_cv  = df_main.groupby(['equipo','jornada']).agg(
        viv_reales=('viv','sum'), carga_ponderada=('carga_pond','sum')).reset_index()

    for jornada in ['Jornada 1','Jornada 2']:
        sub = res_cv[res_cv['jornada']==jornada]
        if len(sub)==0: continue
        if len(sub)==1:
            st.markdown(f"<div class='ibox'><b>{jornada}:</b> 1 equipo.</div>",
                        unsafe_allow_html=True); continue
        cr = cv_pct(sub['viv_reales'])
        cp = cv_pct(sub['carga_ponderada'])
        ccr="#27ae60" if cr<20 else ("#f39c12" if cr<40 else "#e74c3c")
        ccp="#27ae60" if cp<20 else ("#f39c12" if cp<40 else "#e74c3c")
        em = "✓" if cp<20 else ("⚠" if cp<40 else "✗")
        st.markdown(f"""<div class='ibox'>
        <b>{jornada}</b><br>
        &nbsp;&nbsp;CV viviendas reales: <span style='color:{ccr};font-family:monospace;
        font-weight:600'>{cr:.1f}%</span><br>
        &nbsp;&nbsp;CV carga ponderada:&nbsp; <span style='color:{ccp};font-family:monospace;
        font-weight:600'>{cp:.1f}%</span>
        &nbsp;{em} <b>{'Muy bueno' if cp<20 else ('Aceptable' if cp<40 else 'Revisar')}</b>
        </div>""", unsafe_allow_html=True)

    with st.expander("Ver tabla de balance"):
        st.dataframe(res_cv.rename(columns={
            'equipo':'Equipo','jornada':'Jornada',
            'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.'
        }),use_container_width=True)

    # Tarjetas de equipos
    st.markdown("<div class='stitle'>Carga por equipo</div>",unsafe_allow_html=True)
    eq_act = [n for n in nombres if n in df_plan['equipo'].values]
    cols_e = st.columns(len(eq_act))
    for col_e,nombre_eq in zip(cols_e,eq_act):
        sub_e  = df_plan[df_plan['equipo']==nombre_eq]
        vt     = int(sub_e['viv'].sum())
        cv_e   = cv_pct(sub_e['carga_pond'])
        ce     = color_map.get(nombre_eq,'#2e86de')
        ccv    = "#27ae60" if cv_e<20 else ("#f39c12" if cv_e<40 else "#e74c3c")
        with col_e:
            st.markdown(f"""<div class='eq-card' style='border-color:{ce}55'>
              <div style='width:10px;height:10px;background:{ce};border-radius:50%;margin:0 auto 7px'></div>
              <div style='font-family:"IBM Plex Mono",monospace;font-size:12px;
                          color:{ce};font-weight:600'>{nombre_eq}</div>
              <div style='font-size:17px;font-weight:600;color:#d0d8e8;margin:4px 0'>{vt:,}</div>
              <div style='font-size:10px;color:#7a8fa6'>viviendas reales</div>
              <div style='font-size:11px;color:{ccv};margin-top:4px'>CV {cv_e:.1f}%</div>
            </div>""",unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)

    # Drilldown encuestadores
    eq_sel = st.selectbox("Detalle de encuestadores:", eq_act)
    df_sel = df_plan[df_plan['equipo']==eq_sel].copy()
    df_enc = df_sel.groupby(['jornada','encuestador']).agg(
        upms=('id_entidad','count'),
        viv_reales=('viv','sum'),
        carga_pond=('carga_pond','sum')).reset_index()

    cd1,cd2 = st.columns(2)
    with cd1:
        fig=px.bar(df_enc,x='encuestador',y='viv_reales',color='jornada',barmode='group',
                   title=f'Viviendas reales — {eq_sel}',
                   labels={'viv_reales':'Viv. reales','encuestador':'Encuestador','jornada':'Jornada'},
                   template='plotly_dark',color_discrete_sequence=['#2e86de','#27ae60'])
        fig.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",title_font_size=12)
        st.plotly_chart(fig,use_container_width=True)
    with cd2:
        fig2=px.bar(df_enc,x='encuestador',y='carga_pond',color='jornada',barmode='group',
                    title=f'Carga ponderada — {eq_sel}',
                    labels={'carga_pond':'Carga pond.','encuestador':'Encuestador','jornada':'Jornada'},
                    template='plotly_dark',color_discrete_sequence=['#e74c3c','#f39c12'])
        fig2.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",title_font_size=12)
        st.plotly_chart(fig2,use_container_width=True)

    # Distribución por días
    st.markdown("<div class='stitle'>Distribución diaria de viviendas</div>",
                unsafe_allow_html=True)
    jor_opts = ["Jornada 1","Jornada 2","Ambas"]
    jor_idx  = st.radio("Filtrar:",jor_opts,horizontal=True,key="radio_jornada_dias_v4",index=0)
    jor_filtro = jor_idx

    pivot_all = df_plan[df_plan['equipo'].isin(eq_act)].copy()
    if jor_filtro != "Ambas":
        pivot_all = pivot_all[pivot_all['jornada']==jor_filtro]

    rows_exp = []
    for _,row in pivot_all.iterrows():
        d_ini    = int(row.get('dia_inicio',row.get('dia_operativo',1)))
        d_fin    = int(row.get('dia_fin',d_ini))
        dias_dur = max(1,d_fin-d_ini+1)
        viv_d    = row['viv']/dias_dur
        for dd in range(d_ini,d_fin+1):
            rows_exp.append({
                'equipo':row['equipo'],'jornada':row['jornada'],
                'encuestador':int(row.get('encuestador',0)),
                'dia_abs':dd,'viv':viv_d
            })

    if len(rows_exp) > 0:
        df_exp = pd.DataFrame(rows_exp)
        if jor_filtro != "Ambas":
            d_min_jor = df_exp['dia_abs'].min()
            df_exp['dia_rel'] = df_exp['dia_abs'] - d_min_jor + 1
        else:
            df_exp['dia_rel'] = df_exp.groupby('jornada')['dia_abs'].transform(
                lambda s: s - s.min() + 1).astype(int)

        pivot   = df_exp.groupby(['equipo','dia_rel'])['viv'].sum().reset_index()
        fig_d   = px.bar(pivot,x='dia_rel',y='viv',color='equipo',barmode='group',
                         title=f'Viviendas por día operativo — {jor_filtro}',
                         labels={'dia_rel':'Día','viv':'Viviendas','equipo':'Equipo'},
                         template='plotly_dark',color_discrete_map=color_map)
        tot_enc_f = sum(e["enc"] for e in eq_cfg if e["nombre"] in eq_act)
        n_eq_act  = max(1,len(eq_act))
        avg_enc_f = tot_enc_f/n_eq_act
        fig_d.add_hline(y=p["viv_min"]*avg_enc_f,line_dash="dot",line_color="#f39c12",
                        annotation_text=f"Mín ({p['viv_min']} viv/enc)")
        fig_d.add_hline(y=p["viv_max"]*avg_enc_f,line_dash="dot",line_color="#e74c3c",
                        annotation_text=f"Máx ({p['viv_max']} viv/enc)")
        fig_d.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",
                            xaxis=dict(dtick=1,title="Día de la jornada"),
                            yaxis_title="Viviendas")
        st.plotly_chart(fig_d,use_container_width=True)

        if jor_filtro != "Ambas":
            pivot_enc = df_exp.groupby(['encuestador','dia_rel'])['viv'].sum().reset_index()
            pivot_enc['encuestador'] = "Enc. "+pivot_enc['encuestador'].astype(str)
            fig_enc = px.line(pivot_enc,x='dia_rel',y='viv',color='encuestador',
                              markers=True,
                              title=f'Carga diaria por encuestador — {jor_filtro}',
                              labels={'dia_rel':'Día','viv':'Viviendas','encuestador':''},
                              template='plotly_dark')
            fig_enc.add_hline(y=p["viv_min"],line_dash="dot",line_color="#f39c12",
                              annotation_text=f"Mín {p['viv_min']}")
            fig_enc.add_hline(y=p["viv_max"],line_dash="dot",line_color="#e74c3c",
                              annotation_text=f"Máx {p['viv_max']}")
            fig_enc.update_layout(paper_bgcolor="#111827",plot_bgcolor="#0a1020",
                                  xaxis=dict(dtick=1))
            st.plotly_chart(fig_enc,use_container_width=True)

    # Equipo Bombero
    df_bm = df_plan[df_plan['equipo']=='Equipo Bombero']
    n_bm  = st.session_state.n_bombero
    st.markdown("<div class='stitle'>Equipo Bombero</div>",unsafe_allow_html=True)
    if n_bm==0:
        st.markdown("""<div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — 0 UPMs asignadas.<br>
        <span style='font-size:12px;color:#7a5a9a'>
        Ningún punto resultó outlier dentro de su cluster en este mes.
        </span></div>""",unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — {n_bm} UPMs ·
        {int(df_bm['viv'].sum()):,} viviendas
        </div>""",unsafe_allow_html=True)
        st.dataframe(
            df_bm[['id_entidad','tipo_entidad','viv','lat','lon','dist_base_m']]
            .rename(columns={'dist_base_m':'Dist. base (m)'})
            .sort_values('Dist. base (m)',ascending=False).reset_index(drop=True),
            use_container_width=True,height=200)

# ══ TAB 3 — REPORTE Y DESCARGA ════════════════
with tab_reporte:
    st.markdown("<div class='stitle'>Reporte Mensual y Descarga Excel</div>",
                unsafe_allow_html=True)

    fc1,fc2 = st.columns(2)
    with fc1:
        fj1 = st.date_input("Fecha inicio Jornada 1",
                             value=st.session_state.fecha_j1 or date.today(),key="fi_j1")
        st.session_state.fecha_j1 = fj1
        if fj1:
            st.caption(f"Fin: {(fj1+timedelta(days=p['dias_op']-1)).strftime('%d/%m/%Y')}")
    with fc2:
        fj2 = st.date_input("Fecha inicio Jornada 2",
                             value=st.session_state.fecha_j2 or date.today(),key="fi_j2")
        st.session_state.fecha_j2 = fj2
        if fj2:
            st.caption(f"Fin: {(fj2+timedelta(days=p['dias_op']-1)).strftime('%d/%m/%Y')}")

    st.markdown("<div class='stitle'>Personal por equipo</div>",unsafe_allow_html=True)
    for eq in eq_cfg:
        nombre_eq = eq["nombre"]
        n_enc_eq  = eq["enc"]
        pi_prev   = st.session_state.personal_info.get(nombre_eq, {})
        with st.expander(f"👥 {nombre_eq} — {n_enc_eq} encuestador(es)", expanded=False):
            pc1,pc2,pc3 = st.columns(3)
            with pc1: sup_n = st.text_input("Supervisor (nombre)",value=pi_prev.get('supervisor_nombre',''),key=f"sup_n_{nombre_eq}")
            with pc2: sup_c = st.text_input("Cédula supervisor",value=pi_prev.get('supervisor_cedula',''),key=f"sup_c_{nombre_eq}")
            with pc3: sup_t = st.text_input("Celular supervisor",value=pi_prev.get('supervisor_celular',''),key=f"sup_t_{nombre_eq}")
            enc_list_new = []
            for j in range(n_enc_eq):
                prev_enc = pi_prev.get('encuestadores',[{}]*n_enc_eq)
                prev_j   = prev_enc[j] if j < len(prev_enc) else {}
                pe1,pe2,pe3 = st.columns(3)
                with pe1: en = st.text_input(f"Encuestador {j+1}",value=prev_j.get('nombre',''),key=f"enc_n_{nombre_eq}_{j}")
                with pe2: ec = st.text_input(f"Cédula enc. {j+1}",value=prev_j.get('cedula',''),key=f"enc_c_{nombre_eq}_{j}")
                with pe3: et = st.text_input(f"Celular enc. {j+1}",value=prev_j.get('celular',''),key=f"enc_t_{nombre_eq}_{j}")
                enc_list_new.append({'nombre':en,'cedula':ec,'celular':et,'cod':''})
            pch1,pch2,pch3 = st.columns(3)
            with pch1: ch_n = st.text_input("Chofer (nombre)",value=pi_prev.get('chofer_nombre',''),key=f"ch_n_{nombre_eq}")
            with pch2: plca = st.text_input("Placa",value=pi_prev.get('placa',''),key=f"plca_{nombre_eq}")
            with pch3: ch_t = st.text_input("Celular chofer",value=pi_prev.get('chofer_celular',''),key=f"ch_t_{nombre_eq}")
            st.session_state.personal_info[nombre_eq] = {
                'supervisor_nombre':sup_n,'supervisor_cedula':sup_c,
                'supervisor_celular':sup_t,'supervisor_cod':'',
                'encuestadores':enc_list_new,'n_enc':n_enc_eq,
                'chofer_nombre':ch_n,'placa':plca,'chofer_celular':ch_t,
            }

    st.markdown("<div class='stitle'>Resumen de planificación</div>",unsafe_allow_html=True)
    if res_bal is not None and len(res_bal)>0:
        tr = pd.DataFrame([{
            'equipo':'TOTAL','jornada':'—',
            'n_upms':res_bal['n_upms'].sum(),
            'viv_reales':res_bal['viv_reales'].sum(),
            'carga_ponderada':res_bal['carga_ponderada'].sum(),
            'dist_km':res_bal.get('dist_km',pd.Series([0])).sum()
        }])
        rep = pd.concat([res_bal,tr],ignore_index=True)
        st.dataframe(rep.rename(columns={
            'equipo':'Equipo','jornada':'Jornada','n_upms':'UPMs',
            'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.',
            'dist_km':'Dist. (km)'}),use_container_width=True)

    cols_ok=[c for c in ['id_entidad','upm','tipo_entidad','viv','carga_pond',
                          'equipo','jornada','encuestador','dia_inicio','dia_fin','lat','lon']
             if c in df_plan.columns]
    st.dataframe(
        df_plan[cols_ok].sort_values(
            ['equipo','jornada','encuestador','dia_inicio']).reset_index(drop=True),
        use_container_width=True, height=300)

    st.markdown("<div class='stitle'>Descargar Excel formateado</div>",unsafe_allow_html=True)
    if st.button("📋 Generar Excel",use_container_width=True,type="primary"):
        with st.spinner("Generando Excel..."):
            try:
                excel_bytes = generar_excel(
                    df_plan=df_plan, eq_cfg=eq_cfg,
                    personal_info=st.session_state.personal_info,
                    fecha_j1=st.session_state.fecha_j1,
                    fecha_j2=st.session_state.fecha_j2,
                    dias_op=p["dias_op"],
                    j1_num=st.session_state.get("j1_num",1),
                    j2_num=st.session_state.get("j2_num",2),
                    mes_nombre=MESES_N.get(int(df['mes'].iloc[0]),'')
                )
                j1n  = st.session_state.get("j1_num",1)
                j2n  = st.session_state.get("j2_num",2)
                mes_n= MESES_N.get(int(df['mes'].iloc[0]),'mes')
                st.download_button(
                    label=f"⬇️ Descargar J{j1n}+J{j2n}_{mes_n}.xlsx",
                    data=excel_bytes,
                    file_name=f"planificacion_J{j1n}-J{j2n}_{mes_n}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("✓ Excel listo para descargar.")
            except Exception as e:
                st.error(f"Error generando Excel: {e}")
                import traceback
                st.code(traceback.format_exc())
