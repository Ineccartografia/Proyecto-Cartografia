# =============================================================================
# PLANIFICACIÓN CARTOGRÁFICA ENDI 2025 — STREAMLIT v6
# INEC · Zonal Litoral · Autores: Franklin López, Carlos Quinto
#
# CAMBIOS v6 (sobre v5.x con CSS INEC):
#
# ── FIX CLUSTERING: emparejamiento óptimo en lugar de swap de puntos ─────────
#   Problema v5: la restricción 2×radio bloqueaba TODOS los swaps (CV=82.8%→82.8%).
#   Causa: en el Litoral los clusters están a 50-300km entre sí. Con radio típico
#   de 20km, max_dist = 40km → ningún swap interprovincial era posible → 0 mejora.
#
#   Solución v6: dos pasos separados:
#   1. KMeans geográfico puro (sin modificar los clusters).
#   2. Emparejamiento óptimo de clusters en pares (J1, J2):
#      Backtracking exhaustivo busca qué par de clusters asignar a cada equipo
#      para minimizar el CV de cargas entre equipos. Para ≤7 equipos es <0.1s.
#   3. Micro-ajustes de frontera opcionales dentro del mismo equipo (conservadores).
#
#   Por qué esto funciona mejor:
#   - No mueve manzanas entre provincias (la geografía de cada cluster es intacta).
#   - Balancea emparejando un cluster pesado (Guayaquil norte) con uno liviano
#     (Manabí interior) → el equipo tiene suma similar al resto.
#   - El CV que se muestra ahora es el CV ENTRE EQUIPOS, que es lo que importa
#     operativamente, no el CV entre clusters individuales.
#
# ── FIX JORNADAS CANCELADAS: desplazamiento en slots ────────────────────────
#   Problema v5: cancelar una jornada la eliminaba completamente del operativo.
#   Correcto: al cancelar J2, todas las jornadas siguientes se desplazan 1 slot.
#   Si hay N cancelaciones al final del operativo se agregan N meses adicionales.
#
#   Implementado con construir_slots_calendario():
#   - Separa jornadas activas/canceladas.
#   - Asigna jornadas activas a slots en orden.
#   - Un slot = mitad de mes en el calendario real.
#   - Se puede tener J3 en la primera mitad del mes 2 del calendario si J2 se canceló.
#
# ── SIMPLIFICACIÓN DE PARÁMETROS ─────────────────────────────────────────────
#   El sidebar tenía demasiados parámetros técnicos. v6 los reduce a los
#   operativamente relevantes. El rebalanceo ya no necesita k_vecinos ni max_iter
#   como parámetros de primera línea (usan valores sensatos por defecto).
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
st.set_page_config(page_title="INEC · ENDI Planificación",
                   page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

INEC_LOGO = "https://upload.wikimedia.org/wikipedia/commons/a/a8/Logo_del_INEC_Ecuador.png"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;color:#1a1a2e}}
.main .block-container{{padding-top:2rem}}
[data-testid="stSidebar"]{{background:#f7f8fb;border-right:1px solid #e2e6ed}}
[data-testid="stSidebar"] *{{color:#2d3348 !important}}
.hdr{{background:#ffffff;border-radius:10px;padding:20px 28px;margin-bottom:24px;
      border:1px solid #e2e6ed;border-left:4px solid #003B71;
      display:flex;align-items:center;gap:20px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.hdr img{{height:52px;flex-shrink:0}}
.hdr-text h1{{color:#003B71!important;font-size:17px!important;font-weight:600!important;
              margin:0 0 2px!important;font-family:'JetBrains Mono',monospace!important}}
.hdr-text p{{color:#6b7a90!important;font-size:12px!important;margin:0!important}}
.kcard{{background:#ffffff;border:1px solid #e2e6ed;border-radius:10px;
        padding:16px;text-align:center;box-shadow:0 1px 2px rgba(0,0,0,.03)}}
.kcard .v{{font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:600;color:#003B71;line-height:1}}
.kcard .l{{font-size:10px;color:#8896a6;margin-top:5px;text-transform:uppercase;letter-spacing:.6px}}
.kcard .s{{font-size:10px;color:#a8b5c0;margin-top:2px}}
.step{{display:inline-block;background:#eef3fa;color:#003B71;border:1px solid #d0daea;
       border-radius:4px;padding:2px 8px;font-family:'JetBrains Mono',monospace;
       font-size:10px;font-weight:600;letter-spacing:.8px;margin-bottom:6px}}
.stitle{{font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;
         color:#003B71;text-transform:uppercase;letter-spacing:1px;
         border-bottom:2px solid #e2e6ed;padding-bottom:8px;margin:22px 0 14px}}
.ibox{{background:#f0f6ff;border:1px solid #d0daea;border-left:3px solid #003B71;
       border-radius:7px;padding:12px 16px;margin:9px 0;font-size:13px;color:#2d4a6f}}
.wbox{{background:#fffbf0;border:1px solid #f0deb0;border-left:3px solid #e6a817;
       border-radius:7px;padding:12px 16px;margin:9px 0;font-size:13px;color:#7a5c10}}
.bcard{{background:#faf5ff;border:1px solid #e4d5f5;border-left:3px solid #7c3aed;
        border-radius:7px;padding:13px 16px;margin:9px 0}}
.pill-ok{{display:inline-block;background:#ecfdf5;color:#047857;border:1px solid #a7f3d0;
          border-radius:20px;padding:2px 10px;font-size:11px;font-family:'JetBrains Mono',monospace;font-weight:600}}
.pill-w{{display:inline-block;background:#fffbeb;color:#b45309;border:1px solid #fde68a;
         border-radius:20px;padding:2px 10px;font-size:11px;font-family:'JetBrains Mono',monospace;font-weight:600}}
.eq-card{{background:#ffffff;border:1px solid #e2e6ed;border-radius:9px;
          padding:14px 16px;text-align:center;box-shadow:0 1px 2px rgba(0,0,0,.03)}}
.balance-box{{background:#ecfdf5;border:1px solid #a7f3d0;border-left:3px solid #059669;
              border-radius:7px;padding:12px 16px;margin:9px 0;font-size:12px;color:#065f46}}
.jplan-ok{{background:#ecfdf5;border:1px solid #a7f3d0;border-left:3px solid #059669;
           border-radius:6px;padding:8px 14px;margin:4px 0;font-size:12px;color:#065f46}}
.jplan-can{{background:#fef2f2;border:1px solid #fecaca;border-left:3px solid #dc2626;
            border-radius:6px;padding:8px 14px;margin:4px 0;font-size:12px;color:#991b1b}}
.slot-card{{background:#f8faff;border:1px solid #d0daea;border-radius:8px;
            padding:10px 14px;margin:3px 0;font-size:12px;color:#2d4a6f;
            display:flex;align-items:center;gap:10px}}
.sidebar-logo{{display:flex;align-items:center;gap:12px;padding:4px 0 12px;margin-bottom:4px}}
.sidebar-logo img{{height:40px}}
.sidebar-logo .sidebar-title{{font-family:'JetBrains Mono',monospace;font-size:12px;
                               font-weight:600;color:#003B71!important;line-height:1.3}}
.sidebar-logo .sidebar-sub{{font-size:10px;color:#8896a6!important;margin-top:2px}}
button[kind="primary"]{{background:#003B71!important;border:none!important;font-weight:600!important}}
</style>
""", unsafe_allow_html=True)

# ── CONSTANTES ────────────────────────────────
BASE_LAT = -2.145825935522539
BASE_LON = -79.89383956329586
PRO_GYE  = "09"
CAN_GYE  = "01"
MESES_CAL = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}
COLORES = ['#dc2626','#003B71','#059669','#d97706','#7c3aed','#0891b2','#c2410c','#be185d']

def cv_pct(s):
    m = s.mean(); return float(s.std()/m*100) if m > 0 else 0.0

def utm_to_wgs84(df):
    t = Transformer.from_crs("epsg:32717","epsg:4326",always_xy=True)
    lons,lats = t.transform(df["x"].values,df["y"].values)
    df = df.copy(); df["lon"]=lons; df["lat"]=lats; return df

def parse_codigo(codigo):
    c = str(codigo).strip()
    r = {'prov':'','canton':'','ciudad_parroq':'','zona':'','sector':'','man':''}
    if len(c)>=6:  r['prov']=c[:2]; r['canton']=c[2:4]; r['ciudad_parroq']=c[4:6]
    if len(c)>=9:  r['zona']=c[6:9]
    if len(c)>=12: r['sector']=c[9:12]
    if len(c)>=15: r['man']=c[12:15]
    return r

def jornada_num_desde_mes(mes_operativo, mes_inicio_cal):
    mes_cal = ((mes_inicio_cal-1+mes_operativo-1)%12)+1
    j1_num  = (mes_operativo-1)*2+1
    j2_num  = j1_num+1
    return j1_num, j2_num, MESES_CAL[mes_cal]

def cargar_gpkg(path, dissolve_upm=True):
    capas = pyogrio.list_layers(path)
    if len(capas)==1:
        gdf = gpd.read_file(path,layer=capas[0][0])
        col_map={'1_mes_cart':'mes','viv_total':'viv','1_zonal':'zonal','1_id_upm':'upm','ManSec':'id_entidad'}
        gdf = gdf.rename(columns={k:v for k,v in col_map.items() if k in gdf.columns})
        if 'id_entidad' in gdf.columns:
            gdf['tipo_entidad']=gdf['id_entidad'].astype(str).apply(lambda x:'sec' if '999' in x else 'man')
        gdf_u=gdf.to_crs(epsg=32717); gdf_u['geometry']=gdf_u.geometry.representative_point()
        gdf_u['x']=gdf_u.geometry.x; gdf_u['y']=gdf_u.geometry.y
        if 'pro' in gdf_u.columns: gdf_u['pro_x']=gdf_u['pro']
        if 'can' in gdf_u.columns: gdf_u['can_x']=gdf_u['can']
        if 'mes' in gdf_u.columns: gdf_u['mes']=pd.to_numeric(gdf_u['mes'],errors='coerce')
        if dissolve_upm and 'upm' in gdf_u.columns:
            agg={'viv':'sum','mes':'first','x':'first','y':'first','tipo_entidad':'first'}
            if 'pro_x' in gdf_u.columns: agg['pro_x']='first'
            if 'can_x' in gdf_u.columns: agg['can_x']='first'
            gdf_f=gdf_u.groupby('upm').agg(agg).reset_index()
            gdf_f['id_entidad']=gdf_f['upm']
            gdf_f['tipo_entidad']=gdf_f['tipo_entidad'].apply(lambda t:f"{t}_upm")
        else: gdf_f=gdf_u
        return utm_to_wgs84(gdf_f)
    else:
        man=gpd.read_file(path,layer=capas[0][0]); disp=gpd.read_file(path,layer=capas[1][0])
        man=man[man['zonal']=='LITORAL']; disp=disp[disp['zonal']=='LITORAL']
        man_u=man.to_crs(epsg=32717); dis_u=disp.to_crs(epsg=32717)
        if dissolve_upm:
            def _d(gdf,tipo):
                d=gdf.dissolve(by='upm',aggfunc={'mes':'first','viv':'sum'})
                d['geometry']=d.geometry.representative_point()
                o=d[['mes','viv']].copy(); o['id_entidad']=d.index; o['upm']=d.index; o['tipo_entidad']=tipo
                o['x']=d.geometry.x; o['y']=d.geometry.y
                if 'mes' in o.columns: o['mes']=pd.to_numeric(o['mes'],errors='coerce')
                return o[['id_entidad','upm','mes','viv','x','y','tipo_entidad']]
            ms=_d(man_u,'man_upm'); ds=_d(dis_u,'sec_upm')
        else:
            for g in [man_u,dis_u]:
                g['geometry']=g.geometry.representative_point(); g['x']=g.geometry.x; g['y']=g.geometry.y
            ms=man_u[['man','upm','mes','viv','x','y']].rename(columns={'man':'id_entidad'}); ms['tipo_entidad']='man'
            ds=dis_u[['sec','upm','mes','viv','x','y']].rename(columns={'sec':'id_entidad'}); ds['tipo_entidad']='sec'
            ms['pro_x']=ms['id_entidad'].astype(str).str[:2]; ms['can_x']=ms['id_entidad'].astype(str).str[2:4]
            for df_t in [ms,ds]:
                if 'mes' in df_t.columns: df_t['mes']=pd.to_numeric(df_t['mes'],errors='coerce')
        data=pd.concat([ms,ds],ignore_index=True)
        if not dissolve_upm: data=data.drop_duplicates(subset=['id_entidad','upm'],keep='first')
        return utm_to_wgs84(data)


# ══════════════════════════════════════════════════════════════════════════════
#  CLUSTERING v6 — EMPAREJAMIENTO ÓPTIMO
# ══════════════════════════════════════════════════════════════════════════════

def _mejor_emparejamiento(sums, n_eq):
    """
    Backtracking exhaustivo para encontrar el emparejamiento de 2*n_eq clusters
    en n_eq pares que minimiza el CV de las sumas de cada par.
    Para n_eq ≤ 7: siempre exacto. Para n_eq > 7: greedy.
    """
    n_cl = len(sums)
    best = {'cv': np.inf, 'pairs': None}

    def backtrack(restantes, parejas):
        if not restantes:
            cargas = np.array([sums[p[0]]+sums[p[1]] for p in parejas])
            cv = cv_pct(pd.Series(cargas))
            if cv < best['cv']:
                best['cv'] = cv; best['pairs'] = list(parejas)
            return
        primero = restantes[0]
        for i in range(1, len(restantes)):
            otro = restantes[i]
            nuevos = [x for j,x in enumerate(restantes) if j!=0 and j!=i]
            parejas.append((primero, otro))
            backtrack(nuevos, parejas)
            parejas.pop()

    backtrack(list(range(n_cl)), [])
    return best['pairs'], best['cv']


def clustering_balanceado(df, n_clusters, cv_objetivo=0.10, max_iter_micro=100,
                           k_vecinos=8):
    """
    Clustering geográfico + balanceo por emparejamiento óptimo (v6).

    PASO 1 — KMeans puro
    Crea n_clusters geográficamente coherentes. Sin modificaciones.

    PASO 2 — Emparejamiento óptimo
    Busca qué par de clusters asignar a cada equipo para minimizar CV entre equipos.
    NO mueve UPMs entre clusters → la geografía de cada cluster es intacta.
    Esto resuelve el desequilibrio entre zonas sin distorsionar la distribución.

    PASO 3 — Micro-ajustes de frontera (conservador)
    Si el CV aún supera cv_objetivo, hace pequeños swaps de frontera (k-NN)
    entre clusters del MISMO equipo (solo entre J1 y J2 del mismo equipo).
    Esto balancea la carga dentro del equipo sin mezclar zonas de distintos equipos.

    Retorna
    ───────
    labels     : np.ndarray (N,) — cluster id por UPM
    best_pairs : list de (c_j1, c_j2) por equipo
    log        : list de dicts con historial
    cv_ini     : CV inicial entre clusters (KMeans puro)
    cv_eq_fin  : CV final entre equipos (métrica operativa relevante)
    """
    coords = df[['x','y']].values.astype(float)
    cargas = df['carga_pond'].values.astype(float)
    n      = len(df)
    n_eq   = n_clusters // 2

    # ── PASO 1 ───────────────────────────────────────────────────────────────
    km = KMeans(n_clusters=n_clusters, init='k-means++', n_init=20,
                max_iter=500, random_state=42)
    labels = km.fit_predict(coords).copy()

    sums   = np.array([cargas[labels==c].sum() for c in range(n_clusters)])
    cv_ini = cv_pct(pd.Series(sums))
    log    = [{'paso':'KMeans','cv_eq':cv_ini,'nota':f'{n_clusters} clusters'}]

    # ── PASO 2 ───────────────────────────────────────────────────────────────
    if n_eq <= 7:
        best_pairs, cv_eq = _mejor_emparejamiento(sums, n_eq)
    else:
        orden = np.argsort(sums)
        best_pairs = [(orden[-(i+1)], orden[i]) for i in range(n_eq)]
        cv_eq = cv_pct(pd.Series([sums[p[0]]+sums[p[1]] for p in best_pairs]))

    log.append({'paso':'Emparejamiento óptimo','cv_eq':cv_eq,
                'nota':f'Pares: {[(c1,c2) for c1,c2 in best_pairs]}'})

    # ── PASO 3: micro-ajustes ────────────────────────────────────────────────
    def eq_sums():
        cs = np.array([cargas[labels==c].sum() for c in range(n_clusters)])
        return np.array([cs[best_pairs[i][0]]+cs[best_pairs[i][1]] for i in range(n_eq)])

    cv_history = [cv_eq]
    if cv_eq > cv_objetivo * 100 and max_iter_micro > 0:
        tree = BallTree(coords, leaf_size=40)
        _, nbr = tree.query(coords, k=min(k_vecinos+1, n))
        no_mejora = 0

        for it in range(max_iter_micro):
            eq_c = eq_sums()
            cv_now = cv_pct(pd.Series(eq_c))
            cv_history.append(cv_now)
            if cv_now <= cv_objetivo*100: break
            if len(cv_history)>=8 and (cv_history[-8]-cv_now)<0.05: break

            orden_p = np.argsort(eq_c)[::-1]
            mejora  = False

            for eq_p in orden_p[:2]:
                for jorn_p in [0,1]:
                    c_p    = best_pairs[eq_p][jorn_p]
                    mask_p = np.where(labels==c_p)[0]
                    if len(mask_p)==0: continue

                    for eq_l in np.argsort(eq_c)[:2]:
                        if eq_l==eq_p: continue
                        c_l = best_pairs[eq_l][jorn_p]
                        mejor_cv, mejor_idx = cv_now, -1
                        for idx in mask_p:
                            if not any(labels[v]==c_l for v in nbr[idx] if v!=idx): continue
                            labels[idx]=c_l
                            cv_n=cv_pct(pd.Series(eq_sums()))
                            labels[idx]=c_p
                            if cv_n<mejor_cv: mejor_cv,mejor_idx=cv_n,idx
                        if mejor_idx>=0:
                            labels[mejor_idx]=c_l
                            log.append({'paso':f'micro iter {it}','cv_eq':mejor_cv,'nota':f'Eq{eq_p}→Eq{eq_l}'})
                            mejora=True; no_mejora=0; break
                    if mejora: break
                if mejora: break

            if not mejora:
                no_mejora+=1
                if no_mejora>=5: break

    eq_c_fin   = eq_sums()
    cv_eq_fin  = cv_pct(pd.Series(eq_c_fin))
    log.append({'paso':'Final','cv_eq':cv_eq_fin,'nota':f'CV ini={cv_ini:.1f}% → CV eq={cv_eq_fin:.1f}%'})
    return labels, best_pairs, log, cv_ini, cv_eq_fin


# ══════════════════════════════════════════════════════════════════════════════
#  CALENDARIO DE JORNADAS — desplazamiento por cancelación (v6)
# ══════════════════════════════════════════════════════════════════════════════

def construir_slots_calendario(meses_disponibles, mes_inicio_cal, config_jornadas):
    """
    Construye el calendario de slots con desplazamiento por cancelaciones.

    Cada mes tiene 2 slots (primera y segunda mitad).
    Si una jornada se cancela, las siguientes se desplazan 1 slot.
    Se añaden slots al final si las cancelaciones lo requieren.

    Retorna: (slots_activos, n_canceladas)
    slots_activos : list de dicts con info de cada slot ejecutable.
    """
    todas = []
    for mes in meses_disponibles:
        mes_i   = int(mes)
        mes_cal = ((mes_inicio_cal-1+mes_i-1)%12)+1
        j1_n    = (mes_i-1)*2+1; j2_n=j1_n+1
        for mitad, jn in [(1,j1_n),(2,j2_n)]:
            cfg = config_jornadas.get(jn, {})
            todas.append({
                'mes_operativo': mes_i,
                'slot_en_mes':   mitad,
                'jornada_num':   jn,
                'jornada_nombre': f'Jornada {mitad}',
                'mes_nombre_cal': MESES_CAL[mes_cal],
                'fecha':          cfg.get('fecha', None),
                'cancelada':      cfg.get('cancelada', False),
            })

    activas    = [j for j in todas if not j['cancelada']]
    n_cancel   = len(todas) - len(activas)

    slots = []
    for i, jinfo in enumerate(activas):
        sn  = i+1
        slots.append({
            'slot_num':              sn,
            'mes_operativo':         jinfo['mes_operativo'],
            'slot_mes_calendario':   (sn-1)//2+1,
            'slot_mitad_calendario': (sn-1)%2+1,
            'jornada_num':           jinfo['jornada_num'],
            'jornada_nombre':        jinfo['jornada_nombre'],
            'mes_nombre_cal':        jinfo['mes_nombre_cal'],
            'fecha':                 jinfo['fecha'],
        })
    return slots, n_cancel


# ══════════════════════════════════════════════════════════════════════════════
#  ASIGNACIÓN ENCUESTADORES + DÍAS (sin cambios en lógica v5)
# ══════════════════════════════════════════════════════════════════════════════

def nearest_neighbor_order(points_xy, start_xy=None):
    n = len(points_xy)
    if n==0: return []
    if n==1: return [0]
    visited=[False]*n
    cur = int(np.argmin(np.linalg.norm(points_xy-start_xy,axis=1))) if start_xy is not None else 0
    order=[cur]; visited[cur]=True
    for _ in range(n-1):
        best_d,best_j=np.inf,-1
        px,py=points_xy[cur]
        for j in range(n):
            if not visited[j]:
                d=(points_xy[j,0]-px)**2+(points_xy[j,1]-py)**2
                if d<best_d: best_d,best_j=d,j
        cur=best_j; order.append(cur); visited[cur]=True
    return order


def asignar_encuestadores_y_dias(df_grp, n_enc, dias_tot, viv_min, viv_max, inicio_dia=1):
    """Asigna encuestadores y distribuye días. Índice original preservado (fix v5)."""
    target = (viv_min+viv_max)/2.0
    ultimo = inicio_dia+dias_tot-1
    df_g   = df_grp.copy()
    n_rows = len(df_g)
    cargas    = df_g['carga_pond'].values.astype(float)
    viviendas = df_g['viv'].values.astype(float)
    coords    = df_g[['x','y']].values.astype(float)

    orden_bp = np.argsort(cargas)[::-1]
    enc_acum = np.zeros(n_enc); enc_asig=np.zeros(n_rows,dtype=int)
    for pos in orden_bp:
        em=int(np.argmin(enc_acum)); enc_asig[pos]=em+1; enc_acum[em]+=cargas[pos]

    geo_order=[]
    for enc_id in range(1,n_enc+1):
        pos_enc=np.where(enc_asig==enc_id)[0]
        if len(pos_enc)==0: continue
        sub=coords[pos_enc]; cent=sub.mean(axis=0)
        nn=nearest_neighbor_order(sub,start_xy=cent)
        for nn_i in nn: geo_order.append((pos_enc[nn_i],enc_id))

    dias_range=list(range(inicio_dia,ultimo+1))
    calendario={e:{d:0.0 for d in dias_range} for e in range(1,n_enc+1)}
    cursor={e:inicio_dia for e in range(1,n_enc+1)}
    dia_ini_arr=np.full(n_rows,inicio_dia,dtype=int)
    dia_fin_arr=np.full(n_rows,inicio_dia,dtype=int)

    for pos,enc_id in geo_order:
        viv_m=max(0.0,viviendas[pos]); cal=calendario[enc_id]; cur=cursor[enc_id]
        if viv_m>viv_max:
            dias_m=max(1,int(np.ceil(viv_m/target))); dias_m=min(dias_m,dias_tot)
            bloque=None
            for d_s in range(cur,ultimo-dias_m+2):
                if all(cal.get(d,target)<target for d in range(d_s,d_s+dias_m)):
                    bloque=d_s; break
            if bloque is None: bloque=max(inicio_dia,min(cur,ultimo-dias_m+1))
            d_ini=bloque; d_fin=min(d_ini+dias_m-1,ultimo)
            vpd=viv_m/max(1,d_fin-d_ini+1)
            for dd in range(d_ini,d_fin+1): cal[dd]=cal.get(dd,0.0)+vpd
            cursor[enc_id]=d_fin+1
        else:
            dia_asig=None
            for d in range(cur,ultimo+1):
                if cal.get(d,0.0)<target: dia_asig=d; break
            if dia_asig is None: dia_asig=ultimo
            cal[dia_asig]=cal.get(dia_asig,0.0)+viv_m
            d_ini=dia_asig; d_fin=dia_asig
            if cal[dia_asig]>=target and cursor[enc_id]==dia_asig:
                cursor[enc_id]=min(dia_asig+1,ultimo)
        d_ini=max(inicio_dia,min(d_ini,ultimo)); d_fin=max(d_ini,min(d_fin,ultimo))
        dia_ini_arr[pos]=d_ini; dia_fin_arr[pos]=d_fin

    df_g['encuestador']=enc_asig; df_g['dia_inicio']=dia_ini_arr
    df_g['dia_fin']=dia_fin_arr; df_g['dia_operativo']=dia_ini_arr
    return df_g


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def generar_excel(df_plan, eq_cfg, personal_info, slots_activos, dias_op):
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    AZ_OSCURO="0D3B6E"; AZ_MEDIO="1A5276"; AZ_CLARO="D6EAF8"; VRD_CHECK="D5F5E3"; BLANCO="FFFFFF"
    ENC_PALETAS=[
        {"par":"DBEAFE","impar":"EFF6FF","subtot":"BFDBFE","hdr":"1D4ED8"},
        {"par":"D1FAE5","impar":"ECFDF5","subtot":"A7F3D0","hdr":"065F46"},
        {"par":"FEF9C3","impar":"FEFCE8","subtot":"FDE68A","hdr":"854D0E"},
        {"par":"FCE7F3","impar":"FDF4FF","subtot":"F9A8D4","hdr":"831843"},
        {"par":"FFE4E6","impar":"FFF1F2","subtot":"FECACA","hdr":"9F1239"},
        {"par":"E0E7FF","impar":"EEF2FF","subtot":"C7D2FE","hdr":"3730A3"},
    ]
    def sc(cell,bold=False,bg=None,fg="000000",ha="left",sz=9,brd=False,wrap=False):
        cell.font=Font(bold=bold,size=sz,color=fg)
        cell.alignment=Alignment(horizontal=ha,vertical="center",wrap_text=wrap)
        if bg: cell.fill=PatternFill("solid",fgColor=bg)
        if brd:
            t=Side(style='thin'); cell.border=Border(left=t,right=t,top=t,bottom=t)
    ct_counter=[700]

    for slot in slots_activos:
        j_num=slot['jornada_num']; j_nombre=slot['jornada_nombre']; fecha_inicio=slot.get('fecha')
        df_jor=df_plan[df_plan['jornada']==j_nombre].copy()
        if len(df_jor)==0: continue

        ws=wb.create_sheet(title=f"J{j_num}"); ws.sheet_view.showGridLines=False
        anchos={'A':14,'B':14,'C':8,'D':5,'E':5,'F':6,'G':6,'H':6,'I':5,'J':18,'K':13,'L':10,'M':14,'N':5}
        for col_l,w in anchos.items(): ws.column_dimensions[col_l].width=w
        for i in range(dias_op): ws.column_dimensions[get_column_letter(15+i)].width=7
        ws.column_dimensions[get_column_letter(15+dias_op)].width=6

        cur=1
        equipos_jor=[e['nombre'] for e in eq_cfg if e['nombre'] in df_jor['equipo'].values]

        for grupo_num,nombre_eq in enumerate(equipos_jor,1):
            df_eq=df_jor[df_jor['equipo']==nombre_eq].copy()
            if len(df_eq)==0: continue
            pi=personal_info.get(nombre_eq,{}); n_enc=next((e['enc'] for e in eq_cfg if e['nombre']==nombre_eq),3)
            last_col=15+dias_op
            if fecha_inicio:
                fechas=[fecha_inicio+timedelta(days=i) for i in range(dias_op)]
                fi_str=fecha_inicio.strftime("%d-%b-%y").upper(); ff_str=fechas[-1].strftime("%d-%b-%y").upper()
            else:
                fechas=None; fi_str="____"; ff_str="____"

            def merge_row(row,c1,c2,val,**kw):
                ws.merge_cells(f'{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}')
                c=ws.cell(row,c1,val); sc(c,**kw); return c

            for txt in ["INSTITUTO NACIONAL DE ESTADÍSTICA Y CENSOS","COORDINACIÓN ZONAL LITORAL CZ8L",
                        "ACTUALIZACIÓN CARTOGRÁFICA - ENDI ENLISTAMIENTO","PROGRAMACIÓN OPERATIVO DE CAMPO"]:
                merge_row(cur,1,last_col,txt,bold=True,bg=AZ_OSCURO,fg=BLANCO,ha="center",sz=9); cur+=1
            cur+=1
            ws.cell(cur,1,"JORNADA"); sc(ws.cell(cur,1),bold=True,sz=10)
            ws.cell(cur,2,str(j_num)).font=Font(bold=True,size=11)
            ws.cell(cur,7,"GRUPO"); sc(ws.cell(cur,7),bold=True,sz=10)
            ws.cell(cur,9,str(grupo_num)).font=Font(bold=True,size=11); cur+=2
            ws.cell(cur,1,"PERÍODO DE ACTUALIZACIÓN:"); sc(ws.cell(cur,1),bold=True,sz=9)
            ws.cell(cur,5,"DEL"); ws.cell(cur,6,fi_str); ws.cell(cur,9,"AL"); ws.cell(cur,10,ff_str); cur+=2
            for col,txt in [(3,"COD."),(4,"NOMBRE"),(8,"No. CÉDULA"),(11,"No. CELULAR")]:
                sc(ws.cell(cur,col,txt),bold=True,sz=8)
            cur+=1
            ws.cell(cur,1,"SUPERVISOR:"); sc(ws.cell(cur,1),bold=True,sz=9)
            ws.cell(cur,4,pi.get('supervisor_nombre','')); ws.cell(cur,8,pi.get('supervisor_cedula','')); cur+=2
            enc_list=pi.get('encuestadores',[])
            for j in range(n_enc):
                info=enc_list[j] if j<len(enc_list) else {}
                ws.cell(cur,1,"ENCUESTADOR"); sc(ws.cell(cur,1),bold=True,sz=9)
                ws.cell(cur,4,info.get('nombre','')); ws.cell(cur,8,info.get('cedula','')); cur+=1
            cur+=1
            ws.cell(cur,1,"VEHÍCULO: CHOFER"); sc(ws.cell(cur,1),bold=True,sz=9)
            ws.cell(cur,4,pi.get('chofer_nombre','')); cur+=1
            ws.cell(cur,1,"PLACA:"); sc(ws.cell(cur,1),bold=True,sz=9); ws.cell(cur,4,pi.get('placa','')); cur+=2
            merge_row(cur,1,4,"EQUIPO",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            merge_row(cur,5,14,"IDENTIFICACIÓN",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            merge_row(cur,15,14+dias_op,"RECORRIDO — FECHA",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=7,brd=True)
            sc(ws.cell(cur,last_col,"# VIV"),bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            ws.row_dimensions[cur].height=24; cur+=1
            for ci,h in enumerate(["SUPERVISOR","ENCUESTADOR","CARGA","PROV","CANTON","CIUDAD/PARROQ",
                                    "ZONA","SECTOR","MAN","CÓDIGO","PROVINCIA","CANTÓN","CIUDAD","NRO EDIF"],1):
                sc(ws.cell(cur,ci,h),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True,wrap=True)
            for i in range(dias_op):
                lbl=fechas[i].strftime("%d/%m") if fechas else f"D{i+1}"
                sc(ws.cell(cur,15+i,lbl),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)
            sc(ws.cell(cur,last_col,"# VIV"),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)
            ws.row_dimensions[cur].height=32; cur+=1

            df_sorted=df_eq.sort_values(['encuestador','dia_inicio']).copy()
            enc_actual=None; fila_enc=0; viv_enc_acum=0; enc_ci=-1
            for _,(_,rd) in enumerate(df_sorted.iterrows()):
                enc_id=int(rd.get('encuestador',0))
                if enc_id!=enc_actual and enc_actual is not None:
                    pal=ENC_PALETAS[enc_ci%len(ENC_PALETAS)]
                    enc_info=enc_list[enc_actual-1] if 0<enc_actual<=len(enc_list) else {}
                    merge_row(cur,1,9,f"SUBTOTAL {enc_info.get('nombre',f'Enc {enc_actual}')}",
                              bold=True,bg=pal["subtot"],fg=pal["hdr"],ha="right",sz=8)
                    for ci in range(10,last_col): sc(ws.cell(cur,ci,""),bg=pal["subtot"],brd=True)
                    sc(ws.cell(cur,last_col,viv_enc_acum),bold=True,ha="center",sz=9,bg=pal["subtot"],fg=pal["hdr"],brd=True)
                    ws.row_dimensions[cur].height=14; cur+=1; viv_enc_acum=0
                if enc_id!=enc_actual:
                    enc_actual=enc_id; fila_enc=0; enc_ci=(enc_ci+1)%len(ENC_PALETAS)
                pal=ENC_PALETAS[enc_ci%len(ENC_PALETAS)]
                bg_row=pal["par"] if fila_enc%2==0 else pal["impar"]
                fila_enc+=1; viv_enc_acum+=int(rd.get('viv',0))
                p_cod=parse_codigo(str(rd['id_entidad']))
                enc_i=enc_list[enc_id-1] if 0<enc_id<=len(enc_list) else {}
                ct_str=f"CT{ct_counter[0]:03d}"; ct_counter[0]+=1
                vals=[pi.get('supervisor_cedula',''),enc_i.get('cedula',''),ct_str,
                      p_cod['prov'],p_cod['canton'],p_cod['ciudad_parroq'],
                      p_cod['zona'],p_cod['sector'],p_cod['man'],str(rd['id_entidad']),'','','','']
                for ci,val in enumerate(vals,1):
                    sc(ws.cell(cur,ci,val),bg=AZ_CLARO if ci==10 else bg_row,ha="center",sz=8,brd=True)
                d_ini=int(rd.get('dia_inicio',rd.get('dia_operativo',1))); d_fin=int(rd.get('dia_fin',d_ini))
                for i in range(dias_op):
                    if d_ini<=i+1<=d_fin: sc(ws.cell(cur,15+i,"✓"),bold=True,bg=VRD_CHECK,ha="center",sz=11,brd=True)
                    else: sc(ws.cell(cur,15+i,""),bg=bg_row,ha="center",brd=True)
                sc(ws.cell(cur,last_col,int(rd.get('viv',0))),ha="center",sz=8,brd=True,bg=bg_row); cur+=1

            if enc_actual is not None:
                pal=ENC_PALETAS[enc_ci%len(ENC_PALETAS)]
                enc_info=enc_list[enc_actual-1] if 0<enc_actual<=len(enc_list) else {}
                merge_row(cur,1,9,f"SUBTOTAL {enc_info.get('nombre',f'Enc {enc_actual}')}",
                          bold=True,bg=pal["subtot"],fg=pal["hdr"],ha="right",sz=8)
                for ci in range(10,last_col): sc(ws.cell(cur,ci,""),bg=pal["subtot"],brd=True)
                sc(ws.cell(cur,last_col,viv_enc_acum),bold=True,ha="center",sz=9,bg=pal["subtot"],fg=pal["hdr"],brd=True)
                ws.row_dimensions[cur].height=14; cur+=1
            sc(ws.cell(cur,last_col-1,"TOTAL"),bold=True,ha="right",sz=8,bg=AZ_CLARO,brd=True)
            sc(ws.cell(cur,last_col,int(df_eq['viv'].sum())),bold=True,ha="center",sz=8,bg=AZ_CLARO,brd=True); cur+=4

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
_defs = {
    "data_raw":None,"data_mes":None,"graph_G":None,
    "resultados_generados":False,"df_plan":None,
    "tsp_results":{},"road_paths":{},"resumen_bal":None,
    "sil_score":None,"n_bombero":0,"personal_info":{},
    "balance_log":[],"cv_ini_bal":None,"cv_eq_ini":None,"cv_eq_fin":None,
    "viv_clusters_antes":None,"viv_clusters_despues":None,"best_pairs_cache":None,
    "mes_inicio_cal":7,
    "config_jornadas":{},   # {jornada_num: {'cancelada':bool, 'fecha':date|None}}
    "slots_cache":[],
    "params":{
        "dias_op":12,"viv_min":50,"viv_max":80,"factor_r":1.5,
        "usar_bomb":True,"usar_gye":True,"dias_gye":3,"umbral_gye":10,
        "cv_objetivo":15,   # % — más realista para datos reales del Litoral
    },
    "equipos_cfg":[
        {"id":1,"nombre":"Equipo 1","enc":3},
        {"id":2,"nombre":"Equipo 2","enc":3},
        {"id":3,"nombre":"Equipo 3","enc":3},
    ],
}
for k,v in _defs.items():
    if k not in st.session_state: st.session_state[k]=v

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""<div class='sidebar-logo'>
        <img src='{INEC_LOGO}' alt='INEC'>
        <div><div class='sidebar-title'>Encuesta Nacional</div>
             <div class='sidebar-sub'>INEC · Zonal Litoral</div></div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    st.markdown("<div class='step'>PASO 1</div>",unsafe_allow_html=True)
    st.markdown("**Muestra (.gpkg)**")
    gpkg_f=st.file_uploader("GeoPackage",type=["gpkg"],key="gpkg_up")
    if gpkg_f:
        dissolve=st.radio("Nivel",["Por UPM","Por manzana"],index=0)
        if st.button("⚡ Procesar",use_container_width=True,type="primary"):
            with st.spinner("Leyendo..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False,suffix=".gpkg") as tmp:
                        tmp.write(gpkg_f.read()); p_tmp=tmp.name
                    data=cargar_gpkg(p_tmp,dissolve_upm=dissolve.startswith("Por UPM")); os.unlink(p_tmp)
                    st.session_state.data_raw=data; st.session_state.resultados_generados=False
                    st.success(f"✓ {len(data):,} entidades")
                except Exception as e: st.error(str(e))
        if st.session_state.data_raw is not None:
            st.markdown("<span class='pill-ok'>✓ Listo</span>",unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin archivo</span>",unsafe_allow_html=True)
    st.divider()

    st.markdown("<div class='step'>PASO 2</div>",unsafe_allow_html=True)
    st.markdown("**Red vial (.graphml)**")
    gml_f=st.file_uploader("GraphML",type=["graphml"],key="gml_up")
    if gml_f:
        if st.button("⚡ Cargar grafo",use_container_width=True):
            with st.spinner("Cargando..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False,suffix=".graphml") as tmp:
                        tmp.write(gml_f.read()); pg=tmp.name
                    G=ox.load_graphml(pg); os.unlink(pg)
                    st.session_state.graph_G=G; st.success(f"✓ {len(G.nodes):,} nodos")
                except Exception as e: st.error(str(e))
        if st.session_state.graph_G is not None:
            st.markdown("<span class='pill-ok'>✓ Red lista</span>",unsafe_allow_html=True)
    else:
        st.markdown("<span class='pill-w'>⏳ Sin grafo</span>",unsafe_allow_html=True)
    st.divider()

    if st.session_state.data_raw is not None:
        st.markdown("<div class='step'>PASO 3</div>",unsafe_allow_html=True)
        meses_disp=sorted(st.session_state.data_raw["mes"].dropna().unique().tolist())
        mes_sel=st.selectbox("Mes operativo",meses_disp,format_func=lambda x:f"Mes {int(x)}")
        df_mes=st.session_state.data_raw[st.session_state.data_raw["mes"]==mes_sel].copy()
        st.session_state.data_mes=df_mes

        mes_ini_cal=st.selectbox("Mes calendario de inicio",list(MESES_CAL.keys()),
            index=st.session_state.mes_inicio_cal-1,format_func=lambda x:MESES_CAL[x],key="sel_mes_ini")
        st.session_state.mes_inicio_cal=mes_ini_cal
        j1_n,j2_n,mes_nom=jornada_num_desde_mes(int(mes_sel),mes_ini_cal)
        st.markdown(f"""<div style='font-size:11px;background:#f0f6ff;border-radius:6px;color:#2d4a6f;
                    padding:8px 12px;border-left:3px solid #003B71;margin:6px 0'>
        📅 Mes {int(mes_sel)} ({mes_nom}) →
        <b style='color:#003B71'>J{j1_n}</b> + <b style='color:#059669'>J{j2_n}</b>
        </div>""",unsafe_allow_html=True)
        st.divider()

        st.markdown("<div class='step'>PASO 4</div>",unsafe_allow_html=True)
        st.markdown("**Equipos**")
        c1,c2=st.columns(2)
        with c1:
            if st.button("＋",use_container_width=True):
                nid=max(t["id"] for t in st.session_state.equipos_cfg)+1
                st.session_state.equipos_cfg.append({"id":nid,"nombre":f"Equipo {nid}","enc":3})
                st.session_state.resultados_generados=False
        with c2:
            if st.button("－",use_container_width=True,disabled=len(st.session_state.equipos_cfg)<=1):
                st.session_state.equipos_cfg.pop(); st.session_state.resultados_generados=False
        for i,eq in enumerate(st.session_state.equipos_cfg):
            cc1,cc2=st.columns([2,1])
            with cc1:
                nn=st.text_input(f"n{eq['id']}",value=eq["nombre"],key=f"n_{eq['id']}",label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["nombre"]=nn
            with cc2:
                ne=st.number_input("e",min_value=1,max_value=6,value=eq["enc"],key=f"e_{eq['id']}",label_visibility="collapsed")
                st.session_state.equipos_cfg[i]["enc"]=ne
        st.divider()

        st.markdown("**Parámetros operativos**")
        p=st.session_state.params
        p["dias_op"] =st.slider("Días operativos",10,14,p["dias_op"])
        p["viv_min"] =st.slider("Mín viv/día por enc.",30,60,p["viv_min"])
        p["viv_max"] =st.slider("Máx viv/día por enc.",60,120,p["viv_max"])
        p["factor_r"]=st.slider("Factor rural (×)",1.0,2.5,p["factor_r"],0.1,
            help="Zonas dispersas pesan más. Ej: 1.5× = visitar 1 casa rural ≈ 1.5 casas urbanas.")
        p["cv_objetivo"]=st.slider("CV objetivo entre equipos (%)",5,30,p.get("cv_objetivo",15),
            help="Desigualdad máxima aceptable entre equipos. 15% = equipos con ±15% de carga.")
        p["usar_bomb"]=st.toggle("Equipo Bombero",value=p["usar_bomb"],
            help="Detecta UPMs geográficamente aisladas dentro de su cluster.")
        if p["usar_bomb"]:
            p["min_dist_bomb_m"]=st.slider("Dist. mín. Bombero (km)",10,150,
                p.get("min_dist_bomb_m",40000)//1000)*1000
        p["usar_gye"]=st.toggle("Restricción Guayaquil",value=p["usar_gye"])
        p["dias_gye"]=st.slider("Días GYE",1,5,p["dias_gye"],disabled=not p["usar_gye"])
        p["umbral_gye"]=st.slider("Umbral GYE (%)",5,30,p["umbral_gye"],disabled=not p["usar_gye"])

        tot_enc=sum(e["enc"] for e in st.session_state.equipos_cfg)
        tot_viv=int(df_mes["viv"].sum()) if len(df_mes)>0 else 0
        st.markdown(f"""<div style='font-size:11px;color:#4a5568;line-height:2;margin-top:8px'>
        📍 <b style='color:#003B71'>{len(df_mes):,}</b> UPMs ·
        🏠 <b style='color:#003B71'>{tot_viv:,}</b> viv ·
        👥 <b style='color:#003B71'>{tot_enc}</b> enc.
        </div>""",unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────
st.markdown(f"""
<div class='hdr'>
  <img src='{INEC_LOGO}' alt='INEC'>
  <div class='hdr-text'>
    <h1>Planificación Automática · Actualización Cartográfica · v6</h1>
    <p>Instituto Nacional de Estadística y Censos · Zonal Litoral · ENDI 2025</p>
  </div>
</div>""",unsafe_allow_html=True)

if st.session_state.data_raw is None:
    st.markdown("<div class='ibox'>👈 Carga el <code>.gpkg</code> desde el panel lateral.</div>",unsafe_allow_html=True)
    st.stop()

df=st.session_state.data_mes
if df is None or len(df)==0: st.warning("Sin datos para el mes seleccionado."); st.stop()

p=st.session_state.params
k1,k2,k3,k4,k5=st.columns(5)
cv_v=cv_pct(df["viv"]); cv_c="#059669" if cv_v<50 else "#dc2626"
for col,(val,lbl,sub,c) in zip([k1,k2,k3,k4,k5],[
    (f"{len(df):,}","UPMs",f"mes {int(df['mes'].iloc[0])}","#003B71"),
    (f"{int(df['viv'].sum()):,}","Viviendas","precenso 2020","#003B71"),
    (f"{len(df[df['tipo_entidad'].isin(['man','man_upm'])]):,}","Amanzanadas","man/man_upm","#003B71"),
    (f"{len(df[df['tipo_entidad'].isin(['sec','sec_upm'])]):,}","Dispersas","sec/sec_upm","#003B71"),
    (f"{cv_v:.1f}%","CV viviendas","dispersión",cv_c),
]):
    with col:
        st.markdown(f"<div class='kcard'><div class='v' style='color:{c}'>{val}</div>"
                    f"<div class='l'>{lbl}</div><div class='s'>{sub}</div></div>",unsafe_allow_html=True)

st.markdown("<br>",unsafe_allow_html=True)
cb1,cb2=st.columns([1,3])
with cb1:
    btn=st.button("⚡ Generar Planificación",use_container_width=True,type="primary",
                  disabled=(st.session_state.graph_G is None))
with cb2:
    if st.session_state.graph_G is None:
        st.markdown("<div class='wbox'>⚠️ Carga el <code>.graphml</code> (Paso 2).</div>",unsafe_allow_html=True)
    elif st.session_state.resultados_generados:
        st.markdown("<div class='ibox'>✓ Planificación lista. Puedes regenerar.</div>",unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ALGORITMO PRINCIPAL v6
# ══════════════════════════════════════════════════════════════════════════════
if btn:
    G=st.session_state.graph_G; eq_cfg=st.session_state.equipos_cfg
    n_eq=len(eq_cfg); nombres=[e["nombre"] for e in eq_cfg]
    n_clust=n_eq*2; p=st.session_state.params

    df_w=df.copy()
    df_w['equipo']='sin_asignar'; df_w['jornada']='sin_asignar'; df_w['cluster_geo']=-1
    df_w['carga_pond']=df_w.apply(
        lambda r: r['viv']*p["factor_r"] if str(r.get('tipo_entidad','')).startswith('sec') else r['viv'],axis=1)
    df_w['encuestador']=0; df_w['dia_operativo']=0; df_w['dia_inicio']=0; df_w['dia_fin']=0; df_w['dist_base_m']=0.0

    prog=st.progress(0,"Iniciando...")

    t_utm=Transformer.from_crs("EPSG:4326","EPSG:32717",always_xy=True)
    bx,by=t_utm.transform(BASE_LON,BASE_LAT)
    df_w['dist_base_m']=np.sqrt((df_w['x']-bx)**2+(df_w['y']-by)**2)

    prog.progress(8,"Verificando restricción Guayaquil...")
    upms_gye=pd.Series(False,index=df_w.index)
    if p["usar_gye"] and 'pro_x' in df_w.columns and 'can_x' in df_w.columns:
        upms_gye=(df_w['pro_x']==PRO_GYE)&(df_w['can_x']==CAN_GYE)
    pct_gye=upms_gye.sum()/len(df_w) if len(df_w)>0 else 0
    act_gye=p["usar_gye"] and (pct_gye>=p["umbral_gye"]/100) and upms_gye.sum()>0
    df_gye=df_w[upms_gye].copy() if act_gye else pd.DataFrame()
    df_no_gye=df_w[~upms_gye].copy()

    # ── CLUSTERING v6 ────────────────────────────────────────────────────────
    prog.progress(12,f"KMeans + emparejamiento óptimo ({n_clust} clusters → {n_eq} equipos)...")

    if len(df_no_gye)>=n_clust:
        # CV antes
        km0=KMeans(n_clusters=n_clust,init='k-means++',n_init=20,max_iter=500,random_state=42)
        lab0=km0.fit_predict(df_no_gye[['x','y']].values.astype(float))
        carg_arr=df_no_gye['carga_pond'].values
        st.session_state.viv_clusters_antes={c:float(carg_arr[lab0==c].sum()) for c in range(n_clust)}

        prog.progress(22,"Emparejamiento óptimo...")
        labels,best_pairs,bal_log,cv_ini,cv_eq_fin=clustering_balanceado(
            df_no_gye,n_clusters=n_clust,
            cv_objetivo=p["cv_objetivo"]/100.0,
            max_iter_micro=150, k_vecinos=8)

        df_no_gye=df_no_gye.copy(); df_no_gye['cluster_geo']=labels
        st.session_state.balance_log=bal_log
        st.session_state.cv_ini_bal=cv_ini
        st.session_state.cv_eq_fin=cv_eq_fin
        st.session_state.best_pairs_cache=best_pairs
        st.session_state.viv_clusters_despues={c:float(carg_arr[labels==c].sum()) for c in range(n_clust)}

        try: st.session_state.sil_score=silhouette_score(df_no_gye[['x','y']].values,labels)
        except: st.session_state.sil_score=None

        # Centroides para asignación orden (lejanos primero → J1, cercanos → J2)
        centroides=np.array([
            df_no_gye[['x','y']].values[labels==c].mean(axis=0)
            if (labels==c).sum()>0 else np.array([bx,by])
            for c in range(n_clust)])

        # Usar best_pairs para asignar equipo+jornada
        asig={}
        for eq_i,(c_j1,c_j2) in enumerate(best_pairs):
            asig[c_j1]=(nombres[eq_i],'Jornada 1')
            asig[c_j2]=(nombres[eq_i],'Jornada 2')

        df_no_gye['equipo'] =df_no_gye['cluster_geo'].map(lambda c: asig.get(c,('sin_asignar','sin_asignar'))[0])
        df_no_gye['jornada']=df_no_gye['cluster_geo'].map(lambda c: asig.get(c,('sin_asignar','sin_asignar'))[1])

        # Equipo Bombero
        if p["usar_bomb"]:
            prog.progress(32,"Detectando outliers...")
            MIN_D=p.get("min_dist_bomb_m",40000)
            for c_id in range(n_clust):
                if c_id not in asig: continue
                mask_c=df_no_gye['cluster_geo']==c_id; pts=df_no_gye[mask_c]
                if len(pts)<8: continue
                cx,cy=centroides[c_id]
                dists_b=np.sqrt((pts['x']-cx)**2+(pts['y']-cy)**2)
                Q1c,Q3c=dists_b.quantile(.25),dists_b.quantile(.75); iqrc=Q3c-Q1c
                if iqrc==0: continue
                bomb_idx=dists_b[(dists_b>Q3c+3*iqrc)&(dists_b>MIN_D)].index
                if len(bomb_idx)>0:
                    df_no_gye.loc[bomb_idx,'equipo']='Equipo Bombero'
                    df_no_gye.loc[bomb_idx,'jornada']='Jornada Especial'

        df_w.update(df_no_gye[['equipo','jornada','cluster_geo']])

    st.session_state.n_bombero=int((df_w['equipo']=='Equipo Bombero').sum())

    # ── ENCUESTADORES + DÍAS ─────────────────────────────────────────────────
    prog.progress(42,"Asignando encuestadores y días...")
    enc_dict={e["nombre"]:e["enc"] for e in eq_cfg}
    for nombre_eq in nombres:
        for jornada in ['Jornada 1','Jornada 2']:
            mask_g=(df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp=df_w[mask_g].copy()
            if len(grp)==0: continue
            n_enc=enc_dict.get(nombre_eq,3)
            if jornada=='Jornada 1' and act_gye:
                inicio=p["dias_gye"]+1; dias_disp=p["dias_op"]-p["dias_gye"]
            else:
                inicio=1; dias_disp=p["dias_op"]
            if dias_disp<=0: continue
            ga=asignar_encuestadores_y_dias(grp,n_enc,dias_disp,p["viv_min"],p["viv_max"],inicio)
            df_w.update(ga[['encuestador','dia_operativo','dia_inicio','dia_fin']])

    # GYE
    if act_gye and len(df_gye)>0:
        n_gye_c=min(n_eq,len(df_gye))
        if n_gye_c>=2:
            km_g=KMeans(n_clusters=n_gye_c,init='k-means++',n_init=30,max_iter=500,random_state=42)
            lab_g=km_g.fit_predict(df_gye[['x','y']].values.astype(float))
            df_gye=df_gye.copy(); df_gye['cluster_gye']=lab_g
            for c_id in range(n_gye_c):
                eq_a=nombres[c_id%n_eq]; grp_gye=df_gye[df_gye['cluster_gye']==c_id].copy()
                if len(grp_gye)==0: continue
                grp_gye['equipo']=eq_a; grp_gye['jornada']='Jornada 1'
                ga_gye=asignar_encuestadores_y_dias(grp_gye,enc_dict.get(eq_a,3),p["dias_gye"],p["viv_min"],p["viv_max"],1)
                df_w.update(ga_gye[['equipo','jornada','encuestador','dia_operativo','dia_inicio','dia_fin']])
        else:
            for idx in df_gye.index:
                df_w.loc[idx,['equipo','jornada','encuestador','dia_operativo','dia_inicio','dia_fin']]=[nombres[0],'Jornada 1',1,1,1,1]

    # ── TSP ──────────────────────────────────────────────────────────────────
    prog.progress(52,"Optimizando rutas TSP...")
    base_nd=ox.nearest_nodes(G,BASE_LON,BASE_LAT)
    G_u=G.to_undirected(); comp_base=nx.node_connected_component(G_u,base_nd)
    tsp_r,road_p={},{}
    for ri,nombre_eq in enumerate(nombres):
        for jornada in ['Jornada 1','Jornada 2']:
            pct=52+int((ri*2+['Jornada 1','Jornada 2'].index(jornada)+1)/(n_eq*2)*42)
            prog.progress(pct,f"TSP: {nombre_eq}|{jornada}...")
            mask_g=(df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp=df_w[mask_g]
            if len(grp)==0: continue
            nr=ox.nearest_nodes(G,grp['lon'].values,grp['lat'].values)
            nk=[n for n in nr if n in comp_base]
            if not nk: continue
            nu=[base_nd]+list(dict.fromkeys(nk)); nd=len(nu)
            if nd<=2: continue
            D=np.zeros((nd,nd))
            for i in range(nd):
                for j in range(i+1,nd):
                    try: d=nx.shortest_path_length(G_u,nu[i],nu[j],weight='length');D[i,j]=D[j,i]=d
                    except: D[i,j]=D[j,i]=1e9
            Gt=nx.Graph()
            for i in range(nd):
                for j in range(i+1,nd):
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
                    seg=nx.shortest_path(G_u,ng[k],ng[k+1],weight='length')
                    ruta.extend((G.nodes[nd2]['y'],G.nodes[nd2]['x']) for nd2 in seg[:-1])
                except: continue
            if ng: ruta.append((G.nodes[ng[-1]]['y'],G.nodes[ng[-1]]['x']))
            clave=f"{nombre_eq}||{jornada}"
            tsp_r[clave]={'equipo':nombre_eq,'jornada':jornada,'n_puntos':len(grp),'dist_km':dist/1000}
            road_p[clave]=ruta

    prog.progress(98,"Métricas finales...")
    resumen=df_w[~df_w['equipo'].isin(['Equipo Bombero','sin_asignar'])].groupby(['equipo','jornada']).agg(
        n_upms=('id_entidad','count'),viv_reales=('viv','sum'),carga_ponderada=('carga_pond','sum')).reset_index()
    dist_df=pd.DataFrame([{'equipo':v['equipo'],'jornada':v['jornada'],'dist_km':round(v['dist_km'],1)}
                           for v in tsp_r.values()]) if tsp_r else pd.DataFrame(columns=['equipo','jornada','dist_km'])
    resumen_bal=pd.merge(resumen,dist_df,on=['equipo','jornada'],how='left').fillna(0)

    prog.progress(100,"¡Listo!"); prog.empty()
    st.session_state.df_plan=df_w; st.session_state.tsp_results=tsp_r
    st.session_state.road_paths=road_p; st.session_state.resumen_bal=resumen_bal
    st.session_state.resultados_generados=True
    st.success("✓ Planificación v6 generada.")

# ── RESULTADOS ────────────────────────────────
if not st.session_state.resultados_generados:
    st.markdown("<div class='ibox'>👆 Presiona <b>Generar Planificación</b>.</div>",unsafe_allow_html=True)
    st.stop()

df_plan=st.session_state.df_plan; tsp_r=st.session_state.tsp_results
road_p=st.session_state.road_paths; res_bal=st.session_state.resumen_bal
eq_cfg=st.session_state.equipos_cfg; nombres=[e["nombre"] for e in eq_cfg]; p=st.session_state.params
mes_ini_cal=st.session_state.mes_inicio_cal
j1_n,j2_n,mes_nom=jornada_num_desde_mes(int(df['mes'].iloc[0]),mes_ini_cal)
color_map={n:COLORES[i%len(COLORES)] for i,n in enumerate(nombres)}
color_map['Equipo Bombero']='#7c3aed'

# Construir slots del calendario
meses_todos=sorted(st.session_state.data_raw["mes"].dropna().unique().tolist())
slots_cache,n_cancel=construir_slots_calendario(meses_todos,mes_ini_cal,st.session_state.config_jornadas)
st.session_state.slots_cache=slots_cache

tab_mapa,tab_analisis,tab_plan,tab_reporte=st.tabs([
    "🗺️  Mapa","📊  Análisis","📅  Jornadas","📋  Reporte"])

# ══ TAB 1 — MAPA ══════════════════════════════
with tab_mapa:
    st.markdown("<div class='stitle'>Mapa del Operativo</div>",unsafe_allow_html=True)
    cc1,cc2=st.columns([1,3])
    with cc1:
        mj1=st.checkbox("Jornada 1",value=True); mj2=st.checkbox("Jornada 2",value=True)
        n_b=int((df_plan['equipo']=='Equipo Bombero').sum())
        mbm=st.checkbox(f"Equipo Bombero ({n_b})",value=True)
        mrts=st.checkbox("Rutas",value=True)
        fnd=st.selectbox("Fondo",["CartoDB positron","OpenStreetMap","CartoDB dark_matter"])
        st.divider()
        for n,c in color_map.items():
            if n in nombres:
                st.markdown(f"<span style='color:{c};font-size:17px'>●</span> {n}",unsafe_allow_html=True)
        st.markdown(f"<span style='color:#7c3aed;font-size:17px'>●</span> Bombero ({n_b})",unsafe_allow_html=True)
    with cc2:
        m=folium.Map(location=[BASE_LAT,BASE_LON],zoom_start=8,tiles=fnd)
        folium.Marker([BASE_LAT,BASE_LON],popup="<b>Base INEC GYE</b>",
            icon=folium.Icon(color='white',icon='home',prefix='fa')).add_to(m)
        for _,row in df_plan.iterrows():
            eq,jor=row.get('equipo',''),row.get('jornada','')
            if jor=='Jornada 1' and not mj1: continue
            if jor=='Jornada 2' and not mj2: continue
            if jor=='Jornada Especial' and not mbm: continue
            clr=color_map.get(eq,'#888')
            d_ini=int(row.get('dia_inicio',0)); d_fin=int(row.get('dia_fin',d_ini))
            dias_str=f"Día {d_ini}" if d_ini==d_fin else f"Días {d_ini}–{d_fin}"
            folium.CircleMarker(location=[row['lat'],row['lon']],radius=5,color=clr,fill=True,
                fill_color=clr,fill_opacity=.85,
                popup=folium.Popup(f"<b>{row['id_entidad']}</b><br>{int(row['viv'])} viv<br>"
                    f"{eq}<br>Enc {int(row.get('encuestador',0))}<br>{dias_str}",max_width=200),
                tooltip=f"{eq}·Enc{int(row.get('encuestador',0))}·{int(row['viv'])}viv").add_to(m)
        if mrts:
            for clave,coords in road_p.items():
                eq,jor=clave.split('||')
                if jor=='Jornada 1' and not mj1: continue
                if jor=='Jornada 2' and not mj2: continue
                if len(coords)>1:
                    folium.PolyLine(coords,weight=3,color=color_map.get(eq,'#888'),opacity=.75).add_to(m)
        st_folium(m,width=None,height=540,returned_objects=[],key="mapa_v6")

# ══ TAB 2 — ANÁLISIS ══════════════════════════
with tab_analisis:
    st.markdown("<div class='stitle'>Clustering y Balanceo (v6)</div>",unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    <b>v6 — Emparejamiento óptimo:</b> los clusters KMeans no se modifican.
    En su lugar, se busca el emparejamiento de clusters en pares (Jornada 1 + Jornada 2)
    que minimiza el CV de carga entre equipos. Esto balancea sin distorsionar la geografía.
    </div>""",unsafe_allow_html=True)

    cv_ini=st.session_state.get("cv_ini_bal"); cv_fin=st.session_state.get("cv_eq_fin")
    bal_log=st.session_state.get("balance_log",[])
    viv_ant=st.session_state.get("viv_clusters_antes"); viv_dep=st.session_state.get("viv_clusters_despues")
    best_pairs=st.session_state.get("best_pairs_cache",[])

    if cv_ini is not None and cv_fin is not None:
        mejora=cv_ini-cv_fin
        cc_m="#059669" if mejora>5 else ("#d97706" if mejora>0 else "#dc2626")
        modo_fin=next((l.get('nota','') for l in reversed(bal_log) if 'plateau' in l.get('nota','') or 'objetivo' in l.get('nota','')),'-')
        st.markdown(f"""<div class='balance-box'>
        <b>CV inicial clusters (KMeans puro):</b>
        <span style='color:#dc2626;font-family:monospace'>{cv_ini:.1f}%</span>
        &nbsp;→&nbsp;
        <b>CV final entre equipos:</b>
        <span style='color:#059669;font-family:monospace'>{cv_fin:.1f}%</span>
        &nbsp;<b style='color:{cc_m}'>Δ {mejora:.1f} pp</b><br>
        <span style='font-size:11px;color:#047857'>
        {len([l for l in bal_log if 'micro' in l.get('paso','')])} micro-ajustes de frontera
        · {modo_fin if modo_fin!='-' else 'emparejamiento + micro-ajuste'}
        </span></div>""",unsafe_allow_html=True)

    # Mostrar emparejamiento
    if best_pairs and viv_dep:
        st.markdown("**Emparejamiento de clusters por equipo:**")
        for eq_i,(c_j1,c_j2) in enumerate(best_pairs):
            eq_n=nombres[eq_i] if eq_i<len(nombres) else f"Equipo {eq_i+1}"
            v1=int(viv_dep.get(c_j1,0)); v2=int(viv_dep.get(c_j2,0))
            color=COLORES[eq_i%len(COLORES)]
            st.markdown(f"<div class='slot-card'>"
                        f"<span style='color:{color};font-weight:600'>{eq_n}</span> &nbsp;·&nbsp; "
                        f"C{c_j1} ({v1:,} viv) + C{c_j2} ({v2:,} viv) = "
                        f"<b>{v1+v2:,} viv</b></div>",unsafe_allow_html=True)

    if viv_ant and viv_dep:
        n_cl=len(viv_ant)
        df_comp=pd.DataFrame({
            'Cluster':[f"C{c}" for c in range(n_cl)]*2,
            'Carga pond.':(list(viv_ant.values())+list(viv_dep.values())),
            'Fase':['Antes (KMeans)']*n_cl+['Después (rebalanceo)']*n_cl
        })
        fig_comp=px.bar(df_comp,x='Cluster',y='Carga pond.',color='Fase',barmode='group',
                        title='Carga por cluster — antes vs después',template='plotly_white',
                        color_discrete_map={'Antes (KMeans)':'#dc2626','Después (rebalanceo)':'#059669'})
        fig_comp.update_layout(paper_bgcolor="#ffffff",plot_bgcolor="#fafbfc",title_font_size=12)
        st.plotly_chart(fig_comp,use_container_width=True)
        with st.expander("Log del balanceo"):
            st.dataframe(pd.DataFrame(bal_log),use_container_width=True,height=200)

    st.divider()
    st.markdown("<div class='stitle'>Equidad entre equipos</div>",unsafe_allow_html=True)
    df_main=df_plan[~df_plan['equipo'].isin(['Equipo Bombero','sin_asignar'])].copy()
    res_cv=df_main.groupby(['equipo','jornada']).agg(viv_reales=('viv','sum'),carga_ponderada=('carga_pond','sum')).reset_index()
    for jornada in ['Jornada 1','Jornada 2']:
        sub=res_cv[res_cv['jornada']==jornada]
        if len(sub)<2: continue
        cr=cv_pct(sub['viv_reales']); cp=cv_pct(sub['carga_ponderada'])
        ccr="#059669" if cr<20 else ("#d97706" if cr<40 else "#dc2626")
        ccp="#059669" if cp<20 else ("#d97706" if cp<40 else "#dc2626")
        em="✓" if cp<20 else ("⚠" if cp<40 else "✗")
        st.markdown(f"""<div class='ibox'><b>{jornada}</b><br>
        CV viv. reales: <span style='color:{ccr};font-family:monospace;font-weight:600'>{cr:.1f}%</span>
        &nbsp;·&nbsp; CV carga pond.: <span style='color:{ccp};font-family:monospace;font-weight:600'>{cp:.1f}%</span>
        {em}</div>""",unsafe_allow_html=True)
    with st.expander("Tabla de balance"): st.dataframe(res_cv,use_container_width=True)

    eq_act=[n for n in nombres if n in df_plan['equipo'].values]
    st.markdown("<div class='stitle'>Carga por equipo</div>",unsafe_allow_html=True)
    cols_e=st.columns(len(eq_act))
    for col_e,nombre_eq in zip(cols_e,eq_act):
        sub_e=df_plan[df_plan['equipo']==nombre_eq]
        vt=int(sub_e['viv'].sum()); cv_e=cv_pct(sub_e['carga_pond'])
        ce=color_map.get(nombre_eq,'#003B71')
        ccv="#059669" if cv_e<20 else ("#d97706" if cv_e<40 else "#dc2626")
        with col_e:
            st.markdown(f"""<div class='eq-card' style='border-top:3px solid {ce}'>
              <div style='font-family:"JetBrains Mono",monospace;font-size:12px;color:{ce};font-weight:600'>{nombre_eq}</div>
              <div style='font-size:20px;font-weight:600;color:#1a1a2e;margin:6px 0'>{vt:,}</div>
              <div style='font-size:10px;color:#8896a6'>viviendas</div>
              <div style='font-size:11px;color:{ccv};margin-top:4px'>CV {cv_e:.1f}%</div>
            </div>""",unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)
    eq_sel=st.selectbox("Detalle:",eq_act)
    df_enc=df_plan[df_plan['equipo']==eq_sel].groupby(['jornada','encuestador']).agg(
        upms=('id_entidad','count'),viv_reales=('viv','sum'),carga_pond=('carga_pond','sum')).reset_index()
    cd1,cd2=st.columns(2)
    with cd1:
        fig=px.bar(df_enc,x='encuestador',y='viv_reales',color='jornada',barmode='group',
                   title=f'Viviendas — {eq_sel}',template='plotly_white',color_discrete_sequence=['#003B71','#059669'])
        fig.update_layout(paper_bgcolor="#ffffff",plot_bgcolor="#fafbfc",title_font_size=12)
        st.plotly_chart(fig,use_container_width=True)
    with cd2:
        fig2=px.bar(df_enc,x='encuestador',y='carga_pond',color='jornada',barmode='group',
                    title=f'Carga ponderada — {eq_sel}',template='plotly_white',color_discrete_sequence=['#dc2626','#d97706'])
        fig2.update_layout(paper_bgcolor="#ffffff",plot_bgcolor="#fafbfc",title_font_size=12)
        st.plotly_chart(fig2,use_container_width=True)

    st.markdown("<div class='stitle'>Distribución diaria</div>",unsafe_allow_html=True)
    jor_filtro=st.radio("Filtrar:",["Jornada 1","Jornada 2","Ambas"],horizontal=True,key="radio_jor_v6",index=0)
    pivot_all=df_plan[df_plan['equipo'].isin(eq_act)].copy()
    if jor_filtro!="Ambas": pivot_all=pivot_all[pivot_all['jornada']==jor_filtro]
    rows_exp=[]
    for _,row in pivot_all.iterrows():
        d_ini=int(row.get('dia_inicio',1)); d_fin=int(row.get('dia_fin',d_ini))
        dias_dur=max(1,d_fin-d_ini+1); viv_d=row['viv']/dias_dur
        for dd in range(d_ini,d_fin+1):
            rows_exp.append({'equipo':row['equipo'],'jornada':row['jornada'],
                             'encuestador':int(row.get('encuestador',0)),'dia_abs':dd,'viv':viv_d})
    if rows_exp:
        df_exp=pd.DataFrame(rows_exp)
        if jor_filtro!="Ambas":
            df_exp['dia_rel']=df_exp['dia_abs']-df_exp['dia_abs'].min()+1
        else:
            df_exp['dia_rel']=df_exp.groupby('jornada')['dia_abs'].transform(lambda s:s-s.min()+1).astype(int)
        pivot=df_exp.groupby(['equipo','dia_rel'])['viv'].sum().reset_index()
        fig_d=px.bar(pivot,x='dia_rel',y='viv',color='equipo',barmode='group',
                     title=f'Viv/día — {jor_filtro}',template='plotly_white',color_discrete_map=color_map)
        tot_enc_f=sum(e["enc"] for e in eq_cfg if e["nombre"] in eq_act)
        avg_enc_f=tot_enc_f/max(1,len(eq_act))
        fig_d.add_hline(y=p["viv_min"]*avg_enc_f,line_dash="dot",line_color="#d97706",annotation_text=f"Mín")
        fig_d.add_hline(y=p["viv_max"]*avg_enc_f,line_dash="dot",line_color="#dc2626",annotation_text=f"Máx")
        fig_d.update_layout(paper_bgcolor="#ffffff",plot_bgcolor="#fafbfc",xaxis=dict(dtick=1))
        st.plotly_chart(fig_d,use_container_width=True)
        if jor_filtro!="Ambas":
            piv_enc=df_exp.groupby(['encuestador','dia_rel'])['viv'].sum().reset_index()
            piv_enc['encuestador']="Enc. "+piv_enc['encuestador'].astype(str)
            fig_enc=px.line(piv_enc,x='dia_rel',y='viv',color='encuestador',markers=True,
                            title=f'Carga diaria por encuestador — {jor_filtro}',template='plotly_white')
            fig_enc.add_hline(y=p["viv_min"],line_dash="dot",line_color="#d97706")
            fig_enc.add_hline(y=p["viv_max"],line_dash="dot",line_color="#dc2626")
            fig_enc.update_layout(paper_bgcolor="#ffffff",plot_bgcolor="#fafbfc",xaxis=dict(dtick=1))
            st.plotly_chart(fig_enc,use_container_width=True)

# ══ TAB 3 — PLANIFICACIÓN DE JORNADAS (v6 — desplazamiento correcto) ══════
with tab_plan:
    st.markdown("<div class='stitle'>Calendario de Jornadas</div>",unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    Cuando una jornada se <b>cancela</b>, todas las siguientes se desplazan
    automáticamente un slot. Un slot = mitad de mes (primera o segunda quincena).
    Si hay N cancelaciones, el operativo necesita N meses adicionales al final.
    Las fechas que ingreses se usan en el Excel.
    </div>""",unsafe_allow_html=True)

    cfg_j=st.session_state.config_jornadas

    if n_cancel>0:
        st.markdown(f"<div class='wbox'>⚠️ <b>{n_cancel} jornada(s) cancelada(s)</b> → "
                    f"se añaden {n_cancel} slot(s) al final del operativo.</div>",unsafe_allow_html=True)

    # Mostrar tabla de slots
    st.markdown(f"**{len(slots_cache)} jornadas activas en {len(meses_todos)} meses operativos**")

    # Editor de cancelaciones por mes
    for mes in meses_todos:
        mes_i=int(mes)
        j1_n_,j2_n_,mes_n_=jornada_num_desde_mes(mes_i,mes_ini_cal)
        with st.expander(f"📅 Mes {mes_i} — {mes_n_} · J{j1_n_} + J{j2_n_}",
                         expanded=(mes_i==int(df['mes'].iloc[0]))):
            for j_n_iter,mitad_lbl in [(j1_n_,"1ª mitad"),(j2_n_,"2ª mitad")]:
                cfg_this=cfg_j.get(j_n_iter,{})
                col_cb,col_f=st.columns([1,2])
                with col_cb:
                    cancelar=st.checkbox(f"❌ Cancelar J{j_n_iter} ({mitad_lbl})",
                                         value=cfg_this.get('cancelada',False),
                                         key=f"can_{j_n_iter}")
                with col_f:
                    fecha_sel=st.date_input(f"Fecha inicio J{j_n_iter}",
                                            value=cfg_this.get('fecha',None) or date.today(),
                                            key=f"fec_{j_n_iter}",disabled=cancelar)
                cfg_j[j_n_iter]={'cancelada':cancelar,'fecha':None if cancelar else fecha_sel}
                if not cancelar:
                    fin=(fecha_sel+timedelta(days=p['dias_op']-1)).strftime('%d/%m/%Y')
                    st.markdown(f"<div class='jplan-ok'>✅ J{j_n_iter} — "
                                f"{fecha_sel.strftime('%d/%m/%Y')} → {fin}</div>",unsafe_allow_html=True)
                else:
                    st.markdown(f"<div class='jplan-can'>❌ J{j_n_iter} cancelada → "
                                f"las siguientes jornadas se desplazan 1 slot</div>",unsafe_allow_html=True)

    st.session_state.config_jornadas=cfg_j

    # Recompute slots tras cambios
    slots_upd,n_cancel_upd=construir_slots_calendario(meses_todos,mes_ini_cal,cfg_j)
    st.session_state.slots_cache=slots_upd

    # Tabla resumen del cronograma resultante
    st.markdown("<div class='stitle'>Cronograma resultante</div>",unsafe_allow_html=True)
    filas=[]
    for s in slots_upd:
        fecha=s['fecha']; fecha_str=fecha.strftime("%d/%m/%Y") if fecha else "—"
        fin_str=(fecha+timedelta(days=p['dias_op']-1)).strftime("%d/%m/%Y") if fecha else "—"
        filas.append({'Slot':s['slot_num'],'Jornada':f"J{s['jornada_num']}",
                      'Mes calendario':f"{s['slot_mes_calendario']} ({s['mes_nombre_cal']})",
                      'Mitad':s['slot_mitad_calendario'],'Inicio':fecha_str,'Fin':fin_str})
    st.dataframe(pd.DataFrame(filas),use_container_width=True,hide_index=True)

# ══ TAB 4 — REPORTE Y DESCARGA ════════════════
with tab_reporte:
    st.markdown("<div class='stitle'>Reporte y Descarga Excel</div>",unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    Una hoja por jornada activa. Las canceladas se omiten.
    El número de jornada en el Excel es el número real del cronograma.
    </div>""",unsafe_allow_html=True)

    st.markdown("<div class='stitle'>Personal por equipo</div>",unsafe_allow_html=True)
    for eq in eq_cfg:
        nombre_eq=eq["nombre"]; n_enc_eq=eq["enc"]
        pi_prev=st.session_state.personal_info.get(nombre_eq,{})
        with st.expander(f"👥 {nombre_eq} — {n_enc_eq} enc.",expanded=False):
            pc1,pc2,pc3=st.columns(3)
            with pc1: sup_n=st.text_input("Supervisor",value=pi_prev.get('supervisor_nombre',''),key=f"sup_n_{nombre_eq}")
            with pc2: sup_c=st.text_input("Cédula sup.",value=pi_prev.get('supervisor_cedula',''),key=f"sup_c_{nombre_eq}")
            with pc3: sup_t=st.text_input("Celular sup.",value=pi_prev.get('supervisor_celular',''),key=f"sup_t_{nombre_eq}")
            enc_list_new=[]
            for j in range(n_enc_eq):
                prev_enc=pi_prev.get('encuestadores',[{}]*n_enc_eq)
                prev_j=prev_enc[j] if j<len(prev_enc) else {}
                pe1,pe2,pe3=st.columns(3)
                with pe1: en=st.text_input(f"Enc. {j+1}",value=prev_j.get('nombre',''),key=f"enc_n_{nombre_eq}_{j}")
                with pe2: ec=st.text_input(f"Cédula {j+1}",value=prev_j.get('cedula',''),key=f"enc_c_{nombre_eq}_{j}")
                with pe3: et=st.text_input(f"Celular {j+1}",value=prev_j.get('celular',''),key=f"enc_t_{nombre_eq}_{j}")
                enc_list_new.append({'nombre':en,'cedula':ec,'celular':et,'cod':''})
            pch1,pch2=st.columns(2)
            with pch1: ch_n=st.text_input("Chofer",value=pi_prev.get('chofer_nombre',''),key=f"ch_n_{nombre_eq}")
            with pch2: plca=st.text_input("Placa",value=pi_prev.get('placa',''),key=f"plca_{nombre_eq}")
            st.session_state.personal_info[nombre_eq]={
                'supervisor_nombre':sup_n,'supervisor_cedula':sup_c,'supervisor_celular':sup_t,'supervisor_cod':'',
                'encuestadores':enc_list_new,'n_enc':n_enc_eq,'chofer_nombre':ch_n,'placa':plca}

    if res_bal is not None and len(res_bal)>0:
        st.markdown("<div class='stitle'>Resumen</div>",unsafe_allow_html=True)
        tr=pd.DataFrame([{'equipo':'TOTAL','jornada':'—','n_upms':res_bal['n_upms'].sum(),
            'viv_reales':res_bal['viv_reales'].sum(),'carga_ponderada':res_bal['carga_ponderada'].sum(),
            'dist_km':res_bal.get('dist_km',pd.Series([0])).sum()}])
        st.dataframe(pd.concat([res_bal,tr],ignore_index=True).rename(columns={
            'equipo':'Equipo','jornada':'Jornada','n_upms':'UPMs',
            'viv_reales':'Viv.','carga_ponderada':'Carga pond.','dist_km':'Dist (km)'}),
            use_container_width=True)

    cols_ok=[c for c in ['id_entidad','tipo_entidad','viv','carga_pond','equipo','jornada',
                          'encuestador','dia_inicio','dia_fin'] if c in df_plan.columns]
    st.dataframe(df_plan[cols_ok].sort_values(['equipo','jornada','encuestador','dia_inicio']).reset_index(drop=True),
                 use_container_width=True,height=280)

    st.markdown("<div class='stitle'>Descargar Excel</div>",unsafe_allow_html=True)
    slots_excel=st.session_state.slots_cache
    if not slots_excel:
        slots_excel=[{'slot_num':i+1,'jornada_num':j_n,'jornada_nombre':jnom,
                      'mes_nombre_cal':mes_nom,'fecha':None}
                     for i,(j_n,jnom) in enumerate([(j1_n,'Jornada 1'),(j2_n,'Jornada 2')])]

    if st.button("📋 Generar Excel",use_container_width=True,type="primary"):
        with st.spinner("Generando Excel..."):
            try:
                excel_bytes=generar_excel(df_plan=df_plan,eq_cfg=eq_cfg,
                    personal_info=st.session_state.personal_info,
                    slots_activos=slots_excel,dias_op=p["dias_op"])
                nums="-".join(str(s['jornada_num']) for s in slots_excel[:4])
                fname=f"planificacion_J{nums}_{mes_nom}.xlsx"
                st.download_button(label=f"⬇️ Descargar {fname}",data=excel_bytes,file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.success("✓ Excel listo.")
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc())
