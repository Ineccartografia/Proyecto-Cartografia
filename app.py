# =============================================================================
# PLANIFICACIÓN CARTOGRÁFICA ENDI 2025 — STREAMLIT v3
# INEC · Zonal Litoral · Autores: Franklin López, Carlos Quinto
# Se realiza cambios para la automatización
# CAMBIOS v3:
# 1. MANZANAS GRANDES MULTI-DÍA: una manzana con 250 viv se distribuye en
#    ceil(250/meta_dia) días consecutivos para el MISMO encuestador.
#    Sus compañeros trabajan otras manzanas esos mismos días.
#
# 2. DISTRIBUCIÓN EN 12 DÍAS COMPLETOS: antes terminaba en 7-8 días.
#    Ahora meta_dia = total_viv_encuestador / dias_tot → usa TODOS los días.
#
# 3. EQUIPO BOMBERO POR CLUSTER: ya no detecta outliers globales desde la
#    base de Guayaquil, sino outliers DENTRO de cada cluster (IQR de distancias
#    al centroide del cluster). Más útil y geográficamente coherente.
#
# 4. GRÁFICO DÍAS FILTRABLE: radio button Jornada 1 / Jornada 2 en el gráfico
#    de intensidad diaria. Antes mezclaba ambas jornadas en el mismo eje.
#
# 5. EXCEL FORMATEADO CON OPENPYXL: reemplaza CSV.
#    - Una hoja por jornada
#    - Por equipo: encabezado institucional + bloque de personal (supervisor,
#      encuestadores, chofer, placa) + tabla de manzanas
#    - Tabla incluye columnas de fecha con ✓ según dia_inicio/dia_fin
#    - Formato visual similar al ejemplo INEC (Jornada 16)
#
# 6. PERSONAL INFO: formulario en el tab Reporte para ingresar nombres de
#    supervisor, encuestadores, chofer y placa antes de descargar.
#
# 7. FECHAS DE JORNADA: selector de fecha de inicio por jornada → las columnas
#    del cronograma muestran fechas reales (ej: "12/03") en vez de "Día 1".
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
    """
    Parsea código INEC (12 o 15 chars) en componentes geográficos.
    Formato: PP CC ZZ ZZZ SSS [MMM]
      PP=provincia, CC=cantón, ZZ=zona_tipo, ZZZ=zona, SSS=sector, MMM=manzana (opcional)
    """
    c = str(codigo).strip()
    r = {'prov':'','canton':'','ciudad_parroq':'','zona':'','sector':'','man':''}
    if len(c)>=6:  r['prov']=c[:2]; r['canton']=c[2:4]; r['ciudad_parroq']=c[4:6]
    if len(c)>=9:  r['zona']=c[6:9]
    if len(c)>=12: r['sector']=c[9:12]
    if len(c)>=15: r['man']=c[12:15]
    return r

def cargar_gpkg(path, dissolve_upm=True):
    capas = pyogrio.list_layers(path)
    
    # Detección de estructura robusta
    # Caso 1: Una sola capa (Nuevo formato ENCIET)
    if len(capas) == 1:
        gdf = gpd.read_file(path, layer=capas[0][0])
        
        # Mapeo de columnas dinámico
        col_map = {
            '1_mes_cart': 'mes',
            'viv_total': 'viv',
            '1_zonal': 'zonal',
            '1_id_upm': 'upm',
            'ManSec': 'id_entidad'
        }
        gdf = gdf.rename(columns={k: v for k, v in col_map.items() if k in gdf.columns})
        
        # Filtrar solo Litoral si la columna existe
        if 'zonal' in gdf.columns:
            gdf = gdf[gdf['zonal'].str.contains('LITORAL', na=False, case=False)]
        
        # Crear tipo_entidad basado en '999'
        if 'id_entidad' in gdf.columns:
            gdf['tipo_entidad'] = gdf['id_entidad'].astype(str).apply(
                lambda x: 'sec' if '999' in x else 'man'
            )
        
        gdf_u = gdf.to_crs(epsg=32717)
        gdf_u['geometry'] = gdf_u.geometry.representative_point()
        gdf_u['x'] = gdf_u.geometry.x
        gdf_u['y'] = gdf_u.geometry.y
        
        # Columnas de provincia/cantón para restricciones GYE
        if 'pro' in gdf_u.columns: gdf_u['pro_x'] = gdf_u['pro']
        if 'can' in gdf_u.columns: gdf_u['can_x'] = gdf_u['can']

        if dissolve_upm and 'upm' in gdf_u.columns:
            agg_dict = {'viv': 'sum', 'mes': 'first', 'x': 'first', 'y': 'first', 'tipo_entidad': 'first'}
            if 'pro_x' in gdf_u.columns: agg_dict['pro_x'] = 'first'
            if 'can_x' in gdf_u.columns: agg_dict['can_x'] = 'first'
            gdf_final = gdf_u.groupby('upm').agg(agg_dict).reset_index()
            gdf_final['id_entidad'] = gdf_final['upm']
            gdf_final['tipo_entidad'] = gdf_final['tipo_entidad'].apply(lambda t: f"{t}_upm")
        else:
            gdf_final = gdf_u
            
        return utm_to_wgs84(gdf_final)

    # Caso 2: Dos capas (Formato anterior)
    else:
        man   = gpd.read_file(path,layer=capas[0][0])
        disp  = gpd.read_file(path,layer=capas[1][0])
        man   = man[man['zonal']=='LITORAL']
        disp  = disp[disp['zonal']=='LITORAL']
        man_u = man.to_crs(epsg=32717)
        dis_u = disp.to_crs(epsg=32717)

        if dissolve_upm:
            def _d(gdf,tipo):
                d=gdf.dissolve(by='upm',aggfunc={'mes':'first','viv':'sum'})
                d['geometry']=d.geometry.representative_point()
                o=d[['mes','viv']].copy()
                o['id_entidad']=d.index; o['upm']=d.index
                o['tipo_entidad']=tipo
                o['x']=d.geometry.x; o['y']=d.geometry.y
                return o[['id_entidad','upm','mes','viv','x','y','tipo_entidad']]
            ms=_d(man_u,'man_upm'); ds=_d(dis_u,'sec_upm')
        else:
            for g in [man_u,dis_u]:
                g['geometry']=g.geometry.representative_point()
                g['x']=g.geometry.x; g['y']=g.geometry.y
            ms=man_u[['man','upm','mes','viv','x','y']].rename(columns={'man':'id_entidad'})
            ms['tipo_entidad']='man'
            ds=dis_u[['sec','upm','mes','viv','x','y']].rename(columns={'sec':'id_entidad'})
            ds['tipo_entidad']='sec'
            ms['pro_x']=ms['id_entidad'].astype(str).str[:2]
            ms['can_x']=ms['id_entidad'].astype(str).str[2:4]
            ds['pro_x']=ds['id_entidad'].astype(str).str[:2]
            ds['can_x']=ds['id_entidad'].astype(str).str[2:4]

        data=pd.concat([ms,ds],ignore_index=True)
        if not dissolve_upm:
            data=data.drop_duplicates(subset=['id_entidad','upm'],keep='first')
        return utm_to_wgs84(data)


def asignar_encuestadores_y_dias(df_grp, n_enc, dias_tot, viv_min, viv_max, inicio_dia=1):
    """
    Asignación greedy de encuestadores + distribución por target diario (v6).

    LÓGICA CORREGIDA:
    ─────────────────
    target_dia = promedio(viv_min, viv_max) = 65 viv/día por defecto.

    Para cada manzana de un encuestador:
      • Si viv > viv_max → manzana GRANDE: necesita ceil(viv / target) días
        consecutivos. El encuestador trabaja SOLO esa manzana esos días.
        Sus compañeros trabajan otras manzanas en paralelo.
        # Para cambiar el umbral de "manzana grande" modifica viv_max en el slider.
      • Si viv <= viv_max → manzana NORMAL: acumula en el día actual hasta
        superar el target, luego avanza al siguiente día.
        # Para cambiar el ritmo diario, ajusta viv_min/viv_max en los sliders.

    Por qué ya no se sobrecargan día 1 y día 12:
      • Antes: cursor avanzaba manzana a manzana con ceil() → overflow al último día.
      • Ahora: el cursor avanza solo cuando la acumulación supera el target.
        Las manzanas pequeñas restantes van a días intermedios, no al último.

    Por qué ya no aparece "227 viv en 1 día":
      • Manzana con 227 viv con target=65 → ceil(227/65)=4 días consecutivos.
        Antes podía quedar en 1 día si el budget se calculaba con total_viv alto.
        # Si aun así quieres más días para manzanas grandes, baja viv_max.
    """
    target = (viv_min + viv_max) / 2.0   # e.g. (50+80)/2 = 65 viv/día
    ultimo = inicio_dia + dias_tot - 1

    # ── Paso 1: greedy de encuestadores ──────────────────────────────────────
    # Ordenamos DESC por carga_pond para que el greedy balancee primero las
    # manzanas más pesadas. No cambiar este orden sin revisar el balance.
    df_g = df_grp.sort_values('carga_pond', ascending=False).copy()
    cargas = np.zeros(n_enc)
    enc_asig = []
    for _, row in df_g.iterrows():
        em = int(np.argmin(cargas))
        enc_asig.append(em + 1)
        cargas[em] += row['carga_pond']
    df_g['encuestador'] = enc_asig

    # ── Paso 2: asignación de días por target ────────────────────────────────
    n_rows      = len(df_g)
    dia_ini_col = [inicio_dia] * n_rows
    dia_fin_col = [inicio_dia] * n_rows

    for enc_id in range(1, n_enc + 1):
        idx_enc = df_g[df_g['encuestador'] == enc_id].index.tolist()
        if not idx_enc:
            continue

        dia_cursor = inicio_dia   # día en que empieza la manzana actual
        viv_acum   = 0.0          # viviendas acumuladas en el día actual

        for idx in idx_enc:
            loc   = df_g.index.get_loc(idx)
            viv_m = max(0.0, float(df_g.iloc[loc]['viv']))

            if viv_m > viv_max:
                # ── MANZANA GRANDE ──
                # Necesita múltiples días consecutivos.
                # ceil(viv / target) días, clampeado al rango disponible.
                # El encuestador trabaja esta manzana días seguidos;
                # sus compañeros van a otras manzanas en paralelo.
                # Si deseas que sea aún más agresivo (más días), baja viv_max.
                dias_m    = max(1, int(np.ceil(viv_m / target)))
                dias_m    = min(dias_m, ultimo - dia_cursor + 1)
                d_ini     = dia_cursor
                d_fin     = min(d_ini + dias_m - 1, ultimo)
                dia_cursor = d_fin + 1       # siguiente manzana empieza después
                viv_acum  = 0.0
            else:
                # ── MANZANA NORMAL ──
                # Va al día actual. Si acumular supera el target, avanza día.
                d_ini    = min(dia_cursor, ultimo)
                d_fin    = d_ini
                viv_acum += viv_m
                if viv_acum >= target and dia_cursor < ultimo:
                    dia_cursor += 1
                    viv_acum   = 0.0

            # Clamp de seguridad final
            d_ini = max(inicio_dia, min(d_ini, ultimo))
            d_fin = max(d_ini,      min(d_fin, ultimo))

            dia_ini_col[loc] = d_ini
            dia_fin_col[loc] = d_fin

    df_g['dia_inicio']    = dia_ini_col
    df_g['dia_fin']       = dia_fin_col
    df_g['dia_operativo'] = dia_ini_col  # retrocompatibilidad
    return df_g


def generar_excel(df_plan, eq_cfg, personal_info,
                  fecha_j1, fecha_j2, dias_op, j1_num, j2_num, mes_nombre):
    """
    Genera Excel con dos hojas (una por jornada), cada una con su número
    de jornada correcto (j1_num, j2_num) y su fecha de inicio independiente.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Estilos ──
    AZ_OSCURO = "0D3B6E"; AZ_MEDIO="1A5276"; AZ_CLARO="D6EAF8"
    VRD_CHECK  = "D5F5E3"; GRIS="F2F3F4"; BLANCO="FFFFFF"
    # Paletas de color por encuestador (rotativas)
    # Cada encuestador tiene su propia familia de colores para identificarlo visualmente
    ENC_PALETAS = [
        {"par": "DBEAFE", "impar": "EFF6FF", "subtot": "BFDBFE", "hdr": "1D4ED8"},  # azul
        {"par": "D1FAE5", "impar": "ECFDF5", "subtot": "A7F3D0", "hdr": "065F46"},  # verde
        {"par": "FEF9C3", "impar": "FEFCE8", "subtot": "FDE68A", "hdr": "854D0E"},  # amarillo
        {"par": "FCE7F3", "impar": "FDF4FF", "subtot": "F9A8D4", "hdr": "831843"},  # rosa
        {"par": "FFE4E6", "impar": "FFF1F2", "subtot": "FECACA", "hdr": "9F1239"},  # rojo
        {"par": "E0E7FF", "impar": "EEF2FF", "subtot": "C7D2FE", "hdr": "3730A3"},  # índigo
    ]

    def sc(cell, bold=False, bg=None, fg="000000",
           ha="left", sz=9, brd=False, wrap=False, italic=False):
        """Shortcut para estilizar una celda."""
        cell.font      = Font(bold=bold, size=sz, color=fg, italic=italic)
        cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        if brd:
            t = Side(style='thin')
            cell.border = Border(left=t, right=t, top=t, bottom=t)

    ct_counter = [700]  # CT global, empieza en CT700

    # Iteramos con el número de jornada REAL de cada hoja
    # j1_num y j2_num se calculan automáticamente desde el mes seleccionado
    for jornada_nombre, fecha_inicio, n_jornada_hoja in [
        ("Jornada 1", fecha_j1, j1_num),
        ("Jornada 2", fecha_j2, j2_num)
    ]:
        df_jor = df_plan[df_plan['jornada'] == jornada_nombre].copy()
        if len(df_jor) == 0:
            continue

        ws = wb.create_sheet(title=jornada_nombre)
        ws.sheet_view.showGridLines = False

        # Anchos de columnas fijas (A=1 … N=14)
        anchos = {'A':14,'B':14,'C':8,'D':5,'E':5,'F':6,
                  'G':6,'H':6,'I':5,'J':18,'K':13,'L':10,'M':14,'N':5}
        for col_l, w in anchos.items():
            ws.column_dimensions[col_l].width = w
        # Columnas de fecha (O en adelante)
        for i in range(dias_op):
            ws.column_dimensions[get_column_letter(15+i)].width = 7
        # Columna # VIV
        ws.column_dimensions[get_column_letter(15+dias_op)].width = 6

        cur = 1  # fila actual

        equipos_jor = [e['nombre'] for e in eq_cfg
                       if e['nombre'] in df_jor['equipo'].values]

        for grupo_num, nombre_eq in enumerate(equipos_jor, 1):
            df_eq = df_jor[df_jor['equipo'] == nombre_eq].copy()
            if len(df_eq) == 0: continue

            pi    = personal_info.get(nombre_eq, {})
            n_enc = next((e['enc'] for e in eq_cfg if e['nombre']==nombre_eq), 3)

            # Fechas
            if fecha_inicio:
                fechas      = [fecha_inicio + timedelta(days=i) for i in range(dias_op)]
                fi_str      = fecha_inicio.strftime("%d-%b-%y").upper()
                ff_str      = fechas[-1].strftime("%d-%b-%y").upper()
            else:
                fechas      = None
                fi_str      = "____"
                ff_str      = "____"

            last_col_idx = 15 + dias_op  # columna # viv

            def merge_row(row, c1, c2, val, **kw):
                ws.merge_cells(f'{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}')
                c = ws.cell(row, c1, val)
                sc(c, **kw)
                return c

            # ── Encabezado institucional ──────────────────
            for txt in ["INSTITUTO NACIONAL DE ESTADÍSTICA Y CENSOS",
                        "COORDINACIÓN ZONAL LITORAL CZ8L",
                        "ACTUALIZACIÓN CARTOGRÁFICA - ENDI ENLISTAMIENTO",
                        "PROGRAMACIÓN OPERATIVO DE CAMPO"]:
                merge_row(cur,1,last_col_idx,txt,bold=True,bg=AZ_OSCURO,
                          fg=BLANCO,ha="center",sz=9)
                cur += 1
            cur += 1

            # JORNADA / GRUPO — número de jornada real de esta hoja
            ws.cell(cur,1,"JORNADA")
            jorn_cell = ws.cell(cur,2,str(n_jornada_hoja))
            jorn_cell.font = Font(bold=True,size=11)
            ws.cell(cur,7,"GRUPO")
            ws.cell(cur,9,str(grupo_num)).font = Font(bold=True,size=11)
            sc(ws.cell(cur,1),bold=True,sz=10)
            sc(ws.cell(cur,7),bold=True,sz=10)
            cur += 2

            # Período
            ws.cell(cur,1,"PERÍODO DE ACTUALIZACIÓN:")
            sc(ws.cell(cur,1),bold=True,sz=9)
            ws.cell(cur,5,"DEL"); ws.cell(cur,6,fi_str)
            ws.cell(cur,9,"AL"); ws.cell(cur,10,ff_str)
            cur += 2

            # Cabecera de personal
            for col,txt in [(3,"COD."),(4,"NOMBRE"),(8,"No. CÉDULA"),(11,"No. CELULAR")]:
                sc(ws.cell(cur,col,txt),bold=True,sz=8)
            cur += 1

            # Supervisor
            ws.cell(cur,1,"SUPERVISOR:")
            ws.cell(cur,3,pi.get('supervisor_cod',''))
            ws.cell(cur,4,pi.get('supervisor_nombre',''))
            ws.cell(cur,8,pi.get('supervisor_cedula',''))
            ws.cell(cur,11,pi.get('supervisor_celular',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 2

            # Encuestadores
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

            # Vehículo / Chofer
            ws.cell(cur,1,"VEHÍCULO: CHOFER")
            ws.cell(cur,4,pi.get('chofer_nombre',''))
            ws.cell(cur,8,pi.get('chofer_cedula',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 1
            ws.cell(cur,1,"PLACA:")
            ws.cell(cur,4,pi.get('placa',''))
            sc(ws.cell(cur,1),bold=True,sz=9)
            cur += 2

            # ── Encabezado de tabla ──────────────────────
            # Fila 1: secciones principales
            merge_row(cur,1,4,"EQUIPO",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            merge_row(cur,5,14,"IDENTIFICACIÓN",bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            fecha_hdr = "RECORRIDO DE LOS SECTORES EN LA JORNADA — FECHA"
            merge_row(cur,15,14+dias_op,fecha_hdr,bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=7,brd=True)
            sc(ws.cell(cur,last_col_idx,"# VIV"),bold=True,bg=AZ_MEDIO,fg=BLANCO,ha="center",sz=8,brd=True)
            ws.row_dimensions[cur].height = 24
            cur += 1

            # Fila 2: columnas detalle
            sub_hdrs = ["SUPERVISOR","ENCUESTADOR","CARGA DE TRABAJO",
                        "PROV","CANTON","CIUDAD O PARROQ","ZONA","SECTOR","MAN",
                        "CÓDIGO DE LA JURISDICCIÓN","PROVINCIA","CANTÓN",
                        "CIUDAD, PARROQ. O LOC AMAZ.","NRO EDIF"]
            for ci,h in enumerate(sub_hdrs,1):
                sc(ws.cell(cur,ci,h),bold=True,bg=AZ_CLARO,ha="center",
                   sz=7,brd=True,wrap=True)

            for i in range(dias_op):
                lbl = fechas[i].strftime("%d/%m") if fechas else f"Día {i+1}"
                sc(ws.cell(cur,15+i,lbl),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)

            sc(ws.cell(cur,last_col_idx,"# VIV"),bold=True,bg=AZ_CLARO,ha="center",sz=7,brd=True)
            ws.row_dimensions[cur].height = 32
            cur += 1

            # ── Filas de datos agrupadas por encuestador ─
            # Ordenamos por encuestador primero, luego por dia_inicio
            df_sorted = df_eq.sort_values(['encuestador','dia_inicio']).copy()

            enc_actual  = None      # encuestador en curso
            fila_enc    = 0         # contador de fila dentro del encuestador
            viv_enc_acum = 0        # viviendas acumuladas para el subtotal
            enc_color_idx = -1      # índice de paleta del encuestador actual

            for ri, (_, rd) in enumerate(df_sorted.iterrows()):
                enc_id = int(rd.get('encuestador', 0))

                # ¿Cambiamos de encuestador? → insertar fila de subtotal del anterior
                if enc_id != enc_actual and enc_actual is not None:
                    pal_sub = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                    bg_sub  = pal_sub["subtot"]
                    fg_sub  = pal_sub["hdr"]
                    enc_info_prev = enc_list[enc_actual-1] if 0 < enc_actual <= len(enc_list) else {}
                    # Fila separadora / subtotal
                    merge_row(cur, 1, 9,
                              f"SUBTOTAL {enc_info_prev.get('nombre', f'Encuestador {enc_actual}')}",
                              bold=True, bg=bg_sub, fg=fg_sub, ha="right", sz=8)
                    for ci in range(10, last_col_idx):
                        sc(ws.cell(cur, ci, ""), bg=bg_sub, brd=True)
                    sc(ws.cell(cur, last_col_idx, viv_enc_acum),
                       bold=True, ha="center", sz=9, bg=bg_sub, fg=fg_sub, brd=True)
                    ws.row_dimensions[cur].height = 14
                    cur += 1
                    viv_enc_acum = 0

                # Actualizar encuestador actual
                if enc_id != enc_actual:
                    enc_actual    = enc_id
                    fila_enc      = 0
                    enc_color_idx = (enc_color_idx + 1) % len(ENC_PALETAS)

                pal      = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                bg_row   = pal["par"] if fila_enc % 2 == 0 else pal["impar"]
                fila_enc += 1
                viv_enc_acum += int(rd.get('viv', 0))

                p_cod    = parse_codigo(str(rd['id_entidad']))
                enc_i    = enc_list[enc_id-1] if 0 < enc_id <= len(enc_list) else {}
                ct_str   = f"CT{ct_counter[0]:03d}"
                ct_counter[0] += 1

                row_vals = [
                    pi.get('supervisor_cedula', ''),
                    enc_i.get('cedula', ''),
                    ct_str,
                    p_cod['prov'], p_cod['canton'],
                    p_cod['ciudad_parroq'],
                    p_cod['zona'], p_cod['sector'], p_cod['man'],
                    str(rd['id_entidad']),
                    '', '', '', '',
                ]
                for ci, val in enumerate(row_vals, 1):
                    c = ws.cell(cur, ci, val)
                    sc(c, bg=AZ_CLARO if ci == 10 else bg_row,
                       ha="center", sz=8, brd=True)

                d_ini = int(rd.get('dia_inicio', rd.get('dia_operativo', 1)))
                d_fin = int(rd.get('dia_fin', d_ini))
                for i in range(dias_op):
                    dia_num = i + 1
                    in_rng  = (d_ini <= dia_num <= d_fin)
                    if in_rng:
                        c = ws.cell(cur, 15 + i, "✓")
                        sc(c, bold=True, bg=VRD_CHECK, ha="center", sz=11, brd=True)
                    else:
                        sc(ws.cell(cur, 15 + i, ""), bg=bg_row, ha="center", brd=True)

                sc(ws.cell(cur, last_col_idx, int(rd.get('viv', 0))),
                   ha="center", sz=8, brd=True, bg=bg_row)
                cur += 1

            # Subtotal del último encuestador
            if enc_actual is not None:
                pal_sub  = ENC_PALETAS[enc_color_idx % len(ENC_PALETAS)]
                enc_info = enc_list[enc_actual-1] if 0 < enc_actual <= len(enc_list) else {}
                merge_row(cur, 1, 9,
                          f"SUBTOTAL {enc_info.get('nombre', f'Encuestador {enc_actual}')}",
                          bold=True, bg=pal_sub["subtot"], fg=pal_sub["hdr"], ha="right", sz=8)
                for ci in range(10, last_col_idx):
                    sc(ws.cell(cur, ci, ""), bg=pal_sub["subtot"], brd=True)
                sc(ws.cell(cur, last_col_idx, viv_enc_acum),
                   bold=True, ha="center", sz=9,
                   bg=pal_sub["subtot"], fg=pal_sub["hdr"], brd=True)
                ws.row_dimensions[cur].height = 14
                cur += 1

            # Total de viviendas del equipo
            tot_viv_eq = int(df_eq['viv'].sum())
            sc(ws.cell(cur, last_col_idx-1, "TOTAL"),
               bold=True, ha="right", sz=8, bg=AZ_CLARO, brd=True)
            sc(ws.cell(cur, last_col_idx, tot_viv_eq),
               bold=True, ha="center", sz=8, bg=AZ_CLARO, brd=True)
            cur += 4  # espacio entre grupos

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
    "params": {"dias_op":12,"viv_min":50,"viv_max":80,"factor_r":1.5,
               "usar_bomb":True,"usar_gye":True,"dias_gye":3,"umbral_gye":10},
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

    # PASO 1 — GeoPackage
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

    # PASO 2 — GraphML
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
        # PASO 3 — Mes
        st.markdown("<div class='step'>PASO 3</div>", unsafe_allow_html=True)
        st.markdown("**Mes operativo**")
        meses_disp = sorted(st.session_state.data_raw["mes"].dropna().unique().tolist())
        mes_sel = st.selectbox("Mes", meses_disp,
            format_func=lambda x: f"{MESES_N.get(int(x),str(int(x)))} (mes {int(x)})")
        df_mes = st.session_state.data_raw[st.session_state.data_raw["mes"]==mes_sel].copy()
        st.session_state.data_mes = df_mes

        st.divider()

        # PASO 4 — Equipos
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
        st.markdown("**Parámetros**")
        p = st.session_state.params
        p["dias_op"]  = st.slider("Días operativos",10,14,p["dias_op"])
        p["viv_min"]  = st.slider("Mín viv/día",30,60,p["viv_min"])
        p["viv_max"]  = st.slider("Máx viv/día",60,120,p["viv_max"])
        p["factor_r"] = st.slider("Factor rural (×)",1.0,2.5,p["factor_r"],0.1,
            help="Viviendas dispersas pesan X veces más para el balance. "
                 "No cambia la cantidad real a visitar, solo la asignación interna.")
        p["usar_bomb"] = st.toggle("Equipo Bombero",value=p["usar_bomb"],
            help="Detecta UPMs outliers DENTRO de cada cluster y las asigna a un equipo especial.")
        if p["usar_bomb"]:
            p["min_dist_bomb_m"] = st.slider(
                "Distancia mín. Bombero (km)", 10, 150,
                p.get("min_dist_bomb_m", 40000) // 1000) * 1000
            st.caption("Solo va al Bombero si además supera esta distancia al centroide del cluster.")
        p["usar_gye"]  = st.toggle("Restricción Guayaquil",value=p["usar_gye"])
        p["dias_gye"]  = st.slider("Días GYE",1,5,p["dias_gye"],disabled=not p["usar_gye"])
        p["umbral_gye"]= st.slider("Umbral GYE (%)",5,30,p["umbral_gye"],disabled=not p["usar_gye"])

        # ── Número de jornada: auto-calculado desde el mes ───────────────────
        # Cada mes tiene 2 jornadas. Mes 1→J1+J2, Mes 2→J3+J4, Mes N→J(2N-1)+J(2N).
        # Si el mes 1 del gpkg no corresponde a la Jornada 1 del calendario INEC,
        # ajustar OFFSET_JORNADA aquí (por ahora = 0).
        # OFFSET_JORNADA = 0
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
  <h1>Planificación Automática · Actualización Cartográfica</h1>
  <p>Encuesta Nacional &nbsp;·&nbsp; Zonal Litoral &nbsp;·&nbsp; INEC Ecuador</p>
</div>""", unsafe_allow_html=True)

if st.session_state.data_raw is None:
    st.markdown("<div class='ibox'>👈 Carga el <code>.gpkg</code> desde el panel lateral.</div>",
                unsafe_allow_html=True)
    st.stop()

df = st.session_state.data_mes
if df is None or len(df)==0:
    st.warning("Sin datos para el mes seleccionado."); st.stop()

# KPIs
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

# ── BOTÓN GENERAR ─────────────────────────────
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

# ═══════════════════════════════════════════════
#  ALGORITMO PRINCIPAL
# ═══════════════════════════════════════════════
if btn:
    G         = st.session_state.graph_G
    eq_cfg    = st.session_state.equipos_cfg
    n_eq      = len(eq_cfg)
    nombres   = [e["nombre"] for e in eq_cfg]
    n_clust   = n_eq * 2
    p         = st.session_state.params

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

    # 1. Distancias a la base (para referencia y GYE)
    prog.progress(8,"Calculando distancias...")
    t_utm = Transformer.from_crs("EPSG:4326","EPSG:32717",always_xy=True)
    bx,by = t_utm.transform(BASE_LON,BASE_LAT)
    df_w['dist_base_m'] = np.sqrt((df_w['x']-bx)**2+(df_w['y']-by)**2)

    # 2. Restricción Guayaquil
    prog.progress(12,"Verificando restricción Guayaquil...")
    upms_gye = pd.Series(False,index=df_w.index)
    if p["usar_gye"] and 'pro_x' in df_w.columns and 'can_x' in df_w.columns:
        upms_gye = (df_w['pro_x']==PRO_GYE)&(df_w['can_x']==CAN_GYE)
    pct_gye   = upms_gye.sum()/len(df_w) if len(df_w)>0 else 0
    act_gye   = p["usar_gye"] and (pct_gye >= p["umbral_gye"]/100) and upms_gye.sum()>0

    df_gye    = df_w[upms_gye].copy()    if act_gye else pd.DataFrame()
    df_no_gye = df_w[~upms_gye].copy()

    # 3. Clustering
    prog.progress(22,f"Generando {n_clust} clusters...")
    mask_bomb_global = pd.Series(False,index=df_w.index)

    if len(df_no_gye) >= n_clust:
        coords = df_no_gye[['x','y']].values
        km = KMeans(n_clusters=n_clust,init='k-means++',n_init=20,
                    max_iter=500,random_state=42)
        df_no_gye = df_no_gye.copy()
        df_no_gye['cluster_geo'] = km.fit_predict(coords)

        if len(df_no_gye) > n_clust:
            try: st.session_state.sil_score = silhouette_score(coords,df_no_gye['cluster_geo'])
            except: st.session_state.sil_score = None

        centroides = km.cluster_centers_
        dist_c = np.sqrt((centroides[:,0]-bx)**2+(centroides[:,1]-by)**2)
        orden  = np.argsort(dist_c)[::-1]
        asig   = {}
        for i,(cj1,cj2) in enumerate(zip(orden[:n_eq],orden[n_eq:])):
            asig[cj1]=(nombres[i],'Jornada 1')
            asig[cj2]=(nombres[i],'Jornada 2')

        df_no_gye['equipo']  = df_no_gye['cluster_geo'].map(lambda c: asig[c][0])
        df_no_gye['jornada'] = df_no_gye['cluster_geo'].map(lambda c: asig[c][1])

        # EQUIPO BOMBERO POR CLUSTER (v5 — menos agresivo):
        # Usamos 3×IQR en vez de 1.5×IQR, Y añadimos un umbral mínimo
        # absoluto de distancia al centroide. Así evitamos marcar como
        # "bombero" puntos que son estadísticamente outliers pero en la
        # práctica están a sólo unos kilómetros del resto del cluster.
        # El parámetro MIN_DIST_BOMBERO se puede subir/bajar según criterio.
        if p["usar_bomb"]:
            prog.progress(30,"Detectando outliers por cluster (Equipo Bombero)...")
            MIN_DIST_BOMBERO_M = p.get("min_dist_bomb_m", 40000)  # 40 km por defecto
            for c_id in range(n_clust):
                if c_id not in asig: continue
                mask_c = df_no_gye['cluster_geo'] == c_id
                pts = df_no_gye[mask_c]
                # Necesitamos al menos 8 puntos para que el IQR sea significativo
                if len(pts) < 8: continue

                cx, cy = centroides[c_id]
                dists = np.sqrt((pts['x'] - cx)**2 + (pts['y'] - cy)**2)
                Q1c, Q3c = dists.quantile(.25), dists.quantile(.75)
                iqrc = Q3c - Q1c
                if iqrc == 0: continue  # cluster compacto, sin outliers

                # 3×IQR (en vez de 1.5×) = criterio mucho más permisivo
                umbral_iqr = Q3c + 3.0 * iqrc

                # Condición DOBLE: outlier estadístico Y suficientemente lejos
                bomb_cands = dists[(dists > umbral_iqr) & (dists > MIN_DIST_BOMBERO_M)]
                bomb_idx   = bomb_cands.index

                if len(bomb_idx) > 0:
                    df_no_gye.loc[bomb_idx, 'equipo']  = 'Equipo Bombero'
                    df_no_gye.loc[bomb_idx, 'jornada'] = 'Jornada Especial'
                    mask_bomb_global.loc[bomb_idx] = True

        df_w.update(df_no_gye[['equipo','jornada','cluster_geo']])

    n_bomb = int((df_w['equipo']=='Equipo Bombero').sum())
    st.session_state.n_bombero = n_bomb

    # 4. Encuestadores + días
    # ─────────────────────────────────────────────────────────────────────────
    # CORRECCIÓN INICIO_DIA:
    # La restricción de Guayaquil (días_gye primeros días en GYE) SOLO aplica
    # a la Jornada 1. La Jornada 2 es un período operativo COMPLETAMENTE
    # SEPARADO de 12 días que siempre empieza en el día 1.
    #
    # Bug anterior: inicio = dias_gye+1 se aplicaba a AMBAS jornadas.
    # Resultado: Jornada 2 tenía días 1-3 vacíos porque empezaba en día 4.
    #
    # Si en el futuro la Jornada 2 también tiene restricción GYE,
    # cambiar la condición a: act_gye and jornada in ['Jornada 1','Jornada 2']
    prog.progress(42,"Asignando encuestadores y distribuyendo días...")
    enc_dict = {e["nombre"]:e["enc"] for e in eq_cfg}

    for nombre_eq in nombres:
        for jornada in ['Jornada 1','Jornada 2']:
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp = df_w[mask_g].copy()
            if len(grp)==0: continue
            n_enc = enc_dict.get(nombre_eq,3)

            # Jornada 2 siempre empieza en día 1 (período separado)
            # Jornada 1 empieza después de los días GYE si la restricción está activa
            if jornada == 'Jornada 1' and act_gye:
                inicio    = p["dias_gye"] + 1
                dias_disp = p["dias_op"] - p["dias_gye"]
            else:
                inicio    = 1
                dias_disp = p["dias_op"]

            ga = asignar_encuestadores_y_dias(grp, n_enc, dias_disp,
                                               p["viv_min"], p["viv_max"], inicio)
            df_w.update(ga[['encuestador','dia_operativo','dia_inicio','dia_fin']])

    # Fase Guayaquil
    if act_gye and len(df_gye)>0:
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
    tsp_r,road_p = {},{}

    for ri,nombre_eq in enumerate(nombres):
        for jornada in ['Jornada 1','Jornada 2']:
            pct = 52+int((ri*2+['Jornada 1','Jornada 2'].index(jornada)+1)/(n_eq*2)*42)
            prog.progress(pct,f"TSP: {nombre_eq} | {jornada}...")
            mask_g = (df_w['equipo']==nombre_eq)&(df_w['jornada']==jornada)
            grp = df_w[mask_g]
            if len(grp)==0: continue
            nr = ox.nearest_nodes(G,grp['lon'].values,grp['lat'].values)
            nk = [n for n in nr if n in comp_base]
            if not nk: continue
            nu = [base_nd]+list(dict.fromkeys(nk))
            n  = len(nu)
            if n<=2: continue
            D = np.zeros((n,n))
            for i in range(n):
                for j in range(i+1,n):
                    try: d=nx.shortest_path_length(G,nu[i],nu[j],weight='length'); D[i,j]=D[j,i]=d
                    except: D[i,j]=D[j,i]=1e9
            Gt=nx.Graph()
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
            clave=f"{nombre_eq}||{jornada}"
            tsp_r[clave]={'equipo':nombre_eq,'jornada':jornada,
                          'n_puntos':len(grp),'dist_km':dist/1000}
            road_p[clave]=ruta

    prog.progress(98,"Calculando métricas...")

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
    st.session_state.df_plan  = df_w
    st.session_state.tsp_results = tsp_r
    st.session_state.road_paths  = road_p
    st.session_state.resumen_bal = resumen_bal
    st.session_state.resultados_generados = True
    st.success("✓ Planificación generada.")

# ── RESULTADOS ────────────────────────────────
if not st.session_state.resultados_generados:
    st.markdown("<div class='ibox'>👆 Presiona <b>Generar Planificación</b>.</div>",
                unsafe_allow_html=True)
    st.stop()

df_plan   = st.session_state.df_plan
tsp_r     = st.session_state.tsp_results
road_p    = st.session_state.road_paths
res_bal   = st.session_state.resumen_bal
eq_cfg    = st.session_state.equipos_cfg
nombres   = [e["nombre"] for e in eq_cfg]
p         = st.session_state.params

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
                    f"<b>Carga pond.:</b> {row.get('carga_pond',0):.0f}<br>"
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

        st_folium(m,width=None,height=540,returned_objects=[],key="mapa_v3")

# ══ TAB 2 — ANÁLISIS ══════════════════════════
with tab_analisis:
    st.markdown("<div class='stitle'>Análisis Estadístico de Carga</div>",unsafe_allow_html=True)

    with st.expander("ℹ️ Carga real vs carga ponderada — ¿qué significan y cuál usar?"):
        st.markdown(f"""
        **Carga real (viviendas):** casas del precenso 2020 que el encuestador visita.
        Es el número final que aparece en el reporte INEC.

        **Carga ponderada:** número *ficticio* que usa el algoritmo internamente para asignar
        de forma equitativa. No aparece en el operativo; solo sirve para balancear.

        **¿Por qué hay diferencia?**
        Visitar 50 casas urbanas ≈ 4 horas. Visitar 50 casas dispersas rurales ≈ 7-8 horas
        por los desplazamientos entre viviendas. Con factor rural **{p['factor_r']}×**, una
        casa rural "pesa" {p['factor_r']} veces más en el balance. Resultado: el encuestador
        rural recibe *menos* viviendas reales, compensando el mayor tiempo de trabajo.

        **¿Cuál mira el supervisor?** Solo la columna *Viviendas reales* y el cronograma de días.
        El CV que mide equidad entre equipos es el de **carga ponderada**.
        """)

    # Métricas de clustering
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
        st.markdown(f"<div class='kcard'><div class='v'>{len(eq_cfg)*2}</div>"
                    f"<div class='l'>Clusters</div>"
                    f"<div class='s'>{len(eq_cfg)} eq × 2 jornadas</div></div>",
                    unsafe_allow_html=True)
    with mc3:
        bc = "#9b59b6" if n_bm>0 else "#445566"
        st.markdown(f"<div class='kcard'><div class='v' style='color:{bc}'>{n_bm}</div>"
                    f"<div class='l'>UPMs Bombero</div>"
                    f"<div class='s'>{'outliers por cluster' if n_bm>0 else 'ninguno'}</div></div>",
                    unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)

    # CV entre equipos (recalculado desde df_plan para robustez)
    st.markdown("<div class='stitle'>Equidad entre equipos</div>",unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    <b>CV viviendas reales</b> = dispersión observable en campo.<br>
    <b>CV carga ponderada</b> = criterio interno de equidad (incluye factor rural).<br>
    CV &lt;20% muy bueno · 20–40% aceptable · &gt;40% revisar configuración.
    </div>""", unsafe_allow_html=True)

    df_main = df_plan[~df_plan['equipo'].isin(['Equipo Bombero','sin_asignar'])].copy()
    res_cv  = df_main.groupby(['equipo','jornada']).agg(
        viv_reales=('viv','sum'), carga_ponderada=('carga_pond','sum')).reset_index()

    for jornada in ['Jornada 1','Jornada 2']:
        sub = res_cv[res_cv['jornada']==jornada]
        if len(sub)==0:
            st.markdown(f"<div class='ibox'><b>{jornada}:</b> sin UPMs.</div>",
                        unsafe_allow_html=True); continue
        if len(sub)==1:
            st.markdown(f"<div class='ibox'><b>{jornada}:</b> 1 equipo — CV no aplica.</div>",
                        unsafe_allow_html=True); continue
        cr = cv_pct(sub['viv_reales'])
        cp = cv_pct(sub['carga_ponderada'])
        ccr="#27ae60" if cr<20 else ("#f39c12" if cr<40 else "#e74c3c")
        ccp="#27ae60" if cp<20 else ("#f39c12" if cp<40 else "#e74c3c")
        em = "✓" if cp<20 else ("⚠" if cp<40 else "✗")
        st.markdown(f"""<div class='ibox'>
        <b>{jornada}</b><br>
        &nbsp;&nbsp;CV viviendas reales: <span style='color:{ccr};font-family:monospace;
        font-weight:600'>{cr:.1f}%</span>
        <i style='font-size:11px;color:#445566'> (encuestador en campo)</i><br>
        &nbsp;&nbsp;CV carga ponderada:&nbsp; <span style='color:{ccp};font-family:monospace;
        font-weight:600'>{cp:.1f}%</span>
        &nbsp;{em} <b>{'Muy bueno' if cp<20 else ('Aceptable' if cp<40 else 'Revisar')}</b>
        <i style='font-size:11px;color:#445566'> (criterio de equidad)</i>
        </div>""", unsafe_allow_html=True)

    with st.expander("Ver tabla de balance"):
        st.dataframe(res_cv.rename(columns={
            'equipo':'Equipo','jornada':'Jornada',
            'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.'
        }),use_container_width=True)

    # Tarjetas de equipos (horizontal)
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

    # Distribución por días — FILTRABLE POR JORNADA
    st.markdown("<div class='stitle'>Distribución diaria de viviendas</div>",
                unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    Viviendas distribuidas por día (manzanas grandes se reparten proporcionalmente
    entre sus días de duración). Los días del eje son días de la jornada, no del mes.
    Jornada 1 = días lejanos primero. Jornada 2 = días cercanos.
    </div>""", unsafe_allow_html=True)

    # Filtro con índice explícito para evitar el bug del radio
    jor_opts   = ["Jornada 1", "Jornada 2", "Ambas"]
    jor_idx    = st.radio("Filtrar:", jor_opts, horizontal=True,
                          key="radio_jornada_dias_v5",
                          index=0)
    jor_filtro = jor_opts[jor_opts.index(jor_idx)]

    pivot_all = df_plan[df_plan['equipo'].isin(eq_act)].copy()
    if jor_filtro != "Ambas":
        pivot_all = pivot_all[pivot_all['jornada'] == jor_filtro]

    # Normalizamos el eje X por jornada de forma INDEPENDIENTE.
    # Cada jornada tiene su propio "día 1". Si mezclamos ambas y Jornada 1
    # empieza en día 1 mientras Jornada 2 también empieza en día 1 (tras el fix),
    # la normalización ya es correcta. Pero si hubiera un offset (ej. GYE en J1),
    # la normalización por jornada evita que días 1-3 aparezcan vacíos en el eje.
    rows_exp = []
    for _, row in pivot_all.iterrows():
        d_ini = int(row.get('dia_inicio', row.get('dia_operativo', 1)))
        d_fin = int(row.get('dia_fin', d_ini))
        dias_dur = max(1, d_fin - d_ini + 1)
        viv_d    = row['viv'] / dias_dur
        # dia_rel: relativo al primer día DENTRO de la misma jornada
        # Con el fix de inicio_dia, J2 ya empieza en 1, así que esto no cambia nada.
        # Si en el futuro J1 tuviera fase GYE que empiece en día 4, el gráfico
        # mostraría esos días como 4,5... Para normalizar basta con restar d_min
        # pero SOLO dentro de la misma jornada.
        for dd in range(d_ini, d_fin + 1):
            rows_exp.append({
                'equipo'      : row['equipo'],
                'jornada'     : row['jornada'],
                'encuestador' : int(row.get('encuestador', 0)),
                'dia_abs'     : dd,
                'viv'         : viv_d
            })

    if len(rows_exp) > 0:
        df_exp = pd.DataFrame(rows_exp)

        # Normalizar día relativo POR JORNADA (para que siempre empiece en 1)
        if jor_filtro != "Ambas":
            d_min_jor = df_exp['dia_abs'].min()
            df_exp['dia_rel'] = df_exp['dia_abs'] - d_min_jor + 1
        else:
            # Para "Ambas" cada jornada se normaliza independientemente a 1..N
            # Usamos transform dentro de cada jornada
            df_exp['dia_rel'] = df_exp.groupby('jornada')['dia_abs'].transform(
                lambda s: s - s.min() + 1
            ).astype(int)

        pivot   = df_exp.groupby(['equipo', 'dia_rel'])['viv'].sum().reset_index()
        fig_d   = px.bar(pivot, x='dia_rel', y='viv', color='equipo',
                         barmode='group',
                         title=f'Viviendas por día operativo — {jor_filtro}',
                         labels={'dia_rel': 'Día', 'viv': 'Viviendas', 'equipo': 'Equipo'},
                         template='plotly_dark',
                         color_discrete_map=color_map)

        # Líneas de referencia por encuestador promedio
        tot_enc_f = sum(e["enc"] for e in eq_cfg if e["nombre"] in eq_act)
        n_eq_act  = max(1, len(eq_act))
        avg_enc_f = tot_enc_f / n_eq_act
        fig_d.add_hline(y=p["viv_min"] * avg_enc_f, line_dash="dot",
                        line_color="#f39c12",
                        annotation_text=f"Mín referencia ({p['viv_min']} viv/enc)")
        fig_d.add_hline(y=p["viv_max"] * avg_enc_f, line_dash="dot",
                        line_color="#e74c3c",
                        annotation_text=f"Máx referencia ({p['viv_max']} viv/enc)")
        fig_d.update_layout(
            paper_bgcolor="#111827", plot_bgcolor="#0a1020",
            xaxis=dict(dtick=1, title="Día de la jornada"),
            yaxis_title="Viviendas"
        )
        st.plotly_chart(fig_d, use_container_width=True)

        # Gráfico de carga por encuestador — solo cuando se filtra una jornada
        # (en modo "Ambas" los encuestadores se mezclan y el gráfico pierde sentido)
        # Para ver ambas jornadas por encuestador usa el drilldown de equipos arriba.
        if jor_filtro != "Ambas":
            pivot_enc = df_exp.groupby(['encuestador','dia_rel'])['viv'].sum().reset_index()
            pivot_enc['encuestador'] = "Enc. " + pivot_enc['encuestador'].astype(str)
            fig_enc = px.line(pivot_enc, x='dia_rel', y='viv', color='encuestador',
                              markers=True,
                              title=f'Carga diaria por encuestador — {jor_filtro}',
                              labels={'dia_rel':'Día','viv':'Viviendas','encuestador':''},
                              template='plotly_dark')
            fig_enc.add_hline(y=p["viv_min"], line_dash="dot", line_color="#f39c12",
                              annotation_text=f"Mín {p['viv_min']}")
            fig_enc.add_hline(y=p["viv_max"], line_dash="dot", line_color="#e74c3c",
                              annotation_text=f"Máx {p['viv_max']}")
            fig_enc.update_layout(paper_bgcolor="#111827", plot_bgcolor="#0a1020",
                                  xaxis=dict(dtick=1))
            st.plotly_chart(fig_enc, use_container_width=True)

    # Equipo Bombero
    df_bm = df_plan[df_plan['equipo']=='Equipo Bombero']
    n_bm  = st.session_state.n_bombero
    st.markdown("<div class='stitle'>Equipo Bombero</div>",unsafe_allow_html=True)
    if n_bm==0:
        st.markdown("""<div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — 0 UPMs asignadas<br>
        <span style='font-size:12px;color:#7a5a9a'>
        Ningún punto resultó outlier dentro de su cluster en este mes.
        El equipo está disponible como contingencia operativa.
        </span></div>""",unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class='bcard'>
        <b style='color:#9b59b6'>Equipo Bombero</b> — {n_bm} UPMs<br>
        <span style='font-size:12px;color:#7a5a9a'>
        Estas UPMs son outliers dentro de su cluster geográfico (IQR de distancia al centroide).
        Viviendas: {int(df_bm['viv'].sum()):,}
        </span></div>""",unsafe_allow_html=True)
        st.dataframe(
            df_bm[['id_entidad','tipo_entidad','viv','lat','lon','dist_base_m']]
            .rename(columns={'dist_base_m':'Dist. base (m)'})
            .sort_values('Dist. base (m)',ascending=False).reset_index(drop=True),
            use_container_width=True,height=200)

# ══ TAB 3 — REPORTE Y DESCARGA ════════════════
with tab_reporte:
    st.markdown("<div class='stitle'>Reporte Mensual y Descarga Excel</div>",
                unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    El Excel generado replica el formato oficial INEC (Jornada 16, Grupos 1–N).
    Completa los datos de personal y las fechas de inicio antes de descargar.
    </div>""",unsafe_allow_html=True)

    # ── Fechas de jornada ─────────────────────
    st.markdown("<div class='stitle'>Fechas de las jornadas</div>",unsafe_allow_html=True)
    fc1,fc2 = st.columns(2)
    with fc1:
        fj1 = st.date_input("Fecha inicio Jornada 1",
                             value=st.session_state.fecha_j1 or date.today(),
                             key="fi_j1")
        st.session_state.fecha_j1 = fj1
        if fj1:
            fin_j1 = fj1 + timedelta(days=p["dias_op"]-1)
            st.caption(f"Fin: {fin_j1.strftime('%d/%m/%Y')} ({p['dias_op']} días)")
    with fc2:
        fj2 = st.date_input("Fecha inicio Jornada 2",
                             value=st.session_state.fecha_j2 or date.today(),
                             key="fi_j2")
        st.session_state.fecha_j2 = fj2
        if fj2:
            fin_j2 = fj2 + timedelta(days=p["dias_op"]-1)
            st.caption(f"Fin: {fin_j2.strftime('%d/%m/%Y')} ({p['dias_op']} días)")

    # ── Información de personal ───────────────
    st.markdown("<div class='stitle'>Personal por equipo</div>",unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    Ingresa nombres y cédulas. Los campos vacíos quedarán en blanco en el Excel para llenar manualmente.
    </div>""",unsafe_allow_html=True)

    for eq in eq_cfg:
        nombre_eq = eq["nombre"]
        n_enc_eq  = eq["enc"]
        pi_prev   = st.session_state.personal_info.get(nombre_eq, {})

        with st.expander(f"👥 {nombre_eq} — {n_enc_eq} encuestador(es)", expanded=False):
            st.markdown(f"<div class='pi-form'>",unsafe_allow_html=True)
            pc1,pc2,pc3 = st.columns(3)
            with pc1:
                sup_n = st.text_input("Supervisor (nombre)",
                                       value=pi_prev.get('supervisor_nombre',''),
                                       key=f"sup_n_{nombre_eq}")
            with pc2:
                sup_c = st.text_input("Cédula supervisor",
                                       value=pi_prev.get('supervisor_cedula',''),
                                       key=f"sup_c_{nombre_eq}")
            with pc3:
                sup_t = st.text_input("Celular supervisor",
                                       value=pi_prev.get('supervisor_celular',''),
                                       key=f"sup_t_{nombre_eq}")

            enc_list_new = []
            for j in range(n_enc_eq):
                prev_enc = (pi_prev.get('encuestadores',[{}]*n_enc_eq))
                prev_j   = prev_enc[j] if j < len(prev_enc) else {}
                pe1,pe2,pe3 = st.columns(3)
                with pe1:
                    en = st.text_input(f"Encuestador {j+1}",
                                        value=prev_j.get('nombre',''),
                                        key=f"enc_n_{nombre_eq}_{j}")
                with pe2:
                    ec = st.text_input(f"Cédula enc. {j+1}",
                                        value=prev_j.get('cedula',''),
                                        key=f"enc_c_{nombre_eq}_{j}")
                with pe3:
                    et = st.text_input(f"Celular enc. {j+1}",
                                        value=prev_j.get('celular',''),
                                        key=f"enc_t_{nombre_eq}_{j}")
                enc_list_new.append({'nombre':en,'cedula':ec,'celular':et,'cod':''})

            pch1,pch2,pch3 = st.columns(3)
            with pch1:
                ch_n = st.text_input("Chofer (nombre)",
                                      value=pi_prev.get('chofer_nombre',''),
                                      key=f"ch_n_{nombre_eq}")
            with pch2:
                plca = st.text_input("Placa",
                                      value=pi_prev.get('placa',''),
                                      key=f"plca_{nombre_eq}")
            with pch3:
                ch_t = st.text_input("Celular chofer",
                                      value=pi_prev.get('chofer_celular',''),
                                      key=f"ch_t_{nombre_eq}")

            st.session_state.personal_info[nombre_eq] = {
                'supervisor_nombre': sup_n, 'supervisor_cedula': sup_c,
                'supervisor_celular': sup_t, 'supervisor_cod': '',
                'encuestadores': enc_list_new, 'n_enc': n_enc_eq,
                'chofer_nombre': ch_n, 'placa': plca, 'chofer_celular': ch_t,
            }
            st.markdown("</div>",unsafe_allow_html=True)

    # ── Resumen tabular ───────────────────────
    st.markdown("<div class='stitle'>Resumen de planificación</div>",unsafe_allow_html=True)
    if res_bal is not None and len(res_bal)>0:
        tr = pd.DataFrame([{
            'equipo':'TOTAL','jornada':'—',
            'n_upms':res_bal['n_upms'].sum(),
            'viv_reales':res_bal['viv_reales'].sum(),
            'carga_ponderada':res_bal['carga_ponderada'].sum(),
            'dist_km':res_bal.get('dist_km',pd.Series([0])).sum()
        }])
        rep=pd.concat([res_bal,tr],ignore_index=True)
        st.dataframe(rep.rename(columns={
            'equipo':'Equipo','jornada':'Jornada','n_upms':'UPMs',
            'viv_reales':'Viv. reales','carga_ponderada':'Carga pond.',
            'dist_km':'Dist. (km)'}),use_container_width=True)

    # Vista previa de la tabla de asignación
    cols_ok=[c for c in ['id_entidad','upm','tipo_entidad','viv','carga_pond',
                          'equipo','jornada','encuestador','dia_inicio','dia_fin','lat','lon']
             if c in df_plan.columns]
    df_exp_pre=df_plan[cols_ok].sort_values(
        ['equipo','jornada','encuestador','dia_inicio']).reset_index(drop=True)
    st.dataframe(df_exp_pre,use_container_width=True,height=300)

    # ── Botón de descarga Excel ───────────────
    st.markdown("<div class='stitle'>Descargar Excel formateado</div>",
                unsafe_allow_html=True)
    st.markdown("""<div class='ibox'>
    El Excel incluye una hoja por jornada. Cada equipo tiene su bloque de encabezado
    con el personal y su tabla de manzanas con el cronograma de ✓ por día.
    </div>""",unsafe_allow_html=True)

    if st.button("📋 Generar Excel", use_container_width=True, type="primary"):
        with st.spinner("Generando Excel..."):
            try:
                excel_bytes = generar_excel(
                    df_plan       = df_plan,
                    eq_cfg        = eq_cfg,
                    personal_info = st.session_state.personal_info,
                    fecha_j1      = st.session_state.fecha_j1,
                    fecha_j2      = st.session_state.fecha_j2,
                    dias_op       = p["dias_op"],
                    j1_num        = st.session_state.get("j1_num", 1),
                    j2_num        = st.session_state.get("j2_num", 2),
                    mes_nombre    = MESES_N.get(int(df['mes'].iloc[0]),'')
                )
                j1n = st.session_state.get("j1_num", 1)
                j2n = st.session_state.get("j2_num", 2)
                mes_n = MESES_N.get(int(df['mes'].iloc[0]),'mes')
                st.download_button(
                    label     = f"⬇️ Descargar J{j1n}+J{j2n}_{mes_n}.xlsx",
                    data      = excel_bytes,
                    file_name = f"planificacion_J{j1n}-J{j2n}_{mes_n}.xlsx",
                    mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("✓ Excel listo para descargar.")
            except Exception as e:
                st.error(f"Error generando Excel: {e}")
                import traceback
                st.code(traceback.format_exc())
