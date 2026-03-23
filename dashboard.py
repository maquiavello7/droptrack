import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import io, unicodedata, re, json, uuid

st.set_page_config(page_title="DropTrack Pro", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

# ══════════════════════════════════════════════
# CSS GLOBAL
# ══════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif!important}
.main .block-container{padding:1.5rem 2rem 3rem;max-width:1400px}
.stApp{background:#0f172a}
[data-testid="stSidebar"]{background:#1e293b!important;border-right:1px solid #334155}
[data-testid="stSidebar"] *{color:#e2e8f0!important}
[data-testid="stSidebar"] .stMarkdown p{color:#94a3b8!important;font-size:12px!important}
[data-testid="stFileUploader"]{background:#1e3a5f!important;border:2px dashed #3b82f6!important;border-radius:10px!important}
[data-testid="stFileUploader"] *{color:#93c5fd!important}
h1{color:#f1f5f9!important;font-size:28px!important;font-weight:800!important}
h2{color:#e2e8f0!important;font-size:20px!important;font-weight:700!important}
h3{color:#cbd5e1!important;font-size:16px!important;font-weight:600!important}
[data-testid="stMetric"]{background:#1e293b!important;border:1px solid #334155!important;border-radius:14px!important;padding:18px 20px!important}
[data-testid="stMetricLabel"]{color:#94a3b8!important;font-size:11px!important;font-weight:700!important;text-transform:uppercase!important;letter-spacing:.8px!important}
[data-testid="stMetricValue"]{color:#f1f5f9!important;font-size:32px!important;font-weight:800!important}
[data-testid="stDataFrame"]{border-radius:12px!important;overflow:hidden!important;border:1px solid #334155!important}
.stButton>button{background:linear-gradient(135deg,#0ea5e9,#3b82f6)!important;color:white!important;border:none!important;border-radius:8px!important;font-weight:700!important;font-size:13px!important;padding:8px 18px!important}
.stDownloadButton>button{background:linear-gradient(135deg,#10b981,#059669)!important;color:white!important;border:none!important;border-radius:8px!important;font-weight:700!important}
.stSelectbox>div>div{background:#1e293b!important;border:1.5px solid #334155!important;border-radius:8px!important;color:#e2e8f0!important}
.stSelectbox label{color:#94a3b8!important;font-size:11px!important;font-weight:700!important;text-transform:uppercase!important}
.stCheckbox label{color:#cbd5e1!important;font-size:13px!important}
.stDateInput>div>div>input{background:#1e293b!important;border:1.5px solid #3b82f6!important;color:#e2e8f0!important;border-radius:8px!important;font-size:14px!important;padding:10px!important}
.stDateInput label{color:#94a3b8!important;font-size:11px!important;font-weight:700!important;text-transform:uppercase!important}
.stTextInput>div>div>input,.stTextArea>div>div>textarea,.stNumberInput>div>div>input{background:#1e293b!important;border:1.5px solid #334155!important;color:#e2e8f0!important;border-radius:8px!important}
.stTextInput label,.stTextArea label,.stNumberInput label{color:#94a3b8!important;font-size:11px!important;font-weight:700!important;text-transform:uppercase!important}
.streamlit-expanderHeader{background:#1e293b!important;border:1px solid #334155!important;border-radius:10px!important;color:#cbd5e1!important;font-weight:600!important}
.streamlit-expanderContent{background:#162032!important;border:1px solid #334155!important}
hr{border-color:#334155!important}
/* Cards de informes */
.inf-card{background:#1e293b;border:1.5px solid #334155;border-radius:16px;padding:22px;transition:all .18s;position:relative;cursor:pointer;margin-bottom:4px}
.inf-card:hover{border-color:#0ea5e9;box-shadow:0 4px 20px rgba(14,165,233,.15);transform:translateY(-2px)}
.inf-card-title{font-size:16px;font-weight:800;color:#f1f5f9;margin-bottom:3px;padding-right:30px;line-height:1.3}
.inf-card-pub{font-size:11px;color:#475569;margin-bottom:10px;display:flex;align-items:center;gap:6px}
.inf-card-date{font-size:12px;color:#64748b;margin-bottom:3px;display:flex;align-items:center;gap:5px}
.inf-card-sync{font-size:11px;color:#475569;margin-bottom:12px;display:flex;align-items:center;gap:5px}
.inf-card-orders{font-size:15px;font-weight:700;color:#e2e8f0;margin-bottom:2px}
.inf-card-conf{font-size:12px;color:#0ea5e9;font-weight:600;margin-bottom:12px}
.inf-card-4stats{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-bottom:14px}
.inf-stat{text-align:center;padding:7px 4px;background:#162032;border-radius:6px}
.inf-stat-val{font-size:15px;font-weight:800;line-height:1}
.inf-stat-pct{font-size:10px;font-weight:600;margin-top:2px}
.inf-card-actions{display:flex;gap:6px;flex-wrap:wrap}
.btn-historial{background:transparent;border:1.5px solid #334155;color:#94a3b8;padding:5px 12px;border-radius:7px;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s}
.btn-historial:hover{border-color:#0ea5e9;color:#0ea5e9}
.btn-sync{background:transparent;border:1.5px solid #334155;color:#94a3b8;padding:5px 12px;border-radius:7px;font-size:12px;font-weight:600;cursor:pointer}
.btn-detalle{background:linear-gradient(135deg,#0ea5e9,#3b82f6);color:white;border:none;padding:5px 14px;border-radius:7px;font-size:12px;font-weight:700;cursor:pointer}
/* Home header */
.home-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:28px;flex-wrap:wrap;gap:12px}
.home-title{font-size:26px;font-weight:800;color:#f1f5f9;letter-spacing:.5px}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:16px}
/* Crear informe form */
.crear-wrap{max-width:700px;margin:0 auto}
.crear-title{font-size:26px;font-weight:800;color:#f1f5f9;margin-bottom:6px}
.crear-sub{font-size:13px;color:#64748b;margin-bottom:28px}
.form-card{background:#1e293b;border:1px solid #334155;border-radius:14px;padding:28px;margin-bottom:16px}
/* Detalle header */
.det-header{background:#1e293b;border:1px solid #334155;border-radius:14px;padding:20px 24px;margin-bottom:20px;display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:12px}
.det-title{font-size:22px;font-weight:800;color:#f1f5f9}
.det-meta{font-size:12px;color:#64748b;margin-top:3px}
/* Sec title */
.sec-title{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin:20px 0 10px;padding-bottom:8px;border-bottom:1px solid #334155}
.fin-section-title{font-size:20px;font-weight:800;color:#f1f5f9;margin:32px 0 16px;padding-bottom:10px;border-bottom:2px solid #334155;display:flex;align-items:center;gap:10px}
.fin-sub-title{font-size:14px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:1px;margin:20px 0 10px}
.big-kpi{background:#1e293b;border:1px solid #334155;border-radius:14px;padding:20px 22px;margin-bottom:10px}
.big-kpi-val{font-size:32px;font-weight:800;line-height:1.1;margin-bottom:4px}
.big-kpi-sub{font-size:13px;font-weight:600;margin-bottom:6px}
.big-kpi-lbl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#64748b}
.kpi-card{background:#1e293b;border:1px solid #334155;border-radius:14px;padding:20px 22px;text-align:center;margin-bottom:8px}
.kpi-val{font-size:36px;font-weight:800;line-height:1.1;margin-bottom:4px}
.kpi-pct{font-size:13px;font-weight:600;margin-bottom:6px}
.kpi-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#64748b}
.pg-row{display:flex;justify-content:space-between;align-items:center;padding:9px 14px;margin:2px 0;border-radius:7px}
.pg-row-bold{background:#1e293b}
.pg-row-label{color:#cbd5e1;font-size:14px}
.pg-row-label-bold{color:#f1f5f9;font-size:15px;font-weight:800}
.pg-row-val{font-size:14px;font-weight:700}
.pg-row-val-bold{font-size:16px;font-weight:800}
.pg-detalle{padding:2px 0 2px 28px;color:#64748b;font-size:12px}
.pg-sep{height:1px;background:#1e293b;margin:5px 0}
.formula-box{background:#0f172a;border:1px solid #334155;border-radius:8px;padding:12px 14px;margin-top:8px;font-size:12px;color:#94a3b8}
.formula-highlight{color:#0ea5e9;font-family:monospace;font-size:12px}
.fuente-tag{display:inline-block;background:#1e3a5f;color:#93c5fd;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;margin:2px}
.alerta-card{background:#1f0707;border:1.5px solid #ef4444;border-radius:12px;padding:14px 18px;margin:12px 0;color:#fca5a5;font-weight:600;font-size:14px}
.ok-card{background:#052e16;border:1.5px solid #10b981;border-radius:12px;padding:14px 18px;margin:12px 0;color:#6ee7b7;font-weight:600}
.capital-row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #2d1f1f}
/* Scrollbar */
::-webkit-scrollbar{width:5px}
::-webkit-scrollbar-track{background:#0f172a}
::-webkit-scrollbar-thumb{background:#334155;border-radius:3px}
</style>
""", unsafe_allow_html=True)

PLOT_BG = dict(paper_bgcolor='#1e293b', plot_bgcolor='#162032',
    font=dict(family='Inter',color='#cbd5e1',size=12),
    margin=dict(t=36,b=44,l=10,r=10),
    legend=dict(bgcolor='#1e293b',bordercolor='#334155',borderwidth=1,font=dict(color='#cbd5e1',size=11)))
AX = dict(gridcolor='#334155',linecolor='#334155',tickfont=dict(color='#94a3b8',size=11))
C  = {'ent':'#10b981','dev':'#ef4444','can':'#94a3b8','sinf':'#f59e0b','sind':'#3b82f6','teal':'#0ea5e9','purple':'#8b5cf6'}

# ══════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════
def fmt_cop(v):
    try:
        v = float(v or 0)
        if v < 0:
            return f"-${abs(v):,.2f}"
        return f"${v:,.2f}"
    except: return "$0"


MESES_ES = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',
             7:'JUL',8:'AGO',9:'SEP',10:'OCT',11:'NOV',12:'DIC'}

def fmt_fecha_mes(fecha_str):
    """Convierte '15/03/2026' o '2026-03-15' → '15 MAR 2026'"""
    try:
        s = str(fecha_str).strip()
        # Try dd/mm/yyyy
        if '/' in s:
            parts = s.split('/')
            if len(parts) == 3:
                d,m,y = int(parts[0]),int(parts[1]),int(parts[2][:4])
                return f"{d:02d} {MESES_ES[m]} {y}"
        # Try yyyy-mm-dd
        if '-' in s and len(s) >= 10:
            y,m,d = int(s[:4]),int(s[5:7]),int(s[8:10])
            return f"{d:02d} {MESES_ES[m]} {y}"
    except: pass
    return fecha_str

def fmt_mes_anio(fecha_str):
    """Convierte fecha → 'MAR 2026'"""
    try:
        s = str(fecha_str).strip()
        if '/' in s:
            parts = s.split('/')
            if len(parts) == 3:
                m,y = int(parts[1]),int(parts[2][:4])
                return f"{MESES_ES[m]} {y}"
        if '-' in s and len(s) >= 10:
            m,y = int(s[5:7]),int(s[:4])
            return f"{MESES_ES[m]} {y}"
    except: pass
    return fecha_str

def pct(n,d):
    try: return round(n/d*100,2) if d else 0
    except: return 0

def ns(s):
    s=str(s).lower().strip()
    s=unicodedata.normalize('NFD',s)
    return ''.join(c for c in s if unicodedata.category(c)!='Mn')

def find_col(df,cands):
    normed={ns(c):c for c in df.columns}
    for c in cands:
        nc=ns(c)
        if nc in normed: return normed[nc]
    for c in cands:
        nc=ns(c)
        for k,o in normed.items():
            if nc in k or k in nc: return o
    return None

def load_file(f):
    try:
        n=f.name.lower()
        if n.endswith('.csv'):
            b=f.read(); f.seek(0)
            for sep in [',',';','\t']:
                try:
                    df=pd.read_csv(io.BytesIO(b),sep=sep,encoding='utf-8',low_memory=False)
                    if len(df.columns)>1: return df
                except: pass
            return pd.read_csv(f,encoding='latin1',low_memory=False)
        return pd.read_excel(f,engine='openpyxl')
    except Exception as e:
        st.error(f"Error leyendo archivo: {e}"); return None

def parse_dt_dropi(fecha_str,hora_str=""):
    try:
        f=str(fecha_str or '').strip()
        h=str(hora_str or '').strip()
        if not f or f.lower() in ['nan','none','']: return pd.NaT
        if re.match(r'\d{2}-\d{2}-\d{4}',f):
            dt_str=f"{f} {h}" if h and h.lower() not in ['nan','none',''] else f
            fmt='%d-%m-%Y %H:%M' if h and h.lower() not in ['nan','none',''] else '%d-%m-%Y'
            return datetime.strptime(dt_str[:len(fmt)],fmt)
        return pd.to_datetime(f,dayfirst=True,errors='coerce')
    except: return pd.NaT

def norm_estado(s):
    s=str(s).upper().strip()
    s=unicodedata.normalize('NFD',s)
    return ''.join(c for c in s if unicodedata.category(c)!='Mn')

CANCELACIONES={'CANCELADO','RECHAZADO','NUEVO','CONFIRMADO'}
def es_cancelacion(s): return norm_estado(s) in CANCELACIONES
def es_e(s):  return norm_estado(s)=='ENTREGADO'
def es_d(s):  return norm_estado(s)=='DEVOLUCION'
def es_c(s):  return es_cancelacion(s)
def es_despachado(s):
    n=norm_estado(s)
    return n not in CANCELACIONES and n not in {'','NONE','NAN'}
def es_s(s): return es_despachado(s) and not es_e(s) and not es_d(s)

def cat_log(s):
    """Categoriza estado para informe logístico."""
    n=norm_estado(s)
    if n=='ENTREGADO':          return 'ENTREGADO'
    if n=='DEVOLUCION':         return 'DEVOLUCION'
    if n in {'EN PROCESAMIENTO','EN PUNTO DROOP','EN TERMINAL DESTINO',
             'EN TERMINAL ORIGEN','REENVIO'}: return 'EN TRANSITO'
    if n=='EN REPARTO':         return 'EN REPARTO'
    if n=='RECLAME EN OFICINA': return 'RECLAMAR EN OFICINA'
    if n in {'NOVEDAD','NOVEDAD SOLUCIONADA'}: return 'CON NOVEDAD'
    if n in CANCELACIONES:      return 'CANCELACION'
    return 'OTRO'

LOG_CATS = ['ENTREGADO','DEVOLUCION','EN TRANSITO','EN REPARTO',
            'RECLAMAR EN OFICINA','CON NOVEDAD','OTRO']
CAT_CMAP = {
    'ENTREGADO':'#10b981','DEVOLUCION':'#ef4444','EN TRANSITO':'#f59e0b',
    'EN REPARTO':'#3b82f6','RECLAMAR EN OFICINA':'#8b5cf6',
    'CON NOVEDAD':'#f97316','CANCELACION':'#94a3b8','OTRO':'#64748b',
}
def badge(s):
    if es_e(s): return '🟢'
    if es_d(s): return '🔴'
    if es_cancelacion(s): return '⚫'
    if es_s(s): return '🟡'
    return '⚪'
def color_estado(s):
    if es_e(s): return C['ent']
    if es_d(s): return C['dev']
    if es_cancelacion(s): return C['can']
    return C['sinf']

def kpi_card(lbl,val,color,sub=""):
    return f'<div class="kpi-card"><div class="kpi-val" style="color:{color}">{val}</div><div class="kpi-pct" style="color:{color}99">{sub}</div><div class="kpi-lbl">{lbl}</div></div>'

def to_n(s): return pd.to_numeric(s,errors='coerce').fillna(0)

# IDs de tarjetas Dropi
TARJETA_PAUTA       = 'car_031C0CiYRq7jjnfX85EEfg'.upper()
TARJETA_GASTOS_FIJOS= 'car_02yYLGyMzUygINeCYnT6Yn'.upper()

def clasif_desc(desc):
    """Clasifica cada transacción del historial de cartera Dropi."""
    d = str(desc).upper().strip()
    # Normalizar tildes: DEVOLUCIÓN → DEVOLUCION, etc.
    d_norm = unicodedata.normalize('NFD', d)
    d_norm = ''.join(c for c in d_norm if unicodedata.category(c) != 'Mn')

    # ── ENTRADAS ──
    if 'GANANCIA EN LA ORDEN' in d_norm or ('GANANCIA' in d_norm and 'ENTRADA' in d_norm):
        return 'Ganancia cobrada'
    if 'INDEMNIZACION' in d_norm:
        return 'Indemnización'
    if 'GARANTIA' in d_norm and 'ENTRADA' in d_norm:
        return 'Devolución por garantía'
    if 'DEVOLUCION' in d_norm and 'DINERO' in d_norm and 'GARANTIA' in d_norm:
        return 'Devolución por garantía'

    # ── SALIDAS QUE SON COSTOS REALES ──
    if 'COBRO DE FLETE INICIAL' in d_norm or 'FLETE INICIAL' in d_norm:
        return 'Flete envío cobrado'
    if ('DEVOLUCION' in d_norm and 'ENTREGA NO EFECTIVA' in d_norm) or        'COBRO DE DEVOLUCION' in d_norm or        ('DEVOLUCION' in d_norm and 'SALIDA' in d_norm and 'COBRO' in d_norm):
        return 'Flete devolución cobrado'
    if 'NUEVA ORDEN' in d_norm:
        return 'Pago anticipado orden (COD)'
    if 'MANTENIMIENTO MENSUAL TARJETA' in d_norm:
        return 'Mantenimiento tarjeta virtual'

    # ── MOVIMIENTOS (no afectan P&G) ──
    if 'RECARGA DE TARJETA' in d_norm:
        if TARJETA_PAUTA in d:
            return 'Recarga tarjeta — Pauta publicitaria'
        if TARJETA_GASTOS_FIJOS in d:
            return 'Recarga tarjeta — Gastos fijos'
        return 'Recarga tarjeta — Otra'
    if 'RETIRO DE SALDO' in d_norm or 'PETICION DE RETIRO' in d_norm:
        return 'Retiro a cuenta bancaria'
    if 'TRANSFERENCIA DE WALLET' in d_norm:
        return 'Transferencia entre wallets'
    if 'RECARGA DE WALLET' in d_norm or 'RECARGA WALLET' in d_norm:
        return 'Recarga de wallet'

    # ── FALLBACKS ──
    if 'ENTRADA' in d_norm: return 'Otra entrada'
    if 'SALIDA'  in d_norm: return 'Otra salida'
    return 'Otro'

# Categorías que SÍ son costos reales (afectan el P&G)
COSTOS_REALES = {
    'Ganancia cobrada',
    'Flete envío cobrado',
    'Flete devolución cobrado',
    'Pago anticipado orden (COD)',
    'Mantenimiento tarjeta virtual',
    'Devolución por garantía',
    'Indemnización',
}

# Categorías que son solo movimientos (NO afectan el P&G)
SOLO_MOVIMIENTOS = {
    'Recarga tarjeta — Pauta publicitaria',
    'Recarga tarjeta — Gastos fijos',
    'Recarga tarjeta — Otra',
    'Retiro a cuenta bancaria',
    'Transferencia entre wallets',
    'Recarga de wallet',
}

def fmt_dt(iso):
    try:
        d=datetime.fromisoformat(iso)
        return f"{d.day} {MESES_ES[d.month]} {d.year}, {d.strftime('%-I:%M %p').lower()}"
    except: return iso or '—'

def uid(): return str(uuid.uuid4())[:8]

# ══════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════
if 'page'       not in st.session_state: st.session_state.page       = 'home'
if 'informes'   not in st.session_state: st.session_state.informes   = []
if 'cur_id'     not in st.session_state: st.session_state.cur_id     = None
if 'nav_mod'    not in st.session_state: st.session_state.nav_mod    = '📊 Informes'
if 'tareas'     not in st.session_state: st.session_state.tareas     = []
if 'alertas'    not in st.session_state: st.session_state.alertas    = []
if 'cortes'     not in st.session_state:
    st.session_state.cortes=[
        {'id':'c1','hora':'08:00','label':'Corte mañana','done':False},
        {'id':'c2','hora':'14:00','label':'Corte mediodía','done':False},
        {'id':'c3','hora':'20:00','label':'Corte noche','done':False},
    ]
if 'camp_meta'  not in st.session_state: st.session_state.camp_meta  = None
if 'camp_tiktok'not in st.session_state: st.session_state.camp_tiktok= None
if 'nov_data'   not in st.session_state: st.session_state.nov_data   = None
if 'seg_data'   not in st.session_state: st.session_state.seg_data   = None


def nav(page, cur_id=None):
    st.session_state.page   = page
    if cur_id: st.session_state.cur_id = cur_id
    st.rerun()

def get_inf(inf_id):
    return next((i for i in st.session_state.informes if i['id']==inf_id), None)

# ══════════════════════════════════════════════
# SIDEBAR NAV
# ══════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:10px;padding:4px 0 16px">
        <div style="width:34px;height:34px;background:linear-gradient(135deg,#0ea5e9,#3b82f6);
                    border-radius:9px;display:flex;align-items:center;justify-content:center;
                    font-size:16px;font-weight:800;color:white">D</div>
        <span style="font-size:18px;font-weight:800;color:#f1f5f9">Drop<span style="color:#0ea5e9">Track</span></span>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.page in ('home','crear','detalle'):
        st.caption("TIENDA")
        st.markdown('<div style="background:#162032;border:1px solid #334155;border-radius:8px;padding:8px 12px;font-size:13px;color:#e2e8f0;margin-bottom:12px">🏪 Mi Tienda</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.caption("MÓDULOS")

    nav_items = [
        ("📊","Informes rendimiento"),
        ("📣","Monitoreo campañas"),
        ("💵","Flujo de Caja"),
        ("🔄","Novedades"),
        ("🗺","Seguimiento"),
        ("✅","Tareas programadas"),
    ]
    for icon, label in nav_items:
        full = f"{icon} {label}"
        is_active = (st.session_state.nav_mod == full and
                     st.session_state.page not in ('crear','detalle')) or \
                    (label == "Informes rendimiento" and
                     st.session_state.page in ('home','crear','detalle'))
        bg    = "background:#162032;border-left:3px solid #0ea5e9;" if is_active else "border-left:3px solid transparent;"
        color = "#0ea5e9" if is_active else "#94a3b8"
        fw    = "700" if is_active else "500"
        if st.button(f"{icon}  {label}", key=f"nav_{label}", use_container_width=True):
            st.session_state.nav_mod = full
            if label == "Informes rendimiento":
                nav('home')
            else:
                nav('module')
    st.markdown("---")
    st.caption("🔒 Datos solo en tu Mac")

# ══════════════════════════════════════════════
# PÁGINA: HOME — Lista de Informes
# ══════════════════════════════════════════════
if st.session_state.page == 'home':
    # Header
    col_h1, col_h2 = st.columns([1,1])
    with col_h1:
        st.markdown('<div style="font-size:10px;font-weight:700;color:#64748b;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">← INFORMES DE RENDIMIENTO</div>', unsafe_allow_html=True)
    # Title row with button
    th1, th2 = st.columns([3,1])
    with th1:
        st.markdown('<div style="font-size:26px;font-weight:800;color:#f1f5f9;margin-bottom:20px">📊 INFORMES DE RENDIMIENTO</div>', unsafe_allow_html=True)
    with th2:
        if st.button("➕  Crear informe", key="btn_crear", use_container_width=True):
            nav('crear')

    if not st.session_state.informes:
        st.markdown("""
        <div style="text-align:center;padding:80px 20px;color:#475569">
            <div style="font-size:48px;margin-bottom:16px;opacity:.4">📊</div>
            <div style="font-size:16px;font-weight:600;color:#64748b;margin-bottom:6px">Sin informes creados</div>
            <div style="font-size:13px">Presiona <strong style="color:#0ea5e9">➕ Crear informe</strong> para comenzar</div>
        </div>""", unsafe_allow_html=True)
    else:
        cols_per_row = 3
        informes = st.session_state.informes
        rows = [informes[i:i+cols_per_row] for i in range(0,len(informes),cols_per_row)]
        for row in rows:
            cols = st.columns(len(row))
            for col, inf in zip(cols, row):
                last     = inf['history'][-1] if inf.get('history') else None
                sync_str = fmt_dt(last['timestamp']) if last else 'Sin sincronizar'
                with col:
                    if last:
                        total = last.get('totalOrd',0)
                        desp  = last.get('despachados',0)
                        conf  = last.get('pctConf',0)
                        ent   = last.get('entregado',0)
                        dev   = last.get('devolucion',0)
                        sinf  = last.get('sinFinal',0)
                        can   = last.get('cancelados',0)
                        pE    = last.get('pctE',0)
                        pD    = last.get('pctD',0)
                        pS    = last.get('pctS',0)
                        pC    = round(pct(can,total),2)
                        maxDev= inf.get('max_dev',20)
                        dev_color = '#ef4444' if pD > maxDev else '#ef4444'  # siempre rojo dev

                        st.markdown(f"""
<div style="background:#1e293b;border:1.5px solid #334155;border-radius:16px;padding:20px 18px;margin-bottom:4px">
    <div style="font-size:16px;font-weight:800;color:#f1f5f9;margin-bottom:6px">{inf['title']}</div>
    <div style="font-size:11px;color:#475569;margin-bottom:6px">🌐 Público</div>
    <div style="font-size:12px;color:#64748b;margin-bottom:2px">📅 {fmt_fecha_mes(inf['date_from'])} – {fmt_fecha_mes(inf['date_to'])}</div>
    <div style="font-size:12px;color:#64748b;margin-bottom:14px">🔄 {sync_str}</div>
    <div style="font-size:15px;font-weight:700;color:#e2e8f0;margin-bottom:2px">{total:,} ordenes</div>
    <div style="font-size:13px;color:#0ea5e9;font-weight:600;margin-bottom:12px">{desp} ({conf:.2f}%) Confirmación</div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:4px">
        <div style="background:#162032;border-radius:6px;padding:8px 4px;text-align:center">
            <div style="font-size:16px;font-weight:800;color:#10b981">{ent}</div>
            <div style="font-size:10px;font-weight:600;color:#10b981">({pE:.2f}%)</div>
        </div>
        <div style="background:#162032;border-radius:6px;padding:8px 4px;text-align:center">
            <div style="font-size:16px;font-weight:800;color:{dev_color}">{dev}</div>
            <div style="font-size:10px;font-weight:600;color:{dev_color}">({pD:.2f}%)</div>
        </div>
        <div style="background:#162032;border-radius:6px;padding:8px 4px;text-align:center">
            <div style="font-size:16px;font-weight:800;color:#f59e0b">{sinf}</div>
            <div style="font-size:10px;font-weight:600;color:#f59e0b">({pS:.2f}%)</div>
        </div>
        <div style="background:#162032;border-radius:6px;padding:8px 4px;text-align:center">
            <div style="font-size:16px;font-weight:800;color:#94a3b8">{can}</div>
            <div style="font-size:10px;font-weight:600;color:#94a3b8">({pC:.2f}%)</div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
<div style="background:#1e293b;border:1.5px solid #334155;border-radius:16px;padding:20px 18px;margin-bottom:4px">
    <div style="font-size:16px;font-weight:800;color:#f1f5f9;margin-bottom:6px">{inf['title']}</div>
    <div style="font-size:11px;color:#475569;margin-bottom:6px">🌐 Público</div>
    <div style="font-size:12px;color:#64748b;margin-bottom:2px">📅 {fmt_fecha_mes(inf['date_from'])} – {fmt_fecha_mes(inf['date_to'])}</div>
    <div style="font-size:12px;color:#64748b;margin-bottom:14px">🔄 Sin sincronizar</div>
    <div style="height:72px;display:flex;align-items:center;justify-content:center;color:#334155;font-size:12px">
        Sincroniza para ver resultados
    </div>
</div>""", unsafe_allow_html=True)

                    # Buttons
                    b1, b2, b3 = st.columns(3)
                    with b1:
                        if st.button("🔄 Sincronizar", key=f"sync_{inf['id']}", use_container_width=True):
                            nav('detalle', inf['id'])
                    with b2:
                        if st.button("Detalle →", key=f"det_{inf['id']}", use_container_width=True):
                            nav('detalle', inf['id'])
                    with b3:
                        if st.button("🗑 Borrar", key=f"del_{inf['id']}", use_container_width=True):
                            st.session_state[f"confirm_del_{inf['id']}"] = True
                            st.rerun()

                    # Confirmación de borrado
                    if st.session_state.get(f"confirm_del_{inf['id']}", False):
                        st.warning(f"¿Seguro que quieres borrar **{inf['title']}**? Esta acción no se puede deshacer.")
                        cc1, cc2 = st.columns(2)
                        with cc1:
                            if st.button("✅ Sí, borrar", key=f"conf_del_{inf['id']}", use_container_width=True):
                                st.session_state.informes = [i for i in st.session_state.informes if i['id'] != inf['id']]
                                for k in [f"raw_dropi_{inf['id']}", f"raw_panc_{inf['id']}", f"raw_fin_{inf['id']}"]:
                                    if k in st.session_state: del st.session_state[k]
                                st.session_state.pop(f"confirm_del_{inf['id']}", None)
                                st.rerun()
                        with cc2:
                            if st.button("❌ Cancelar", key=f"canc_del_{inf['id']}", use_container_width=True):
                                st.session_state.pop(f"confirm_del_{inf['id']}", None)
                                st.rerun()

    # Modal historial
    if st.session_state.get('show_hist') and st.session_state.cur_id:
        inf = get_inf(st.session_state.cur_id)
        if inf:
            with st.expander(f"📋 Historial: {inf['title']}", expanded=True):
                hist = inf.get('history',[])
                if not hist:
                    st.info("Sin historial — sincroniza el informe primero")
                else:
                    h_df = pd.DataFrame([{
                        'Fecha':      fmt_dt(h['timestamp']),
                        'Órdenes':    h.get('totalOrd',0),
                        '% Conf.':    f"{h.get('pctConf',0)}%",
                        'Entregadas': f"{h.get('entregado',0)} ({h.get('pctE',0)}%)",
                        'Devolucion': f"{h.get('devolucion',0)} ({h.get('pctD',0)}%)",
                        'Sin estado final': f"{h.get('sinFinal',0)} ({h.get('pctS',0)}%)",
                        'Ganancia':   fmt_cop(h.get('totalGanancia',0)),
                    } for h in reversed(hist)])
                    st.dataframe(h_df, use_container_width=True, hide_index=True)
                if st.button("Cerrar historial", key="close_hist"):
                    st.session_state.show_hist = False
                    st.rerun()

# ══════════════════════════════════════════════
# PÁGINA: CREAR INFORME
# ══════════════════════════════════════════════
elif st.session_state.page == 'crear':
    if st.button("← Volver a Informes", key="back_crear"):
        nav('home')

    st.markdown("## Crear informe de rendimiento")
    st.markdown('<div style="font-size:13px;color:#64748b;margin-bottom:24px">Define el período del informe. Los archivos se cargan dentro del informe.</div>', unsafe_allow_html=True)

    import calendar as _cal

    MESES_ES = {
        'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
        'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12,
        'ene':1,'feb':2,'mar':3,'abr':4,'jun':6,'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12,
    }

    def detectar_mes_anio(texto):
        t = texto.lower().strip()
        mes = None
        anio = date.today().year
        for nombre_mes, num in sorted(MESES_ES.items(), key=lambda x: -len(x[0])):
            if nombre_mes in t:
                mes = num
                break
        m = re.search(r'\b(202\d)\b', t)
        if m: anio = int(m.group(1))
        return mes, anio

    def fechas_del_mes(mes, anio):
        ini = date(anio, mes, 1)
        fin = date(anio, mes, _cal.monthrange(anio, mes)[1])
        return ini, fin

    # Inicializar session_state para las fechas si no existen
    if 'ci_desde' not in st.session_state:
        st.session_state['ci_desde'] = date.today().replace(day=1)
    if 'ci_hasta' not in st.session_state:
        st.session_state['ci_hasta'] = date.today()
    if 'ci_nombre_prev' not in st.session_state:
        st.session_state['ci_nombre_prev'] = ''

    # Nombre del informe — FUERA del form para detectar en tiempo real
    nombre = st.text_input(
        "Nombre del informe *",
        placeholder="Ej: Enero General 2026",
        key="ci_nombre"
    )

    # Detectar mes y actualizar session_state ANTES de renderizar el form
    mes_det, anio_det = detectar_mes_anio(nombre) if nombre else (None, date.today().year)

    if mes_det and nombre != st.session_state.get('ci_nombre_prev',''):
        # Nombre cambió y tiene mes → actualizar fechas
        fi_new, ft_new = fechas_del_mes(mes_det, anio_det)
        st.session_state['ci_desde'] = fi_new
        st.session_state['ci_hasta'] = ft_new
        st.session_state['ci_nombre_prev'] = nombre

    if mes_det:
        fi_show, ft_show = fechas_del_mes(mes_det, anio_det)
        nombre_mes_display = [k for k,v in MESES_ES.items() if v==mes_det and len(k)>3][0].capitalize()
        st.markdown(f'''
        <div style="background:#052e16;border:1px solid #10b981;border-radius:8px;
                    padding:10px 16px;margin-bottom:12px;font-size:13px;color:#6ee7b7">
            ✅ Mes detectado: <strong>{nombre_mes_display} {anio_det}</strong> —
            fechas: {fi_show.day:02d} {MESES_ES[fi_show.month]} {fi_show.year} → {ft_show.day:02d} {MESES_ES[ft_show.month]} {ft_show.year}
        </div>''', unsafe_allow_html=True)

    with st.form("form_crear"):
        st.markdown('<div style="font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.5px;margin:0 0 12px">📅 RANGO DE FECHAS</div>', unsafe_allow_html=True)
        dc1, dc2 = st.columns(2)
        with dc1:
            fecha_ini = st.date_input(
                "Fecha inicial *",
                value=st.session_state['ci_desde'],
                min_value=date(2020,1,1),
                max_value=date(2030,12,31),
                format="DD/MM/YYYY",
                key="ci_desde_form"
            )
        with dc2:
            fecha_fin = st.date_input(
                "Fecha final *",
                value=st.session_state['ci_hasta'],
                min_value=date(2020,1,1),
                max_value=date(2030,12,31),
                format="DD/MM/YYYY",
                key="ci_hasta_form"
            )

        st.markdown("---")
        max_dev = st.slider("Máximo nivel de devolución aceptable (%)", 0, 50, 20, key="ci_maxdev",
            help="Se resaltará en rojo cuando las devoluciones superen este porcentaje")

        submitted = st.form_submit_button("Crear informe ➕", use_container_width=True)
        if submitted:
            if not nombre:
                st.error("El nombre del informe es obligatorio")
            elif fecha_ini >= fecha_fin:
                st.error("La fecha final debe ser posterior a la inicial")
            else:
                # Limpiar session_state del form
                for k in ['ci_desde','ci_hasta','ci_nombre_prev']:
                    if k in st.session_state: del st.session_state[k]
                nuevo = {
                    'id':        uid(),
                    'title':     nombre,
                    'date_from': fecha_ini.strftime('%d/%m/%Y'),
                    'date_to':   fecha_fin.strftime('%d/%m/%Y'),
                    'date_from_dt': fecha_ini.isoformat(),
                    'date_to_dt':   fecha_fin.isoformat(),
                    'max_dev':   max_dev,
                    'history':   [],
                    'created_at':datetime.now().isoformat(),
                }
                st.session_state.informes.insert(0, nuevo)
                st.success(f"✅ Informe '{nombre}' creado")
                nav('detalle', nuevo['id'])

# ══════════════════════════════════════════════
# PÁGINA: DETALLE DEL INFORME
# ══════════════════════════════════════════════
elif st.session_state.page == 'detalle':
    inf = get_inf(st.session_state.cur_id)
    if not inf:
        st.error("Informe no encontrado"); nav('home')
        st.stop()

    # Back button
    if st.button("← Informes de Rendimiento", key="back_det"):
        nav('home')

    # Header card
    last = inf['history'][-1] if inf.get('history') else None
    sync_str = fmt_dt(last['timestamp']) if last else 'Sin sincronizar'
    st.markdown(f"""
    <div class="det-header">
        <div>
            <div class="det-title">{inf['title']}</div>
            <div class="det-meta">📅 {fmt_fecha_mes(inf['date_from'])} – {fmt_fecha_mes(inf['date_to'])}  ·  🔄 {sync_str}</div>
        </div>
    </div>""", unsafe_allow_html=True)

    # Tipo de informe selector
    tipo_inf = st.selectbox("Ver", ["📦 Informe Logístico","💰 Informe Financiero"], key="tipo_inf_sel")

    # ── CARGA DE ARCHIVOS (dentro del informe, colapsados) ──
    with st.expander("📁 Cargar / actualizar archivos del informe", expanded=not bool(last)):

        # ── Guía de columnas requeridas ──
        with st.expander("📋 ¿Qué columnas necesita cada archivo?", expanded=False):
            c_pan, c_dlog, c_dfin, c_cart = st.columns(4)
            with c_pan:
                st.markdown("""**🥞 Pancake**
- `Número de seguimiento` ← guía
- `Cronograma de actualización de estado` ← para días entrega y novedades
- `Creado en` ← fecha
- `Estado del pedido`
- `Número de teléfono`
- `Día de creación`
- `Etiqueta` ← para detectar órdenes en oficina (`pick up at office`)
- `Día de actualización` ← para calcular días en oficina
- `Hora de actualización de etiquetas`

> Exportar desde Pancake → Pedidos → Exportar Excel  
> ⚠️ Asegúrate de incluir las columnas **Etiqueta** y **Día de actualización**""")
            with c_dlog:
                st.markdown("""**📦 Dropi Logístico**
- `NÚMERO GUIA` ← llave
- `ESTATUS`
- `TRANSPORTADORA`
- `CIUDAD DESTINO`
- `DEPARTAMENTO DESTINO`
- `NOVEDAD`
- `FECHA DE ÚLTIMO MOVIMIENTO`
- `VENDEDOR`
- `TOTAL DE LA ORDEN`
- `GANANCIA`
- `PRECIO PROVEEDOR X CANTIDAD`
- `PRECIO FLETE`
- `COSTO DEVOLUCION FLETE`

> Dropi → Mis órdenes → Exportar (producto por fila)""")
            with c_dfin:
                st.markdown("""**💰 Dropi Financiero**
- `NÚMERO GUIA` ← llave
- `ESTATUS`
- `TRANSPORTADORA`
- `PRECIO FLETE`
- `COSTO DEVOLUCION FLETE`
- `FECHA GENERACION DE GUIA`
- `FECHA DE ÚLTIMO MOVIMIENTO`
- `GANANCIA`
- `VALOR FACTURADO`
- `VALOR DE COMPRA EN PRODUCTOS`

> Dropi → Mis órdenes → Exportar (orden por fila)""")
            with c_cart:
                st.markdown("""**💳 Cartera Dropi**
- `FECHA`
- `TIPO`
- `MONTO`
- `DESCRIPCIÓN`
- `ORDEN ID`
- `NUMERO DE GUIA`
- `CONCEPTO DE RETIRO`

> Dropi → Cartera → Historial → Exportar""")

        st.markdown('<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:#94a3b8">📋 <strong style="color:#f1f5f9">4 archivos:</strong> Pancake · Dropi Logístico · Dropi Financiero · Cartera Dropi — <em>Presiona "¿Qué columnas necesita cada archivo?" para ver el detalle</em></div>', unsafe_allow_html=True)

        fa, fb = st.columns(2)
        with fa:
            pf_det = st.file_uploader("🥞 Pancake — Órdenes", type=['csv','xlsx','xls'], key=f"pf_{inf['id']}",
                help="Columnas clave: 'Número de seguimiento' y 'Cronograma de actualización de estado'")
        with fb:
            drf_det = st.file_uploader("📦 Dropi Logístico — 1 fila por producto", type=['csv','xlsx','xls'], key=f"df_{inf['id']}",
                help="Columnas clave: NÚMERO GUIA, ESTATUS, TRANSPORTADORA, CIUDAD DESTINO, NOVEDAD")

        fc2, fd2 = st.columns(2)
        with fc2:
            drf_fin = st.file_uploader("💰 Dropi Financiero — 1 fila por orden", type=['csv','xlsx','xls'], key=f"dfin_{inf['id']}",
                help="Columnas clave: NÚMERO GUIA, PRECIO FLETE, COSTO DEVOLUCION FLETE, FECHA GENERACION DE GUIA")
        with fd2:
            crf_det = st.file_uploader("💳 Cartera Dropi — Historial transacciones", type=['csv','xlsx','xls'], key=f"cf_{inf['id']}",
                help="Columnas clave: FECHA, TIPO, MONTO, DESCRIPCIÓN")

        # Manual column mapping for Dropi
        dp_raw_det = dd_raw_det = dc_raw_det = dd_fin_raw = None
        dp_det = dd_det = dc_det = None
        p_map_det = d_map_det = {}

        # Cache raw files in session_state — persiste al hacer st.rerun()
        raw_cache_key = f"raw_dropi_{inf['id']}"
        raw_panc_key  = f"raw_panc_{inf['id']}"
        raw_fin_key   = f"raw_fin_{inf['id']}"

        if drf_det:
            dd_raw_det = load_file(drf_det)
            if dd_raw_det is not None:
                st.session_state[raw_cache_key] = dd_raw_det
        elif raw_cache_key in st.session_state:
            dd_raw_det = st.session_state[raw_cache_key]

        if pf_det:
            dp_raw_det = load_file(pf_det)
            if dp_raw_det is not None:
                st.session_state[raw_panc_key] = dp_raw_det
        elif raw_panc_key in st.session_state:
            dp_raw_det = st.session_state[raw_panc_key]

        # Dropi Financiero (1 orden por fila) — para fletes, costos, días entrega
        if drf_fin:
            dd_fin_raw = load_file(drf_fin)
            if dd_fin_raw is not None:
                st.session_state[raw_fin_key] = dd_fin_raw
                st.success(f"✅ Dropi Financiero cargado: {len(dd_fin_raw):,} órdenes · columnas: {len(dd_fin_raw.columns)}")
        elif raw_fin_key in st.session_state:
            dd_fin_raw = st.session_state[raw_fin_key]
            st.caption(f"💰 Dropi Financiero en caché: {len(dd_fin_raw):,} órdenes")

        if crf_det:
            dc_raw_det = load_file(crf_det)
            if dc_raw_det is not None:
                st.session_state[f"raw_cart_{inf['id']}"] = dc_raw_det
        elif f"raw_cart_{inf['id']}" in st.session_state:
            dc_raw_det = st.session_state[f"raw_cart_{inf['id']}"]

        def sel_det(df,key,label,cands):
            if df is None: return None
            opts=["— no usar —"]+list(df.columns)
            det=find_col(df,cands)
            idx=0
            if det:
                try: idx=opts.index(det)
                except: pass
            ch=st.selectbox(label,opts,index=idx,key=key)
            return None if ch=="— no usar —" else ch

        if dd_raw_det is not None:
            with st.expander("🗂 Mapeo de columnas Dropi", expanded=False):
                st.caption("El sistema pre-selecciona las columnas más probables. Ajusta si es necesario.")
                m1,m2,m3 = st.columns(3)
                with m1:
                    dc_est  = sel_det(dd_raw_det,f"e_est_{inf['id']}", "ESTATUS ⭐",           ['estatus','estado','status','ESTATUS'])
                    dc_fecha= sel_det(dd_raw_det,f"e_fec_{inf['id']}", "Fecha",                ['fecha','date','FECHA'])
                    dc_guia = sel_det(dd_raw_det,f"e_gui_{inf['id']}", "Número guía",          ['numero guia','numero guia','NÚMERO GUIA','guia','tracking'])
                    dc_trans= sel_det(dd_raw_det,f"e_tra_{inf['id']}", "Transportadora",       ['transportadora','TRANSPORTADORA','carrier'])
                    dc_ciudad=sel_det(dd_raw_det,f"e_ciu_{inf['id']}", "Ciudad destino",       ['ciudad destino','CIUDAD DESTINO','ciudad'])
                    dc_depto= sel_det(dd_raw_det,f"e_dep_{inf['id']}", "Departamento",         ['departamento destino','DEPARTAMENTO DESTINO','departamento'])
                    dc_nom  = sel_det(dd_raw_det,f"e_nom_{inf['id']}", "Nombre cliente",       ['nombre cliente','NOMBRE CLIENTE','nombre del cliente'])
                    dc_tel  = sel_det(dd_raw_det,f"e_tel_{inf['id']}", "Teléfono",             ['telefono','TELÉFONO','celular'])
                with m2:
                    dc_val  = sel_det(dd_raw_det,f"e_val_{inf['id']}", "Valor COD",            ['total de la orden','TOTAL DE LA ORDEN','valor facturado','VALOR FACTURADO','valor cod','valor de venta'])
                    dc_gan  = sel_det(dd_raw_det,f"e_gan_{inf['id']}", "Ganancia",             ['ganancia','GANANCIA','utilidad'])
                    dc_cos  = sel_det(dd_raw_det,f"e_cos_{inf['id']}", "Costo producto",       ['precio proveedor x cantidad','PRECIO PROVEEDOR X CANTIDAD','valor de compra en productos','costo producto'])
                    dc_fle  = sel_det(dd_raw_det,f"e_fle_{inf['id']}", "Precio flete",         ['precio flete','PRECIO FLETE','flete'])
                    dc_cdev = sel_det(dd_raw_det,f"e_cde_{inf['id']}", "Costo dev. flete",     ['costo devolucion flete','COSTO DEVOLUCION FLETE','costo devolucion'])
                    dc_prod = sel_det(dd_raw_det,f"e_pro_{inf['id']}", "Producto",             ['producto','PRODUCTO','nombre del producto'])
                    dc_vend = sel_det(dd_raw_det,f"e_ven_{inf['id']}", "Vendedor",             ['vendedor','VENDEDOR','seller'])
                with m3:
                    dc_fumov= sel_det(dd_raw_det,f"e_fum_{inf['id']}", "Fecha último mov.",    ['fecha de ultimo movimiento','FECHA DE ÚLTIMO MOVIMIENTO','fecha ultimo movimiento'])
                    dc_humov= sel_det(dd_raw_det,f"e_hum_{inf['id']}", "Hora último mov.",     ['hora de ultimo movimiento','HORA DE ÚLTIMO MOVIMIENTO','hora ultimo movimiento'])
                    dc_umov = sel_det(dd_raw_det,f"e_umv_{inf['id']}", "Último movimiento",    ['ultimo movimiento','ÚLTIMO MOVIMIENTO'])
                    dc_fguia= sel_det(dd_raw_det,f"e_fgu_{inf['id']}", "Fecha generación guía",['fecha guia generada','FECHA GUIA GENERADA','fecha generacion de guia','FECHA GENERACION DE GUIA'])
                    dc_nov  = sel_det(dd_raw_det,f"e_nov_{inf['id']}", "Novedad",              ['novedad','NOVEDAD','observacion'])
                    dc_novr = sel_det(dd_raw_det,f"e_nvr_{inf['id']}", "¿Fue solucionada?",   ['fue solucionada la novedad','FUE SOLUCIONADA LA NOVEDAD'])
                    dc_fnov = sel_det(dd_raw_det,f"e_fno_{inf['id']}", "Fecha de novedad",     ['fecha de novedad','FECHA DE NOVEDAD'])
                    dc_sol  = sel_det(dd_raw_det,f"e_sol_{inf['id']}", "Solución",             ['solucion','SOLUCIÓN','solucion novedad'])
                    dc_tipo_env = sel_det(dd_raw_det,f"e_ten_{inf['id']}", "Tipo de envío",         ['tipo de envio','TIPO DE ENVIO','tipo envio'])

            # Build dd_det
            manual_d = {'estatus':dc_est,'fecha':dc_fecha,'guia':dc_guia,'transportadora':dc_trans,
                        'ciudad':dc_ciudad,'departamento':dc_depto,'nombre':dc_nom,'telefono':dc_tel,
                        'valor':dc_val,'ganancia':dc_gan,'costo_prod':dc_cos,'flete':dc_fle,'costo_dev':dc_cdev,
                        'producto':dc_prod,'vendedor':dc_vend,'fecha_umov':dc_fumov,'hora_umov':dc_humov,
                        'ult_mov':dc_umov,'fecha_guia':dc_fguia,'novedad':dc_nov,'nov_resuelta':dc_novr,
                        'fecha_nov':dc_fnov,'solucion':dc_sol,'tipo_envio':dc_tipo_env}
            dd_det = pd.DataFrame()
            for std,col_real in manual_d.items():
                dd_det[std] = dd_raw_det[col_real].astype(str).str.strip() if col_real else None
            if 'fecha' in dd_det.columns:
                dd_det['_fecha_dt'] = dd_det['fecha'].apply(lambda x: parse_dt_dropi(x))
            if 'fecha_umov' in dd_det.columns:
                dd_det['_dt_umov'] = dd_det.apply(lambda r: parse_dt_dropi(r.get('fecha_umov',''),r.get('hora_umov','')), axis=1)
            if 'fecha_guia' in dd_det.columns:
                dd_det['_fecha_guia_dt'] = dd_det['fecha_guia'].apply(lambda x: parse_dt_dropi(x))

            # ── DEDUPLICACIÓN DE ÓRDENES ──
            # Cuando hay upsell/productos extra, Dropi crea filas duplicadas con el mismo
            # ID + FECHA + HORA + NOMBRE + TELÉFONO + GUÍA. Hay que consolidarlas en una sola
            # sumando las columnas financieras y conservando el resto de la primera fila.

            COLS_SUMA   = ['ganancia','costo_prod','flete','valor','costo_dev']   # se suman
            COLS_CONCAT = ['producto']                                              # se concatenan
            COLS_CLAVE  = ['id','fecha','hora_umov','nombre','telefono','guia']    # definen duplicado

            # Solo deduplicar si hay columnas clave disponibles
            claves_presentes = [c for c in COLS_CLAVE if c in dd_det.columns
                                and dd_det[c].notna().any()
                                and dd_det[c].astype(str).str.strip().ne('None').any()]

            if len(claves_presentes) >= 3:  # mínimo 3 claves para deduplicar
                n_antes = len(dd_det)

                # Normalizar claves para comparación
                for c in claves_presentes:
                    dd_det[f'_key_{c}'] = dd_det[c].astype(str).str.strip().str.upper()

                key_cols = [f'_key_{c}' for c in claves_presentes]

                # Convertir cols financieras a numérico antes de agrupar
                for col in COLS_SUMA:
                    if col in dd_det.columns:
                        dd_det[col] = pd.to_numeric(dd_det[col], errors='coerce').fillna(0)

                # Construir función de agregación
                agg_dict = {}
                for col in dd_det.columns:
                    if col.startswith('_key_') or col in key_cols:
                        continue
                    elif col in COLS_SUMA and col in dd_det.columns:
                        agg_dict[col] = 'sum'
                    elif col in COLS_CONCAT and col in dd_det.columns:
                        agg_dict[col] = lambda x: ' + '.join(
                            [v for v in x.astype(str).str.strip() if v not in ('','nan','None')]
                        )
                    else:
                        agg_dict[col] = 'first'

                dd_det = dd_det.groupby(key_cols, as_index=False).agg(agg_dict)
                # Limpiar columnas temporales de clave
                dd_det = dd_det.drop(columns=[c for c in dd_det.columns if c.startswith('_key_')], errors='ignore')

                # Reconstruir columnas datetime después de deduplicación
                # (groupby 'first' puede convertirlas a object)
                if 'fecha' in dd_det.columns:
                    dd_det['_fecha_dt'] = dd_det['fecha'].apply(lambda x: parse_dt_dropi(str(x)))
                if 'fecha_guia' in dd_det.columns:
                    dd_det['_fecha_guia_dt'] = dd_det['fecha_guia'].apply(lambda x: parse_dt_dropi(str(x)))
                if 'fecha_umov' in dd_det.columns:
                    dd_det['_dt_umov'] = dd_det.apply(
                        lambda r: parse_dt_dropi(str(r.get('fecha_umov','')), str(r.get('hora_umov',''))), axis=1
                    )
                # Asegurar que las columnas financieras sigan siendo numéricas
                for col in COLS_SUMA:
                    if col in dd_det.columns:
                        dd_det[col] = pd.to_numeric(dd_det[col], errors='coerce').fillna(0)

                n_desp = len(dd_det)
                n_dup  = n_antes - n_desp
                if n_dup > 0:
                    st.success(f"✅ Se detectaron y consolidaron **{n_dup} filas duplicadas** (upsells). "
                               f"De {n_antes:,} filas → {n_desp:,} órdenes únicas.")
                else:
                    st.info("✅ No se detectaron órdenes duplicadas en el archivo.")
            else:
                st.warning("⚠️ No hay suficientes columnas clave para detectar duplicados "
                           f"(se necesitan ID, FECHA, GUÍA — detectadas: {claves_presentes})")

        if dp_raw_det is not None:
            # Auto-map Pancake
            mp = {'id':find_col(dp_raw_det,['id del pedido','id pedido','id']),
                  'fecha':find_col(dp_raw_det,['creado en','fecha creacion','fecha','created at']),
                  'nombre':find_col(dp_raw_det,['nombre del cliente','nombre cliente']),
                  'telefono':find_col(dp_raw_det,['numero de telefono','telefono','celular']),
                  'estado':find_col(dp_raw_det,['estado del pedido','estado','status']),
                  'precio':find_col(dp_raw_det,['precio total','valor del pedido'])}
            dp_det = pd.DataFrame()
            for std,col in mp.items():
                dp_det[std] = dp_raw_det[col].astype(str).str.strip() if col else None
            if 'fecha' in dp_det.columns:
                dp_det['_fecha_dt'] = dp_det['fecha'].apply(lambda x: parse_dt_dropi(x))

        if dc_raw_det is not None:
            # Auto-map Cartera
            cc_tipo    = find_col(dc_raw_det,['tipo'])
            cc_monto   = find_col(dc_raw_det,['monto'])
            cc_desc    = find_col(dc_raw_det,['descripción','descripcion'])
            cc_fecha_c = find_col(dc_raw_det,['fecha'])
            dc_det = dc_raw_det.copy()
            dc_det['_tipo']   = dc_det[cc_tipo].str.upper().str.strip()  if cc_tipo  else ''
            dc_det['_monto']  = to_n(dc_det[cc_monto])                   if cc_monto else 0
            dc_det['_desc']   = dc_det[cc_desc].astype(str)              if cc_desc  else ''
            dc_det['_clasif'] = dc_det['_desc'].apply(clasif_desc)
            dc_det['_fecha']  = dc_det[cc_fecha_c].astype(str)           if cc_fecha_c else ''
            if cc_fecha_c:
                dc_det['_fecha_dt'] = dc_det[cc_fecha_c].apply(lambda x: parse_dt_dropi(str(x)[:10]))

        # Sincronizar button
        if dp_raw_det is not None and dd_raw_det is not None:
            # Show detected date ranges before syncing
            if dd_det is not None and '_fecha_dt' in dd_det.columns and dd_det['_fecha_dt'].notna().any():
                min_d = dd_det['_fecha_dt'].min()
                max_d = dd_det['_fecha_dt'].max()
                st.caption(f"📅 Fechas detectadas en Dropi: {f"{min_d.day:02d} {MESES_ES[min_d.month]} {min_d.year}" if pd.notna(min_d) else '?'} → {f"{max_d.day:02d} {MESES_ES[max_d.month]} {max_d.year}" if pd.notna(max_d) else '?'}  ·  Rango: {fmt_fecha_mes(inf['date_from'])} – {fmt_fecha_mes(inf['date_to'])}")
            st.markdown("---")
            st.markdown("---")
            if st.button("🔄 Sincronizar ahora", key=f"run_sync_{inf['id']}", use_container_width=True):
                # Apply date filter
                dF = pd.Timestamp(inf['date_from_dt'])
                dT = pd.Timestamp(inf['date_to_dt'])

                p_filt = dp_det.copy() if dp_det is not None else pd.DataFrame()
                d_filt = dd_det.copy()

                # Filtro Pancake por fecha
                if '_fecha_dt' in p_filt.columns and p_filt['_fecha_dt'].notna().any():
                    p_filt_dated = p_filt[(p_filt['_fecha_dt']>=dF)&(p_filt['_fecha_dt']<=dT)]
                    if len(p_filt_dated) > 0:
                        p_filt = p_filt_dated

                # Filtro Dropi por fecha — solo si hay fechas válidas y hay filas en el rango
                if '_fecha_dt' in d_filt.columns and d_filt['_fecha_dt'].notna().any():
                    d_filt_dated = d_filt[(d_filt['_fecha_dt']>=dF)&(d_filt['_fecha_dt']<=dT)]
                    if len(d_filt_dated) > 0:
                        d_filt = d_filt_dated
                        st.info(f"📅 Filtro de fechas aplicado: {len(d_filt):,} filas · {fmt_fecha_mes(inf['date_from'])} – {fmt_fecha_mes(inf['date_to'])}")
                    else:
                        # No hay filas en el rango — usar todo el archivo con aviso
                        st.warning(f"⚠️ No hay filas de Dropi en el rango {inf['date_from']} – {inf['date_to']}. "
                                   f"Usando el archivo completo ({len(d_filt):,} filas). "
                                   f"Verifica que la columna FECHA esté mapeada correctamente.")

                totalOrd = len(p_filt)
                dTot     = len(d_filt)

                # ══════════════════════════════════════════════════════
                # OPCIÓN A: DataFrame RAW con nombres EXACTOS del Excel
                # Para todos los cálculos financieros se usa raw_filt
                # que lee directamente de dd_raw_det sin mapeo ni dedup
                # ══════════════════════════════════════════════════════
                RAW_COLS = {
                    'estatus':       ['ESTATUS'],
                    'fecha':         ['FECHA'],
                    'transportadora':['TRANSPORTADORA'],
                    'ciudad':        ['CIUDAD DESTINO'],
                    'departamento':  ['DEPARTAMENTO DESTINO'],
                    # Valor COD — distintos nombres según tipo de reporte Dropi
                    'valor':         ['TOTAL DE LA ORDEN','VALOR FACTURADO','VALOR COD'],
                    'ganancia':      ['GANANCIA'],
                    # Costo producto — distintos nombres según reporte
                    'costo_prod':    ['PRECIO PROVEEDOR X CANTIDAD','VALOR DE COMPRA EN PRODUCTOS','TOTAL EN PRECIOS DE PROVEEDOR'],
                    'flete':         ['PRECIO FLETE'],
                    'costo_dev':     ['COSTO DEVOLUCION FLETE'],
                    # Fecha guía — DISTINTO nombre en cada reporte
                    'fecha_guia':    ['FECHA GUIA GENERADA','FECHA GENERACION DE GUIA'],
                    'fecha_umov':    ['FECHA DE ÚLTIMO MOVIMIENTO','FECHA DE ULTIMO MOVIMIENTO'],
                    'hora_umov':     ['HORA DE ÚLTIMO MOVIMIENTO','HORA DE ULTIMO MOVIMIENTO'],
                    'vendedor':      ['VENDEDOR'],
                    'guia':          ['NÚMERO GUIA','NUMERO GUIA'],
                    'nombre':        ['NOMBRE CLIENTE'],
                    'novedad':       ['NOVEDAD'],
                }

                raw_filt = pd.DataFrame()
                # Usar el archivo financiero (1 orden por fila) si está disponible
                # Si no, intentar con el logístico como fallback
                raw_source = dd_fin_raw if dd_fin_raw is not None else dd_raw_det
                if raw_source is not None:
                    st.info(f"💰 Calculando métricas financieras desde: {'Dropi Financiero' if dd_fin_raw is not None else 'Dropi Logístico'} ({len(raw_source):,} filas)")
                    raw_filt = raw_source.copy()
                    # Normalizar nombres de columnas
                    raw_filt.columns = [c.strip() for c in raw_filt.columns]

                    # Crear columnas estándar desde nombres exactos
                    for std, cands in RAW_COLS.items():
                        col_found = find_col(raw_filt, cands)
                        if col_found:
                            raw_filt[f'_r_{std}'] = raw_filt[col_found]
                        else:
                            raw_filt[f'_r_{std}'] = None

                    # Parsear estatus para filtrar
                    raw_filt['_r_estatus_str'] = raw_filt['_r_estatus'].astype(str).str.strip().str.upper()

                    # Parsear fecha para filtro de rango
                    raw_filt['_r_fecha_dt'] = raw_filt['_r_fecha'].apply(
                        lambda x: parse_dt_dropi(str(x)) if pd.notna(x) and str(x).strip() not in ('','nan','None') else pd.NaT
                    )

                    # Parsear columnas financieras a numérico
                    for fc in ['valor','ganancia','costo_prod','flete','costo_dev']:
                        raw_filt[f'_r_{fc}_n'] = pd.to_numeric(raw_filt[f'_r_{fc}'], errors='coerce').fillna(0)

                    # Parsear fechas para días de entrega
                    raw_filt['_r_fguia_dt'] = raw_filt['_r_fecha_guia'].apply(
                        lambda x: parse_dt_dropi(str(x)) if pd.notna(x) and str(x).strip() not in ('','nan','None') else pd.NaT
                    )
                    raw_filt['_r_fumov_dt'] = raw_filt.apply(
                        lambda r: parse_dt_dropi(
                            str(r['_r_fecha_umov']) if pd.notna(r.get('_r_fecha_umov')) else '',
                            str(r['_r_hora_umov'])  if pd.notna(r.get('_r_hora_umov'))  else ''
                        ), axis=1
                    )

                    # Deduplicar sobre RAW también (sumar columnas financieras)
                    raw_key_cands = ['NÚMERO GUIA','numero guia','NUMERO GUIA','FECHA','fecha']
                    raw_guia_col = find_col(raw_filt, ['NÚMERO GUIA','NUMERO GUIA','numero guia'])
                    raw_fecha_col = find_col(raw_filt, ['FECHA','fecha'])
                    if raw_guia_col and raw_fecha_col:
                        raw_filt['_dedup_key'] = (
                            raw_filt[raw_guia_col].astype(str).str.strip() + '|' +
                            raw_filt[raw_fecha_col].astype(str).str.strip()
                        )
                        # Sumar columnas financieras por grupo ANTES de deduplicar
                        fin_cols = [f'_r_{fc}_n' for fc in ['valor','ganancia','costo_prod','flete','costo_dev']
                                    if f'_r_{fc}_n' in raw_filt.columns]
                        fin_sums = raw_filt.groupby('_dedup_key')[fin_cols].sum()
                        # Keep first row per key for non-financial columns
                        raw_filt = raw_filt.drop_duplicates(subset=['_dedup_key'], keep='first').copy()
                        # Overwrite financial cols with summed values
                        for fc_col in fin_cols:
                            raw_filt[fc_col] = raw_filt['_dedup_key'].map(fin_sums[fc_col])

                    # Aplicar filtro de fechas al raw
                    if raw_filt['_r_fecha_dt'].notna().any():
                        raw_dated = raw_filt[(raw_filt['_r_fecha_dt']>=dF)&(raw_filt['_r_fecha_dt']<=dT)]
                        if len(raw_dated) > 0:
                            raw_filt = raw_dated

                    # Máscaras por estado
                    raw_filt['_r_es_e'] = raw_filt['_r_estatus_str'].apply(lambda s: norm_estado(s)=='ENTREGADO')
                    raw_filt['_r_es_d'] = raw_filt['_r_estatus_str'].apply(lambda s: norm_estado(s)=='DEVOLUCION')

                def raw_num(col_n):
                    """Lee columna numérica del raw_filt por nombre estándar."""
                    c = f'_r_{col_n}_n'
                    if c in raw_filt.columns:
                        return pd.to_numeric(raw_filt[c], errors='coerce').fillna(0)
                    return pd.Series([0.0]*len(raw_filt))

                def raw_flete_by_carrier():
                    """Flete promedio por transportadora — solo ENTREGADAS, post-dedup."""
                    result = {}
                    if raw_filt.empty or '_r_transportadora' not in raw_filt.columns: return result
                    df_e = raw_filt[raw_filt['_r_es_e']==True]
                    for t, g in df_e.groupby('_r_transportadora'):
                        ts = str(t).strip()
                        if ts in ('nan','None',''): continue
                        vals = pd.to_numeric(g['_r_flete_n'], errors='coerce').dropna()
                        result[ts] = round(float(vals.mean()), 2) if len(vals) > 0 else 0
                    return result

                def raw_cdev_by_carrier():
                    """Costo devolución promedio — solo DEVOLUCION con costo > 0, post-dedup."""
                    result = {}
                    if raw_filt.empty or '_r_transportadora' not in raw_filt.columns: return result
                    # Solo órdenes DEVOLUCION con costo > 0 (los $0 son órdenes sin costo real)
                    df_d = raw_filt[raw_filt['_r_es_d']==True]
                    df_d_real = df_d[df_d['_r_costo_dev_n'] > 0]
                    for t, g in df_d_real.groupby('_r_transportadora'):
                        ts = str(t).strip()
                        if ts in ('nan','None',''): continue
                        vals = g['_r_costo_dev_n']
                        result[ts] = round(float(vals.mean()), 2) if len(vals) > 0 else 0
                    return result

                def raw_dias_by_carrier():
                    """
                    Días promedio de entrega por transportadora.
                    Usa el cronograma de Pancake: fecha 'Enviado' → fecha 'Entregado'.
                    Cruza por número de guía con Dropi para obtener transportadora.
                    """
                    result = {}

                    # ── Intentar con Pancake cronograma ──
                    panc_src = dp_raw_det if dp_raw_det is not None else st.session_state.get(f"raw_panc_{inf['id']}")
                    dropi_src = raw_filt if not raw_filt.empty else None

                    if panc_src is not None and dropi_src is not None:
                        try:
                            # Encontrar columna cronograma en Pancake
                            cron_col  = find_col(panc_src, ['cronograma de actualización de estado',
                                                            'cronograma de actualizacion de estado',
                                                            'cronograma','status schedule'])
                            guia_panc = find_col(panc_src, ['número de seguimiento','numero de seguimiento',
                                                            'tracking','guia','número guia'])

                            if cron_col and guia_panc:
                                def _parse_cron(text):
                                    events = []
                                    if not text or str(text).strip() in ('nan','None',''): return events
                                    for entry in str(text).split(';'):
                                        m = re.match(
                                            r'^(.+?)\s*-\s*(\d{2}:\d{2}\s+\d{2}/\d{2}/\d{4})\s*-\s*(.+)$',
                                            entry.strip()
                                        )
                                        if m:
                                            try:
                                                dt = datetime.strptime(m.group(2).strip(), '%H:%M %d/%m/%Y')
                                                events.append((m.group(1).strip().lower(), dt))
                                            except: pass
                                    return events

                                def _dias_cron(text):
                                    events = _parse_cron(text)
                                    env_dt = ent_dt = None
                                    for estado, dt in events:
                                        if 'envi' in estado and 'devolu' not in estado:
                                            env_dt = dt
                                        if 'entregado' in estado:
                                            ent_dt = dt
                                    if env_dt and ent_dt and ent_dt >= env_dt:
                                        d = (ent_dt - env_dt).days
                                        return d if d > 0 else None
                                    return None

                                # Calcular días por guía en Pancake
                                panc_df = panc_src.copy()
                                panc_df['_guia_n'] = pd.to_numeric(panc_df[guia_panc], errors='coerce')
                                panc_df['_dias']   = panc_df[cron_col].apply(_dias_cron)
                                guia_dias = panc_df[panc_df['_dias'].notna() & panc_df['_guia_n'].notna()]\
                                            .set_index('_guia_n')['_dias'].to_dict()

                                if guia_dias:
                                    # Cruzar con Dropi para obtener transportadora
                                    guia_trans_col = find_col(dropi_src, ['NÚMERO GUIA','numero guia','guia'])
                                    trans_col      = find_col(dropi_src, ['TRANSPORTADORA','transportadora'])

                                    if guia_trans_col and trans_col:
                                        dropi_df = dropi_src.copy()
                                        dropi_df['_guia_n'] = pd.to_numeric(dropi_df[guia_trans_col], errors='coerce')
                                        dropi_df['_dias']   = dropi_df['_guia_n'].map(guia_dias)
                                        df_match = dropi_df[dropi_df['_dias'].notna() & (dropi_df['_dias'] > 0)]

                                        for trans, g in df_match.groupby(trans_col):
                                            ts = str(trans).strip()
                                            if ts in ('nan','None',''): continue
                                            result[ts] = round(float(g['_dias'].mean()), 1)

                                        if result:
                                            return result  # éxito con Pancake
                        except Exception as e_cron:
                            pass  # fallback a método por fechas

                    # ── Fallback: usar fechas de Dropi (FECHA GUIA → FECHA ÚLTIMO MOV) ──
                    if raw_filt.empty or '_r_transportadora' not in raw_filt.columns:
                        return result
                    df_e = raw_filt[raw_filt.get('_r_es_e', pd.Series(dtype=bool))==True].copy() \
                           if '_r_es_e' in raw_filt.columns else pd.DataFrame()
                    if df_e.empty: return result
                    for t, g in df_e.groupby('_r_transportadora'):
                        ts = str(t).strip()
                        if ts in ('nan','None',''): continue
                        try:
                            fg = pd.to_datetime(g['_r_fguia_dt'], errors='coerce')
                            fm = pd.to_datetime(g['_r_fumov_dt'],  errors='coerce')
                            dias = (fm - fg).dt.days.dropna()
                            dias = dias[(dias > 0) & (dias < 365)]
                            result[ts] = round(float(dias.mean()), 1) if len(dias) > 0 else 0
                        except: result[ts] = 0
                    return result

                # Helper: acceso seguro a columna — devuelve Serie numérica o ceros
                def safe_col(df, col):
                    if col not in df.columns:
                        return pd.Series([0.0]*len(df), index=df.index)
                    try:
                        serie = df[col]
                        if not isinstance(serie, pd.Series):
                            return pd.Series([0.0]*len(df), index=df.index)
                        return to_n(serie)
                    except Exception:
                        return pd.Series([0.0]*len(df), index=df.index)

                def has_col(df, col):
                    if col not in df.columns: return False
                    try: return isinstance(df[col], pd.Series)
                    except: return False

                mask_e   = d_filt['estatus'].apply(es_e)
                mask_d   = d_filt['estatus'].apply(es_d)
                mask_can = d_filt['estatus'].apply(es_cancelacion)
                mask_s   = d_filt['estatus'].apply(es_s)
                mask_des = d_filt['estatus'].apply(es_despachado)

                entregado  = int(mask_e.sum())
                devolucion = int(mask_d.sum())
                cancelados = int(mask_can.sum())
                sinFinal   = int(mask_s.sum())
                despachados= int(mask_des.sum())
                pendGuia   = max(0, totalOrd - dTot)

                # Usar RAW para valores financieros correctos
                if not raw_filt.empty:
                    gan_ent = float(raw_filt[raw_filt['_r_es_e']==True]['_r_ganancia_n'].sum())
                    cos_ent = float(raw_filt[raw_filt['_r_es_e']==True]['_r_costo_prod_n'].sum())
                    fle_ent = float(raw_filt[raw_filt['_r_es_e']==True]['_r_flete_n'].sum())
                    val_ent = float(raw_filt[raw_filt['_r_es_e']==True]['_r_valor_n'].sum())
                    cdev    = float(raw_filt[raw_filt['_r_es_d']==True]['_r_costo_dev_n'].sum())
                else:
                    gan_ent = float(safe_col(d_filt,'ganancia')[mask_e].sum())
                    cos_ent = float(safe_col(d_filt,'costo_prod')[mask_e].sum())
                    fle_ent = float(safe_col(d_filt,'flete')[mask_e].sum())
                    val_ent = float(safe_col(d_filt,'valor')[mask_e].sum())
                    cdev    = float(safe_col(d_filt,'costo_dev')[mask_d].sum())
                ingresos = val_ent if val_ent > 0 else (gan_ent + cos_ent + fle_ent)

                estcount={}
                for k,v in d_filt['estatus'].fillna('Sin estado').value_counts().items():
                    estcount[str(k)]=int(v)

                byCarrier={}
                if has_col(d_filt, 'transportadora'):
                    try:
                        for t,g in d_filt.groupby('transportadora'):
                            if str(t) in ('nan','None',''): continue
                            byCarrier[str(t)]={
                                'total':len(g),
                                'entregado':int(g['estatus'].apply(es_e).sum()),
                                'devolucion':int(g['estatus'].apply(es_d).sum()),
                                'sinFinal':int(g['estatus'].apply(es_s).sum()),
                                'ganancia':float(safe_col(g,'ganancia')[g['estatus'].apply(es_e)].sum()),
                                'costo_dev':float(safe_col(g,'costo_dev')[g['estatus'].apply(es_d)].sum()),
                            }
                    except Exception: pass

                byCiudad={}
                if has_col(d_filt, 'ciudad'):
                    try:
                        for c,g in d_filt.groupby('ciudad'):
                            if str(c) in ('nan','None',''): continue
                            byCiudad[str(c)]={'total':len(g),'entregado':int(g['estatus'].apply(es_e).sum()),'devolucion':int(g['estatus'].apply(es_d).sum())}
                    except Exception: pass

                byDepto={}
                if has_col(d_filt, 'departamento'):
                    try:
                        for c,g in d_filt.groupby('departamento'):
                            if str(c) in ('nan','None',''): continue
                            byDepto[str(c)]={'total':len(g),'entregado':int(g['estatus'].apply(es_e).sum()),'devolucion':int(g['estatus'].apply(es_d).sum())}
                    except Exception: pass

                def es_nov_act(s): return norm_estado(s) in {'NOVEDAD','NOVEDAD SOLUCIONADA'}
                def tiene_nov(row):
                    return es_nov_act(str(row.get('estatus',''))) or (str(row.get('novedad','')).strip().lower() not in ['','nan','none'])
                try:
                    novedades_list = d_filt[d_filt.apply(tiene_nov, axis=1)]
                    totalNov    = len(novedades_list)
                    novAbiertas = int((novedades_list['estatus'].apply(lambda s: norm_estado(s)!='NOVEDAD SOLUCIONADA')).sum()) if totalNov else 0
                except Exception:
                    totalNov = 0; novAbiertas = 0

                snap = {
                    'timestamp': datetime.now().isoformat(),
                    'totalOrd':totalOrd, 'dTot':dTot, 'despachados':despachados,
                    'pendGuia':pendGuia, 'cancelados':cancelados,
                    'pctConf':round(pct(despachados,totalOrd),2),
                    'entregado':entregado,'devolucion':devolucion,'sinFinal':sinFinal,
                    'pctE':round(pct(entregado,despachados),2),
                    'pctD':round(pct(devolucion,despachados),2),
                    'pctS':round(pct(sinFinal,despachados),2),
                    'pctP':round(pct(pendGuia,totalOrd),2),
                    'totalGanancia':float(gan_ent),'totalCosto':float(cos_ent),
                    'totalFlete':float(fle_ent),'totalCostoDev':float(cdev),
                    'ingresos':float(ingresos),
                    'totalNovedades':totalNov,'novAbiertas':novAbiertas,
                    'estatusCount':estcount,'byCarrier':byCarrier,
                    'byCiudad':byCiudad,'byDepto':byDepto,
                    'ordenesLite': d_filt.head(500).to_dict('records'),
                }

                # ── DATOS ENRIQUECIDOS PARA INFORME LOGÍSTICO ──
                def _build_geo(df_in, gcol, ref_tot):
                    result = {}
                    if not has_col(df_in, gcol): return result
                    try:
                        dfc = df_in.copy()
                        dfc['__cat'] = dfc['estatus'].apply(cat_log)

                        # Parsear fechas DIRECTO desde columnas de texto
                        has_fg = 'fecha_guia' in dfc.columns and dfc['fecha_guia'].notna().any()
                        has_fu = 'fecha_umov' in dfc.columns and dfc['fecha_umov'].notna().any()
                        if has_fg:
                            dfc['__fguia'] = dfc['fecha_guia'].apply(
                                lambda x: parse_dt_dropi(str(x)) if pd.notna(x) and str(x).strip() not in ('','nan','None') else pd.NaT
                            )
                        if has_fu:
                            dfc['__fumov'] = dfc.apply(
                                lambda r: parse_dt_dropi(
                                    str(r['fecha_umov']) if pd.notna(r.get('fecha_umov')) else '',
                                    str(r['hora_umov'])  if pd.notna(r.get('hora_umov'))  else ''
                                ), axis=1
                            )

                        for name, g in dfc.groupby(gcol):
                            ns = str(name)
                            if ns in ('nan','None',''): continue
                            tot = len(g)
                            row = {'total':tot,'valor':float(safe_col(g,'valor').sum()),
                                   'pct_total':round(pct(tot,ref_tot),2)}
                            for cat in LOG_CATS:
                                cnt = int((g['__cat']==cat).sum())
                                row[cat]=cnt; row[f'pct_{cat}']=round(pct(cnt,tot),2)
                            g_ent = g[g['__cat']=='ENTREGADO']
                            g_dev = g[g['__cat']=='DEVOLUCION']

                            # Flete promedio — ENTREGADAS
                            fle_s = pd.to_numeric(g_ent['flete'], errors='coerce').dropna() if ('flete' in g_ent.columns and len(g_ent)>0) else pd.Series([], dtype=float)
                            fle_s = fle_s[fle_s > 0]
                            row['flete_avg'] = round(float(fle_s.mean()), 2) if len(fle_s) > 0 else 0

                            # Costo devolución — DEVOLUCION
                            cdev_s = pd.to_numeric(g_dev['costo_dev'], errors='coerce').dropna() if ('costo_dev' in g_dev.columns and len(g_dev)>0) else pd.Series([], dtype=float)
                            cdev_s = cdev_s[cdev_s > 0]
                            row['cdev_avg'] = round(float(cdev_s.mean()), 2) if len(cdev_s) > 0 else 0

                            # Días entrega — FECHA GUIA → FECHA ÚLTIMO MOV (ENTREGADO)
                            dias_avg = 0
                            if len(g_ent) > 0 and has_fg and has_fu:
                                try:
                                    fguia = pd.to_datetime(g_ent['__fguia'], errors='coerce')
                                    fumov = pd.to_datetime(g_ent['__fumov'], errors='coerce')
                                    dias  = (fumov - fguia).dt.total_seconds() / 86400
                                    dias  = dias.dropna()
                                    dias  = dias[(dias > 0) & (dias < 365)]  # excluir 0 y outliers
                                    if len(dias) > 0:
                                        dias_avg = round(float(dias.mean()), 1)
                                except: pass
                            row['dias_avg'] = dias_avg
                            result[ns] = row
                    except: pass
                    return result

                byEstadoCat_s = {}
                try:
                    dfc2=d_filt.copy(); dfc2['__cat']=dfc2['estatus'].apply(cat_log)
                    for cat,g in dfc2.groupby('__cat'):
                        byEstadoCat_s[cat]={'count':len(g),'valor':float(safe_col(g,'valor').sum()),'pct':round(pct(len(g),dTot),2)}
                except: pass

                byCiudadFull_s  = _build_geo(d_filt,'ciudad',dTot)
                byDeptoFull_s   = _build_geo(d_filt,'departamento',dTot)

                byCarrierFull_s = {}
                if not raw_filt.empty and '_r_transportadora' in raw_filt.columns:
                    try:
                        flete_by_carrier = raw_flete_by_carrier()
                        cdev_by_carrier  = raw_cdev_by_carrier()
                        dias_by_carrier  = raw_dias_by_carrier()

                        # Usar d_filt para counts de estados (es más confiable para estatus)
                        cat_log_map = {}
                        if has_col(d_filt,'transportadora'):
                            dfc3 = d_filt.copy()
                            dfc3['__cat'] = dfc3['estatus'].apply(cat_log)
                            for t, g in dfc3.groupby('transportadora'):
                                ns = str(t).strip()
                                if ns in ('nan','None',''): continue
                                tot = len(g)
                                row = {'total':tot,'valor':float(safe_col(g,'valor').sum())}
                                for cat in LOG_CATS:
                                    cnt = int((g['__cat']==cat).sum())
                                    row[cat]=cnt; row[f'pct_{cat}']=round(pct(cnt,tot),2)
                                row['ganancia']  = float(safe_col(g[g['__cat']=='ENTREGADO'],'ganancia').sum())
                                row['flete_avg'] = flete_by_carrier.get(ns, 0)
                                row['cdev_avg']  = cdev_by_carrier.get(ns, 0)
                                row['dias_avg']  = dias_by_carrier.get(ns, 0)
                                byCarrierFull_s[ns] = row
                    except Exception as e_carrier:
                        st.warning(f"⚠️ Error transportadoras: {e_carrier}")

                novData_s={'total_con':0,'total_sin':0,'entregadas':0,'devueltas':0,'pendientes':0}
                try:
                    def _tnov(row): return norm_estado(str(row.get('estatus',''))) in {'NOVEDAD','NOVEDAD SOLUCIONADA'} or str(row.get('novedad','')).strip().lower() not in ['','nan','none']
                    mnv=d_filt.apply(_tnov,axis=1); dfnv=d_filt[mnv]
                    novData_s.update({'total_con':len(dfnv),'total_sin':dTot-len(dfnv),
                        'entregadas':int(dfnv['estatus'].apply(es_e).sum()),
                        'devueltas':int(dfnv['estatus'].apply(es_d).sum()),
                        'pendientes':len(dfnv)-int(dfnv['estatus'].apply(es_e).sum())-int(dfnv['estatus'].apply(es_d).sum())})
                except: pass

                pendPorFecha_s={}
                pendPorFechaLite_s={}
                try:
                    dfpf=d_filt[d_filt['estatus'].apply(es_s)].copy()
                    dcol='_fecha_guia_dt' if '_fecha_guia_dt' in dfpf.columns else '_fecha_dt'
                    if dcol in dfpf.columns and dfpf[dcol].notna().any():
                        dfpf['_fd']=dfpf[dcol].dt.strftime('%d/%m')
                        pendPorFecha_s={k:int(v) for k,v in dfpf.groupby('_fd').size().items()}
                        # Guardar detalle de órdenes por fecha para drill-down
                        cols_lite=['guia','nombre','telefono','ciudad','transportadora',
                                   'estatus','novedad','ult_mov','fecha_umov','vendedor','valor','flete']
                        cols_ok=[c for c in cols_lite if c in dfpf.columns]
                        for fecha_k, grp in dfpf.groupby('_fd'):
                            pendPorFechaLite_s[fecha_k] = grp[cols_ok].head(200).to_dict('records')
                except: pass

                devSinNovLite_s=[]
                try:
                    dfdsn=d_filt[d_filt['estatus'].apply(es_d)].copy()
                    mno=dfdsn.apply(lambda r:str(r.get('novedad','')).strip().lower() in ['','nan','none'],axis=1)
                    for _,row in dfdsn[mno].head(300).iterrows():
                        devSinNovLite_s.append({'guia':str(row.get('guia','')),'fecha':str(row.get('fecha','')),
                            'ciudad':str(row.get('ciudad','')),'transportadora':str(row.get('transportadora','')),
                            'valor':float(to_n(pd.Series([row.get('valor','0')])).iloc[0]),
                            'flete':float(to_n(pd.Series([row.get('flete','0')])).iloc[0])})
                except: pass

                byVendedor_s={}
                if has_col(d_filt,'vendedor'):
                    try:
                        dfvv=d_filt.copy(); dfvv['__cat']=dfvv['estatus'].apply(cat_log)
                        for v,g in dfvv.groupby('vendedor'):
                            ns=str(v)
                            if ns in ('nan','None',''): continue
                            row={'total':len(g)}
                            for cat in LOG_CATS: row[cat]=int((g['__cat']==cat).sum())
                            row['ganancia']=float(safe_col(g[g['__cat']=='ENTREGADO'],'ganancia').sum())
                            byVendedor_s[ns]=row
                    except: pass

                byCarrierCiudad_s={}; byCarrierDepto_s={}
                if has_col(d_filt,'transportadora'):
                    try:
                        for carrier,gC in d_filt.groupby('transportadora'):
                            cs=str(carrier)
                            if cs in ('nan','None',''): continue
                            byCarrierCiudad_s[cs]=_build_geo(gC,'ciudad',len(gC))
                            byCarrierDepto_s[cs] =_build_geo(gC,'departamento',len(gC))
                    except: pass

                # ── DEBUG: mostrar valores calculados para verificar mapeo ──
                with st.expander("🔍 Debug — Valores calculados en sincronización", expanded=True):
                    st.markdown("**Columnas disponibles en Dropi:**")
                    st.write(list(d_filt.columns))

                    st.markdown("**Muestra PRECIO FLETE (primeras 5 ENTREGADAS):**")
                    df_ent_dbg = d_filt[d_filt['estatus'].apply(es_e)]
                    if 'flete' in df_ent_dbg.columns:
                        st.write(df_ent_dbg['flete'].head(5).tolist())
                        st.write(f"→ Numérico: {pd.to_numeric(df_ent_dbg['flete'], errors='coerce').dropna().head(5).tolist()}")
                    else:
                        st.warning("Columna 'flete' NO encontrada")

                    st.markdown("**Muestra COSTO DEVOLUCION FLETE (primeras 5 DEVOLUCION):**")
                    df_dev_dbg = d_filt[d_filt['estatus'].apply(es_d)]
                    if 'costo_dev' in df_dev_dbg.columns:
                        st.write(df_dev_dbg['costo_dev'].head(5).tolist())
                        st.write(f"→ Numérico: {pd.to_numeric(df_dev_dbg['costo_dev'], errors='coerce').dropna().head(5).tolist()}")
                    else:
                        st.warning("Columna 'costo_dev' NO encontrada")

                    st.markdown("**Muestra FECHA GUIA GENERADA (primeras 5 ENTREGADAS):**")
                    if 'fecha_guia' in df_ent_dbg.columns:
                        st.write(df_ent_dbg['fecha_guia'].head(5).tolist())
                    else:
                        st.warning("Columna 'fecha_guia' NO encontrada")

                    st.markdown("**Muestra FECHA ÚLTIMO MOVIMIENTO (primeras 5 ENTREGADAS):**")
                    if 'fecha_umov' in df_ent_dbg.columns:
                        st.write(df_ent_dbg['fecha_umov'].head(5).tolist())
                    else:
                        st.warning("Columna 'fecha_umov' NO encontrada")

                    st.markdown("**Resultado byCarrierFull:**")
                    for carrier, d in byCarrierFull_s.items():
                        st.write(f"**{carrier}**: flete_avg={d.get('flete_avg',0)}, cdev_avg={d.get('cdev_avg',0)}, dias_avg={d.get('dias_avg',0)}")

                snap.update({'byEstadoCat':byEstadoCat_s,'byCiudadFull':byCiudadFull_s,
                    'byDeptoFull':byDeptoFull_s,'byCarrierFull':byCarrierFull_s,
                    'novData':novData_s,'pendPorFecha':pendPorFecha_s,'pendPorFechaLite':pendPorFechaLite_s,
                    'devSinNovLite':devSinNovLite_s,'byVendedor':byVendedor_s,
                    'byCarrierCiudad':byCarrierCiudad_s,'byCarrierDepto':byCarrierDepto_s})

                inf['history'].append(snap)
                st.success(f"✅ Sincronizado — {totalOrd} Pancake · {dTot} Dropi · {entregado} entregadas")
                st.rerun()

    # ── USAR ÚLTIMO SNAPSHOT ──
    snap = inf['history'][-1] if inf.get('history') else None

    # Use files if loaded in session for display
    d_  = dd_det  if dd_det  is not None and len(dd_det)>0  else None
    p_  = dp_det  if dp_det  is not None and len(dp_det)>0  else None
    dc_ = dc_det  if dc_det  is not None and len(dc_det)>0  else None

    if not snap:
        st.markdown("""
        <div style="text-align:center;padding:60px 20px;color:#475569">
            <div style="font-size:40px;margin-bottom:14px;opacity:.4">🔄</div>
            <div style="font-size:15px;font-weight:600;color:#64748b;margin-bottom:6px">Sin datos aún</div>
            <div style="font-size:13px">Carga los archivos arriba y presiona <strong style="color:#0ea5e9">Sincronizar ahora</strong></div>
        </div>""", unsafe_allow_html=True)
        st.stop()

    st.markdown("---")

    if tipo_inf == "📦 Informe Logístico":

        # ── Datos del snapshot ──
        byEstadoCat   = snap.get('byEstadoCat',   {})
        byCiudadFull  = snap.get('byCiudadFull',  snap.get('byCiudad', {}))
        byDeptoFull   = snap.get('byDeptoFull',   snap.get('byDepto', {}))
        byCarrierFull = snap.get('byCarrierFull', snap.get('byCarrier', {}))
        novData       = snap.get('novData', {})
        pendPorFecha     = snap.get('pendPorFecha', {})
        pendPorFechaLite = snap.get('pendPorFechaLite', {})
        devSinNovLite = snap.get('devSinNovLite', [])
        byVendedor    = snap.get('byVendedor', {})
        byCarrierCiudad = snap.get('byCarrierCiudad', {})
        byCarrierDepto  = snap.get('byCarrierDepto',  {})
        maxDev        = inf.get('max_dev', 20)

        desp_valor = sum(v.get('valor',0) for k,v in byEstadoCat.items() if k != 'CANCELACION')
        total_dTot = snap.get('dTot', snap.get('despachados',0))

        if snap['pctD'] > maxDev:
            st.markdown(f'<div class="alerta-card">⚠️ Devoluciones en <strong>{snap["pctD"]:.1f}%</strong> — superan el límite del {maxDev}%</div>', unsafe_allow_html=True)

        # ── Helper tablas geo ──
        def build_geo_df(geo_dict, col_name, ref_tot, maxDev_threshold=20):
            rows = []
            for name, d in sorted(geo_dict.items(), key=lambda x: -x[1].get('total',0)):
                tot = d.get('total',0)
                row = {col_name: name, 'TOTAL': tot, '%': f"{pct(tot,ref_tot):.2f}%"}
                for cat in LOG_CATS:
                    cnt = d.get(cat, 0)
                    pp  = d.get(f'pct_{cat}', pct(cnt,tot))
                    row[cat]        = cnt
                    pct_str = f"{pp:.2f}%"
                    # Mark devolución % in red if above maxDev
                    if cat == 'DEVOLUCION' and pp > maxDev_threshold:
                        pct_str = f"⚠️ {pp:.2f}%"
                    row[f'% {cat}'] = pct_str
                rows.append(row)
            return pd.DataFrame(rows) if rows else pd.DataFrame()

        # ════════════════════════════
        # 1. ESTADO DE ÓRDENES
        # ════════════════════════════
        st.markdown('<div class="fin-section-title">📊 Estado de órdenes</div>', unsafe_allow_html=True)
        eo_l, eo_r = st.columns([1, 1.4])

        with eo_l:
            cats_donut  = [k for k in LOG_CATS if k in byEstadoCat and byEstadoCat[k].get('count',0)>0]
            vals_donut  = [byEstadoCat[k]['count'] for k in cats_donut]
            cols_donut  = [CAT_CMAP.get(k,'#64748b') for k in cats_donut]

            fig_eo = go.Figure(go.Pie(
                labels=cats_donut or ['Sin datos'], values=vals_donut or [1], hole=0.65,
                marker=dict(colors=cols_donut, line=dict(color='#0f172a',width=2)),
                textinfo='none',
                hovertemplate='<b>%{label}</b><br>%{value} · %{percent}<extra></extra>'
            ))
            fig_eo.add_annotation(text=f"<b>{snap['pctConf']:.2f}%</b>",
                x=0.5,y=0.58,showarrow=False,font=dict(size=24,color='#f1f5f9',family='Inter'))
            fig_eo.update_layout(**PLOT_BG,height=290,xaxis=AX,yaxis=AX,showlegend=False,
                title=dict(text=f"Monto total: {fmt_cop(desp_valor)}  ·  {snap['despachados']:,} órdenes ({snap['pctConf']:.2f}% confirmadas)",
                    font=dict(color='#94a3b8',size=11),x=0.01))
            st.plotly_chart(fig_eo, use_container_width=True)

            mc1,mc2,mc3 = st.columns(3)
            can_d = byEstadoCat.get('CANCELACION',{})
            with mc1: st.markdown(f'<div class="big-kpi" style="text-align:center"><div style="font-size:22px;font-weight:800;color:{C["teal"]}">{snap["despachados"]:,}</div><div style="font-size:10px;color:#64748b;margin:4px 0">{fmt_cop(desp_valor)}</div><div style="font-size:10px;color:#64748b">{snap["despachados"]:,} con guía</div></div>', unsafe_allow_html=True)
            with mc2: st.markdown(f'<div class="big-kpi" style="text-align:center"><div style="font-size:22px;font-weight:800;color:{C["can"]}">{snap.get("pendGuia",0):,}</div><div style="font-size:10px;color:#64748b;margin:4px 0">&nbsp;</div><div style="font-size:10px;color:#64748b">{snap.get("pendGuia",0):,} fuera plataforma</div></div>', unsafe_allow_html=True)
            with mc3: st.markdown(f'<div class="big-kpi" style="text-align:center"><div style="font-size:22px;font-weight:800;color:{C["dev"]}">{snap.get("cancelados",0):,}</div><div style="font-size:10px;color:#64748b;margin:4px 0">{fmt_cop(can_d.get("valor",0))}</div><div style="font-size:10px;color:#64748b">{snap.get("cancelados",0):,} canceladas</div></div>', unsafe_allow_html=True)

        with eo_r:
            st.markdown(f'<div style="font-size:13px;font-weight:700;color:#94a3b8;margin-bottom:12px">Estados por guía ({snap["despachados"]:,} registros / {fmt_cop(desp_valor)})</div>', unsafe_allow_html=True)
            STATE_ORDER = ['ENTREGADO','DEVOLUCION','EN TRANSITO','RECLAMAR EN OFICINA','EN REPARTO','CON NOVEDAD','OTRO']
            for row_cats in [STATE_ORDER[:3], STATE_ORDER[3:6], STATE_ORDER[6:]]:
                cols_sc = st.columns(len(row_cats))
                for col_sc, cat in zip(cols_sc, row_cats):
                    d = byEstadoCat.get(cat, {'count':0,'valor':0,'pct':0})
                    cnt = d.get('count',0); pp = d.get('pct',0); val = d.get('valor',0)
                    color = CAT_CMAP.get(cat,'#64748b')
                    with col_sc:
                        st.markdown(f'''
                        <div style="background:#1e293b;border:1px solid #334155;border-radius:10px;padding:12px 8px;text-align:center;margin-bottom:8px">
                            <div style="font-size:19px;font-weight:800;color:{color}">{cnt:,} ({pp:.2f}%)</div>
                            <div style="font-size:11px;color:#64748b;margin:3px 0">{fmt_cop(val)}</div>
                            <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#64748b">{cat}</div>
                        </div>''', unsafe_allow_html=True)

        # ════════════════════════════
        # 2. RESULTADOS POR ESTADO + PENDIENTES POR FECHA
        # ════════════════════════════
        st.markdown("---")
        r2l, r2r = st.columns(2)

        with r2l:
            st.markdown(f'<div class="fin-section-title" style="font-size:15px">📊 Resultados por estado de guía</div>', unsafe_allow_html=True)
            st.caption(f"{total_dTot:,} registros en total.")
            if byEstadoCat:
                cats_b = [c for c in STATE_ORDER if c in byEstadoCat and byEstadoCat[c].get('count',0)>0]
                fig_bar = go.Figure(go.Bar(
                    x=cats_b, y=[byEstadoCat[c]['count'] for c in cats_b],
                    marker=dict(color=[CAT_CMAP.get(c,'#64748b') for c in cats_b], cornerradius=6),
                    text=[f"{byEstadoCat[c].get('pct',0):.1f}%" for c in cats_b],
                    textposition='inside', textfont=dict(color='white',size=12,family='Inter')
                ))
                fig_bar.update_layout(**PLOT_BG,height=300,showlegend=False,
                    xaxis=dict(**AX,tickangle=-20),yaxis=AX)
                st.plotly_chart(fig_bar, use_container_width=True)

        with r2r:
            st.markdown(f'<div class="fin-section-title" style="font-size:15px">📅 Guías pendientes por fecha</div>', unsafe_allow_html=True)
            if snap.get('sinFinal',0) > 0:
                st.markdown(f'''<div style="background:#1e3a5f;border:1px solid #3b82f6;border-radius:8px;padding:10px 14px;margin-bottom:10px;color:#93c5fd;font-size:11px">
                ℹ️ Tienes <strong>{snap["sinFinal"]:,} guías</strong> pendientes. Las de más de 7 días tienen alto riesgo de devolución.</div>''', unsafe_allow_html=True)

            if pendPorFecha:
                fechas = sorted(pendPorFecha.keys())
                max_n  = max(pendPorFecha.values()) if pendPorFecha else 1

                # Init session state para fecha seleccionada
                sk = f"pend_fecha_sel_{inf['id']}"
                if sk not in st.session_state:
                    st.session_state[sk] = None

                # Dibujar barras HTML interactivas
                st.markdown('<div style="overflow-x:auto;padding-bottom:8px">', unsafe_allow_html=True)
                cols_f = st.columns(len(fechas))
                for col_f, fecha in zip(cols_f, fechas):
                    n = pendPorFecha[fecha]
                    h = max(20, int(n / max_n * 120))  # altura proporcional px
                    is_sel = (st.session_state[sk] == fecha)
                    bar_color = '#0ea5e9' if is_sel else C['dev']
                    with col_f:
                        st.markdown(f'''
                        <div style="display:flex;flex-direction:column;align-items:center;gap:3px">
                            <div style="font-size:11px;font-weight:700;color:{bar_color}">{n}</div>
                            <div style="width:100%;min-width:28px;height:{h}px;background:{bar_color};
                                        border-radius:4px 4px 0 0;opacity:{0.9 if is_sel else 0.75}">
                            </div>
                            <div style="font-size:9px;color:#64748b;white-space:nowrap">{fecha}</div>
                        </div>''', unsafe_allow_html=True)
                        btn_label = "✕" if is_sel else "+"
                        if st.button(btn_label, key=f"pf_{inf['id']}_{fecha}", use_container_width=True):
                            st.session_state[sk] = None if is_sel else fecha
                            st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

                # Detalle al hacer clic en +
                fecha_sel = st.session_state.get(sk)
                if fecha_sel and fecha_sel in pendPorFechaLite:
                    ords_fecha = pendPorFechaLite[fecha_sel]
                    st.markdown(f'''
                    <div style="background:#162032;border:1.5px solid #0ea5e9;border-radius:10px;
                                padding:12px 16px;margin-top:8px">
                        <div style="font-size:13px;font-weight:700;color:#0ea5e9;margin-bottom:8px">
                            📋 {len(ords_fecha)} guías pendientes el {fecha_sel}
                        </div>
                    </div>''', unsafe_allow_html=True)
                    df_fecha = pd.DataFrame(ords_fecha)
                    if not df_fecha.empty:
                        cols_show = [c for c in ['guia','nombre','telefono','ciudad',
                                                  'transportadora','estatus','novedad',
                                                  'ult_mov','fecha_umov','vendedor']
                                     if c in df_fecha.columns]
                        df_fecha.insert(0,'', df_fecha['estatus'].apply(badge) if 'estatus' in df_fecha.columns else '🟡')
                        st.dataframe(df_fecha[[''] + cols_show],
                                     use_container_width=True, hide_index=True, height=280)
                        st.download_button(
                            f"⬇️ Descargar órdenes del {fecha_sel}",
                            data=df_fecha[cols_show].to_csv(index=False).encode('utf-8-sig'),
                            file_name=f"pendientes_{fecha_sel.replace('/','_')}_{inf['id']}.csv",
                            mime='text/csv',
                            key=f"dl_pf_{inf['id']}_{fecha_sel}"
                        )
                elif fecha_sel and fecha_sel not in pendPorFechaLite:
                    st.info(f"Re-sincroniza el informe para ver el detalle del {fecha_sel}")
            else:
                st.info("Sin guías pendientes o sincroniza para ver")

        # ════════════════════════════
        # 3 & 4. CIUDAD / DEPARTAMENTO
        # ════════════════════════════
        for geo_data, geo_col, geo_title, dl_key in [
            (byCiudadFull,  'CIUDAD',       '🏙 Resultados por Ciudad',       f"dl_ciudad_{inf['id']}"),
            (byDeptoFull,   'DEPARTAMENTO', '🗺 Resultados por Departamento',  f"dl_depto_{inf['id']}")
        ]:
            if not geo_data: continue
            st.markdown("---")
            st.markdown(f'<div class="fin-section-title">{geo_title}</div>', unsafe_allow_html=True)
            st.caption(f"{total_dTot:,} guías en total")
            df_geo = build_geo_df(geo_data, geo_col, total_dTot)
            if not df_geo.empty:
                st.dataframe(df_geo, use_container_width=True, hide_index=True, height=380)
                st.download_button(f"⬇️ CSV {geo_col.title()}",
                    data=df_geo.to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"{geo_col.lower()}_{datetime.today().strftime('%Y%m%d')}.csv",
                    mime='text/csv', key=dl_key)

        # ════════════════════════════
        # 5. NOVEDADES
        # ════════════════════════════
        st.markdown("---")
        st.markdown('<div class="fin-section-title">🔔 Análisis de Novedades</div>', unsafe_allow_html=True)
        n5l, n5r = st.columns(2)

        with n5l:
            st.markdown("**GUÍAS QUE PRESENTARON NOVEDADES**")
            con_nov  = novData.get('total_con', snap.get('totalNovedades',0))
            sin_nov  = novData.get('total_sin', total_dTot - con_nov)
            nov_pct_v = pct(con_nov, con_nov+sin_nov)
            if nov_pct_v > 20:
                st.markdown(f'''<div style="background:#3d2a00;border:1px solid #f59e0b;border-radius:8px;padding:8px 12px;margin-bottom:8px;color:#fcd34d;font-size:11px">
                ⚠️ Nivel de novedades elevado ({nov_pct_v:.1f}%). Se recomienda revisar el proceso.</div>''', unsafe_allow_html=True)
            st.caption(f"{con_nov+sin_nov:,} registros en total.")
            if con_nov+sin_nov > 0:
                fn1 = go.Figure(go.Pie(
                    labels=[f'{con_nov} con novedad', f'{sin_nov} sin novedad'],
                    values=[con_nov, sin_nov], hole=0.6,
                    marker=dict(colors=['#f97316', C['ent']], line=dict(color='#0f172a',width=2)),
                    textinfo='percent', textfont=dict(color='white',size=12)
                ))
                fn1.update_layout(**PLOT_BG,height=260,xaxis=AX,yaxis=AX,showlegend=False)
                st.plotly_chart(fn1, use_container_width=True)

        with n5r:
            st.markdown("**RESULTADOS DE GUÍAS CON NOVEDADES**")
            n_ent_nv  = novData.get('entregadas',0)
            n_pend_nv = novData.get('pendientes',0)
            n_dev_nv  = novData.get('devueltas',0)
            tot_nv    = n_ent_nv + n_pend_nv + n_dev_nv
            if tot_nv > 0:
                dev_nv_pct = pct(n_dev_nv, tot_nv)
                if dev_nv_pct > 50:
                    st.markdown(f'''<div style="background:#3d2a00;border:1px solid #f59e0b;border-radius:8px;padding:8px 12px;margin-bottom:8px;color:#fcd34d;font-size:11px">
                    ⚠️ Las guías con novedad y devueltas presentan {dev_nv_pct:.1f}%. Se recomienda gestionar mejor las novedades.</div>''', unsafe_allow_html=True)
                st.caption(f"{tot_nv:,} registros en total.")
                fn2 = go.Figure(go.Pie(
                    labels=[f'{n_ent_nv} Entregadas', f'{n_pend_nv} Pendientes', f'{n_dev_nv} Devoluciones'],
                    values=[n_ent_nv, n_pend_nv, n_dev_nv], hole=0.6,
                    marker=dict(colors=[C['ent'],'#f97316',C['dev']], line=dict(color='#0f172a',width=2)),
                    textinfo='percent', textfont=dict(color='white',size=12)
                ))
                fn2.update_layout(**PLOT_BG,height=260,xaxis=AX,yaxis=AX,showlegend=False)
                st.plotly_chart(fn2, use_container_width=True)

        # ════════════════════════════
        # 6. TRANSPORTADORAS
        # ════════════════════════════
        if byCarrierFull:
            st.markdown("---")
            st.markdown('<div class="fin-section-title">🚛 Resultados por Transportadora</div>', unsafe_allow_html=True)

            carrier_list = sorted(byCarrierFull.items(), key=lambda x: -x[1].get('total',0))

            # 6a. Promedios en 3 columnas
            all_fle = [d.get('flete_avg',0) for _,d in carrier_list if d.get('flete_avg',0)>0]
            all_dev = [d.get('cdev_avg',0)  for _,d in carrier_list if d.get('cdev_avg',0)>0]
            all_dia = [d.get('dias_avg',0)  for _,d in carrier_list if d.get('dias_avg',0)>0]
            avg_fle = sum(all_fle)/len(all_fle) if all_fle else 0
            avg_dev = sum(all_dev)/len(all_dev) if all_dev else 0
            avg_dia = sum(all_dia)/len(all_dia) if all_dia else 0

            pr1, pr2, pr3 = st.columns(3)

            def prom_panel(col, titulo, caption_txt, avg_val, fmt_fn, suffix, key_color):
                with col:
                    st.markdown(f"**{titulo}**")
                    st.caption(caption_txt)
                    st.markdown(f'<div class="big-kpi"><div style="font-size:26px;font-weight:800;color:{key_color}">{fmt_fn(avg_val)}{suffix}</div><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">PROMEDIO GENERAL</div></div>', unsafe_allow_html=True)
                    crs = st.columns(2)
                    for i,(carrier,d) in enumerate(carrier_list):
                        v = d.get({'flete_avg':True}.get(titulo,''),0) or (d.get('flete_avg',0) if 'FLETE' in titulo else (d.get('cdev_avg',0) if 'DEVOL' in titulo else d.get('dias_avg',0)))
                        with crs[i%2]:
                            st.markdown(f'<div class="big-kpi"><div style="font-size:16px;font-weight:800;color:{key_color}">{fmt_fn(v)}{suffix}</div><div style="font-size:9px;font-weight:700;color:#64748b;text-transform:uppercase">{carrier}</div></div>', unsafe_allow_html=True)

            with pr1:
                st.markdown("**VALOR PROMEDIO DE FLETES**")
                st.caption("Calculado con guías entregadas al destinatario.")
                st.markdown(f'<div class="big-kpi"><div style="font-size:26px;font-weight:800;color:{C["teal"]}">{fmt_cop(avg_fle)}</div><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">PROMEDIO GENERAL</div></div>', unsafe_allow_html=True)
                # Grid 2 cols per carrier
                for chunk_i in range(0, len(carrier_list), 2):
                    chunk = carrier_list[chunk_i:chunk_i+2]
                    crs1 = st.columns(len(chunk))
                    for ci,(c,d) in enumerate(chunk):
                        with crs1[ci]:
                            st.markdown(f'<div class="big-kpi" style="padding:14px 10px;text-align:center"><div style="font-size:18px;font-weight:800;color:{C["teal"]}">{fmt_cop(d.get("flete_avg",0))}</div><div style="font-size:9px;font-weight:700;color:#64748b;text-transform:uppercase;margin-top:4px">{c}</div></div>', unsafe_allow_html=True)

            with pr2:
                st.markdown("**VALOR PROMEDIO DE DEVOLUCIÓN**")
                st.caption("Calculado con guías devueltas y liquidadas.")
                st.markdown(f'<div class="big-kpi"><div style="font-size:26px;font-weight:800;color:{C["dev"]}">{fmt_cop(avg_dev)}</div><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">PROMEDIO GENERAL</div></div>', unsafe_allow_html=True)
                for chunk_i in range(0, len(carrier_list), 2):
                    chunk = carrier_list[chunk_i:chunk_i+2]
                    crs2 = st.columns(len(chunk))
                    for ci,(c,d) in enumerate(chunk):
                        with crs2[ci]:
                            st.markdown(f'<div class="big-kpi" style="padding:14px 10px;text-align:center"><div style="font-size:18px;font-weight:800;color:{C["dev"]}">{fmt_cop(d.get("cdev_avg",0))}</div><div style="font-size:9px;font-weight:700;color:#64748b;text-transform:uppercase;margin-top:4px">{c}</div></div>', unsafe_allow_html=True)

            with pr3:
                st.markdown("**DÍAS DE ENTREGA PROMEDIO**")
                st.caption("Desde FECHA GUIA GENERADA hasta FECHA ÚLTIMO MOVIMIENTO (ENTREGADO).")
                st.markdown(f'<div class="big-kpi"><div style="font-size:26px;font-weight:800;color:{C["teal"]}">{avg_dia:.1f} días</div><div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase">PROMEDIO GENERAL</div></div>', unsafe_allow_html=True)
                for chunk_i in range(0, len(carrier_list), 2):
                    chunk = carrier_list[chunk_i:chunk_i+2]
                    crs3 = st.columns(len(chunk))
                    for ci,(c,d) in enumerate(chunk):
                        with crs3[ci]:
                            st.markdown(f'<div class="big-kpi" style="padding:14px 10px;text-align:center"><div style="font-size:18px;font-weight:800;color:{C["teal"]}">{d.get('dias_avg',0):.1f} días</div><div style="font-size:9px;font-weight:700;color:#64748b;text-transform:uppercase;margin-top:4px">{c}</div></div>', unsafe_allow_html=True)

            # 6b. Bar charts por transportadora
            st.markdown("---")
            st.markdown("**ESTADOS POR TRANSPORTADORA**")
            for row_i in range(0, len(carrier_list), 2):
                row_c = carrier_list[row_i:row_i+2]
                cols_rc = st.columns(len(row_c))
                for col_rc,(carrier,d) in zip(cols_rc, row_c):
                    with col_rc:
                        tot_c = d.get('total',1)
                        cats_c = [c for c in LOG_CATS if d.get(c,0)>0]
                        st.markdown(f"<small style='font-size:11px;font-weight:700;color:#0ea5e9;text-transform:uppercase'>RESULTADOS CON {carrier}</small>", unsafe_allow_html=True)
                        st.caption(f"{tot_c:,} registros en total.")
                        fig_rc = go.Figure(go.Bar(
                            x=cats_c, y=[d.get(c,0) for c in cats_c],
                            marker=dict(color=[CAT_CMAP.get(c,'#64748b') for c in cats_c],cornerradius=4),
                            text=[f"{d.get(f'pct_{c}',0):.1f}%" for c in cats_c],
                            textposition='inside', textfont=dict(color='white',size=11)
                        ))
                        fig_rc.update_layout(**PLOT_BG,height=240,showlegend=False,
                            xaxis=dict(**AX,tickangle=-20),yaxis=AX)
                        st.plotly_chart(fig_rc, use_container_width=True)

            # 6c. Ciudad + Transportadora
            if byCarrierCiudad:
                st.markdown("---")
                st.markdown('<div class="fin-section-title" style="font-size:16px">🏙🚛 Resultados por Ciudad y Transportadora</div>', unsafe_allow_html=True)
                for carrier, ciudad_data in sorted(byCarrierCiudad.items()):
                    if not ciudad_data: continue
                    tot_c = sum(d.get('total',0) for d in ciudad_data.values())
                    st.markdown(f"**RESULTADOS POR CIUDAD CON {carrier.upper()}** — {tot_c:,} guías en total")
                    df_cc = build_geo_df(ciudad_data, 'CIUDAD', tot_c, maxDev)
                    if not df_cc.empty:
                        sorted_cities = sorted(ciudad_data.items(), key=lambda x: -x[1].get('total',0))
                        df_cc['Promedio Flete']        = [fmt_cop(d.get('flete_avg',0)) for _,d in sorted_cities]
                        df_cc['Promedio Días Entrega'] = [f"{d.get('dias_avg',0):.1f} día(s)" for _,d in sorted_cities]
                        st.dataframe(df_cc, use_container_width=True, hide_index=True, height=320)
                    else:
                        st.info("Sin datos de ciudad")

            # 6d. Departamento + Transportadora
            if byCarrierDepto:
                st.markdown("---")
                st.markdown('<div class="fin-section-title" style="font-size:16px">🗺🚛 Resultados por Departamento y Transportadora</div>', unsafe_allow_html=True)
                for carrier, depto_data in sorted(byCarrierDepto.items()):
                    if not depto_data: continue
                    tot_d = sum(d.get('total',0) for d in depto_data.values())
                    st.markdown(f"**RESULTADOS POR DEPARTAMENTO CON {carrier.upper()}** — {tot_d:,} guías en total")
                    df_dc = build_geo_df(depto_data, 'DEPARTAMENTO', tot_d, maxDev)
                    if not df_dc.empty:
                        sorted_deptos = sorted(depto_data.items(), key=lambda x: -x[1].get('total',0))
                        df_dc['Promedio Flete']        = [fmt_cop(d.get('flete_avg',0)) for _,d in sorted_deptos]
                        df_dc['Promedio Días Entrega'] = [f"{d.get('dias_avg',0):.1f} día(s)" for _,d in sorted_deptos]
                        st.dataframe(df_dc, use_container_width=True, hide_index=True, height=280)

        # ════════════════════════════
        # 7. GUÍAS DEVUELTAS SIN NOVEDAD
        # ════════════════════════════
        if devSinNovLite:
            st.markdown("---")
            st.markdown('<div class="fin-section-title">📋 Guías Devueltas sin Novedad</div>', unsafe_allow_html=True)
            st.markdown('''<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;padding:10px 14px;margin-bottom:12px;color:#94a3b8;font-size:12px">
            ℹ️ Analiza si estas guías debieron ser devueltas sin novedad o si fue un error evitable. Toma acciones para evitar que esto suceda en el futuro.</div>''', unsafe_allow_html=True)
            df_dsn = pd.DataFrame(devSinNovLite)
            if not df_dsn.empty:
                df_dsn.columns = [c.upper() for c in df_dsn.columns]
                for col in ['VALOR','FLETE']:
                    if col in df_dsn.columns:
                        df_dsn[col] = df_dsn[col].apply(fmt_cop)
                st.dataframe(df_dsn, use_container_width=True, hide_index=True, height=280)

        # ════════════════════════════
        # 8. VENDEDORES
        # ════════════════════════════
        if byVendedor:
            st.markdown("---")
            st.markdown('<div class="fin-section-title">👥 Resultados por Vendedor / Colaborador</div>', unsafe_allow_html=True)
            vend_rows = []
            for v,d in sorted(byVendedor.items(), key=lambda x: -x[1].get('total',0)):
                tot = d.get('total',1)
                vend_rows.append({'Vendedor':v,'Total':tot,
                    'Entregadas':d.get('ENTREGADO',0), '% Entrega':f"{pct(d.get('ENTREGADO',0),tot):.1f}%",
                    'Devol.':d.get('DEVOLUCION',0), '% Dev.':f"{pct(d.get('DEVOLUCION',0),tot):.1f}%",
                    'En tránsito':d.get('EN TRANSITO',0),'Con novedad':d.get('CON NOVEDAD',0),
                    'Ganancia':fmt_cop(d.get('ganancia',0))})
            st.dataframe(pd.DataFrame(vend_rows), use_container_width=True, hide_index=True)

        # ════════════════════════════
        # 9. TABLA COMPLETA
        # ════════════════════════════
        st.markdown("---")
        st.markdown('<div class="fin-section-title">📋 Tabla completa de órdenes</div>', unsafe_allow_html=True)
        ords = snap.get('ordenesLite',[])
        if ords:
            df_ords = pd.DataFrame(ords)
            est_opts = ['Todos']+sorted(df_ords['estatus'].dropna().unique().tolist()) if 'estatus' in df_ords.columns else ['Todos']
            fe = st.selectbox("Filtrar por estado", est_opts, key=f"fe_log_{inf['id']}")
            df_show = df_ords.copy()
            if fe != 'Todos' and 'estatus' in df_show.columns:
                df_show = df_show[df_show['estatus']==fe]
            cs = [c for c in ['guia','nombre','ciudad','departamento','estatus','transportadora',
                               'novedad','fecha_umov','hora_umov','ult_mov','vendedor']
                  if c in df_show.columns]
            if 'estatus' in df_show.columns:
                df_show = df_show.copy()
                df_show.insert(0,'',df_show['estatus'].apply(badge))
            st.dataframe(df_show, use_container_width=True, hide_index=True, height=380)
            st.download_button("⬇️ CSV logístico completo",
                data=df_ords.to_csv(index=False).encode('utf-8-sig'),
                file_name=f"logistico_{inf['title'].replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.csv",
                mime='text/csv', key=f"dl_log_{inf['id']}")
        else:
            st.info("Sincroniza el informe para ver la tabla de órdenes")



    # ══ FINANCIERO ══
    else:
        st.markdown("# 💰 Informe Financiero")

        # ─── helpers visuales ───
        def fin_title(icon, text):
            st.markdown(f'<div class="fin-section-title">{icon} {text}</div>', unsafe_allow_html=True)
        def fin_sub(text):
            st.markdown(f'<div class="fin-sub-title">{text}</div>', unsafe_allow_html=True)
        def metric_card(lbl, val, color=None, size=32):
            c = color or '#f1f5f9'
            return f'''<div class="big-kpi">
                <div style="font-size:{size}px;font-weight:800;color:{c};line-height:1.1;margin-bottom:4px">{val}</div>
                <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#64748b">{lbl}</div>
            </div>'''

        # ─── Parámetros manuales ───
        fin_title("⚙️","Parámetros del período")
        with st.expander("✏️ Gastos manuales y parámetros fiscales", expanded=False):
            pg1,pg2,pg3=st.columns(3)
            with pg1:
                st.markdown("**👥 Nómina**")
                nomina     =st.number_input("Salarios ($)",    min_value=0.0,value=0.0,step=100000.0,format="%.0f",key=f"g_nom_{inf['id']}")
                comisiones =st.number_input("Comisiones ($)",  min_value=0.0,value=0.0,step=50000.0, format="%.0f",key=f"g_com_{inf['id']}")
            with pg2:
                st.markdown("**🖥 Herramientas**")
                pancake_fee=st.number_input("Pancake ($)",     min_value=0.0,value=0.0,step=10000.0, format="%.0f",key=f"g_pan_{inf['id']}")
                dropi_fee  =st.number_input("Dropi ($)",       min_value=0.0,value=0.0,step=10000.0, format="%.0f",key=f"g_dro_{inf['id']}")
                otros_tools=st.number_input("Otras ($)",       min_value=0.0,value=0.0,step=10000.0, format="%.0f",key=f"g_oth_{inf['id']}")
                arriendo   =st.number_input("Arriendo ($)",    min_value=0.0,value=0.0,step=50000.0, format="%.0f",key=f"g_arr_{inf['id']}")
            with pg3:
                st.markdown("**📣 Pauta**")
                ads_meta   =st.number_input("Meta Ads ($)",    min_value=0.0,value=0.0,step=100000.0,format="%.0f",key=f"g_meta_{inf['id']}")
                ads_tiktok =st.number_input("TikTok Ads ($)",  min_value=0.0,value=0.0,step=100000.0,format="%.0f",key=f"g_tik_{inf['id']}")
            st.markdown("---")
            st.markdown("**🏛 Fiscal RST Colombia**")
            tf1,tf2,tf3=st.columns(3)
            with tf1: tasa_rst=st.number_input("Tasa RST (%)",min_value=0.0,max_value=10.0,value=0.4, step=0.1, format="%.2f",key=f"t_rst_{inf['id']}")
            with tf2: tasa_ica=st.number_input("Tasa ICA (%)",min_value=0.0,max_value=3.0, value=1.04,step=0.01,format="%.2f",key=f"t_ica_{inf['id']}")
            with tf3:
                n_transf  =st.number_input("Nº transferencias",min_value=0,value=0,step=1,key=f"t_gn_{inf['id']}")
                val_transf=st.number_input("Valor prom. trans. ($)",min_value=0.0,value=0.0,step=100000.0,format="%.0f",key=f"t_gv_{inf['id']}")

        if d_ is None:
            st.info("👆 Carga el archivo de Dropi en la barra lateral")
            st.stop()

        # ─── Datos base de Dropi ───
        df_ent  = d_[d_['estatus'].apply(es_e)].copy()
        df_dev  = d_[d_['estatus'].apply(es_d)].copy()
        df_sinf = d_[d_['estatus'].apply(es_s)].copy()
        df_desp = d_[d_['estatus'].apply(es_despachado)].copy()

        n_ent   = len(df_ent)
        n_dev   = len(df_dev)
        n_sinf  = len(df_sinf)
        n_desp  = len(df_desp)
        n_panc  = snap.get('totalOrd',0) if snap else (len(p_) if p_ is not None else 0)

        # ── Calcular valores financieros directo del archivo raw ──
        # FUENTES EXACTAS (verificadas con archivos reales):
        # Logístico: TOTAL DE LA ORDEN=recaudo, PRECIO PROVEEDOR X CANTIDAD=costo, PRECIO FLETE=flete
        # Financiero: COSTO DEVOLUCION FLETE=costo devolución (suma correcta)

        raw_log = dd_raw_det   # logístico
        raw_fin = dd_fin_raw   # financiero

        def _raw_col_exact(src, col_cands):
            """Busca columna por nombre EXACTO (sin match parcial)."""
            if src is None: return None
            cols_norm = {c.strip().upper(): c for c in src.columns}
            for cand in col_cands:
                if cand.strip().upper() in cols_norm:
                    return cols_norm[cand.strip().upper()]
            return None

        def _raw_sum(src, col_cands, mask_estatus):
            """Suma columna financiera del raw, filtrando por estatus."""
            if src is None: return 0.0
            col = _raw_col_exact(src, col_cands)
            if not col: return 0.0
            est_col = _raw_col_exact(src, ['ESTATUS'])
            if not est_col: return 0.0
            mask = src[est_col].astype(str).str.upper().str.strip().apply(mask_estatus)
            return float(pd.to_numeric(src[col], errors='coerce').fillna(0)[mask].sum())

        if raw_log is not None:
            val_ent = _raw_sum(raw_log, ['TOTAL DE LA ORDEN'], es_e)
            gan_ent = _raw_sum(raw_log, ['GANANCIA'], es_e)
            cos_ent = _raw_sum(raw_log, ['PRECIO PROVEEDOR X CANTIDAD'], es_e)
            fle_ent = _raw_sum(raw_log, ['PRECIO FLETE'], es_e)
            # Costo devolución: preferir financiero (suma exacta), fallback logístico
            cdev = _raw_sum(raw_fin, ['COSTO DEVOLUCION FLETE'], es_d) if raw_fin is not None \
                   else _raw_sum(raw_log, ['COSTO DEVOLUCION FLETE'], es_d)
        else:
            val_ent = snap.get('ingresos', 0)
            gan_ent = snap.get('totalGanancia', 0)
            cos_ent = snap.get('totalCosto', 0)
            fle_ent = snap.get('totalFlete', 0)
            cdev    = snap.get('totalCostoDev', 0)

        if val_ent == 0 and gan_ent == 0:
            val_ent = snap.get('ingresos', 0)
            gan_ent = snap.get('totalGanancia', 0)
            cos_ent = snap.get('totalCosto', 0)
            fle_ent = snap.get('totalFlete', 0)
            cdev    = snap.get('totalCostoDev', 0)

        ingresos = val_ent if val_ent > 0 else snap.get('ingresos', 0)

        # Promedios por orden (para simulación)
        avg_ticket    = ingresos / n_ent if n_ent else 0
        avg_ganancia  = gan_ent  / n_ent if n_ent else 0
        avg_cos       = cos_ent  / n_ent if n_ent else 0
        avg_fle_ent   = fle_ent  / n_ent if n_ent else 0
        avg_fle_dev   = cdev     / n_dev  if n_dev  else 0

        # ── Garantías e indemnizaciones (entradas reales que afectan P&G) ──
        garantia_total = 0.0
        n_garantia = 0
        indemn_total = 0.0
        n_indemn = 0
        if dc_ is not None:
            df_gar = dc_[dc_['_clasif'] == 'Devolución por garantía']
            garantia_total = float(df_gar['_monto'].sum())
            n_garantia     = len(df_gar)
            df_ind = dc_[dc_['_clasif'] == 'Indemnización']
            indemn_total = float(df_ind['_monto'].sum())
            n_indemn     = len(df_ind)

        # Gastos manuales y fiscales
        inversion_ads = ads_meta + ads_tiktok
        gastos_op     = nomina + comisiones + pancake_fee + dropi_fee + otros_tools + arriendo
        utilidad_bruta        = gan_ent
        utilidad_neta_log     = utilidad_bruta - cdev
        utilidad_antes_gastos = utilidad_neta_log - inversion_ads
        ebitda                = utilidad_antes_gastos - gastos_op
        rst_val  = ingresos * tasa_rst / 100
        ica_val  = ingresos * tasa_ica / 100
        gmf_val  = n_transf * val_transf * 4 / 1000
        total_imp = rst_val + ica_val + gmf_val
        utilidad_neta = ebitda - total_imp

        # ── Utilidad cobrada de cartera ──
        util_cobrada = 0.0
        if dc_ is not None:
            util_cobrada = float(dc_[dc_['_clasif']=='Ganancia cobrada']['_monto'].sum())

        # ── Utilidad actual = Total transacciones + sin recaudo - gastos ──
        # Se calcula DESPUÉS de tener los fletes (se recalcula abajo tras tener fle_ent_cod/nocod)
        # Se deja como placeholder aquí y se actualiza tras calcular recaudo

        # ── Utilidad pendiente = GANANCIA de órdenes entregadas HOY en Pancake ──
        util_pendiente_pancake = 0.0
        n_pendiente_pancake    = 0
        fecha_sync = datetime.now()

        if dp_raw_det is not None and raw_log is not None:
            try:
                cron_col      = find_col(dp_raw_det, ['cronograma de actualización de estado',
                                                       'cronograma de actualizacion de estado','cronograma'])
                guia_panc_col = find_col(dp_raw_det, ['número de seguimiento','numero de seguimiento','tracking'])

                if cron_col and guia_panc_col:
                    def _tiene_entregado_hoy(text):
                        if not text or str(text).strip() in ('nan','None',''): return False
                        for entry in str(text).split(';'):
                            m = re.match(r'^(.+?)\s*-\s*(\d{2}:\d{2}\s+\d{2}/\d{2}/\d{4})', entry.strip())
                            if m:
                                if 'entregado' in m.group(1).strip().lower():
                                    try:
                                        dt = datetime.strptime(m.group(2).strip(), '%H:%M %d/%m/%Y')
                                        if dt.date() == fecha_sync.date():
                                            return True
                                    except: pass
                        return False

                    # Paso 1: guías entregadas hoy en Pancake → convertir a int para comparar
                    panc_df = dp_raw_det.copy()
                    panc_df['_entregado_hoy'] = panc_df[cron_col].apply(_tiene_entregado_hoy)
                    panc_df['_guia_int'] = pd.to_numeric(panc_df[guia_panc_col], errors='coerce').dropna().astype('int64')
                    guias_hoy = set(panc_df[panc_df['_entregado_hoy']]['_guia_int'].dropna().tolist())

                    if guias_hoy:
                        # Paso 2: en el logístico, filtrar guías que están en guias_hoy
                        guia_log_col = _raw_col_exact(raw_log, ['NÚMERO GUIA','NUMERO GUIA'])
                        gan_col_l    = _raw_col_exact(raw_log, ['GANANCIA'])
                        tot_col_l    = _raw_col_exact(raw_log, ['TOTAL DE LA ORDEN'])
                        cos_col_l    = _raw_col_exact(raw_log, ['PRECIO PROVEEDOR X CANTIDAD'])
                        fle_col_l    = _raw_col_exact(raw_log, ['PRECIO FLETE'])

                        if guia_log_col:
                            log_df = raw_log.copy()
                            log_df['_guia_int'] = pd.to_numeric(log_df[guia_log_col], errors='coerce')
                            log_df = log_df[log_df['_guia_int'].notna()].copy()
                            log_df['_guia_int'] = log_df['_guia_int'].astype('int64')
                            mask_hoy = log_df['_guia_int'].isin(guias_hoy)
                            df_hoy_log = log_df[mask_hoy].copy()

                            if len(df_hoy_log) > 0 and tot_col_l and cos_col_l and fle_col_l:
                                # Dropi pone GANANCIA=NaN para órdenes recién entregadas (aún no liquidadas)
                                # Cuando GANANCIA existe usarla, si no: TOTAL - COSTO - FLETE
                                gan_vals = pd.to_numeric(df_hoy_log[gan_col_l], errors='coerce') if gan_col_l else pd.Series(dtype=float)
                                tot_vals = pd.to_numeric(df_hoy_log[tot_col_l], errors='coerce').fillna(0)
                                cos_vals = pd.to_numeric(df_hoy_log[cos_col_l], errors='coerce').fillna(0)
                                fle_vals = pd.to_numeric(df_hoy_log[fle_col_l], errors='coerce').fillna(0)
                                gan_calc = tot_vals - cos_vals - fle_vals
                                gan_final = gan_vals.where(gan_vals.notna(), gan_calc)
                                util_pendiente_pancake = float(gan_final.sum())
                                n_pendiente_pancake = len(df_hoy_log['_guia_int'].unique())
            except: pass

        # Si Pancake está cargado pero no hay entregadas hoy → $0
        # Si no hay Pancake → fallback a ganancia bruta - cobrada
        util_pendiente = util_pendiente_pancake if dp_raw_det is not None else float(utilidad_bruta - util_cobrada)

        dias_periodo = 1
        if d_ is not None and '_fecha_dt' in d_.columns and d_['_fecha_dt'].notna().any():
            mn=d_['_fecha_dt'].min(); mx=d_['_fecha_dt'].max()
            dias_periodo=max(1,(mx-mn).days+1)

        cpa_total  = inversion_ads/n_panc  if inversion_ads and n_panc  else 0
        cpa_desp   = inversion_ads/n_desp  if inversion_ads and n_desp  else 0
        cpa_ent    = inversion_ads/n_ent   if inversion_ads and n_ent   else 0
        roas       = ingresos/inversion_ads if inversion_ads else 0
        ticket_prom  = avg_ticket
        util_x_antes = avg_ganancia - (avg_fle_dev * n_dev / n_ent if n_ent else 0)
        util_x_desp  = util_x_antes - (gastos_op / n_ent if n_ent else 0)
        ventas_dia   = n_desp / 30

        # ════════════════════════════════════════
        # SECCIÓN 1: TRANSACCIONES (de cartera)
        # ════════════════════════════════════════
        fin_title("💳","Transacciones")

        if n_sinf > 0 or n_dev > 0:
            st.markdown(f'''
            <div style="background:#3d2a00;border:1px solid #f59e0b;border-radius:10px;padding:12px 16px;margin-bottom:16px;color:#fcd34d;font-size:13px">
                ⚠️ Aún existen transacciones pendientes de liquidar y guías pendientes de finalizar.
                Se recomienda revisar estos datos y/o esperar su actualización para considerar los resultados
                de este informe como resultados finales. Por ejemplo, si las guías pendientes terminan en
                devolución la utilidad puede disminuir considerablemente.
            </div>''', unsafe_allow_html=True)

        # Tabla de transacciones
        # Para flete usar FINANCIERO (1 fila por orden, sin duplicados por upsell)
        # Para TOTAL DE LA ORDEN y TIPO DE ENVIO usar logístico (tiene esas columnas)
        sin_recaudo_n   = 0
        sin_recaudo_val = 0.0
        fle_ent_cod     = fle_ent
        fle_ent_nocod   = 0.0

        # Fuente: financiero para flete (sin duplicados), logístico para valor/tipo
        raw_src_flete = raw_fin if raw_fin is not None else raw_log
        raw_src_tipo  = raw_log  # logístico tiene TIPO DE ENVIO y TOTAL DE LA ORDEN

        if raw_src_flete is not None and raw_src_tipo is not None:
            try:
                tipo_col = _raw_col_exact(raw_src_tipo,  ['TIPO DE ENVIO'])
                est_col  = _raw_col_exact(raw_src_flete, ['ESTATUS'])
                fle_col  = _raw_col_exact(raw_src_flete, ['PRECIO FLETE'])
                val_col  = _raw_col_exact(raw_src_tipo,  ['TOTAL DE LA ORDEN'])
                guia_fin = _raw_col_exact(raw_src_flete, ['NÚMERO GUIA','NUMERO GUIA'])
                guia_log = _raw_col_exact(raw_src_tipo,  ['NÚMERO GUIA','NUMERO GUIA'])

                if tipo_col and est_col and fle_col and guia_fin and guia_log:
                    # Financiero: 1 fila por orden entregada
                    fin_ent = raw_src_flete[raw_src_flete[est_col].astype(str).str.upper().str.strip()=='ENTREGADO'].copy()
                    fin_ent['_guia_n'] = pd.to_numeric(fin_ent[guia_fin], errors='coerce')

                    # Logístico: obtener TIPO DE ENVIO por guía (dedup)
                    log_tipo = raw_src_tipo[[guia_log, tipo_col]].copy()
                    log_tipo['_guia_n'] = pd.to_numeric(log_tipo[guia_log], errors='coerce')
                    log_tipo = log_tipo.drop_duplicates('_guia_n')

                    # Cruzar financiero con tipo de envío del logístico
                    fin_ent = fin_ent.merge(log_tipo[['_guia_n', tipo_col]], on='_guia_n', how='left', suffixes=('','_log'))
                    tipo_col_m = tipo_col if tipo_col in fin_ent.columns else tipo_col+'_log'

                    mask_recaudo = fin_ent[tipo_col_m].astype(str).str.upper().str.strip() == 'CON RECAUDO'
                    mask_sinrec  = fin_ent[tipo_col_m].astype(str).str.upper().str.strip() == 'SIN RECAUDO'

                    sin_recaudo_n = int(mask_sinrec.sum())
                    fle_ent_cod   = float(pd.to_numeric(fin_ent[mask_recaudo][fle_col], errors='coerce').fillna(0).sum())
                    fle_ent_nocod = float(pd.to_numeric(fin_ent[mask_sinrec][fle_col],  errors='coerce').fillna(0).sum())

                    # Valor sin recaudo desde logístico
                    if val_col:
                        log_ent = raw_src_tipo[raw_src_tipo[_raw_col_exact(raw_src_tipo,['ESTATUS'])].astype(str).str.upper().str.strip()=='ENTREGADO'] if _raw_col_exact(raw_src_tipo,['ESTATUS']) else pd.DataFrame()
                        if not log_ent.empty:
                            log_ent = log_ent.drop_duplicates(guia_log)
                            log_ent['_guia_n'] = pd.to_numeric(log_ent[guia_log], errors='coerce')
                            mask_sin_log = log_ent[tipo_col].astype(str).str.upper().str.strip() == 'SIN RECAUDO'
                            sin_recaudo_val = float(pd.to_numeric(log_ent[mask_sin_log][val_col], errors='coerce').fillna(0).sum())
            except: pass
        elif 'tipo_envio' in df_ent.columns and df_ent['tipo_envio'].notna().any():
            mask_cod = df_ent['tipo_envio'].astype(str).str.upper().str.strip() == 'CON RECAUDO'
            sin_recaudo_n   = int((~mask_cod).sum())
            sin_recaudo_val = float(to_n(df_ent[~mask_cod]['valor']).sum()) if 'valor' in df_ent.columns else 0
            fle_ent_cod     = float(to_n(df_ent[mask_cod]['flete']).sum())  if 'flete' in df_ent.columns else fle_ent
            fle_ent_nocod   = float(to_n(df_ent[~mask_cod]['flete']).sum()) if 'flete' in df_ent.columns else 0

        # Costo devolución: usar FINANCIERO (1 fila por orden, sin duplicados por upsell)
        # Fallback al logístico si no hay financiero
        cdev_real = cdev
        n_dev_real = n_dev
        n_dev_con_costo = 0
        src_cdev_display = raw_fin if raw_fin is not None else raw_src_fin
        if src_cdev_display is not None:
            try:
                est_col  = _raw_col_exact(src_cdev_display, ['ESTATUS'])
                cdev_col = _raw_col_exact(src_cdev_display, ['COSTO DEVOLUCION FLETE'])
                if est_col and cdev_col:
                    raw_dev_d = src_cdev_display[src_cdev_display[est_col].astype(str).str.upper().str.strip() == 'DEVOLUCION']
                    cdev_vals = pd.to_numeric(raw_dev_d[cdev_col], errors='coerce').fillna(0)
                    n_dev_con_costo = int((cdev_vals > 0).sum())
                    cdev_real  = float(cdev_vals.sum())
                    n_dev_real = len(raw_dev_d)
            except: pass

        # Flete devolución pendiente (desde cartera)
        fle_dev_pendiente = 0.0
        n_fle_dev_pend    = 0
        if dc_ is not None:
            # Pendiente = flete dev en órdenes sinf que aún no se cobró
            fle_dev_total_cart = float(dc_[dc_['_clasif']=='Flete devolución cobrado']['_monto'].sum())
            fle_dev_pendiente  = max(0, cdev - fle_dev_total_cart)
            n_fle_dev_pend     = int(round(fle_dev_pendiente / avg_fle_dev)) if avg_fle_dev else 0

        # Mantenimiento tarjetas
        mant_total = 0.0
        n_mant = 0
        if dc_ is not None:
            df_mant = dc_[dc_['_clasif']=='Mantenimiento tarjeta virtual']
            mant_total = float(df_mant['_monto'].sum())
            n_mant     = len(df_mant)

        transacciones = [
            ("Recaudo de venta",                    f"{n_ent} venta(s)",        ingresos,      None, None, True),
            ("Costo de mercancía dropshipping",      f"{n_ent} compra(s)",       -cos_ent,      None, None, False),
            ("Flete guías con recaudo (entregadas)", f"{n_ent - sin_recaudo_n} flete(s)", -fle_ent_cod, None, None, False),
        ]
        if sin_recaudo_n > 0:
            transacciones.append(("Flete guías sin recaudo (entregadas)", f"{sin_recaudo_n} flete(s)", -fle_ent_nocod, None, None, False))

        transacciones += [
            ("Descuento por flete de devolución",
             f"{n_dev_real} flete(s) con costo" if n_dev_con_costo > 0 else f"{n_dev} flete(s)",
             -cdev_real,
             n_fle_dev_pend if n_fle_dev_pend else None,
             -fle_dev_pendiente if fle_dev_pendiente else None,
             False),
        ]
        if garantia_total > 0:
            transacciones.append(("Devolución de dinero por garantía", f"{n_garantia} registro(s)", garantia_total, None, None, True))
        if indemn_total > 0:
            transacciones.append(("Entrada por indemnización", f"{n_indemn} registro(s)", indemn_total, None, None, True))
        if mant_total:
            transacciones.append(("SALIDA POR MANTENIMIENTO MENSUAL TARJETA VIRTUAL", f"{n_mant} registro(s)", -mant_total, None, None, False))

        # Utilidad actual = total transacciones + sin recaudo - gastos manuales y fiscales
        total_trans = ingresos - cos_ent - fle_ent_cod - fle_ent_nocod - cdev_real + garantia_total + indemn_total - mant_total
        util_actual = total_trans + sin_recaudo_val - gastos_op - inversion_ads - total_imp
        roi = util_actual / (inversion_ads + gastos_op) * 100 if (inversion_ads + gastos_op) else 0

        total_pend  = -fle_dev_pendiente if fle_dev_pendiente else None

        # Render tabla
        th = st.columns([4,2,2,2,2])
        th[0].markdown("**Transacción**")
        th[1].markdown("**Registros**")
        th[2].markdown("**Total**")
        th[3].markdown("**Reg. pendientes**")
        th[4].markdown("**Total pendiente**")
        st.markdown('<hr style="border-color:#334155;margin:6px 0">', unsafe_allow_html=True)

        for lbl, regs, val, pend_n, pend_v, is_ing in transacciones:
            pct_val = abs(val)/ingresos*100 if ingresos else 0
            color   = C['ent'] if is_ing else C['dev']
            pct_str = f"<br><small style='color:#64748b'>({pct_val:.2f}%)</small>" if not is_ing else ""
            row = st.columns([4,2,2,2,2])
            row[0].markdown(f"<span style='font-size:13px;color:#cbd5e1'>{lbl}</span>", unsafe_allow_html=True)
            row[1].markdown(f"<span style='font-size:12px;color:#64748b'>{regs}</span>", unsafe_allow_html=True)
            row[2].markdown(f"<span style='font-size:13px;font-weight:700;color:{color}'>{fmt_cop(val)}{pct_str}</span>", unsafe_allow_html=True)
            row[3].markdown(f"<span style='font-size:12px;color:#64748b'>{f'{pend_n} flete(s)' if pend_n else ''}</span>", unsafe_allow_html=True)
            dev_color = C['dev']
            row[4].markdown(f"<span style='font-size:13px;font-weight:700;color:{dev_color}'>{fmt_cop(pend_v) if pend_v else ''}</span>", unsafe_allow_html=True)
            st.markdown('<hr style="border-color:#1e293b;margin:4px 0">', unsafe_allow_html=True)

        # Fila TOTAL
        tot_cols = st.columns([4,2,2,2,2])
        tot_cols[0].markdown("**TOTAL**")
        tot_cols[2].markdown(f"<span style='font-size:15px;font-weight:800;color:{C['ent'] if total_trans>=0 else C['dev']}'>{fmt_cop(total_trans)}</span>", unsafe_allow_html=True)
        if total_pend:
            tot_cols[4].markdown(f"<span style='font-size:15px;font-weight:800;color:{dev_color}'>{fmt_cop(total_pend)}</span>", unsafe_allow_html=True)

        # ── + Ver registros por categoría (Transacciones) ──
        if dc_ is not None:
            CLASIFS_TRANS = {'Ganancia cobrada','Flete envío cobrado','Flete devolución cobrado',
                             'Pago anticipado orden (COD)','Mantenimiento tarjeta virtual',
                             'Devolución por garantía','Indemnización'}
            df_trans_det = dc_[dc_['_clasif'].isin(CLASIFS_TRANS)].copy()
            if not df_trans_det.empty:
                with st.expander(f"＋ Ver todos los registros de transacciones ({len(df_trans_det)})"):
                    cols_dt = [c for c in ['_fecha','_desc','_clasif','_tipo','_monto'] if c in df_trans_det.columns]
                    rename_dt = {'_fecha':'Fecha','_desc':'Descripción','_tipo':'Tipo','_monto':'Monto','_clasif':'Categoría'}
                    df_show = df_trans_det[cols_dt].rename(columns=rename_dt)
                    df_show['Monto'] = df_show['Monto'].apply(fmt_cop)
                    st.dataframe(df_show, use_container_width=True, hide_index=True, height=380)

        # ════════════════════════════════════════
        # SECCIÓN 2: ÓRDENES SIN RECAUDO
        # ════════════════════════════════════════
        if sin_recaudo_n > 0:
            fin_title("📦","Órdenes sin recaudo")
            sr1, sr2 = st.columns(2)
            with sr1:
                st.markdown(f'''
                <div class="big-kpi">
                    <div style="font-size:13px;font-weight:700;color:#94a3b8;margin-bottom:8px">Registros</div>
                    <div style="font-size:28px;font-weight:800;color:#f1f5f9">{sin_recaudo_n}</div>
                </div>''', unsafe_allow_html=True)
            with sr2:
                st.markdown(f'''
                <div class="big-kpi">
                    <div style="font-size:13px;font-weight:700;color:#94a3b8;margin-bottom:8px">Total</div>
                    <div style="font-size:28px;font-weight:800;color:#f1f5f9">{fmt_cop(sin_recaudo_val)}</div>
                </div>''', unsafe_allow_html=True)

        # ════════════════════════════════════════
        # SECCIÓN 3: MOVIMIENTOS OMITIDOS
        # ════════════════════════════════════════
        if dc_ is not None:
            movimientos_omitidos = []
            for clasif in ['Pago anticipado orden (COD)',
                           'Recarga tarjeta — Pauta publicitaria',
                           'Recarga tarjeta — Gastos fijos',
                           'Recarga tarjeta — Otra',
                           'Retiro a cuenta bancaria',
                           'Transferencia entre wallets',
                           'Recarga de wallet']:
                rows_c = dc_[dc_['_clasif']==clasif]
                if len(rows_c):
                    movimientos_omitidos.append((
                        clasif.upper().replace(' — ',' '),
                        len(rows_c),
                        float(rows_c['_monto'].sum()),
                        rows_c['_tipo'].iloc[0]=='ENTRADA'
                    ))
            # También incluir otras entradas/salidas no clasificadas — mostrar nombre real
            for clasif in ['Otra entrada','Otra salida','Otro']:
                rows_c = dc_[dc_['_clasif']==clasif]
                if len(rows_c):
                    # Agrupar por descripción real y mostrar cada una por separado
                    for desc_real, grp in rows_c.groupby('_desc'):
                        nombre_real = str(desc_real).strip()
                        if not nombre_real or nombre_real.lower() in ('nan','none',''):
                            nombre_real = clasif.upper()
                        is_ing = grp['_tipo'].iloc[0] == 'ENTRADA'
                        movimientos_omitidos.append((
                            nombre_real,
                            len(grp),
                            float(grp['_monto'].sum()),
                            is_ing
                        ))

            if movimientos_omitidos:
                fin_title("🔄","Movimientos omitidos")
                st.markdown('''
                <div style="background:#3d2a00;border:1px solid #f59e0b;border-radius:8px;padding:10px 14px;margin-bottom:14px;color:#fcd34d;font-size:12px">
                    ⚠️ Los siguientes movimientos <strong>no se tienen en cuenta en el cálculo de la utilidad</strong> y se muestran aquí de manera informativa.
                </div>''', unsafe_allow_html=True)

                mh = st.columns([5,2,2])
                mh[0].markdown("**Movimiento**")
                mh[1].markdown("**Registros**")
                mh[2].markdown("**Total**")
                st.markdown('<hr style="border-color:#334155;margin:6px 0">', unsafe_allow_html=True)

                for lbl, cnt, val, is_ing in movimientos_omitidos:
                    color = C['ent'] if is_ing else C['dev']
                    mr = st.columns([5,2,2])
                    mr[0].markdown(f"<span style='font-size:12px;color:#94a3b8'>{lbl}</span>", unsafe_allow_html=True)
                    mr[1].markdown(f"<span style='font-size:12px;color:#64748b'>{cnt}</span>", unsafe_allow_html=True)
                    mr[2].markdown(f"<span style='font-size:13px;font-weight:700;color:{color}'>{fmt_cop(val)}</span>", unsafe_allow_html=True)
                    st.markdown('<hr style="border-color:#1e293b;margin:3px 0">', unsafe_allow_html=True)

                # ── + Ver todos los registros de movimientos omitidos ──
                df_omit_det = dc_[dc_['_clasif'].isin(SOLO_MOVIMIENTOS)].copy()
                df_omit_det = pd.concat([
                    df_omit_det,
                    dc_[dc_['_clasif'].isin({'Otra entrada','Otra salida','Otro'})]
                ], ignore_index=True)
                if not df_omit_det.empty:
                    with st.expander(f"＋ Ver todos los registros de movimientos omitidos ({len(df_omit_det)})"):
                        cols_do = [c for c in ['_fecha','_desc','_clasif','_tipo','_monto'] if c in df_omit_det.columns]
                        rename_do = {'_fecha':'Fecha','_desc':'Descripción','_tipo':'Tipo','_monto':'Monto','_clasif':'Categoría'}
                        df_show_o = df_omit_det[cols_do].rename(columns=rename_do)
                        df_show_o['Monto'] = df_show_o['Monto'].apply(fmt_cop)
                        st.dataframe(df_show_o, use_container_width=True, hide_index=True, height=380)

        # ════════════════════════════════════════
        # SECCIÓN 4: UTILIDAD
        # ════════════════════════════════════════
        fin_title("💵","Utilidad")
        u1, u2, u3 = st.columns(3)
        with u1:
            st.markdown(metric_card("UTILIDAD ACTUAL", fmt_cop(util_actual), C['ent'] if util_actual>=0 else C['dev']), unsafe_allow_html=True)
        with u2:
            pend_label = (f"UTILIDAD PENDIENTE ({n_pendiente_pancake} ENTREGADAS HOY)" if n_pendiente_pancake > 0
                         else "UTILIDAD PENDIENTE (0 ENTREGADAS HOY)" if dp_raw_det is not None
                         else "UTILIDAD PENDIENTE (SIN ARCHIVO PANCAKE)")
            st.markdown(metric_card(pend_label, fmt_cop(util_pendiente), C['ent'] if util_pendiente>=0 else C['dev']), unsafe_allow_html=True)
        with u3:
            st.markdown(metric_card("UTILIDAD TOTAL", fmt_cop(util_actual+util_pendiente), C['ent'] if (util_actual+util_pendiente)>=0 else C['dev']), unsafe_allow_html=True)

        # ROI
        if inversion_ads > 0:
            st.markdown("---")
            st.markdown("#### RETORNO DE INVERSIÓN PUBLICITARIA")
            roi_c = st.columns([2,2,1])
            roi_c[0].markdown(f"<div class='big-kpi'><div style='font-size:12px;color:#64748b;margin-bottom:4px'>Inversión publicitaria</div><div style='font-size:20px;font-weight:800;color:#f1f5f9'>{fmt_cop(inversion_ads)}</div></div>", unsafe_allow_html=True)
            roi_c[1].markdown(f"<div class='big-kpi'><div style='font-size:12px;color:#64748b;margin-bottom:4px'>Utilidad actual</div><div style='font-size:20px;font-weight:800;color:#f1f5f9'>{fmt_cop(util_actual)}</div></div>", unsafe_allow_html=True)
            roi_color = C['ent'] if roi>=0 else C['dev']
            roi_c[2].markdown(f"<div class='big-kpi'><div style='font-size:12px;color:#64748b;margin-bottom:4px'>ROI</div><div style='font-size:28px;font-weight:800;color:{roi_color}'>{roi:.2f}%</div></div>", unsafe_allow_html=True)

        # Otros datos
        st.markdown("---")
        st.markdown("#### Otros datos")
        od = st.columns(4)
        otros_datos = [
            ("CPA",                             fmt_cop(cpa_total),   C['teal']),
            ("CPA ORDEN DESPACHADA",             fmt_cop(cpa_desp),    C['teal']),
            ("CPA ORDEN ENTREGADA",              fmt_cop(cpa_ent),     C['teal']),
            ("TICKET PROMEDIO",                  fmt_cop(ticket_prom), C['teal']),
        ]
        for col, (lbl,val,color) in zip(od, otros_datos):
            with col: st.markdown(metric_card(lbl,val,color,22), unsafe_allow_html=True)

        od2 = st.columns(4)
        otros_datos2 = [
            ("UTILIDAD PROMEDIO POR VENTA\nANTES DE GASTOS",   fmt_cop(util_x_antes),  C['ent'] if util_x_antes>=0 else C['dev']),
            ("UTILIDAD PROMEDIO POR VENTA\nDESPUÉS DE GASTOS", fmt_cop(util_x_desp),   C['ent'] if util_x_desp>=0  else C['dev']),
            ("PROMEDIO DE VENTAS DIARIAS\nDESPACHADAS",        f"{ventas_dia:.1f}",     C['teal']),
            ("ROAS REAL",                                       f"{roas:.2f}x",          C['purple']),
        ]
        for col, (lbl,val,color) in zip(od2, otros_datos2):
            with col: st.markdown(metric_card(lbl.replace('\n','<br>'),val,color,22), unsafe_allow_html=True)

        # ════════════════════════════════════════
        # SECCIÓN 5: GUÍAS PENDIENTES (simulador)
        # ════════════════════════════════════════
        fin_title("⏳","Guías pendientes")

        if n_sinf > 0:
            st.markdown(f'''
            <div style="background:#1e293b;border:1px solid #334155;border-radius:10px;padding:14px 18px;margin-bottom:18px">
                <div style="font-size:13px;color:#cbd5e1;margin-bottom:10px">
                    Tienes <strong style="color:#f59e0b">{n_sinf} guías</strong> pendientes por finalizar.
                    Mueve la barra para simular los resultados aproximados de acuerdo a la cantidad de guías
                    pendientes que terminen siendo entregadas al destinatario.
                </div>
                <div style="font-size:11px;color:#64748b">Guías entregadas →</div>
            </div>''', unsafe_allow_html=True)

            sim_ent = st.slider(
                "Guías pendientes que terminarán ENTREGADAS",
                min_value=0, max_value=n_sinf, value=n_sinf//2,
                key=f"sim_slider_{inf['id']}",
                label_visibility="collapsed"
            )
            sim_dev = n_sinf - sim_ent

            # Cálculos de simulación
            total_ent_sim  = n_ent + sim_ent
            total_dev_sim  = n_dev + sim_dev
            total_des_sim  = n_desp

            pct_ent_sim = pct(total_ent_sim, total_des_sim)
            pct_dev_sim = pct(total_dev_sim, total_des_sim)

            sim_recaudo      = sim_ent * avg_ticket
            sim_cos          = sim_ent * avg_cos
            sim_fle_ent      = sim_ent * avg_fle_ent
            sim_fle_dev      = sim_dev * avg_fle_dev
            sim_ganancia     = sim_ent * avg_ganancia
            sim_util_pendientes = sim_ganancia - sim_fle_dev
            utilidad_final_sim  = utilidad_bruta + sim_util_pendientes

            # Fila 1: % totales
            sc1, sc2 = st.columns(2)
            with sc1:
                st.markdown(f'''
                <div style="background:#1e293b;border:1px solid #334155;border-radius:12px;padding:20px;text-align:center">
                    <div style="font-size:36px;font-weight:800;color:{C['ent']}">{pct_ent_sim:.2f}%</div>
                    <div style="font-size:13px;font-weight:700;color:{C['ent']};margin:4px 0">{total_ent_sim:,} GUÍAS</div>
                    <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#64748b">ENTREGA TOTAL</div>
                </div>''', unsafe_allow_html=True)
            with sc2:
                st.markdown(f'''
                <div style="background:#1e293b;border:1px solid #334155;border-radius:12px;padding:20px;text-align:center">
                    <div style="font-size:36px;font-weight:800;color:{C['dev']}">{pct_dev_sim:.2f}%</div>
                    <div style="font-size:13px;font-weight:700;color:{C['dev']};margin:4px 0">{total_dev_sim:,} GUÍAS</div>
                    <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#64748b">DEVOLUCIÓN TOTAL</div>
                </div>''', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Fila 2: detalle pendientes
            sp1, sp2, sp3 = st.columns(3)
            with sp1: st.markdown(metric_card("PENDIENTES A ENTREGAR",   f"{sim_ent:,}",           C['ent'],  28), unsafe_allow_html=True)
            with sp2: st.markdown(metric_card("PENDIENTES A DEVOLVER",   f"{sim_dev:,}",           C['dev'],  28), unsafe_allow_html=True)
            with sp3: st.markdown(metric_card("PENDIENTE POR RECOLECTAR",fmt_cop(sim_recaudo),     C['teal'], 28), unsafe_allow_html=True)

            sp4, sp5, sp6 = st.columns(3)
            with sp4: st.markdown(metric_card("COSTO MERCANCÍA DROPSHIPPING", fmt_cop(sim_cos),     C['dev'],  22), unsafe_allow_html=True)
            with sp5: st.markdown(metric_card("PENDIENTE FLETE ENTREGAS",     fmt_cop(sim_fle_ent), C['sinf'], 22), unsafe_allow_html=True)
            with sp6: st.markdown(metric_card("PENDIENTE FLETE DEVOLUCIONES", fmt_cop(sim_fle_dev), C['dev'],  22), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            sf1, sf2 = st.columns(2)
            with sf1: st.markdown(metric_card("UTILIDAD GUÍAS PENDIENTES", fmt_cop(sim_util_pendientes), C['ent'] if sim_util_pendientes>=0 else C['dev'], 28), unsafe_allow_html=True)
            with sf2: st.markdown(metric_card("UTILIDAD FINAL",            fmt_cop(utilidad_final_sim),  C['ent'] if utilidad_final_sim>=0  else C['dev'], 28), unsafe_allow_html=True)

        else:
            st.markdown('<div class="ok-card">✅ Sin guías pendientes — todas las órdenes tienen estado final</div>', unsafe_allow_html=True)

        # ════════════════════════════════════════
        # SECCIÓN 6: EXPORTAR
        # ════════════════════════════════════════
        st.markdown("---")
        fin_title("⬇️","Exportar")
        ex1, ex2 = st.columns(2)
        with ex1:
            pg_df = pd.DataFrame({
                'Concepto': ['Recaudo de venta','Costo Mercancía','Flete Entregado',
                             'Flete Devolución','Mantenimiento tarjeta',
                             'UTILIDAD BRUTA','Ads','Gastos operacionales',
                             'EBITDA','RST','ICA','GMF','UTILIDAD NETA'],
                'Valor ($)': [ingresos,-cos_ent,-fle_ent,-cdev,-mant_total,
                              utilidad_bruta,-inversion_ads,-gastos_op,
                              ebitda,-rst_val,-ica_val,-gmf_val,utilidad_neta]
            })
            st.download_button("⬇️ Descargar P&G (CSV)", data=pg_df.to_csv(index=False).encode('utf-8-sig'),
                file_name=f"PG_{inf['title'].replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.csv",
                mime='text/csv', key=f"dl_pg_{inf['id']}")
        with ex2:
            if dc_ is not None:
                cols_e = [c for c in dc_.columns if not c.startswith('_')]
                st.download_button("⬇️ Descargar cartera CSV", data=dc_[cols_e].to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"cartera_{datetime.today().strftime('%Y%m%d')}.csv",
                    mime='text/csv', key=f"dl_cart_e_{inf['id']}")




# ══════════════════════════════════════════════
# MÓDULO: OTROS (Seguimiento, Novedades, etc.)
# ══════════════════════════════════════════════
elif st.session_state.page == 'module':
    mod = st.session_state.nav_mod

    if mod == "📣 Monitoreo campañas":
        st.markdown("# 📣 Monitoreo de Campañas")
        c1,c2=st.columns(2)
        def camp_ui(f,name,emoji,color,key):
            uf=st.file_uploader(f"{emoji} {name} (CSV/Excel)",type=['csv','xlsx','xls'],key=key)
            if uf:
                dfr=load_file(uf)
                if dfr is not None:
                    st.success(f"✅ {name}: {len(dfr):,} filas")
                    with st.expander(f"Columnas — {name}"): st.write(list(dfr.columns))
                    def gc2(df,names):
                        c_=find_col(df,names)
                        return to_n(df[c_]) if c_ else pd.Series([0]*len(df))
                    sp=gc2(dfr,['importe gastado','spend','inversion']).sum()
                    rs=gc2(dfr,['resultados','conversiones','compras','results']).sum()
                    im=gc2(dfr,['impresiones','impressions']).sum()
                    cl=gc2(dfr,['clics en el enlace','link clicks','clics']).sum()
                    k1,k2,k3,k4=st.columns(4)
                    k1.metric("💸 Inversión",fmt_cop(sp))
                    k2.metric("🎯 Resultados",f"{int(rs):,}",delta=f"CPR: {fmt_cop(sp/rs) if rs else '—'}")
                    k3.metric("👁 Impresiones",f"{int(im):,}",delta=f"CTR: {pct(cl,im):.2f}%")
                    k4.metric("🖱 Clics",f"{int(cl):,}",delta=f"CPC: {fmt_cop(sp/cl) if cl else '—'}")
                    st.dataframe(dfr,use_container_width=True,hide_index=True,height=280)
            else:
                st.info(f"Carga el CSV de {name}")
        with c1: camp_ui(None,"Meta Ads","📘","#3b82f6","meta_c")
        with c2: camp_ui(None,"TikTok Ads","🎵","#8b5cf6","tik_c")

    elif mod == "🔄 Novedades":
        st.markdown("# 🔔 Novedades")

        # ── Obtener datos del último informe sincronizado ──
        inf_list = [i for i in st.session_state.informes if i.get('history')]
        if not inf_list:
            st.info("Sincroniza un informe primero para ver las novedades.")
            st.stop()

        inf_sel_n = st.selectbox("Informe", [i['title'] for i in inf_list], key="nov_inf_sel")
        inf_n = next(i for i in inf_list if i['title'] == inf_sel_n)

        # Obtener dataframes del informe seleccionado
        dropi_raw = st.session_state.get(f"raw_dropi_{inf_n['id']}")
        panc_raw  = st.session_state.get(f"raw_panc_{inf_n['id']}")

        if dropi_raw is None:
            st.warning("⚠️ No hay datos de Dropi para este informe. Abre el informe y carga los archivos.")
            st.stop()

        # Columnas Dropi
        dc_e   = find_col(dropi_raw, ['ESTATUS','estatus'])
        dc_gu  = find_col(dropi_raw, ['NÚMERO GUIA','numero guia','NUMERO GUIA'])
        dc_nov = find_col(dropi_raw, ['NOVEDAD','novedad'])
        dc_sol = find_col(dropi_raw, ['SOLUCIÓN','solucion','SOLUCION'])
        dc_nr  = find_col(dropi_raw, ['FUE SOLUCIONADA LA NOVEDAD','fue solucionada'])
        dc_tr  = find_col(dropi_raw, ['TRANSPORTADORA','transportadora'])
        dc_ci  = find_col(dropi_raw, ['CIUDAD DESTINO','ciudad destino'])
        dc_no  = find_col(dropi_raw, ['NOMBRE CLIENTE','nombre cliente'])
        dc_tel = find_col(dropi_raw, ['TELÉFONO','TELEFONO','telefono'])
        dc_fn  = find_col(dropi_raw, ['FECHA DE NOVEDAD','fecha de novedad'])
        dc_dep = find_col(dropi_raw, ['DEPARTAMENTO DESTINO','departamento destino'])
        dc_val = find_col(dropi_raw, ['TOTAL DE LA ORDEN','VALOR FACTURADO'])
        dc_fle = find_col(dropi_raw, ['PRECIO FLETE','precio flete'])
        dc_prod= find_col(dropi_raw, ['PRODUCTO','producto'])
        dc_vend= find_col(dropi_raw, ['VENDEDOR','vendedor'])
        dc_umov= find_col(dropi_raw, ['ÚLTIMO MOVIMIENTO','ULTIMO MOVIMIENTO'])
        dc_fum = find_col(dropi_raw, ['FECHA DE ÚLTIMO MOVIMIENTO','FECHA DE ULTIMO MOVIMIENTO'])

        # Filtrar órdenes con novedad
        mask_nov = pd.Series([False]*len(dropi_raw), index=dropi_raw.index)
        if dc_nov:
            mask_nov = mask_nov | (dropi_raw[dc_nov].notna() &
                       dropi_raw[dc_nov].astype(str).str.strip().str.lower().ne('nan') &
                       dropi_raw[dc_nov].astype(str).str.strip().ne(''))
        if dc_e:
            mask_nov = mask_nov | dropi_raw[dc_e].astype(str).str.upper().str.strip().isin({'NOVEDAD','NOVEDAD SOLUCIONADA'})

        df_nov = dropi_raw[mask_nov].copy()

        if df_nov.empty:
            st.markdown('<div class="ok-card">✅ Sin novedades en el informe</div>', unsafe_allow_html=True)
            st.stop()

        # Determinar si está resuelta
        df_nov['_resuelta'] = False
        if dc_nr:
            df_nov['_resuelta'] = df_nov[dc_nr].astype(str).str.lower().str.strip().isin(['si','sí','yes','true'])
        elif dc_e:
            df_nov['_resuelta'] = df_nov[dc_e].astype(str).str.upper().str.strip() == 'NOVEDAD SOLUCIONADA'

        total_n = len(df_nov)
        resueltas = int(df_nov['_resuelta'].sum())
        abiertas  = total_n - resueltas

        # KPIs
        nk1, nk2, nk3 = st.columns(3)
        nk1.metric("🔔 Total novedades", f"{total_n:,}")
        nk2.metric("🔴 Abiertas", f"{abiertas:,}")
        nk3.metric("🟢 Resueltas", f"{resueltas:,}")
        st.markdown("---")

        # Sugerencias de solución por tipo de novedad
        SUGERENCIAS = {
            'cliente ausente':         '📞 Llamar al cliente e intentar reagendar entrega. Si no contesta en 24h, gestionar devolución.',
            'direccion':               '📍 Verificar dirección con el cliente por WhatsApp y actualizar en Dropi.',
            'destinatario se rehusa':  '💬 Contactar al cliente para entender el motivo. Ofrecer solución o gestionar devolución.',
            'coordinar':               '📅 Contactar al cliente para coordinar horario de entrega.',
            'mercancia':               '📦 Verificar con el proveedor el estado del producto. Gestionar reenvío si es necesario.',
            'indemnizacion':           '⚖️ Escalar a soporte de Dropi con evidencias fotográficas del producto.',
            'pedido cancelado':        '❌ Gestionar devolución con la transportadora.',
            'default':                 '📞 Contactar al cliente y a la transportadora para gestionar la novedad.'
        }

        def get_sugerencia(nov_texto):
            if not nov_texto or str(nov_texto).strip().lower() in ('nan','none',''): return SUGERENCIAS['default']
            n = str(nov_texto).lower()
            for key, sug in SUGERENCIAS.items():
                if key in n: return sug
            return SUGERENCIAS['default']

        # ── Historial desde Pancake cronograma ──
        def get_historial_pancake(guia_num):
            if panc_raw is None: return []
            guia_col = find_col(panc_raw, ['número de seguimiento','numero de seguimiento'])
            cron_col = find_col(panc_raw, ['cronograma de actualización de estado','cronograma de actualizacion de estado'])
            if not guia_col or not cron_col: return []
            try:
                guia_n = pd.to_numeric(guia_num, errors='coerce')
                panc_raw['_gn'] = pd.to_numeric(panc_raw[guia_col], errors='coerce')
                row = panc_raw[panc_raw['_gn'] == guia_n]
                if row.empty: return []
                cron = str(row.iloc[0][cron_col])
                events = []
                for entry in cron.split(';'):
                    m = re.match(r'^(.+?)\s*-\s*(\d{2}:\d{2}\s+\d{2}/\d{2}/\d{4})\s*-\s*(.+)$', entry.strip())
                    if m:
                        events.append({'estado': m.group(1).strip(), 'fecha': m.group(2).strip(), 'actor': m.group(3).strip()})
                return events
            except: return []

        # ── Render tabla de novedades ──
        st.markdown(f'<div style="font-size:13px;color:#64748b;margin-bottom:12px">20 de {total_n} resultados</div>', unsafe_allow_html=True)

        # Inicializar estado de expansión
        if f'nov_expanded_{inf_n["id"]}' not in st.session_state:
            st.session_state[f'nov_expanded_{inf_n["id"]}'] = {}
        if f'nov_hidden_{inf_n["id"]}' not in st.session_state:
            st.session_state[f'nov_hidden_{inf_n["id"]}'] = set()

        expanded_map = st.session_state[f'nov_expanded_{inf_n["id"]}']
        hidden_set   = st.session_state[f'nov_hidden_{inf_n["id"]}']

        # Header tabla
        hcols = st.columns([2,2,3,3,1])
        for col, lbl in zip(hcols, ['ORDEN','GUÍA','NOVEDAD','DESTINATARIO','OPCIONES']):
            col.markdown(f'<div style="font-size:11px;font-weight:700;color:#0ea5e9;text-transform:uppercase;letter-spacing:.8px;border-bottom:2px solid #0ea5e9;padding-bottom:6px">{lbl}</div>', unsafe_allow_html=True)

        for idx, row in df_nov.iterrows():
            guia    = str(row[dc_gu]).strip() if dc_gu else '—'
            if idx in hidden_set:
                continue

            nov_txt = str(row[dc_nov]).strip() if dc_nov else ''
            if nov_txt.lower() in ('nan','none',''): nov_txt = '—'
            est_txt = str(row[dc_e]).strip() if dc_e else ''
            nom_txt = str(row[dc_no]).strip() if dc_no else '—'
            tel_txt = str(row[dc_tel]).strip() if dc_tel else '—'
            trans   = str(row[dc_tr]).strip() if dc_tr else '—'
            ciudad  = str(row[dc_ci]).strip() if dc_ci else '—'
            resuelta= bool(row['_resuelta'])
            estado_color = '#10b981' if resuelta else '#f59e0b'
            estado_lbl   = '✅ Solucionada' if resuelta else '🔴 Abierta'

            st.markdown('<hr style="border-color:#1e293b;margin:6px 0">', unsafe_allow_html=True)
            rcols = st.columns([2,2,3,3,1])

            with rcols[0]:
                orden_id = str(row.get('ID','') if 'ID' in row.index else idx)
                st.markdown(f'<div style="font-size:11px;color:#64748b">ID</div><div style="font-size:13px;font-weight:700;color:#f1f5f9">{orden_id}</div>', unsafe_allow_html=True)
                fn_raw = str(row[dc_fn]).strip() if dc_fn and str(row[dc_fn]).strip() not in ('nan','None','') else ''
                try:
                    from datetime import datetime as _dt2
                    if '/' in fn_raw:
                        _p = fn_raw.split('/')
                        fn_txt = f"{int(_p[0]):02d} {MESES_ES[int(_p[1])]} {_p[2][:4]}"
                    elif '-' in fn_raw and len(fn_raw) >= 10:
                        fn_txt = f"{int(fn_raw[8:10]):02d} {MESES_ES[int(fn_raw[5:7])]} {fn_raw[:4]}"
                    else:
                        fn_txt = fn_raw
                except:
                    fn_txt = fn_raw
                if fn_txt:
                    st.markdown(f'<div style="font-size:10px;color:#64748b">F. Novedad: {fn_txt}</div>', unsafe_allow_html=True)

            with rcols[1]:
                st.markdown(f'<div style="font-size:11px;color:#64748b">Transportadora</div><div style="font-size:13px;font-weight:600;color:#f1f5f9">{trans}</div>', unsafe_allow_html=True)
                st.markdown(f'<div style="font-size:11px;color:#64748b;margin-top:4px">Número de guía</div><div style="font-size:12px;color:#cbd5e1">{guia}</div>', unsafe_allow_html=True)
                st.markdown(f'<div style="font-size:11px;color:#64748b;margin-top:4px">Estado</div><div style="font-size:12px;font-weight:600;color:{estado_color}">{est_txt}</div>', unsafe_allow_html=True)

            with rcols[2]:
                st.markdown(f'<div style="font-size:11px;color:#64748b">Novedad</div><div style="font-size:13px;font-weight:600;color:#f59e0b">{nov_txt}</div>', unsafe_allow_html=True)
                st.markdown(f'<div style="font-size:11px;color:#64748b;margin-top:6px">Estado resolución</div><div style="font-size:12px;color:{estado_color}">{estado_lbl}</div>', unsafe_allow_html=True)

            with rcols[3]:
                st.markdown(f'<div style="font-size:11px;color:#64748b">Nombre destinatario</div><div style="font-size:13px;font-weight:600;color:#f1f5f9">{nom_txt}</div>', unsafe_allow_html=True)
                st.markdown(f'<div style="font-size:11px;color:#64748b;margin-top:4px">Teléfono</div><div style="font-size:12px;color:#cbd5e1">{tel_txt}</div>', unsafe_allow_html=True)

            with rcols[4]:
                exp_key = f"nov_exp_{inf_n['id']}_{idx}"
                if st.button("ver más" if not expanded_map.get(idx) else "ver menos", key=exp_key, use_container_width=True):
                    expanded_map[idx] = not expanded_map.get(idx, False)
                    st.rerun()

            # ── Panel expandido ──
            if expanded_map.get(idx, False):
                prod_txt = str(row[dc_prod]).strip() if dc_prod else '—'
                dep_txt  = str(row[dc_dep]).strip()  if dc_dep  else '—'
                val_txt  = fmt_cop(pd.to_numeric(row[dc_val], errors='coerce')) if dc_val else '—'
                fle_txt  = fmt_cop(pd.to_numeric(row[dc_fle], errors='coerce')) if dc_fle else '—'
                sol_txt  = str(row[dc_sol]).strip() if dc_sol and str(row[dc_sol]).strip() not in ('nan','None','') else '—'
                vend_txt = str(row[dc_vend]).strip() if dc_vend else '—'
                umov_txt = str(row[dc_umov]).strip() if dc_umov else '—'
                fum_txt  = str(row[dc_fum]).strip()  if dc_fum  else '—'
                sugerencia = get_sugerencia(nov_txt)

                with st.container():
                    st.markdown(f'''
                    <div style="background:#162032;border:1.5px solid #0ea5e9;border-radius:10px;padding:16px;margin:8px 0">
                    ''', unsafe_allow_html=True)

                    d1, d2, d3, d4 = st.columns(4)
                    with d1:
                        st.markdown("**📦 Orden**")
                        st.markdown(f"**Transportadora:** {trans}  \n**Número guía:** {guia}  \n**Estado:** {est_txt}  \n**Últ. movimiento:** {umov_txt}  \n**F. últ. movimiento:** {fum_txt}")
                    with d2:
                        st.markdown("**🏷 Producto**")
                        st.markdown(f"**Producto:** {prod_txt}  \n**Total orden:** {val_txt}  \n**Flete:** {fle_txt}  \n**Vendedor:** {vend_txt}")
                    with d3:
                        st.markdown("**👤 Destinatario**")
                        st.markdown(f"**Nombre:** {nom_txt}  \n**Teléfono:** {tel_txt}  \n**Ciudad:** {ciudad}  \n**Departamento:** {dep_txt}")
                    with d4:
                        st.markdown("**🔔 Novedad**")
                        st.markdown(f"**Novedad:** {nov_txt}  \n**Solución registrada:** {sol_txt}  \n**Resuelta:** {'Sí ✅' if resuelta else 'No 🔴'}")

                    # Sugerencia de solución
                    st.markdown(f'''
                    <div style="background:#1e3a5f;border:1px solid #3b82f6;border-radius:8px;
                                padding:10px 14px;margin-top:12px;font-size:13px;color:#93c5fd">
                        💡 <strong>¿Cómo solucionar?</strong> {sugerencia}
                    </div>''', unsafe_allow_html=True)

                    # Historial del destinatario desde Pancake
                    historial = get_historial_pancake(guia)
                    if historial:
                        st.markdown("**📋 Historial del destinatario (Pancake)**")
                        for ev in historial:
                            st.markdown(f'<div style="font-size:12px;color:#94a3b8;padding:2px 0">• <strong style="color:#f1f5f9">{ev["estado"]}</strong> — {ev["fecha"]} — {ev["actor"]}</div>', unsafe_allow_html=True)

                    # Botones de acción
                    st.markdown('<div style="margin-top:12px"></div>', unsafe_allow_html=True)
                    ba1, ba2, ba3, ba4 = st.columns(4)
                    with ba1:
                        st.markdown(f'<a href="tel:{tel_txt}" style="display:block;text-align:center;background:#1e293b;border:1px solid #334155;border-radius:6px;padding:6px;font-size:12px;color:#f1f5f9;text-decoration:none">📞 Llamar</a>', unsafe_allow_html=True)
                    with ba2:
                        wa_msg = f"Hola {nom_txt}, te contactamos sobre tu pedido {guia}. ¿Podemos ayudarte?"
                        st.markdown(f'<a href="https://wa.me/57{tel_txt}?text={wa_msg}" target="_blank" style="display:block;text-align:center;background:#1e293b;border:1px solid #334155;border-radius:6px;padding:6px;font-size:12px;color:#25d366;text-decoration:none">💬 WhatsApp</a>', unsafe_allow_html=True)
                    with ba3:
                        if st.button("✅ Marcar resuelta", key=f"res_{inf_n['id']}_{idx}", use_container_width=True):
                            st.toast(f"Novedad {guia} marcada como resuelta")
                    with ba4:
                        if st.button("🙈 Ocultar fila", key=f"hide_{inf_n['id']}_{idx}", use_container_width=True):
                            hidden_set.add(idx)
                            expanded_map[idx] = False
                            st.rerun()

                    st.markdown('</div>', unsafe_allow_html=True)

    elif mod == "🗺 Seguimiento":
        st.markdown("# 🗺 Seguimiento de Guías")

        # ── Obtener datos del último informe ──
        inf_list_s = [i for i in st.session_state.informes if i.get('history')]
        if not inf_list_s:
            st.info("Sincroniza un informe primero.")
            st.stop()

        inf_sel_s = st.selectbox("Informe", [i['title'] for i in inf_list_s], key="seg_inf_sel")
        inf_s = next(i for i in inf_list_s if i['title'] == inf_sel_s)

        dropi_s = st.session_state.get(f"raw_dropi_{inf_s['id']}")
        panc_s  = st.session_state.get(f"raw_panc_{inf_s['id']}")

        if dropi_s is None:
            st.warning("⚠️ No hay datos de Dropi para este informe. Abre el informe y carga los archivos.")
            st.stop()

        # Columnas
        dc_e  = find_col(dropi_s, ['ESTATUS','estatus'])
        dc_gu = find_col(dropi_s, ['NÚMERO GUIA','numero guia','NUMERO GUIA'])
        dc_no = find_col(dropi_s, ['NOMBRE CLIENTE','nombre cliente'])
        dc_te = find_col(dropi_s, ['TELÉFONO','TELEFONO','telefono'])
        dc_tr = find_col(dropi_s, ['TRANSPORTADORA','transportadora'])
        dc_ci = find_col(dropi_s, ['CIUDAD DESTINO','ciudad destino'])
        dc_fu = find_col(dropi_s, ['FECHA DE ÚLTIMO MOVIMIENTO','FECHA DE ULTIMO MOVIMIENTO'])
        dc_hu = find_col(dropi_s, ['HORA DE ÚLTIMO MOVIMIENTO','HORA DE ULTIMO MOVIMIENTO'])
        dc_um = find_col(dropi_s, ['ÚLTIMO MOVIMIENTO','ULTIMO MOVIMIENTO'])
        dc_pr = find_col(dropi_s, ['PRODUCTO','producto'])
        dc_va = find_col(dropi_s, ['TOTAL DE LA ORDEN','VALOR FACTURADO'])
        dc_fl = find_col(dropi_s, ['PRECIO FLETE','precio flete'])
        dc_nov= find_col(dropi_s, ['NOVEDAD','novedad'])

        if not dc_e:
            st.warning("No se detectó columna ESTATUS"); st.stop()

        now = datetime.now()
        hoy = now.date()

        # ── Detectar órdenes en oficina desde Pancake ──
        oficina_guias = {}  # guia → dias_en_oficina
        if panc_s is not None:
            try:
                etiq_col = find_col(panc_s, ['etiqueta','label'])
                guia_p   = find_col(panc_s, ['número de seguimiento','numero de seguimiento'])
                dia_act  = find_col(panc_s, ['día de actualización','dia de actualizacion','día de actualizacion'])
                cron_col = find_col(panc_s, ['cronograma de actualización de estado','cronograma de actualizacion de estado'])

                ESTADOS_FINALES_S = {'entregado','devuelto','devolucion','devolución','cancelado','rechazado'}

                def _tiene_final_s(cron):
                    if not cron or str(cron).strip() in ('nan','None',''): return False
                    for entry in str(cron).split(';'):
                        m = re.match(r'^(.+?)\s*-\s*\d{2}:\d{2}', entry.strip())
                        if m:
                            est = m.group(1).strip().lower()
                            if any(ef in est for ef in ESTADOS_FINALES_S): return True
                    return False

                def parse_f_s(s):
                    for fmt in ['%d/%m/%Y','%d-%m-%Y','%Y-%m-%d']:
                        try: return datetime.strptime(str(s).strip()[:10], fmt).date()
                        except: pass
                    return None

                if etiq_col and guia_p and dia_act:
                    df_ofic = panc_s[panc_s[etiq_col].astype(str).str.lower().str.strip() == 'pick up at office'].copy()
                    if cron_col:
                        df_ofic = df_ofic[~df_ofic[cron_col].apply(_tiene_final_s)]
                    df_ofic['_dias'] = df_ofic[dia_act].apply(
                        lambda x: (hoy - parse_f_s(x)).days if parse_f_s(x) else None
                    )
                    df_ofic['_guia_n'] = pd.to_numeric(df_ofic[guia_p], errors='coerce')
                    for _, row in df_ofic.iterrows():
                        if pd.notna(row['_guia_n']) and pd.notna(row['_dias']):
                            oficina_guias[float(row['_guia_n'])] = int(row['_dias'])
            except: pass

        # Alerta de oficina al entrar
        oficina_retraso = {g: d for g, d in oficina_guias.items() if d > 2}
        if oficina_retraso:
            st.markdown(f'''
            <div style="background:#3d2a00;border:2px solid #f59e0b;border-radius:10px;
                        padding:14px 18px;margin-bottom:16px;color:#fcd34d;font-size:14px">
                ⚠️ <strong>¡Atención!</strong> Hay <strong>{len(oficina_retraso)} órdenes</strong>
                en reclamar en oficina con más de 2 días de retraso.
                La más antigua lleva <strong>{max(oficina_retraso.values())} días</strong>.
            </div>''', unsafe_allow_html=True)

        # ── Selector de lista (imagen 4) ──
        LISTAS_SEG = {
            "Pendientes de confirmación (+2 días)": lambda df: df[dc_e].astype(str).str.upper().str.strip().isin({'NUEVO','CONFIRMADO'}) if dc_e else pd.Series([False]*len(df)),
            "Pendientes de guía (+2 días)":         lambda df: df[dc_e].astype(str).str.upper().str.strip() == 'CONFIRMADO' if dc_e else pd.Series([False]*len(df)),
            "Guía generada (+2 días)":              lambda df: df[dc_e].astype(str).str.upper().str.strip().isin({'EMPACANDO','GUIA GENERADA'}) if dc_e else pd.Series([False]*len(df)),
            "Reclamar en oficina (+2 días)":        None,  # especial — desde Pancake etiqueta
            "En proceso (+7 días)":                 lambda df: df[dc_e].apply(es_s) if dc_e else pd.Series([False]*len(df)),
            "Otros estados":                        lambda df: pd.Series([True]*len(df)),
        }

        lista_sel = st.selectbox(
            "Lista",
            list(LISTAS_SEG.keys()),
            key="seg_lista_sel"
        )

        # Construir DataFrame según lista
        df_seg = dropi_s.copy()

        if lista_sel == "Reclamar en oficina (+2 días)":
            # Filtrar por guías en oficina con más de 2 días
            df_seg['_guia_n'] = pd.to_numeric(df_seg[dc_gu], errors='coerce') if dc_gu else None
            mask_ofic = df_seg['_guia_n'].apply(lambda g: g in oficina_retraso if pd.notna(g) else False)
            df_seg = df_seg[mask_ofic].copy()
            if dc_gu:
                df_seg['_dias_oficina'] = df_seg['_guia_n'].map(oficina_retraso)
        else:
            fn = LISTAS_SEG[lista_sel]
            if fn is not None:
                df_seg = df_seg[fn(df_seg)].copy()

        # Calcular horas sin movimiento
        def plm_s(row):
            fu = str(row[dc_fu]) if dc_fu else ''
            hu = str(row[dc_hu]) if dc_hu else ''
            if not fu or fu == 'nan': return None
            try: return pd.to_datetime(f"{fu.strip()[:10]} {hu.strip()[:5] if hu and hu!='nan' else '00:00'}", dayfirst=True, errors='coerce')
            except: return None

        df_seg['_dt']    = df_seg.apply(plm_s, axis=1)
        df_seg['_horas'] = df_seg['_dt'].apply(lambda dt: round((now - dt).total_seconds()/3600, 1) if pd.notna(dt) else None)
        df_seg['_dias_s']= df_seg['_horas'].apply(lambda h: round(h/24, 1) if h is not None else None)

        # Marcar novedades de oficina
        if lista_sel == "Reclamar en oficina (+2 días)":
            df_seg['_alerta'] = True
        else:
            df_seg['_alerta'] = df_seg['_horas'].apply(lambda h: h is not None and h >= 48)

        # KPIs
        ts = len(df_seg)
        ta = int(df_seg['_alerta'].sum())
        st.markdown(f'<div style="font-size:13px;font-weight:600;color:#94a3b8;margin-bottom:12px">{ts:,} registros — {ta:,} con alerta</div>', unsafe_allow_html=True)

        if ta > 0:
            alerta_msg = (f"⚠️ {ta} órdenes llevan más de 2 días en oficina sin ser reclamadas"
                         if lista_sel == "Reclamar en oficina (+2 días)"
                         else f"⚠️ {ta} guías llevan más de 48 horas sin movimiento")
            st.markdown(f'<div class="alerta-card">{alerta_msg}</div>', unsafe_allow_html=True)

        # Filtros adicionales
        f1, f2 = st.columns(2)
        with f1:
            trans_opts = ['Todas'] + sorted(dropi_s[dc_tr].dropna().unique().tolist()) if dc_tr else ['Todas']
            ftrans = st.selectbox("Transportadora", trans_opts, key="seg_trans")
        with f2:
            solo_alerta = st.checkbox("Solo con alerta", key="seg_alerta")

        dm = df_seg.copy()
        if ftrans != 'Todas' and dc_tr: dm = dm[dm[dc_tr] == ftrans]
        if solo_alerta: dm = dm[dm['_alerta']]

        # Mostrar tabla
        cols_show = [c for c in [dc_gu, dc_no, dc_te, dc_tr, dc_e, dc_nov, dc_um, dc_fu, '_horas', '_dias_s', '_alerta'] if c and c in dm.columns]
        if lista_sel == "Reclamar en oficina (+2 días)" and '_dias_oficina' in dm.columns:
            cols_show = ['_dias_oficina'] + cols_show

        rename_map = {'_horas': 'Horas sin mov.', '_dias_s': 'Días sin mov.',
                      '_alerta': '⚠️ Alerta', '_dias_oficina': '⏳ Días en oficina'}
        st.dataframe(dm[cols_show].rename(columns=rename_map),
                     use_container_width=True, hide_index=True, height=450)

        st.download_button("⬇️ Descargar CSV",
            data=dm[cols_show].rename(columns=rename_map).to_csv(index=False).encode('utf-8-sig'),
            file_name=f"seguimiento_{lista_sel[:20].replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.csv",
            mime='text/csv', key="dl_seg")

    elif mod == "💵 Flujo de Caja":
        st.markdown("# 💵 Flujo de Caja")

        # ── Seleccionar informe ──
        inf_list_fc = [i for i in st.session_state.informes if i.get('history')]
        if not inf_list_fc:
            st.info("Sincroniza un informe con Cartera Dropi para ver el flujo de caja.")
            st.stop()

        inf_sel_fc = st.selectbox("Informe", [i['title'] for i in inf_list_fc], key="fc_inf_sel")
        inf_fc = next(i for i in inf_list_fc if i['title'] == inf_sel_fc)

        cart_raw = st.session_state.get(f"raw_cart_{inf_fc['id']}")
        if cart_raw is None:
            st.warning("⚠️ Carga el archivo de Cartera Dropi en el informe para ver el flujo de caja.")
            st.stop()

        # ── Parsear cartera ──
        import unicodedata as _uc

        def _norm_fc(s):
            s = str(s).upper().strip()
            s = _uc.normalize('NFD', s)
            return ''.join(c for c in s if _uc.category(c) != 'Mn')

        dc = cart_raw.copy()
        dc.columns = [c.strip() for c in dc.columns]
        cc_fecha = find_col(dc, ['FECHA','fecha'])
        cc_monto = find_col(dc, ['MONTO','monto'])
        cc_desc  = find_col(dc, ['DESCRIPCIÓN','DESCRIPCION','descripcion'])

        if not cc_fecha or not cc_monto:
            st.error("No se encontraron columnas FECHA o MONTO en la Cartera Dropi.")
            st.stop()

        dc['_fecha_dt'] = pd.to_datetime(dc[cc_fecha].astype(str).str[:10], format='%d-%m-%Y', errors='coerce')
        dc['_monto']    = pd.to_numeric(dc[cc_monto], errors='coerce').fillna(0)
        dc['_norm']     = dc[cc_desc].apply(_norm_fc) if cc_desc else ''

        def _clasif_fc(n):
            if 'GANANCIA EN LA ORDEN' in n or ('GANANCIA' in n and 'ENTRADA' in n): return 'entrada'
            if 'DEVOLUCION' in n and 'ENTREGA NO EFECTIVA' in n: return 'salida_flete'
            if 'COBRO DE FLETE INICIAL' in n or 'FLETE INICIAL' in n: return 'salida_flete'
            if 'NUEVA ORDEN' in n: return 'salida_flete'
            return 'otro'

        dc['_clasif'] = dc['_norm'].apply(_clasif_fc)
        dc['_dia']    = dc['_fecha_dt'].dt.date
        dc['_semana'] = dc['_fecha_dt'].dt.day.apply(lambda d:
            1 if d <= 6 else 7 if d <= 13 else 14 if d <= 20 else 21 if d <= 27 else 28)

        SEMANAS  = [1, 7, 14, 21, 28]
        SEM_LBLS = ['01 a 06', '07 a 13', '14 a 20', '21 a 27', '28 a 31']

        # ── Session state key for manual values (persist between semanal/diario) ──
        manual_key = f"fc_manual_{inf_fc['id']}"
        if manual_key not in st.session_state:
            st.session_state[manual_key] = {}

        # ── Tipo de flujo ──
        st.markdown("---")
        st.markdown("**✏️ Valores manuales por período**")
        tipo_flujo = st.radio("Tipo de flujo", ["📅 Semanal", "📆 Diario"], horizontal=True, key="fc_tipo")

        # Get periods
        dias_con_data = sorted(dc[dc['_dia'].notna()]['_dia'].unique())
        if not dias_con_data:
            st.warning("No hay datos de fechas en la cartera.")
            st.stop()

        def semana_de(d):
            day = d.day if hasattr(d, 'day') else int(str(d).split('-')[-1])
            return 1 if day<=6 else 7 if day<=13 else 14 if day<=20 else 21 if day<=27 else 28

        if tipo_flujo == "📅 Semanal":
            periodos     = SEMANAS
            periodo_lbls = SEM_LBLS
        else:
            periodos     = dias_con_data
            periodo_lbls = [f"{d.day:02d} {MESES_ES[d.month]}" for d in dias_con_data]

        # ── Manual inputs ──
        n_per = len(periodos)
        st.caption(f"Ingresa los valores manuales para cada {'semana' if tipo_flujo == '📅 Semanal' else 'día'}. Los valores del diario se acumulan automáticamente en el semanal.")

        manual_vals = {}

        if tipo_flujo == "📆 Diario":
            # Show daily inputs — store each day in session_state
            cols_per_row = 7
            for row_start in range(0, n_per, cols_per_row):
                row_dias = list(zip(periodos, periodo_lbls))[row_start:row_start+cols_per_row]
                row_cols = st.columns(len(row_dias))
                for col_i, (per, lbl) in zip(row_cols, row_dias):
                    per_str = str(per)
                    saved = st.session_state[manual_key].get(per_str, {'transf': 0.0, 'pub': 0.0})
                    with col_i:
                        st.markdown(f"<div style='font-size:11px;font-weight:700;color:#0ea5e9;text-align:center'>{lbl}</div>", unsafe_allow_html=True)
                        transf = st.number_input("Transf ($)", min_value=0.0,
                            value=float(saved['transf']), step=10000.0, format="%.0f",
                            key=f"fc_transf_{inf_fc['id']}_{per_str}", label_visibility="collapsed")
                        pub = st.number_input("Pub ($)", min_value=0.0,
                            value=float(saved['pub']), step=10000.0, format="%.0f",
                            key=f"fc_pub_{inf_fc['id']}_{per_str}", label_visibility="collapsed")
                        st.session_state[manual_key][per_str] = {'transf': transf, 'pub': pub}
                        manual_vals[per] = {'transf': transf, 'pub': pub}

        else:
            # Semanal — aggregate from daily stored values
            # Also allow direct weekly override if no daily data entered
            cols_inp = st.columns(5)
            for i, (per, lbl) in enumerate(zip(periodos, periodo_lbls)):
                # Sum daily values for this week
                daily_transf = sum(
                    st.session_state[manual_key].get(str(d), {}).get('transf', 0)
                    for d in dias_con_data if semana_de(d) == per
                )
                daily_pub = sum(
                    st.session_state[manual_key].get(str(d), {}).get('pub', 0)
                    for d in dias_con_data if semana_de(d) == per
                )
                with cols_inp[i % 5]:
                    st.markdown(f"**{lbl}**")
                    # If daily values exist, show aggregated (read-only style)
                    if daily_transf > 0 or daily_pub > 0:
                        st.markdown(f"<div style='font-size:11px;color:#64748b'>Transf: <strong style='color:#f1f5f9'>{fmt_cop(daily_transf)}</strong> (suma diaria)</div>", unsafe_allow_html=True)
                        st.markdown(f"<div style='font-size:11px;color:#64748b'>Pub: <strong style='color:#f1f5f9'>{fmt_cop(daily_pub)}</strong> (suma diaria)</div>", unsafe_allow_html=True)
                        manual_vals[per] = {'transf': daily_transf, 'pub': daily_pub}
                    else:
                        # Allow direct weekly entry if no daily data
                        transf = st.number_input("Transf. pag. adelantados ($)",
                            min_value=0.0, value=0.0, step=10000.0, format="%.0f",
                            key=f"fc_transf_w_{inf_fc['id']}_{i}")
                        pub = st.number_input("Consumo publicidad ($)",
                            min_value=0.0, value=0.0, step=10000.0, format="%.0f",
                            key=f"fc_pub_w_{inf_fc['id']}_{i}")
                        manual_vals[per] = {'transf': transf, 'pub': pub}

        # ── Calcular flujo y renderizar como tabla HTML con scroll ──
        st.markdown("---")
        titulo_flujo = "FLUJO DE CAJA SEMANAL" if tipo_flujo == "📅 Semanal" else "FLUJO DE CAJA DIARIO"
        st.markdown(f'<div style="background:#0ea5e9;color:#0f172a;font-size:15px;font-weight:800;padding:10px 16px;border-radius:8px;margin-bottom:16px;text-transform:uppercase;letter-spacing:1px">{titulo_flujo}</div>', unsafe_allow_html=True)

        FILAS = [
            ('entrada',      'Entrada (Ganancias) Dropi',          True),
            ('salida_flete', 'Salida (Fletes + Devoluciones + Nuevas órdenes)', False),
            ('transf',       'Transferencia de pagos adelantados', False),
            ('pub',          'Consumo Publicidad',                  False),
            ('4xmil',        '4 X MIL',                            False),
            ('utilidad',     'UTILIDAD BRUTA',                     True),
        ]

        # Build data
        matrix = {f[0]: [] for f in FILAS}
        for per in periodos:
            if tipo_flujo == "📅 Semanal":
                mask = dc['_semana'] == per
            else:
                mask = dc['_dia'] == per
            ent    = dc[mask & (dc['_clasif']=='entrada')]['_monto'].sum()
            sal    = dc[mask & (dc['_clasif']=='salida_flete')]['_monto'].sum()
            transf = manual_vals.get(per, {}).get('transf', 0)
            pub    = manual_vals.get(per, {}).get('pub', 0)
            xmil   = (ent - sal) * 4 / 1000
            util   = ent - sal - transf - pub - xmil
            matrix['entrada'].append(ent)
            matrix['salida_flete'].append(sal)
            matrix['transf'].append(transf)
            matrix['pub'].append(pub)
            matrix['4xmil'].append(xmil)
            matrix['utilidad'].append(util)

        # Build HTML table
        def cell_color(val, is_util, is_pos):
            if is_util: return '#0f172a'
            if is_pos:  return '#10b981'
            return '#f1f5f9'

        col_w = max(120, 1000 // max(len(periodos), 1))

        html = f'''
        <div style="overflow-x:auto;margin-bottom:16px;border-radius:8px;border:1px solid #1e293b">
        <table style="width:100%;border-collapse:collapse;min-width:{max(800, col_w*(len(periodos)+2))}px">
        <thead>
          <tr style="background:#0f172a;border-bottom:2px solid #334155">
            <th style="text-align:left;padding:10px 14px;font-size:12px;color:#94a3b8;font-weight:700;min-width:220px;position:sticky;left:0;background:#0f172a;z-index:2;border-right:1px solid #334155">Concepto</th>
            <th style="text-align:center;padding:10px 6px;font-size:11px;color:#64748b;font-weight:700;min-width:45px">COP</th>'''

        for lbl in periodo_lbls:
            html += f'<th style="text-align:right;padding:10px 8px;font-size:11px;color:#0ea5e9;font-weight:700;min-width:{col_w}px;white-space:nowrap">{lbl}</th>'

        html += '<th style="text-align:right;padding:10px 12px;font-size:12px;color:#f1f5f9;font-weight:800;min-width:140px;border-left:2px solid #334155">TOTAL</th></tr></thead><tbody>'

        for key, label, is_pos in FILAS:
            is_util = (key == 'utilidad')
            vals  = matrix[key]
            total = sum(vals)
            row_bg   = '#0ea5e9' if is_util else ('rgba(255,255,255,0.03)' if FILAS.index((key,label,is_pos)) % 2 == 0 else 'transparent')
            lbl_col  = '#0f172a' if is_util else '#cbd5e1'
            lbl_w    = '800' if is_util else '600'
            fs       = '14px' if is_util else '12px'

            html += f'<tr style="background:{row_bg}">'
            html += f'<td style="padding:10px 14px;font-size:{fs};font-weight:{lbl_w};color:{lbl_col};position:sticky;left:0;background:{row_bg};z-index:1">{label}</td>'
            html += f'<td style="text-align:center;padding:10px 6px;font-size:10px;color:{"#0f172a" if is_util else "#64748b"}">COP</td>'

            for v in vals:
                c = cell_color(v, is_util, is_pos)
                if is_util and v < 0: c = '#ef4444'
                html += f'<td style="text-align:right;padding:10px 8px;font-size:{fs};font-weight:{lbl_w};color:{c}">{fmt_cop(v)}</td>'

            tc = cell_color(total, is_util, is_pos)
            if is_util and total < 0: tc = '#ef4444'
            html += f'<td style="text-align:right;padding:10px 10px;font-size:{"15px" if is_util else "13px"};font-weight:800;color:{tc};border-left:2px solid #334155">{fmt_cop(total)}</td>'
            html += '</tr>'

        html += '</tbody></table></div>'
        st.markdown(html, unsafe_allow_html=True)

    elif mod == "✅ Tareas programadas":
        st.markdown("# ✅ Tareas Programadas")
        # Cortes
        st.markdown('<div class="sec-title">⏰ CORTES DEL DÍA</div>',unsafe_allow_html=True)
        for i,c in enumerate(st.session_state.cortes):
            cc1,cc2,cc3=st.columns([1,5,2])
            with cc1:
                done=st.checkbox("",value=c['done'],key=f"co_{c['id']}")
                st.session_state.cortes[i]['done']=done
            with cc2: st.markdown(f"~~**{c['hora']}** · {c['label']}~~" if done else f"**{c['hora']}** · {c['label']}")
            with cc3: st.markdown("✅ Listo" if done else "⏳ Pendiente")
        st.markdown("---")
        st.markdown('<div class="sec-title">🔔 ALERTAS DE KPI</div>',unsafe_allow_html=True)
        with st.expander("➕ Agregar alerta"):
            a1,a2,a3=st.columns([2,1,2])
            with a1: ak=st.selectbox("KPI",['% Entrega','% Devolución','% Confirmación','Novedades abiertas'])
            with a2: ac=st.selectbox("Condición",['Menor que','Mayor que'])
            with a3: av=st.number_input("Umbral",min_value=0.0,value=70.0)
            ad=st.text_input("Descripción")
            if st.button("Guardar alerta"):
                st.session_state.alertas.append({'kpi':ak,'cond':ac,'val':av,'desc':ad}); st.success("✅"); st.rerun()
        for i,a in enumerate(st.session_state.alertas):
            al1,al2=st.columns([6,1])
            with al1: st.markdown(f"🔔 **{a.get('desc') or a['kpi']}** — {a['kpi']} {a['cond'].lower()} {a['val']}")
            with al2:
                if st.button("✕",key=f"da_{i}"): st.session_state.alertas.pop(i); st.rerun()
        if not st.session_state.alertas: st.info("Sin alertas")
        st.markdown("---")
        st.markdown('<div class="sec-title">📝 NOTAS Y PENDIENTES</div>',unsafe_allow_html=True)
        with st.expander("➕ Nueva tarea"):
            t1,t2=st.columns([3,2])
            with t1: tt=st.text_input("Título *")
            with t2: tp=st.selectbox("Tipo",['Nota del equipo','Recordatorio','Urgente'])
            tf=st.date_input("Fecha",value=date.today())
            tn=st.text_area("Detalles")
            if st.button("Guardar") and tt:
                st.session_state.tareas.append({'tipo':tp,'titulo':tt,'fecha':str(tf),'nota':tn,'done':False}); st.success("✅"); st.rerun()
        for i,t in enumerate(st.session_state.tareas):
            tc1,tc2,tc3=st.columns([1,6,1])
            with tc1:
                done2=st.checkbox("",value=t['done'],key=f"tk_{i}")
                st.session_state.tareas[i]['done']=done2
            with tc2:
                p="~~" if done2 else ""
                st.markdown(f"{p}**{t['titulo']}** · {t['fecha']} · _{t['tipo']}_{p}")
                if t.get('nota'): st.caption(t['nota'])
            with tc3:
                if st.button("✕",key=f"dt_{i}"): st.session_state.tareas.pop(i); st.rerun()
        if not st.session_state.tareas: st.info("Sin tareas")
