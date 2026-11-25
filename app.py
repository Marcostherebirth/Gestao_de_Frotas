
import streamlit as st
import sqlite3
from pathlib import Path
from datetime import date, datetime, timedelta
import html
import io
import math
import re
import zipfile
import altair as alt
import pandas as pd

APP_DIR = Path(__file__).parent
DB = APP_DIR / "frota.db"
SOURCE_XLSX = APP_DIR / "Controle de Manutenção Preventiva_RL Connect.xlsx"

st.set_page_config(page_title="Gestão de Frota | RL Connect", page_icon="🚘", layout="wide")

if "sidebar_collapsed" not in st.session_state:
    st.session_state.sidebar_collapsed = False


# Paleta inspirada no painel de referência da RL Connect
st.markdown("""
<style>
:root {
  --rl-sidebar: #102f35;
  --rl-sidebar-active: #174b55;
  --rl-teal: #08798a;
  --rl-teal-hover: #066777;
  --rl-bg: #f3f5f6;
  --rl-border: #d7e0e4;
  --rl-muted: #71808b;
}
.stApp { background: var(--rl-bg); color: #17242a; }
[data-testid="stHeader"] { background: rgba(243,245,246,.96); }
.block-container {
  max-width: 1500px;
  padding-top: 1.45rem;
  padding-bottom: 2rem;
  padding-left: 2.1rem;
  padding-right: 2.1rem;
}
h1 { font-size: 1.85rem !important; font-weight: 650 !important; color: #17242a; }
h2 { font-size: 1.35rem !important; font-weight: 620 !important; color: #17242a; }
h3 { font-size: 1.05rem !important; font-weight: 620 !important; color: #17242a; }
p, label, .stCaption { font-size: .88rem; }
[data-testid="stSidebar"] {
  background: var(--rl-sidebar);
  min-width: 220px;
}
[data-testid="stSidebar"] * { color: #e7f0f2; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
  color: #d7e6e8;
}
[data-testid="stMetric"] {
  background: #fff;
  border: 1px solid var(--rl-border);
  border-radius: 7px;
  padding: 13px 15px;
  min-height: 88px;
}
[data-testid="stMetricLabel"] p {
  color: #647987 !important;
  font-size: .72rem !important;
  text-transform: uppercase;
  letter-spacing: .045em;
}
[data-testid="stMetricValue"] {
  font-size: 1.55rem !important;
  color: #15272d;
}
.stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {
  background: var(--rl-teal);
  color: white;
  border: 1px solid var(--rl-teal);
  border-radius: 5px;
  min-height: 2.35rem;
  font-size: .84rem;
  font-weight: 550;
}
.stButton > button:hover, .stFormSubmitButton > button:hover, .stDownloadButton > button:hover {
  background: var(--rl-teal-hover);
  border-color: var(--rl-teal-hover);
  color: white;
}
div[data-baseweb="input"] input, div[data-baseweb="select"] > div,
.stTextArea textarea {
  border-color: var(--rl-border);
  border-radius: 5px;
  background: #fff;
  font-size: .86rem;
}
[data-testid="stTabs"] [data-baseweb="tab-list"] {
  gap: 1.15rem;
  border-bottom: 1px solid var(--rl-border);
}
[data-testid="stTabs"] button[role="tab"] {
  padding: .62rem .1rem;
  color: #53636d;
  font-size: .84rem;
  border-bottom: 2px solid transparent;
}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
  color: var(--rl-teal);
  border-bottom: 2px solid var(--rl-teal);
  font-weight: 650;
}
[data-testid="stDataFrame"] {
  border: 1px solid var(--rl-border);
  border-radius: 6px;
  overflow: hidden;
}
[data-testid="stAlert"] { border-radius: 6px; font-size: .86rem; }
hr { border-color: var(--rl-border); }

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {
  color: #f4fbfc !important;
  font-size: 1.18rem !important;
  letter-spacing: .06em;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h4 {
  color: #91aeb4 !important;
  font-size: .68rem !important;
  letter-spacing: .10em;
  margin-top: 1.1rem;
  margin-bottom: .35rem;
}
[data-testid="stSidebar"] .stButton > button {
  text-align: left;
  justify-content: flex-start;
  min-height: 2.45rem;
  padding-left: .8rem;
  border-radius: 6px;
  background: transparent;
  border-color: transparent;
  color: #dcebed;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
  background: #174b55;
  border-color: #174b55;
  color: #ffffff;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: #1a424a;
  border-color: #1a424a;
  color: #ffffff;
}


/* CONTRASTE DOS CAMPOS — tema claro */
.stApp [data-testid="stWidgetLabel"],
.stApp [data-testid="stWidgetLabel"] *,
.stApp [data-testid="stWidgetLabel"] p,
.stApp .stTextInput label, .stApp .stNumberInput label,
.stApp .stTextArea label, .stApp .stSelectbox label,
.stApp .stDateInput label, .stApp .stMultiSelect label,
.stApp [data-testid="stForm"] label {
  color: #263840 !important;
  opacity: 1 !important;
  font-weight: 600 !important;
}
.stApp [data-baseweb="input"],
.stApp [data-baseweb="textarea"],
.stApp [data-baseweb="select"] > div {
  background: #ffffff !important;
  border-color: #cbd6dc !important;
}
.stApp input, .stApp textarea,
.stApp [data-baseweb="select"] *,
.stApp [data-baseweb="input"] *,
.stApp [data-baseweb="textarea"] * {
  color: #17242a !important;
  -webkit-text-fill-color: #17242a !important;
  opacity: 1 !important;
}
.stApp input, .stApp textarea { background-color: #ffffff !important; }
.stApp input::placeholder, .stApp textarea::placeholder {
  color: #788790 !important;
  -webkit-text-fill-color: #788790 !important;
  opacity: 1 !important;
}
.stApp [data-testid="stHelp"],
.stApp [data-testid="stCaptionContainer"] { color: #647680 !important; }


    /* Legibilidade dos seletores Adicionar/Editar em tema claro e escuro */
    div[data-testid="stRadio"] label,
    div[data-testid="stRadio"] label p,
    div[data-testid="stRadio"] [role="radiogroup"] label,
    div[data-testid="stRadio"] [role="radiogroup"] label p {
        color: #263238 !important;
        opacity: 1 !important;
    }
    div[data-testid="stRadio"] [role="radiogroup"] label {
        background: rgba(8, 121, 138, 0.07) !important;
        border: 1px solid rgba(8, 121, 138, 0.20) !important;
        border-radius: 8px !important;
        padding: 6px 10px !important;
    }
    div[data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {
        background: rgba(8, 121, 138, 0.16) !important;
        border-color: rgba(8, 121, 138, 0.48) !important;
    }

/* ===== Painel inteligente ===== */
.rl-ring{--p:0;--c:#2e9e6b;width:150px;height:150px;border-radius:50%;margin:.2rem auto;position:relative;
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  background:conic-gradient(var(--c) calc(var(--p)*1%),#dde6ea 0)}
.rl-ring::before{content:"";position:absolute;inset:12px;border-radius:50%;background:#fff}
.rl-ring span,.rl-ring small{position:relative;line-height:1.1}
.rl-ring span{font-size:2.5rem;font-weight:700;color:var(--c)}
.rl-ring small{font-size:.68rem;letter-spacing:.08em;text-transform:uppercase;color:#647987}
.rl-act{display:flex;justify-content:space-between;align-items:center;gap:1rem;background:#fff;
  border:1px solid var(--rl-border);border-left:4px solid var(--rl-teal);border-radius:7px;
  padding:.6rem .9rem;margin-bottom:.45rem;color:#17242a}
.rl-act small{display:block;color:#647987;font-size:.78rem;margin-top:.1rem}
.rl-chip{display:inline-block;border-radius:999px;padding:.12rem .6rem;font-size:.7rem;font-weight:650;
  margin:.25rem .25rem 0 0;white-space:nowrap}
.rl-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(235px,1fr));gap:.7rem}
.rl-card{background:#fff;border:1px solid var(--rl-border);border-radius:9px;padding:.8rem .95rem;color:#17242a}
.rl-card-h{display:flex;justify-content:space-between;font-size:1.05rem}
.rl-card small{color:#647987;font-size:.76rem}
.rl-bar{height:6px;background:#e3eaed;border-radius:99px;margin:.5rem 0 .2rem;overflow:hidden}
.rl-bar i{display:block;height:100%;border-radius:99px}
</style>
""", unsafe_allow_html=True)


# Menu lateral: largura compacta e botões somente com ícones quando recolhido.
_sidebar_w = "76px" if st.session_state.sidebar_collapsed else "220px"
st.markdown(
    f"""
    <style>
    section[data-testid="stSidebar"] {{
        width: {_sidebar_w} !important;
        min-width: {_sidebar_w} !important;
        max-width: {_sidebar_w} !important;
        transition: width .18s ease;
    }}
    /* Oculta apenas a seta nativa do Streamlit no cabeçalho da barra lateral.
       O botão personalizado da RL Connect continua funcionando normalmente. */
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"],
    section[data-testid="stSidebar"] button[aria-label="Collapse sidebar"],
    section[data-testid="stSidebar"] button[aria-label="Recolher barra lateral"] {{
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"],
    section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"],
    section[data-testid="stSidebar"] .stButton > button {{
        background: rgba(255, 255, 255, 0.055) !important;
        border: 1px solid rgba(190, 220, 225, 0.18) !important;
        color: #dcebed !important;
    }}
    section[data-testid="stSidebar"] .stButton > button {{
        display: flex !important;
        align-items: center !important;
        gap: .65rem !important;
        min-height: 2.55rem !important;
        padding: .45rem .7rem !important;
        background: transparent !important;
        border: 1px solid transparent !important;
        color: #dcebed !important;
        box-shadow: none !important;
    }}
    section[data-testid="stSidebar"] .stButton > button:hover {{
        background: rgba(39, 133, 145, 0.20) !important;
        border-color: rgba(111, 205, 215, 0.22) !important;
        color: #ffffff !important;
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"],
    section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {{
        background: rgba(8, 121, 138, 0.24) !important;
        border: 1px solid rgba(111, 205, 215, 0.30) !important;
        color: #ffffff !important;
        box-shadow: inset 3px 0 0 #36a7b5 !important;
    }}
    {"section[data-testid='stSidebar'] [data-testid='stMarkdownContainer'], section[data-testid='stSidebar'] [data-testid='stImage'] {display:none !important;}" if st.session_state.sidebar_collapsed else ""}
    {"section[data-testid='stSidebar'] .stButton > button {justify-content:center !important; padding:.45rem 0 !important;} section[data-testid='stSidebar'] .stButton > button p {display:none !important;}" if st.session_state.sidebar_collapsed else ""}
    </style>
    """,
    unsafe_allow_html=True,
)


def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def save_uploaded_file(uploaded, subfolder, prefix):
    """Salva anexos em diretório do aplicativo e retorna o caminho relativo."""
    if uploaded is None:
        return None
    folder = APP_DIR / "documentos" / subfolder
    folder.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", uploaded.name)
    target = folder / f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{safe_name}"
    target.write_bytes(uploaded.getbuffer())
    return str(target.relative_to(APP_DIR))


def licensing_deadline(plate, year):
    """Calendário oficial RJ 2026. Para outros exercícios, não presume prazo."""
    if int(year) != 2026:
        return None
    final = str(plate).strip()[-1:]
    if not final.isdigit():
        return None
    if final in "012":
        return date(2026, 7, 31)
    if final in "345":
        return date(2026, 8, 31)
    if final in "6789":
        return date(2026, 9, 30)
    return None


def licensing_status(row, plate, year):
    if row is None:
        due = licensing_deadline(plate, year)
        return "Pendente" if due is None or date.today() <= due else "Prazo vencido"
    complete = bool(row["ipva_paid"] and row["grt_paid"] and row["crlv_issued"])
    if complete:
        return "Licenciado / em dia"
    due = licensing_deadline(plate, year)
    return "Pendente" if due is None or date.today() <= due else "Prazo vencido"


def init_db():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS vehicles(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plate TEXT UNIQUE NOT NULL, model TEXT, year INTEGER,
            driver TEXT, current_km INTEGER NOT NULL DEFAULT 0,
            notes TEXT, active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS catalog(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL, unit TEXT NOT NULL,
            interval_value INTEGER, recommendation TEXT
        );
        CREATE TABLE IF NOT EXISTS maintenance(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER NOT NULL, service TEXT NOT NULL,
            performed_on TEXT, km INTEGER,
            cost REAL DEFAULT 0, provider TEXT, notes TEXT,
            next_due_date TEXT, next_due_km INTEGER,
            source_sheet TEXT, source_row INTEGER,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id)
        );
        CREATE TABLE IF NOT EXISTS suppliers(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT, address TEXT, phone TEXT, service TEXT, service_done TEXT
        );
        CREATE TABLE IF NOT EXISTS vehicle_licensing(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            ipva_paid INTEGER NOT NULL DEFAULT 0,
            ipva_paid_on TEXT,
            ipva_amount REAL DEFAULT 0,
            grt_paid INTEGER NOT NULL DEFAULT 0,
            grt_paid_on TEXT,
            grt_amount REAL DEFAULT 0,
            grt_file TEXT,
            crlv_issued INTEGER NOT NULL DEFAULT 0,
            crlv_issue_date TEXT,
            crlv_file TEXT,
            notes TEXT,
            UNIQUE(vehicle_id, year),
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id)
        );
        CREATE TABLE IF NOT EXISTS traffic_fines(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER NOT NULL,
            infraction_date TEXT NOT NULL,
            infraction_time TEXT,
            infraction_type TEXT NOT NULL,
            driver_name TEXT,
            amount REAL NOT NULL DEFAULT 0,
            location TEXT,
            status TEXT NOT NULL DEFAULT 'Em aberto',
            due_date TEXT,
            paid_on TEXT,
            notice_number TEXT,
            notes TEXT,
            proof_file TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id)
        );
        CREATE INDEX IF NOT EXISTS idx_licensing_vehicle_year ON vehicle_licensing(vehicle_id, year);
        CREATE INDEX IF NOT EXISTS idx_fines_vehicle_date ON traffic_fines(vehicle_id, infraction_date);
        """)
        # Migração segura para bancos já existentes: preserva todas as multas cadastradas.
        fine_columns = [row["name"] for row in c.execute("PRAGMA table_info(traffic_fines)").fetchall()]
        if "driver_name" not in fine_columns:
            c.execute("ALTER TABLE traffic_fines ADD COLUMN driver_name TEXT")
        # Importa o catálogo e o veículo-base da planilha original somente na primeira execução.
        if c.execute("SELECT COUNT(*) FROM catalog").fetchone()[0] == 0 and SOURCE_XLSX.exists():
            try:
                import openpyxl
                book = openpyxl.load_workbook(SOURCE_XLSX, data_only=False)
                sh = book["Principais Manutenções"]
                for row in sh.iter_rows(min_row=2, values_only=True):
                    _, name, unit, interval_value, recommendation = (list(row) + [None]*5)[:5]
                    if name and str(name).strip():
                        iv = interval_value if isinstance(interval_value, (int, float)) else None
                        c.execute("INSERT OR IGNORE INTO catalog(name,unit,interval_value,recommendation) VALUES(?,?,?,?)",
                                  (str(name).strip(), str(unit or "indeterminado"), iv, str(recommendation or "")))
            except Exception as e:
                st.warning(f"Não foi possível importar o catálogo automaticamente: {e}")
        # Importa o veículo-base mesmo se o catálogo já tiver sido carregado em uma execução anterior.
        if c.execute("SELECT COUNT(*) FROM vehicles").fetchone()[0] == 0 and SOURCE_XLSX.exists():
            try:
                import openpyxl
                book = openpyxl.load_workbook(SOURCE_XLSX, data_only=True)
                shv = book["Manutenções"]
                plate = shv["E6"].value
                km = shv["I6"].value
                if plate:
                    c.execute("INSERT OR IGNORE INTO vehicles(plate,current_km) VALUES(?,?)",
                              (str(plate).strip().upper(), int(km or 0)))
            except Exception as e:
                st.warning(f"Não foi possível importar o veículo-base automaticamente: {e}")

def df(sql, params=()):
    with conn() as c:
        return pd.read_sql_query(sql, c, params=params)

def money(v):
    return f"R$ {float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ───────────────────────── Inteligência da frota ─────────────────────────
STATUS_COLOR = {"Vencida": "#c8453b", "Próxima": "#e08a00", "Em dia": "#08798a"}
SEV_STYLE = {0: ("#c8453b", "Crítico"), 1: ("#e08a00", "Atenção"), 2: ("#08798a", "Acompanhar")}
UP_COLS = ["Placa", "Serviço", "Critério", "Falta", "Previsão", "Situação"]
_RANK = {"Vencida": 0, "Próxima": 1, "Em dia": 2}


def esc(value):
    """Escapa para HTML e neutraliza '$' (o Markdown trataria pares de $ como fórmula)."""
    return html.escape(str(value)).replace("$", "&#36;")


def fmt_km(n):
    return f"{int(n):,}".replace(",", ".")


def fmt_date(d):
    return d.strftime("%d/%m/%Y")


def score_style(score):
    if score >= 80:
        return "#2e9e6b", "Saudável"
    return ("#e08a00", "Atenção") if score >= 55 else ("#c8453b", "Crítica")


def usage_rate(vh):
    """Ritmo médio de uso (km/dia) pelo histórico: exige 2+ registros com km e 30+ dias entre eles."""
    h = vh.dropna(subset=["_d"]).copy()
    h["_k"] = pd.to_numeric(h["km"], errors="coerce")
    h = h[h["_k"] > 0]
    if len(h) < 2:
        return None
    days = (h["_d"].max() - h["_d"].min()).days
    km = h["_k"].max() - h["_k"].min()
    return km / days if days >= 30 and km > 0 else None


def compute_upcoming(vehicles, history, catalog, today):
    """Vencimentos por veículo/serviço numa única passada (o original fazia uma consulta por combinação).
    Serviços por km ganham uma data estimada a partir do ritmo de uso do próprio veículo."""
    cat = catalog[catalog["interval_value"].notna() & catalog["unit"].isin(["km", "dias"])]
    if vehicles.empty or history.empty or cat.empty:
        return pd.DataFrame(), {}
    h = history.copy()
    h["_d"] = pd.to_datetime(h["performed_on"], errors="coerce")
    h = h.sort_values(["_d", "id"], ascending=False, na_position="last")
    rates = {int(v.id): usage_rate(h[h["vehicle_id"] == v.id]) for v in vehicles.itertuples()}
    known = [r for r in rates.values() if r]
    fallback = float(pd.Series(known).median()) if known else 50.0
    ts_today, rows = pd.Timestamp(today), []
    for v in vehicles.itertuples():
        last = h[h["vehicle_id"] == v.id].drop_duplicates("service", keep="first").set_index("service")
        rate = rates[int(v.id)]
        for c in cat.itertuples():
            if c.name not in last.index:
                continue
            rec, interval = last.loc[c.name], int(c.interval_value)
            if c.unit == "km":
                base_km = 0 if pd.isna(rec["km"]) else int(rec["km"])
                due_km = base_km + interval
                remaining = due_km - int(v.current_km or 0)
                status = "Vencida" if remaining <= 0 else ("Próxima" if remaining <= 1000 else "Em dia")
                if remaining <= 0:
                    due = ts_today
                elif rate:
                    due = ts_today + pd.Timedelta(days=math.ceil(remaining / rate))
                else:
                    due = pd.NaT
                rows.append({
                    "Placa": v.plate, "Serviço": c.name, "Critério": f"{fmt_km(due_km)} km",
                    "Falta": f"{fmt_km(remaining)} km" if remaining > 0 else f"{fmt_km(-remaining)} km de atraso",
                    "Previsão": "Agora" if remaining <= 0 else (f"≈ {fmt_date(due)}" if rate else "—"),
                    "Situação": status, "_due": due, "_urgency": remaining / (rate or fallback)})
            else:
                if pd.isna(rec["_d"]):
                    continue
                due = rec["_d"].normalize() + pd.Timedelta(days=interval)
                remaining = (due - ts_today).days
                status = "Vencida" if remaining < 0 else ("Próxima" if remaining <= 30 else "Em dia")
                rows.append({
                    "Placa": v.plate, "Serviço": c.name, "Critério": fmt_date(due),
                    "Falta": f"{remaining} dias" if remaining >= 0 else f"{-remaining} dias de atraso",
                    "Previsão": fmt_date(due), "Situação": status, "_due": due, "_urgency": remaining})
    if not rows:
        return pd.DataFrame(), rates
    out = pd.DataFrame(rows)
    out["_rank"] = out["Situação"].map(_RANK)
    return out.sort_values(["_rank", "_urgency"]).drop(columns="_rank").reset_index(drop=True), rates


def vehicle_health(plate, up_raw, lic_status, open_fines):
    """Nota 0-100 e pendências. Regra: -20 por manutenção vencida (máx. -50), -8 por próxima (máx. -16),
    -20 licenciamento vencido / -8 pendente, -8 por multa em aberto (máx. -16)."""
    score, issues = 100, []
    u = up_raw[up_raw["Placa"] == plate] if not up_raw.empty else up_raw
    venc = int((u["Situação"] == "Vencida").sum()) if not u.empty else 0
    prox = int((u["Situação"] == "Próxima").sum()) if not u.empty else 0
    if venc:
        score -= min(venc * 20, 50)
        issues.append(("bad", f"{venc} manutenção(ões) vencida(s)"))
    if prox:
        score -= min(prox * 8, 16)
        issues.append(("warn", f"{prox} próxima(s)"))
    lic = lic_status.get(plate)
    if lic == "Prazo vencido":
        score -= 20
        issues.append(("bad", "Licenciamento vencido"))
    elif lic == "Pendente":
        score -= 8
        issues.append(("warn", "Licenciamento pendente"))
    nf = open_fines.get(plate, 0)
    if nf:
        score -= min(nf * 8, 16)
        issues.append(("warn", f"{nf} multa(s) em aberto"))
    return max(score, 0), issues


def build_actions(up_raw, license_alerts, fine_alerts, today):
    """Central de ação: manutenção, licenciamento e multas na mesma fila, por gravidade e prazo."""
    acts = []
    for r in ([] if up_raw.empty else up_raw.to_dict("records")):
        if r["Situação"] == "Em dia":
            continue
        late = r["Situação"] == "Vencida"
        det = f"{'Vencida' if late else 'Próxima'} · {r['Falta']}"
        if not late and r["Previsão"] not in ("—", "Agora"):
            det += f" · prevista {r['Previsão']}"
        acts.append({"sev": 0 if late else 1, "tipo": "Manutenção", "placa": r["Placa"],
                     "titulo": r["Serviço"], "detalhe": det, "ord": r["_urgency"]})
    for a in license_alerts:
        due = licensing_deadline(a["Placa"], a["Exercício"])
        late = a["Situação"] == "Prazo vencido"
        if due:
            det = f"Prazo {'vencido em' if late else 'até'} {fmt_date(due)}"
        else:
            det = "Confira IPVA, GRT e CRLV-e"
        acts.append({"sev": 0 if late else 1, "tipo": "Licenciamento", "placa": a["Placa"],
                     "titulo": f"Licenciamento {a['Exercício']}", "detalhe": det,
                     "ord": (due - today).days if due else 0})
    for f in ([] if fine_alerts.empty else fine_alerts.to_dict("records")):
        due = pd.to_datetime(f["due_date"], errors="coerce")
        left = (due.date() - today).days if pd.notna(due) else None
        sev = 2 if left is None or left > 7 else (0 if left < 0 else 1)
        if left is None:
            when = "sem vencimento informado"
        elif left < 0:
            when = f"venceu há {-left} dias"
        else:
            when = "vence hoje" if left == 0 else f"vence em {left} dias"
        acts.append({"sev": sev, "tipo": "Multa", "placa": f["plate"], "titulo": f["infraction_type"],
                     "detalhe": f"{money(f['amount'])} · {when}", "ord": 9999 if left is None else left})
    acts.sort(key=lambda a: (a["sev"], a["ord"]))
    return acts


def action_card(a):
    color, label = SEV_STYLE[a["sev"]]
    return (f'<div class="rl-act" style="border-left-color:{color}"><div><b>{esc(a["placa"])}</b> · {esc(a["titulo"])}'
            f'<small>{esc(a["tipo"])} — {esc(a["detalhe"])}</small></div>'
            f'<span class="rl-chip" style="background:{color}1a;color:{color}">{label}</span></div>')


def forecast_budget(up_raw, history, today, horizon=90):
    """Gasto previsto: serviços vencidos, próximos ou previstos até o horizonte × mediana histórica do custo."""
    empty = pd.DataFrame(columns=["Serviço", "Veículos", "Custo típico", "Estimativa"])
    if up_raw.empty or history.empty:
        return empty, []
    limit = pd.Timestamp(today) + pd.Timedelta(days=horizon)
    due = up_raw[(up_raw["Situação"] != "Em dia") | (up_raw["_due"].notna() & (up_raw["_due"] <= limit))]
    if due.empty:
        return empty, []
    cost = pd.to_numeric(history["cost"], errors="coerce")
    typical = history.assign(_c=cost)[cost > 0].groupby("service")["_c"].median()
    grouped = due.groupby("Serviço").size().rename("Veículos").reset_index()
    grouped["Custo típico"] = grouped["Serviço"].map(typical)
    missing = grouped.loc[grouped["Custo típico"].isna(), "Serviço"].tolist()
    grouped = grouped.dropna(subset=["Custo típico"])
    grouped["Estimativa"] = grouped["Veículos"] * grouped["Custo típico"]
    return grouped.sort_values("Estimativa", ascending=False).reset_index(drop=True), missing


def cost_per_km(history):
    """R$/km por placa e da frota: custo total ÷ km rodados entre o primeiro e o último registro."""
    if history.empty:
        return {}, None
    h = history.copy()
    h["km"] = pd.to_numeric(h["km"], errors="coerce")
    h["cost"] = pd.to_numeric(h["cost"], errors="coerce").fillna(0)
    per, tot_cost, tot_km = {}, 0.0, 0.0
    for plate, g in h.groupby("plate"):
        k = g["km"].dropna()
        k = k[k > 0]
        span = float(k.max() - k.min()) if len(k) >= 2 else 0.0
        if span > 0:
            per[plate] = float(g["cost"].sum()) / span
            tot_cost += float(g["cost"].sum())
            tot_km += span
    return per, (tot_cost / tot_km if tot_km else None)


def cost_outliers(history, factor=1.6, min_n=3):
    """Lançamentos bem acima da mediana do mesmo serviço (precisa de 3+ registros do serviço)."""
    cols = ["Placa", "Serviço", "Data", "Custo", "Mediana do serviço", "Vezes acima"]
    if history.empty:
        return pd.DataFrame(columns=cols)
    h = history.copy()
    h["cost"] = pd.to_numeric(h["cost"], errors="coerce")
    h = h[h["cost"] > 0].copy()
    g = h.groupby("service")["cost"]
    h["_med"], h["_n"] = g.transform("median"), g.transform("count")
    o = h[(h["_n"] >= min_n) & (h["cost"] >= h["_med"] * factor)].copy()
    o["Vezes acima"] = (o["cost"] / o["_med"]).round(1)
    o = o.rename(columns={"plate": "Placa", "service": "Serviço", "performed_on": "Data",
                          "cost": "Custo", "_med": "Mediana do serviço"})
    return o[cols].sort_values("Vezes acima", ascending=False)


def timeline_chart(up_raw, today, months=12):
    """Linha do tempo (Altair): cada vencimento como um ponto, ligado ao 'hoje'."""
    if up_raw.empty:
        return None
    limit = pd.Timestamp(today) + pd.DateOffset(months=months)
    d = up_raw[up_raw["_due"].notna() & (up_raw["_due"] <= limit)].copy()
    if d.empty:
        return None
    d["Item"] = d["Placa"] + " · " + d["Serviço"]
    d["Data"] = d["_due"].dt.strftime("%Y-%m-%d")
    d["Hoje"] = today.isoformat()
    d["Quando"] = d["Previsão"]
    y = alt.Y("Item:N", sort=d.sort_values("_due")["Item"].tolist(), title=None)
    color = alt.Color("Situação:N", legend=alt.Legend(orient="top", title=None),
                      scale=alt.Scale(domain=list(STATUS_COLOR), range=list(STATUS_COLOR.values())))
    stem = alt.Chart(d).mark_rule(strokeWidth=3, opacity=0.35).encode(x="Hoje:T", x2="Data:T", y=y, color=color)
    dot = alt.Chart(d).mark_circle(size=170, opacity=1).encode(
        x=alt.X("Data:T", title=None), y=y, color=color,
        tooltip=["Placa", "Serviço", "Situação", "Quando", "Falta"])
    now = alt.Chart(pd.DataFrame({"Data": [today.isoformat()]})).mark_rule(
        color="#17242a", strokeDash=[4, 4]).encode(x="Data:T")
    return (now + stem + dot).properties(height=max(140, 30 * len(d)))


def whatsapp_summary(score, actions, budget_total, today):
    lines = [f"*RL Connect · Resumo da frota — {fmt_date(today)}*", f"Saúde da frota: *{score}/100*", ""]
    for sev, title in ((0, "🔴 Críticos"), (1, "🟠 Atenção")):
        items = [a for a in actions if a["sev"] == sev]
        if items:
            lines.append(f"*{title} ({len(items)})*")
            lines += [f"• {a['placa']} — {a['titulo']}: {a['detalhe']}" for a in items[:8]]
            if len(items) > 8:
                lines.append(f"… e mais {len(items) - 8}")
            lines.append("")
    if len(lines) <= 3:
        lines.append("✅ Nada crítico ou próximo do vencimento.")
    if budget_total:
        lines.append(f"💰 Gastos previstos (90 dias): {money(budget_total)}")
    return "\n".join(lines).strip()


def build_excel(sheets):
    """Gera um .xlsx com uma aba por DataFrame, colunas ajustadas e cabeçalho congelado."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, frame in sheets.items():
            name = name[:31]
            frame.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for i, col in enumerate(frame.columns, 1):
                width = max([len(str(col))] + [len(str(x)) for x in frame[col].head(200)]) + 2
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width, 48)
            ws.freeze_panes = "A2"
    return buf.getvalue()


def make_backup_zip():
    """Backup completo: banco SQLite (cópia consistente) + pasta de documentos anexados."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        c = conn()
        try:
            z.writestr(DB.name, c.serialize() if hasattr(c, "serialize") else DB.read_bytes())
        finally:
            c.close()
        docs = APP_DIR / "documentos"
        if docs.exists():
            for p in docs.rglob("*"):
                if p.is_file():
                    z.write(p, p.relative_to(APP_DIR).as_posix())
    return buf.getvalue()


init_db()

st.markdown("## Gestão de Manutenção de Frota")
st.caption("RL CONNECT  /  CONTROLE PREVENTIVO, CORRETIVO E CUSTOS")

vehicles = df("SELECT * FROM vehicles WHERE active=1 ORDER BY plate")
history = df("""SELECT m.*, v.plate FROM maintenance m JOIN vehicles v ON v.id=m.vehicle_id
                ORDER BY m.performed_on DESC, m.id DESC""")

# Vencimentos calculados em uma passada; serviços por km ganham data estimada pelo ritmo de uso.
today = date.today()
up_raw, usage_rates = compute_upcoming(vehicles, history, df("SELECT * FROM catalog"), today)
up = up_raw[UP_COLS].copy() if not up_raw.empty else pd.DataFrame()

p1,p2,p3,p4 = st.columns(4)
p1.metric("Veículos ativos", len(vehicles))
p2.metric("Manutenções registradas", len(history))
p3.metric("Alertas vencidos", int((up["Situação"]=="Vencida").sum()) if not up.empty else 0)
p4.metric("Custo registrado", money(history["cost"].sum()) if not history.empty else "R$ 0,00")

# Alertas documentais e de multas: situação baseada nos dados informados no sistema.
current_year = date.today().year
license_alerts = []
fine_alerts = df("SELECT f.id, v.plate, f.infraction_type, f.amount, f.due_date FROM traffic_fines f JOIN vehicles v ON v.id=f.vehicle_id WHERE f.status='Em aberto' ORDER BY f.due_date, f.infraction_date")
lic_year = df("SELECT * FROM vehicle_licensing WHERE year=?", (current_year,)).set_index("vehicle_id")
license_status_by_plate = {}
for vehicle_row in vehicles.itertuples():
    record = lic_year.loc[vehicle_row.id] if vehicle_row.id in lic_year.index else None
    status = licensing_status(record, vehicle_row.plate, current_year)
    license_status_by_plate[vehicle_row.plate] = status
    if status != "Licenciado / em dia":
        license_alerts.append({"Placa": vehicle_row.plate, "Exercício": current_year, "Situação": status})


# Navegação lateral por módulos — estrutura inspirada no painel de referência.
PAGES = {
    "Painel": "Painel",
    "Registrar manutenção": "Registrar manutenção",
    "Agenda de manutenção": "Agenda de manutenção",
    "Editar manutenção": "Editar manutenção",
    "Licenciamento e documentos": "Licenciamento e documentos",
    "Multas": "Multas",
    "Veículos": "Veículos",
    "Ficha do veículo": "Ficha do veículo",
    "Relatórios": "Relatórios",
    "Configurações": "Configurações",
    "Fornecedores": "Fornecedores",
}
if "current_page" not in st.session_state:
    st.session_state.current_page = "Painel"

with st.sidebar:
    # Controle de expandir/recolher sempre visível e com alto contraste.
    st.button(
        "»" if st.session_state.sidebar_collapsed else "«",
        key="toggle_sidebar_collapsed",
        help="Expandir menu" if st.session_state.sidebar_collapsed else "Recolher menu",
        type="secondary",
        use_container_width=True,
        on_click=lambda: st.session_state.update(
            sidebar_collapsed=not st.session_state.sidebar_collapsed
        ),
    )

    if not st.session_state.sidebar_collapsed:
        # A logo é opcional: o aplicativo continua funcionando caso o PNG não esteja na pasta.
        logo = APP_DIR / "logo_rl_connect.png"
        if logo.is_file():
            st.image(str(logo), width=220)
        else:
            st.markdown("### R&L Connect")
        st.caption("GESTÃO DE FROTA")
        st.markdown("---")

    def nav_button(label, page, icon_name):
        # Recolhido: apenas ícone; o nome aparece exclusivamente no tooltip ao passar o cursor.
        # Expandido: ícone e nome lado a lado.
        st.button(
            label if not st.session_state.sidebar_collapsed else "",
            key=f"nav_{page}",
            icon=icon_name,
            help=label,
            use_container_width=True,
            type="primary" if st.session_state.current_page == page else "secondary",
            on_click=lambda p=page: st.session_state.update(current_page=p),
        )

    if not st.session_state.sidebar_collapsed:
        st.markdown("#### VISÃO GERAL")
    nav_button("Painel", "Painel", ":material/dashboard:")

    if not st.session_state.sidebar_collapsed:
        st.markdown("#### OPERACIONAL")
    nav_button("Registrar manutenção", "Registrar manutenção", ":material/build:")
    nav_button("Agenda de manutenção", "Agenda de manutenção", ":material/calendar_month:")
    nav_button("Editar manutenção", "Editar manutenção", ":material/edit_note:")
    nav_button("Licenciamento e documentos", "Licenciamento e documentos", ":material/description:")
    nav_button("Multas", "Multas", ":material/confirmation_number:")

    if not st.session_state.sidebar_collapsed:
        st.markdown("#### FROTA")
    nav_button("Veículos", "Veículos", ":material/directions_car:")
    nav_button("Ficha do veículo", "Ficha do veículo", ":material/assignment:")

    if not st.session_state.sidebar_collapsed:
        st.markdown("#### FINANCEIRO")
    nav_button("Relatórios e custos", "Relatórios", ":material/monitoring:")

    if not st.session_state.sidebar_collapsed:
        st.markdown("#### SISTEMA")
    nav_button("Configurações", "Configurações", ":material/settings:")
    nav_button("Fornecedores", "Fornecedores", ":material/business:")

page = st.session_state.current_page


if page == "Painel" and vehicles.empty:
    st.subheader("Visão geral da frota")
    st.info("Cadastre o primeiro veículo na aba Veículos para ativar o painel inteligente.")

if page == "Painel" and not vehicles.empty:
    open_by_plate = {} if fine_alerts.empty else fine_alerts.groupby("plate").size().to_dict()
    health = {v.plate: vehicle_health(v.plate, up_raw, license_status_by_plate, open_by_plate)
              for v in vehicles.itertuples()}
    fleet_score = round(sum(sc for sc, _ in health.values()) / len(health))
    actions = build_actions(up_raw, license_alerts, fine_alerts, today)
    budget, budget_missing = forecast_budget(up_raw, history, today)
    budget_total = float(budget["Estimativa"].sum()) if not budget.empty else 0.0
    cpk_by_plate, fleet_cpk = cost_per_km(history)
    n_crit = sum(a["sev"] == 0 for a in actions)
    n_warn = sum(a["sev"] == 1 for a in actions)
    score_color, score_label = score_style(fleet_score)

    st.subheader("Visão geral da frota")
    ring, kpis = st.columns([1, 3])
    ring.markdown(
        f'<div class="rl-ring" style="--p:{fleet_score};--c:{score_color}"><span>{fleet_score}</span>'
        f'<small>{score_label}</small></div>', unsafe_allow_html=True)
    with kpis:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Itens críticos", n_crit)
        k2.metric("Exigem atenção", n_warn)
        k3.metric("Gastos previstos · 90 dias", money(budget_total))
        k4.metric("Custo médio por km", money(fleet_cpk) if fleet_cpk else "—")
        st.caption("Nota da frota = média da nota de cada veículo (100 menos: manutenção vencida −20 cada, "
                   "próxima −8, licenciamento vencido −20 ou pendente −8, multa em aberto −8).")

    st.markdown("#### O que exige ação agora")
    if not actions:
        st.success("Tudo em dia: nenhuma manutenção, documento ou multa pedindo ação.")
    else:
        st.markdown("".join(action_card(a) for a in actions[:6]), unsafe_allow_html=True)
        if len(actions) > 6:
            with st.expander(f"Ver todos os {len(actions)} itens"):
                st.markdown("".join(action_card(a) for a in actions[6:]), unsafe_allow_html=True)

    st.markdown("#### Saúde por veículo")
    cards = []
    for v in vehicles.itertuples():
        sc, issues = health[v.plate]
        c_, _ = score_style(sc)
        chips = "".join(
            f'<span class="rl-chip" style="background:{"#c8453b" if k == "bad" else "#e08a00"}1a;'
            f'color:{"#c8453b" if k == "bad" else "#e08a00"}">{esc(t)}</span>' for k, t in issues
        ) or '<span class="rl-chip" style="background:#2e9e6b1a;color:#2e9e6b">Tudo em dia</span>'
        rate = usage_rates.get(int(v.id))
        meta = f"{v.model or 'Modelo não informado'} · {fmt_km(v.current_km)} km" + (f" · ~{rate:.0f} km/dia" if rate else "")
        cards.append(
            f'<div class="rl-card"><div class="rl-card-h"><b>{esc(v.plate)}</b>'
            f'<span style="color:{c_};font-weight:700">{sc}</span></div><small>{esc(meta)}</small>'
            f'<div class="rl-bar"><i style="width:{sc}%;background:{c_}"></i></div><div>{chips}</div></div>')
    st.markdown(f'<div class="rl-grid">{"".join(cards)}</div>', unsafe_allow_html=True)

    st.markdown("#### Gastos previstos para os próximos 90 dias")
    if budget.empty:
        st.caption("Sem serviços vencidos ou previstos no período com histórico de custo para estimar.")
    else:
        shown = budget.copy()
        shown["Custo típico"] = shown["Custo típico"].map(money)
        shown["Estimativa"] = shown["Estimativa"].map(money)
        st.dataframe(shown, use_container_width=True, hide_index=True)
        note = "Estimativa = mediana histórica de custo do serviço × veículos com vencimento no período."
        if budget_missing:
            note += " Sem custo histórico (fora da soma): " + ", ".join(budget_missing) + "."
        st.caption(note)

    outliers = cost_outliers(history)
    with st.expander(f"Eficiência de custos · R$/km e lançamentos fora do padrão ({len(outliers)})"):
        if cpk_by_plate:
            st.dataframe(pd.DataFrame(
                [{"Placa": p, "R$ por km": money(val)} for p, val in sorted(cpk_by_plate.items(), key=lambda x: -x[1])]),
                use_container_width=True, hide_index=True)
            st.caption("Custo total ÷ km rodados entre o primeiro e o último registro de cada veículo.")
        else:
            st.caption("Precisa de 2+ manutenções com quilometragem por veículo para calcular R$/km.")
        if outliers.empty:
            st.caption("Nenhum lançamento muito acima da mediana do próprio serviço (precisa de 3+ registros do serviço).")
        else:
            show_out = outliers.copy()
            show_out["Custo"] = show_out["Custo"].map(money)
            show_out["Mediana do serviço"] = show_out["Mediana do serviço"].map(money)
            st.dataframe(show_out, use_container_width=True, hide_index=True)

    st.markdown("#### Compartilhar")
    actions_df = pd.DataFrame(
        [{"Gravidade": SEV_STYLE[a["sev"]][1], "Tipo": a["tipo"], "Placa": a["placa"],
          "Item": a["titulo"], "Detalhe": a["detalhe"]} for a in actions],
        columns=["Gravidade", "Tipo", "Placa", "Item", "Detalhe"])
    health_df = pd.DataFrame(
        [{"Placa": p, "Nota": sc, "Pendências": "; ".join(t for _, t in iss) or "Tudo em dia"}
         for p, (sc, iss) in health.items()])
    hist_df = history[["plate", "service", "performed_on", "km", "cost", "provider", "notes"]].rename(columns={
        "plate": "Placa", "service": "Serviço", "performed_on": "Data", "km": "Km",
        "cost": "Custo (R$)", "provider": "Oficina", "notes": "Observações"})
    st.download_button(
        "Baixar relatório executivo (Excel)",
        build_excel({"Ações": actions_df, "Agenda": up, "Saúde da frota": health_df,
                     "Gastos previstos": budget, "Histórico": hist_df}),
        file_name=f"relatorio_frota_{today:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("Mensagem pronta para WhatsApp"):
        st.code(whatsapp_summary(fleet_score, actions, budget_total, today), language=None)

    with st.expander("Detalhe de documentação e multas"):
        a1, a2 = st.columns(2)
        a1.metric("Veículos com licenciamento pendente", len(license_alerts))
        a2.metric("Multas em aberto", len(fine_alerts))
        if license_alerts:
            st.dataframe(pd.DataFrame(license_alerts), use_container_width=True, hide_index=True)
        if not fine_alerts.empty:
            st.dataframe(fine_alerts.rename(columns={"id": "ID", "plate": "Placa", "infraction_type": "Infração",
                                                     "amount": "Valor (R$)", "due_date": "Vencimento"}),
                         use_container_width=True, hide_index=True)
    st.markdown("#### Últimas manutenções")
    if history.empty:
        st.info("Nenhuma manutenção registrada.")
    else:
        st.dataframe(history.head(10), use_container_width=True, hide_index=True)

if page == "Veículos":
    # Inclui inativos: uma placa desativada continua reservada no banco.
    all_vehicles = df("SELECT * FROM vehicles ORDER BY active DESC, plate")
    st.subheader("Cadastrar veículo")
    with st.form("vehicle_form", clear_on_submit=True):
        a,b,c = st.columns(3)
        plate = a.text_input("Placa *").upper().strip()
        model = b.text_input("Modelo")
        year = c.number_input("Ano", min_value=1900, max_value=2100, value=2020)
        d,e = st.columns(2)
        driver = d.text_input("Motorista responsável")
        km = e.number_input("Quilometragem atual", min_value=0, step=100)
        notes = st.text_area("Observações")
        submitted = st.form_submit_button("Salvar veículo")
        if submitted:
            if not plate:
                st.error("Informe a placa.")
            else:
                try:
                    with conn() as db:
                        db.execute("INSERT INTO vehicles(plate,model,year,driver,current_km,notes) VALUES(?,?,?,?,?,?)",
                                   (plate,model,int(year),driver,int(km),notes))
                    st.success("Veículo cadastrado.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Já existe um veículo com essa placa.")
    st.subheader("Frota cadastrada")
    display_vehicles = all_vehicles.rename(columns={
        "plate": "Placa", "model": "Modelo", "year": "Ano",
        "driver": "Motorista", "current_km": "Quilometragem",
        "notes": "Observações", "active": "Ativo"
    })
    st.dataframe(display_vehicles, use_container_width=True, hide_index=True)
    inactive_count = int((all_vehicles["active"] == 0).sum()) if not all_vehicles.empty else 0
    if inactive_count:
        st.warning(f"Existem {inactive_count} veículo(s) inativo(s). Eles continuam reservando a placa até serem excluídos permanentemente.")

    if not all_vehicles.empty:
        st.markdown("### Editar dados do veículo")
        vehicle_options = {
            f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id)
            for _, r in all_vehicles.iterrows()
        }
        selected_vehicle = st.selectbox(
            "Selecione o veículo que deseja corrigir",
            list(vehicle_options.keys()),
            key="edit_vehicle_selector"
        )
        selected_id = vehicle_options[selected_vehicle]
        vrow = all_vehicles[all_vehicles.id == selected_id].iloc[0]

        with st.form("edit_vehicle_form"):
            st.caption("Altere os campos necessários e salve. Os registros de manutenção existentes serão preservados.")
            a, b, c = st.columns(3)
            edit_plate = a.text_input("Placa *", value=str(vrow.plate or "")).strip().upper()
            edit_model = b.text_input("Modelo / versão", value=str(vrow.model or ""))
            edit_year = c.number_input(
                "Ano de fabricação", min_value=1900, max_value=2100,
                value=int(vrow.year) if pd.notna(vrow.year) else 2020
            )
            d, e = st.columns(2)
            edit_driver = d.text_input("Motorista responsável", value=str(vrow.driver or ""))
            edit_km = e.number_input(
                "Quilometragem atual", min_value=0,
                value=int(vrow.current_km or 0), step=100
            )
            edit_notes = st.text_area("Observações", value=str(vrow.notes or ""))
            save_edit = st.form_submit_button("Salvar alterações")

            if save_edit:
                if not edit_plate:
                    st.error("A placa é obrigatória.")
                else:
                    try:
                        with conn() as db:
                            db.execute(
                                """UPDATE vehicles
                                   SET plate=?, model=?, year=?, driver=?, current_km=?, notes=?
                                   WHERE id=?""",
                                (
                                    edit_plate, edit_model.strip() or None, int(edit_year),
                                    edit_driver.strip() or None, int(edit_km),
                                    edit_notes.strip() or None, selected_id
                                )
                            )
                        st.success("Dados do veículo atualizados com sucesso.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Já existe outro veículo cadastrado com essa placa. Confira a informação.")

        st.markdown("### Atualizar quilometragem")
        km_options = {
            f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id)
            for _, r in all_vehicles.iterrows()
        }
        km_selected = st.selectbox(
            "Veículo", list(km_options.keys()), key="km_vehicle_selector"
        )
        km_id = km_options[km_selected]
        km_row = all_vehicles[all_vehicles.id == km_id].iloc[0]
        with st.form("km_form"):
            newkm = st.number_input(
                "Nova quilometragem", min_value=0,
                value=int(km_row.current_km or 0), step=100
            )
            if st.form_submit_button("Atualizar odômetro"):
                if newkm < int(km_row.current_km or 0):
                    st.error("A nova quilometragem não pode ser menor que a atual. Para corrigir um erro de digitação, use Editar dados do veículo.")
                else:
                    with conn() as db:
                        db.execute("UPDATE vehicles SET current_km=? WHERE id=?", (int(newkm), km_id))
                    st.success("Quilometragem atualizada.")
                    st.rerun()

        st.markdown("### Desativar veículo")
        st.caption("Use esta opção para retirar um cadastro incorreto da frota ativa sem apagar o histórico.")
        deactivate_confirm = st.checkbox(
            f"Confirmo que desejo desativar {vrow.plate}",
            key=f"deactivate_confirm_{selected_id}"
        )
        if st.button("Desativar veículo selecionado", disabled=not deactivate_confirm):
            with conn() as db:
                db.execute("UPDATE vehicles SET active=0 WHERE id=?", (selected_id,))
            st.success("Veículo desativado. O histórico foi preservado.")
            st.rerun()

        st.markdown("### Excluir veículo cadastrado incorretamente")
        st.warning(
            "A exclusão é permanente. Os registros de manutenção vinculados a este veículo "
            "também serão excluídos, assim como seus registros de licenciamento e multas. "
            "Use somente se o cadastro foi feito por engano."
        )
        with conn() as db:
            linked_maintenances = db.execute(
                "SELECT COUNT(*) FROM maintenance WHERE vehicle_id=?", (selected_id,)
            ).fetchone()[0]

        if linked_maintenances:
            st.caption(
                f"Este veículo possui {linked_maintenances} manutenção(ões) registrada(s). "
                "Elas também serão removidas se você confirmar a exclusão."
            )
        delete_confirm = st.checkbox(
            f"Confirmo a exclusão permanente do veículo {vrow.plate}"
            + (f" e de suas {linked_maintenances} manutenção(ões)" if linked_maintenances else ""),
            key=f"delete_vehicle_confirm_{selected_id}"
        )
        if st.button(
            "Excluir veículo permanentemente",
            type="secondary",
            disabled=not delete_confirm,
            key=f"delete_vehicle_button_{selected_id}"
        ):
            with conn() as db:
                db.execute("DELETE FROM maintenance WHERE vehicle_id=?", (selected_id,))
                db.execute("DELETE FROM vehicle_licensing WHERE vehicle_id=?", (selected_id,))
                db.execute("DELETE FROM traffic_fines WHERE vehicle_id=?", (selected_id,))
                db.execute("DELETE FROM vehicles WHERE id=?", (selected_id,))
            st.success(
                f"Veículo {vrow.plate} excluído. Agora você poderá cadastrá-lo novamente."
            )
            st.rerun()

if page == "Registrar manutenção":
    st.subheader("Lançar serviço realizado")
    if vehicles.empty:
        st.warning("Cadastre um veículo antes de registrar uma manutenção.")
    else:
        catalog = df("SELECT * FROM catalog ORDER BY name")
        with st.form("maintenance_form", clear_on_submit=True):
            vmap = {f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id) for _,r in vehicles.iterrows()}
            vehicle_label = st.selectbox("Veículo", list(vmap))
            service_options = catalog["name"].tolist() if not catalog.empty else []
            service = st.selectbox("Serviço realizado", service_options) if service_options else st.text_input("Descrição do serviço")
            a,b,c = st.columns(3)
            performed = a.date_input("Data da manutenção", value=date.today())
            km_done = b.number_input("Quilometragem na manutenção", min_value=0, step=100)
            cost = c.number_input("Custo (R$)", min_value=0.0, step=50.0, format="%.2f")
            provider = st.text_input("Oficina / fornecedor")
            notes = st.text_area("Observações / peças utilizadas")
            catrow = catalog[catalog.name==service] if not catalog.empty else pd.DataFrame()
            due_km, due_date = None, None
            if not catrow.empty and pd.notna(catrow.iloc[0]["interval_value"]):
                interval = int(catrow.iloc[0]["interval_value"])
                if catrow.iloc[0]["unit"] == "km":
                    due_km = int(km_done) + interval
                elif catrow.iloc[0]["unit"] == "dias":
                    due_date = (performed + timedelta(days=interval)).isoformat()
            st.caption("A próxima previsão será calculada automaticamente conforme o intervalo cadastrado no catálogo.")
            if st.form_submit_button("Registrar manutenção"):
                if not service:
                    st.error("Informe o serviço.")
                else:
                    vid = vmap[vehicle_label]
                    with conn() as db:
                        db.execute("""INSERT INTO maintenance(vehicle_id,service,performed_on,km,cost,provider,notes,next_due_date,next_due_km)
                                      VALUES(?,?,?,?,?,?,?,?,?)""",
                                   (vid,service,performed.isoformat(),int(km_done),float(cost),provider,notes,due_date,due_km))
                        db.execute("UPDATE vehicles SET current_km=MAX(current_km,?) WHERE id=?", (int(km_done),vid))
                    st.success("Manutenção registrada e histórico atualizado.")
                    st.rerun()

if page == "Editar manutenção":
    st.subheader("Editar manutenção registrada")
    st.caption("Selecione um lançamento para corrigir veículo, serviço, data, quilometragem, custo, oficina ou observações. As alterações são gravadas no histórico original.")

    edit_rows = df("""SELECT m.id, m.vehicle_id, m.service, m.performed_on, m.km,
                           m.cost, m.provider, m.notes, v.plate, v.model
                    FROM maintenance m
                    LEFT JOIN vehicles v ON v.id = m.vehicle_id
                    ORDER BY CASE WHEN m.performed_on IS NULL OR m.performed_on='' THEN 1 ELSE 0 END,
                             m.performed_on DESC, m.id DESC""")
    if edit_rows.empty:
        st.info("Não existem manutenções cadastradas para editar.")
    else:
        vehicle_labels = {
            int(r.id): f'{r.plate} — {r.model or "Modelo não informado"}'
            for _, r in df("SELECT id, plate, model FROM vehicles ORDER BY plate").iterrows()
        }
        edit_rows["_label"] = edit_rows.apply(
            lambda r: f'ID {int(r.id)} | {r.plate or "Veículo não localizado"} | '
                      f'{r.service} | {r.performed_on or "Data não informada"} | {money(r.cost)}', axis=1
        )
        selected_label = st.selectbox(
            "Selecione a manutenção pelo ID, veículo e serviço",
            options=edit_rows["_label"].tolist(), key="maintenance_edit_selector"
        )
        selected = edit_rows.loc[edit_rows["_label"] == selected_label].iloc[0]
        selected_id = int(selected["id"])
        st.info(f"Editando o registro ID {selected_id}. O ID será mantido para preservar o histórico.")

        all_vehicle_ids = list(vehicle_labels.keys())
        current_vehicle_id = int(selected["vehicle_id"])
        if current_vehicle_id not in all_vehicle_ids:
            all_vehicle_ids.append(current_vehicle_id)
            vehicle_labels[current_vehicle_id] = f'Veículo ID {current_vehicle_id} — não localizado'
        current_date = pd.to_datetime(selected["performed_on"], errors="coerce")
        default_date = current_date.date() if pd.notna(current_date) else date.today()
        current_km = int(selected["km"]) if pd.notna(selected["km"]) else 0
        current_cost = float(selected["cost"]) if pd.notna(selected["cost"]) else 0.0

        with st.form(f"edit_maintenance_form_{selected_id}"):
            vehicle_id = st.selectbox(
                "Veículo", options=all_vehicle_ids,
                index=all_vehicle_ids.index(current_vehicle_id),
                format_func=lambda value: vehicle_labels.get(value, str(value))
            )
            service = st.text_input("Tipo de serviço", value=str(selected["service"] or ""))
            a, b, c = st.columns(3)
            performed = a.date_input("Data da manutenção", value=default_date)
            km_done = b.number_input("Quilometragem (km)", min_value=0, step=100, value=max(0, current_km))
            cost = c.number_input("Custo (R$)", min_value=0.0, step=50.0, value=max(0.0, current_cost), format="%.2f")
            provider = st.text_input(
                "Oficina / fornecedor",
                value="" if pd.isna(selected["provider"]) else str(selected["provider"] or "")
            )
            notes = st.text_area(
                "Observações / peças utilizadas",
                value="" if pd.isna(selected["notes"]) else str(selected["notes"] or "")
            )
            submitted = st.form_submit_button("Salvar alterações", type="primary")

        if submitted:
            if not service.strip():
                st.error("Informe o tipo de serviço antes de salvar.")
            else:
                catalog_row = df("SELECT interval_value, unit FROM catalog WHERE name=? LIMIT 1", (service.strip(),))
                next_due_km, next_due_date = None, None
                if not catalog_row.empty and pd.notna(catalog_row.iloc[0]["interval_value"]):
                    interval = int(catalog_row.iloc[0]["interval_value"])
                    if catalog_row.iloc[0]["unit"] == "km":
                        next_due_km = int(km_done) + interval
                    elif catalog_row.iloc[0]["unit"] == "dias":
                        next_due_date = (performed + timedelta(days=interval)).isoformat()
                with conn() as db:
                    db.execute("""UPDATE maintenance
                                  SET vehicle_id=?, service=?, performed_on=?, km=?, cost=?,
                                      provider=?, notes=?, next_due_date=?, next_due_km=?
                                  WHERE id=?""",
                               (int(vehicle_id), service.strip(), performed.isoformat(),
                                int(km_done), float(cost), provider.strip(), notes.strip(),
                                next_due_date, next_due_km, selected_id))
                    # Atualiza a quilometragem exibida sem reduzi-la abaixo do valor já registrado no veículo.
                    db.execute("""UPDATE vehicles
                                  SET current_km = MAX(current_km, COALESCE(
                                      (SELECT MAX(km) FROM maintenance WHERE vehicle_id=?), 0))
                                  WHERE id=?""", (int(vehicle_id), int(vehicle_id)))
                st.success(f"Manutenção ID {selected_id} atualizada com sucesso.")
                st.rerun()

if page == "Licenciamento e documentos":
    st.subheader("Licenciamento anual e documentos")
    st.caption("Controle interno por exercício. O sistema não consulta automaticamente os sistemas do Detran/SEFAZ; atualize os pagamentos e a emissão do CRLV-e após confirmar nos canais oficiais.")
    if vehicles.empty:
        st.warning("Cadastre um veículo antes de registrar o licenciamento.")
    else:
        vehicle_map = {f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id) for _, r in vehicles.iterrows()}
        selected_vehicle = st.selectbox("Veículo", list(vehicle_map), key="lic_vehicle")
        vid = vehicle_map[selected_vehicle]
        plate = selected_vehicle.split(" — ")[0]
        year = st.selectbox("Exercício do licenciamento", options=list(range(date.today().year, date.today().year-5, -1)), key="lic_year")
        existing_df = df("SELECT * FROM vehicle_licensing WHERE vehicle_id=? AND year=?", (vid, year))
        existing = existing_df.iloc[0] if not existing_df.empty else None
        current_ipva = bool(existing["ipva_paid"]) if existing is not None else False
        current_grt = bool(existing["grt_paid"]) if existing is not None else False
        current_crlv = bool(existing["crlv_issued"]) if existing is not None else False
        st.markdown("#### Situação do exercício")
        due = licensing_deadline(plate, year)
        if due:
            st.info(f"Prazo oficial Detran-RJ para final de placa {plate[-1]}: {due.strftime('%d/%m/%Y')}.")
        elif year == 2026:
            st.info("Consulte o calendário oficial para confirmar o prazo desta placa.")
        else:
            st.caption("Prazo deste exercício não foi pré-cadastrado. Consulte o calendário oficial vigente.")
        with st.form(f"licensing_form_{vid}_{year}"):
            c1, c2 = st.columns(2)
            ipva_paid = c1.checkbox("IPVA totalmente pago", value=current_ipva)
            grt_paid = c2.checkbox("GRT paga", value=current_grt)
            c3, c4 = st.columns(2)
            ipva_amount = c3.number_input("Valor do IPVA (R$)", min_value=0.0, value=float(existing["ipva_amount"] or 0) if existing is not None else 0.0, step=50.0, format="%.2f")
            grt_amount = c4.number_input("Valor da GRT (R$)", min_value=0.0, value=float(existing["grt_amount"] or 0) if existing is not None else 0.0, step=10.0, format="%.2f")
            c5, c6 = st.columns(2)
            ipva_paid_on = c5.date_input("Data do pagamento do IPVA", value=pd.to_datetime(existing["ipva_paid_on"]).date() if existing is not None and existing["ipva_paid_on"] else date.today(), disabled=not ipva_paid)
            grt_paid_on = c6.date_input("Data do pagamento da GRT", value=pd.to_datetime(existing["grt_paid_on"]).date() if existing is not None and existing["grt_paid_on"] else date.today(), disabled=not grt_paid)
            crlv_issued = st.checkbox("CRLV-e emitido / disponível", value=current_crlv)
            crlv_issue_date = st.date_input("Data de emissão do CRLV-e", value=pd.to_datetime(existing["crlv_issue_date"]).date() if existing is not None and existing["crlv_issue_date"] else date.today(), disabled=not crlv_issued)
            grt_upload = st.file_uploader("Anexar comprovante / GRT (PDF ou imagem)", type=["pdf", "png", "jpg", "jpeg"], key=f"grt_file_{vid}_{year}")
            crlv_upload = st.file_uploader("Anexar CRLV-e (PDF ou imagem)", type=["pdf", "png", "jpg", "jpeg"], key=f"crlv_file_{vid}_{year}")
            notes = st.text_area("Observações", value=str(existing["notes"] or "") if existing is not None else "")
            save_license = st.form_submit_button("Salvar licenciamento", type="primary")
        if save_license:
            old_grt = existing["grt_file"] if existing is not None else None
            old_crlv = existing["crlv_file"] if existing is not None else None
            grt_path = save_uploaded_file(grt_upload, f"veiculo_{vid}", f"GRT_{year}") or old_grt
            crlv_path = save_uploaded_file(crlv_upload, f"veiculo_{vid}", f"CRLV_{year}") or old_crlv
            with conn() as db:
                db.execute("""INSERT INTO vehicle_licensing(vehicle_id,year,ipva_paid,ipva_paid_on,ipva_amount,grt_paid,grt_paid_on,grt_amount,grt_file,crlv_issued,crlv_issue_date,crlv_file,notes)
                             VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                             ON CONFLICT(vehicle_id,year) DO UPDATE SET
                             ipva_paid=excluded.ipva_paid, ipva_paid_on=excluded.ipva_paid_on, ipva_amount=excluded.ipva_amount,
                             grt_paid=excluded.grt_paid, grt_paid_on=excluded.grt_paid_on, grt_amount=excluded.grt_amount,
                             grt_file=excluded.grt_file, crlv_issued=excluded.crlv_issued, crlv_issue_date=excluded.crlv_issue_date,
                             crlv_file=excluded.crlv_file, notes=excluded.notes""",
                          (vid, year, int(ipva_paid), ipva_paid_on.isoformat() if ipva_paid else None, ipva_amount,
                           int(grt_paid), grt_paid_on.isoformat() if grt_paid else None, grt_amount, grt_path,
                           int(crlv_issued), crlv_issue_date.isoformat() if crlv_issued else None, crlv_path, notes.strip()))
            st.success("Dados do licenciamento salvos.")
            st.rerun()
        latest = df("SELECT * FROM vehicle_licensing WHERE vehicle_id=? AND year=?", (vid, year))
        if not latest.empty:
            status = licensing_status(latest.iloc[0], plate, year)
            if status == "Licenciado / em dia":
                st.success("Licenciamento registrado como em dia: IPVA, GRT e CRLV-e confirmados.")
            elif status == "Prazo vencido":
                st.error("Prazo de licenciamento vencido: verifique os débitos e a emissão do CRLV-e.")
            else:
                st.warning("Licenciamento pendente: confira IPVA, GRT e emissão do CRLV-e.")
            for label, path in [("Comprovante / GRT", latest.iloc[0]["grt_file"]), ("CRLV-e", latest.iloc[0]["crlv_file"])]:
                if path and (APP_DIR / path).exists():
                    st.download_button(f"Baixar {label}", (APP_DIR / path).read_bytes(), file_name=Path(path).name, key=f"download_{label}_{vid}_{year}")
        st.markdown("#### Resumo da frota")
        licensing_rows = []
        for _, vr in vehicles.iterrows():
            lr = df("SELECT * FROM vehicle_licensing WHERE vehicle_id=? AND year=?", (int(vr.id), year))
            rec = lr.iloc[0] if not lr.empty else None
            licensing_rows.append({"Placa": vr.plate, "Exercício": year, "IPVA": "Pago" if rec is not None and rec["ipva_paid"] else "Pendente", "GRT": "Paga" if rec is not None and rec["grt_paid"] else "Pendente", "CRLV-e": "Emitido" if rec is not None and rec["crlv_issued"] else "Não emitido", "Situação": licensing_status(rec, vr.plate, year)})
        st.dataframe(pd.DataFrame(licensing_rows), use_container_width=True, hide_index=True)


if page == "Multas":
    st.subheader("Controle de multas de trânsito")
    st.caption("Registre notificações, acompanhe vencimentos e marque a quitação após confirmar o pagamento.")
    if vehicles.empty:
        st.warning("Cadastre um veículo antes de lançar uma multa.")
    else:
        vmap = {f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id) for _, r in vehicles.iterrows()}
        with st.expander("＋ Registrar nova multa", expanded=True):
            with st.form("fine_form", clear_on_submit=True):
                vehicle_label = st.selectbox("Veículo", list(vmap), key="fine_vehicle")
                driver_name = st.text_input("Nome do motorista", placeholder="Nome completo do motorista responsável")
                c1, c2, c3 = st.columns(3)
                infraction_date = c1.date_input("Data da infração", value=date.today())
                infraction_time = c2.time_input("Hora da infração", value=datetime.now().time().replace(second=0, microsecond=0))
                amount = c3.number_input("Valor da multa (R$)", min_value=0.0, step=25.0, format="%.2f")
                infraction_type = st.text_input("Tipo de infração", placeholder="Ex.: excesso de velocidade, estacionamento irregular")
                location = st.text_input("Local da infração", placeholder="Rua, avenida, município ou rodovia")
                d1, d2 = st.columns(2)
                notice_number = d1.text_input("Número do auto / notificação (opcional)")
                due_date = d2.date_input("Data de vencimento", value=date.today())
                notes = st.text_area("Observações")
                proof = st.file_uploader("Anexar notificação ou comprovante (PDF ou imagem)", type=["pdf", "png", "jpg", "jpeg"], key="fine_proof")
                submit_fine = st.form_submit_button("Registrar multa", type="primary")
            if submit_fine:
                if not infraction_type.strip():
                    st.error("Informe o tipo de infração.")
                else:
                    proof_path = save_uploaded_file(proof, "multas", "multa")
                    with conn() as db:
                        db.execute("""INSERT INTO traffic_fines(vehicle_id,infraction_date,infraction_time,infraction_type,driver_name,amount,location,status,due_date,notice_number,notes,proof_file)
                                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                                   (vmap[vehicle_label], infraction_date.isoformat(), infraction_time.strftime("%H:%M"), infraction_type.strip(), driver_name.strip(), float(amount), location.strip(), "Em aberto", due_date.isoformat(), notice_number.strip(), notes.strip(), proof_path))
                    st.success("Multa registrada como em aberto.")
                    st.rerun()

        st.markdown("#### Situação das multas")
        fines = df("""SELECT f.*, v.plate FROM traffic_fines f JOIN vehicles v ON v.id=f.vehicle_id
                     ORDER BY CASE WHEN f.status='Em aberto' THEN 0 ELSE 1 END, f.due_date, f.infraction_date DESC, f.id DESC""")
        if fines.empty:
            st.info("Nenhuma multa cadastrada.")
        else:
            open_fines = fines[fines["status"] == "Em aberto"]
            paid_fines = fines[fines["status"] == "Paga"]
            m1, m2, m3 = st.columns(3)
            m1.metric("Multas em aberto", len(open_fines))
            m2.metric("Total em aberto", money(open_fines["amount"].sum()))
            m3.metric("Total pago", money(paid_fines["amount"].sum()))
            view = fines.rename(columns={"id":"ID", "plate":"Placa", "infraction_date":"Data", "infraction_time":"Hora", "infraction_type":"Tipo de infração", "driver_name":"Nome do motorista", "amount":"Valor (R$)", "location":"Local", "status":"Situação", "due_date":"Vencimento", "paid_on":"Data do pagamento", "notice_number":"Auto / notificação", "notes":"Observações"})
            st.dataframe(view[[c for c in ["ID","Placa","Nome do motorista","Data","Hora","Tipo de infração","Valor (R$)","Local","Situação","Vencimento","Data do pagamento","Auto / notificação","Observações"] if c in view.columns]], use_container_width=True, hide_index=True)
            st.markdown("#### Atualizar situação de uma multa")
            fine_options = {f'ID {int(r.id)} | {r.plate} | {r.infraction_type} | {money(r.amount)} | {r.status}': int(r.id) for _, r in fines.iterrows()}
            chosen_fine = st.selectbox("Selecione a multa", list(fine_options), key="fine_status_select")
            fine_id = fine_options[chosen_fine]
            fine_row = fines.loc[fines["id"] == fine_id].iloc[0]
            with st.form(f"fine_update_{fine_id}"):
                status = st.selectbox("Situação", ["Em aberto", "Paga"], index=0 if fine_row["status"] == "Em aberto" else 1)
                paid_on = st.date_input("Data do pagamento", value=pd.to_datetime(fine_row["paid_on"]).date() if pd.notna(fine_row["paid_on"]) and fine_row["paid_on"] else date.today(), disabled=(status != "Paga"))
                save_status = st.form_submit_button("Atualizar multa")
            if save_status:
                with conn() as db:
                    db.execute("UPDATE traffic_fines SET status=?, paid_on=? WHERE id=?", (status, paid_on.isoformat() if status == "Paga" else None, fine_id))
                st.success("Situação da multa atualizada.")
                st.rerun()
            if fine_row["proof_file"] and (APP_DIR / fine_row["proof_file"]).exists():
                st.download_button("Baixar notificação / comprovante", (APP_DIR / fine_row["proof_file"]).read_bytes(), file_name=Path(fine_row["proof_file"]).name)


if page == "Agenda de manutenção":
    st.subheader("Próximas manutenções")
    st.caption("Alertas por quilometragem: vencida quando atingiu/passou o limite; próxima quando faltam até 1.000 km. "
               "Por tempo: próxima quando faltam até 30 dias. A data dos serviços por km (≈) usa o ritmo médio de "
               "uso de cada veículo, calculado pelo próprio histórico.")
    if up.empty:
        st.info("Sem previsões calculadas. Registre os serviços realizados para iniciar o acompanhamento.")
    else:
        chart = timeline_chart(up_raw, today)
        if chart is not None:
            st.markdown("##### Linha do tempo · próximos 12 meses")
            st.altair_chart(chart, use_container_width=True)
        plate_of = dict(zip(vehicles["id"].astype(int), vehicles["plate"]))
        rate_txt = " · ".join(f"{plate_of[i]}: ~{r:.0f} km/dia" for i, r in usage_rates.items() if r and i in plate_of)
        if rate_txt:
            st.caption("Ritmo de uso estimado — " + rate_txt)
        filt = st.multiselect("Filtrar situação", ["Vencida", "Próxima", "Em dia"], default=["Vencida", "Próxima", "Em dia"])
        st.dataframe(up[up["Situação"].isin(filt)], use_container_width=True, hide_index=True)

if page == "Relatórios":
    st.subheader("Histórico e custos")
    if history.empty:
        st.info("Ainda não existem despesas registradas.")
    else:
        # Preparação dos campos para filtros sem alterar os dados originais do banco.
        report_data = history.copy()
        report_data["_data"] = pd.to_datetime(report_data["performed_on"], errors="coerce")
        report_data["_mes"] = report_data["_data"].dt.strftime("%Y-%m")
        report_data["_id"] = pd.to_numeric(report_data["id"], errors="coerce")
        report_data["cost"] = pd.to_numeric(report_data["cost"], errors="coerce").fillna(0)

        st.markdown("#### Filtros do histórico")
        f1, f2, f3, f4 = st.columns(4)
        id_filter = f1.number_input("ID da manutenção", min_value=0, step=1, value=0,
                                    help="Digite 0 para exibir todos os IDs.")
        services = sorted([str(x) for x in report_data["service"].dropna().unique() if str(x).strip()])
        service_filter = f2.multiselect("Tipo de serviço", options=services,
                                        placeholder="Todos os serviços")
        months = sorted(report_data["_mes"].dropna().unique().tolist(), reverse=True)
        month_labels = {m: pd.to_datetime(m + "-01").strftime("%m/%Y") for m in months}
        month_filter = f3.multiselect("Mês", options=months,
                                      format_func=lambda m: month_labels.get(m, m),
                                      placeholder="Todos os meses")

        # Filtro adicional de veículo: por padrão, mantém o relatório geral.
        vehicle_options = sorted(
            [str(v) for v in vehicles["plate"].dropna().unique() if str(v).strip()]
        )
        vehicle_filter = f4.selectbox(
            "Veículo",
            options=["Todos os veículos"] + vehicle_options,
            index=0,
            help="Selecione uma placa para consultar somente os gastos daquele veículo."
        )

        use_period = st.checkbox("Filtrar por intervalo de datas", value=False)
        date_filter = None
        valid_dates = report_data["_data"].dropna()
        if use_period and not valid_dates.empty:
            d1, d2 = st.columns(2)
            start_date = d1.date_input("Data inicial", value=valid_dates.min().date(), key="report_start_date")
            end_date = d2.date_input("Data final", value=valid_dates.max().date(), key="report_end_date")
            if start_date > end_date:
                st.error("A data inicial não pode ser posterior à data final.")
            else:
                date_filter = (pd.Timestamp(start_date), pd.Timestamp(end_date) + pd.Timedelta(days=1))

        filtered = report_data.copy()
        if id_filter > 0:
            filtered = filtered[filtered["_id"] == int(id_filter)]
        if service_filter:
            filtered = filtered[filtered["service"].isin(service_filter)]
        if vehicle_filter != "Todos os veículos":
            filtered = filtered[filtered["plate"] == vehicle_filter]
        if month_filter:
            filtered = filtered[filtered["_mes"].isin(month_filter)]
        if date_filter is not None:
            filtered = filtered[filtered["_data"].notna() &
                                (filtered["_data"] >= date_filter[0]) &
                                (filtered["_data"] < date_filter[1])]

        def format_brl(value):
            return f"R$ {float(value or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        total_filtered = float(filtered["cost"].sum()) if not filtered.empty else 0.0
        m1, m2, m3 = st.columns(3)
        m1.metric("Total gasto no filtro", format_brl(total_filtered))
        m2.metric("Manutenções encontradas", f"{len(filtered)}")
        m3.metric("Tipos de serviço", f"{filtered['service'].nunique() if not filtered.empty else 0}")

        # Painel visual: evolução mensal, comparação por veículo e distribuição por serviço.
        st.markdown("#### Indicadores gráficos de manutenção")
        if filtered.empty:
            st.info("Nenhum registro corresponde aos filtros selecionados.")
        else:
            monthly_rows = filtered.dropna(subset=["_mes"]).copy()
            # Ranking de gastos: inclui todos os veículos ativos cadastrados,
            # inclusive os que ainda não possuem manutenção no período/filtros.
            ranking_vehicles = vehicles[["plate"]].drop_duplicates()
            if vehicle_filter != "Todos os veículos":
                ranking_vehicles = ranking_vehicles[
                    ranking_vehicles["plate"] == vehicle_filter
                ]

            vehicle_costs = (
                ranking_vehicles
                .merge(
                    filtered.groupby("plate", as_index=False)
                    .agg(Custo=("cost", "sum"), Manutenções=("id", "count")),
                    on="plate", how="left"
                )
            )
            vehicle_costs[["Custo", "Manutenções"]] = (
                vehicle_costs[["Custo", "Manutenções"]].fillna(0)
            )
            vehicle_costs["Manutenções"] = vehicle_costs["Manutenções"].astype(int)
            vehicle_costs = vehicle_costs.sort_values(
                ["Custo", "plate"], ascending=[False, True]
            ).reset_index(drop=True)
            service_costs = (
                filtered.groupby("service", as_index=False)
                .agg(Custo=("cost", "sum"), Quantidade=("id", "count"))
                .sort_values("Custo", ascending=False)
            )

            # Destaque do veículo com maior custo no conjunto filtrado.
            if not vehicle_costs.empty:
                highest = vehicle_costs.iloc[0]
                st.markdown("##### Veículo com maior gasto em manutenção")
                k1, k2, k3 = st.columns(3)
                k1.metric("Carro com maior custo", str(highest["plate"]))
                k2.metric("Total gasto", format_brl(highest["Custo"]))
                k3.metric("Manutenções registradas", f'{int(highest["Manutenções"])}')

            chart_left, chart_right = st.columns(2)
            with chart_left:
                st.markdown("##### Gasto por veículo — do maior para o menor")
                if not vehicle_costs.empty:
                    vehicle_chart = alt.Chart(vehicle_costs).mark_bar().encode(
                        x=alt.X("Custo:Q", title="Total gasto (R$)", axis=alt.Axis(format=",.2f")),
                        y=alt.Y("plate:N", sort="-x", title="Veículo"),
                        tooltip=[
                            alt.Tooltip("plate:N", title="Placa"),
                            alt.Tooltip("Custo:Q", title="Total gasto", format=",.2f"),
                            alt.Tooltip("Manutenções:Q", title="Manutenções")
                        ]
                    ).properties(
                        height=max(260, min(720, 34 * len(vehicle_costs)))
                    )
                    st.altair_chart(vehicle_chart, use_container_width=True)
                    st.caption(
                        "Ranking de todos os veículos ativos cadastrados. "
                        "Veículos sem manutenção nos filtros selecionados aparecem com R$ 0,00."
                    )
                else:
                    st.caption("Não há veículos ativos cadastrados.")

            with chart_right:
                st.markdown("##### Custo por tipo de serviço")
                if not service_costs.empty:
                    service_chart = service_costs.head(10).set_index("service")[["Custo"]]
                    st.bar_chart(service_chart, use_container_width=True)
                    st.caption("Exibindo os 10 serviços com maior custo.")
                else:
                    st.caption("Não há custos por serviço para os filtros selecionados.")

            st.markdown("##### Evolução dos gastos por mês")
            if not monthly_rows.empty:
                by_month_chart = (
                    monthly_rows.groupby("_mes", as_index=False)["cost"].sum()
                    .rename(columns={"_mes": "Mês", "cost": "Custo"})
                    .sort_values("Mês")
                )
                by_month_chart["Mês"] = by_month_chart["Mês"].map(
                    lambda m: month_labels.get(m, m)
                )
                st.bar_chart(by_month_chart.set_index("Mês")[["Custo"]], use_container_width=True)
            else:
                st.caption("Os registros filtrados não possuem data informada para agrupamento mensal.")

            # Tabelas-resumo permanecem disponíveis para consulta exata dos valores.
            with st.expander("Ver valores detalhados por serviço e por mês"):
                st.markdown("**Gastos por tipo de serviço**")
                by_service_table = service_costs.copy()
                by_service_table["Custo"] = by_service_table["Custo"].map(format_brl)
                st.dataframe(
                    by_service_table.rename(columns={"service": "Tipo de serviço"}),
                    use_container_width=True, hide_index=True
                )
                st.markdown("**Gastos por mês**")
                if not monthly_rows.empty:
                    by_month_table = (
                        monthly_rows.groupby("_mes", as_index=False)["cost"].sum()
                        .rename(columns={"_mes": "Mês", "cost": "Custo"})
                        .sort_values("Mês")
                    )
                    by_month_table["Mês"] = by_month_table["Mês"].map(
                        lambda m: month_labels.get(m, m)
                    )
                    by_month_table["Custo"] = by_month_table["Custo"].map(format_brl)
                    st.dataframe(by_month_table, use_container_width=True, hide_index=True)

        st.markdown("#### Registros de manutenção")
        display = filtered.drop(columns=["_data", "_mes", "_id"], errors="ignore").copy()
        display = display.rename(columns={
            "id": "ID", "vehicle_id": "ID do veículo", "service": "Serviço",
            "performed_on": "Data da manutenção", "km": "Quilometragem (km)",
            "cost": "Custo (R$)", "provider": "Oficina / fornecedor",
            "notes": "Observações", "plate": "Placa"
        })
        if "Custo (R$)" in display.columns:
            display["Custo (R$)"] = display["Custo (R$)"].map(format_brl)
        st.dataframe(display, use_container_width=True, hide_index=True)

        csv_data = filtered.drop(columns=["_data", "_mes", "_id"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
        st.download_button("Baixar histórico filtrado em CSV", csv_data,
                           file_name="historico_manutencoes_filtrado.csv", mime="text/csv")


if page == "Ficha do veículo":
    st.subheader("Ficha completa do veículo")
    if vehicles.empty:
        st.info("Cadastre um veículo para consultar sua ficha individual.")
    else:
        vchoices = {f'{r.plate} — {r.model or "Modelo não informado"}': int(r.id) for _, r in vehicles.iterrows()}
        chosen = st.selectbox("Selecione o veículo", list(vchoices), key="vehicle_detail_select")
        vid = vchoices[chosen]
        v = vehicles[vehicles.id == vid].iloc[0]
        vh = df("""SELECT service AS Serviço, performed_on AS Data, km AS 'Km na manutenção',
                          cost AS Custo, provider AS Oficina, notes AS Observações
                   FROM maintenance WHERE vehicle_id=? ORDER BY (performed_on IS NULL), performed_on DESC, id DESC""", (vid,))
        total_cost = float(vh["Custo"].fillna(0).sum()) if not vh.empty else 0.0
        count = len(vh)
        last_date = vh.iloc[0]["Data"] if not vh.empty else "Sem registros"
        a,b,c,d = st.columns(4)
        a.metric("Quilometragem atual", f'{int(v.current_km):,} km'.replace(",", "."))
        b.metric("Manutenções realizadas", count)
        c.metric("Custo acumulado registrado", money(total_cost))
        d.metric("Última manutenção", pd.to_datetime(last_date, errors="coerce").strftime("%d/%m/%Y") if last_date != "Sem registros" and pd.notna(pd.to_datetime(last_date, errors="coerce")) else "Não informada")

        _rate = usage_rates.get(vid)
        _cpk = cost_per_km(history)[0].get(v.plate)
        st.caption(
            (f"Ritmo de uso estimado: ~{_rate:.0f} km/dia" if _rate else
             "Ritmo de uso: dados insuficientes (precisa de 2+ manutenções com km, com 30+ dias entre elas)")
            + (f" · Custo: {money(_cpk)} por km" if _cpk else ""))
        st.markdown("#### Dados do veículo")
        st.write(f"**Placa:** {v.plate}  |  **Modelo:** {v.model or 'Não informado'}  |  **Ano:** {int(v.year) if pd.notna(v.year) else 'Não informado'}")
        st.write(f"**Motorista:** {v.driver or 'Não informado'}")
        if v.notes:
            st.write(f"**Observações:** {v.notes}")

        st.markdown("#### Próximas manutenções deste veículo")
        vehicle_up = up[up["Placa"] == v.plate] if not up.empty else pd.DataFrame()
        if vehicle_up.empty:
            st.info("Ainda não há previsões. Registre as manutenções realizadas e mantenha o catálogo de intervalos atualizado.")
        else:
            st.dataframe(vehicle_up.drop(columns=["_sort"], errors="ignore"), use_container_width=True, hide_index=True)

        st.markdown("#### Histórico completo")
        if vh.empty:
            st.info("Nenhuma manutenção cadastrada para este veículo.")
        else:
            st.dataframe(vh, use_container_width=True, hide_index=True)
            st.download_button("Baixar histórico deste veículo (CSV)",
                               vh.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"historico_{v.plate}.csv", mime="text/csv",
                               key=f"download_vehicle_{vid}")
            st.markdown("#### Custos por serviço")
            costs = vh.groupby("Serviço", as_index=False)["Custo"].sum().sort_values("Custo", ascending=False)
            st.dataframe(costs, use_container_width=True, hide_index=True)
            st.markdown("#### Possíveis problemas recorrentes")
            # Heurística transparente: destaca descrições repetidas; não diagnostica defeitos automaticamente.
            text_values = vh["Observações"].fillna("").astype(str).str.strip()
            combined = pd.concat([vh["Serviço"].fillna("").astype(str).str.strip(), text_values], axis=1)
            terms = pd.concat([combined.iloc[:,0], combined.iloc[:,1]]).replace("", pd.NA).dropna()
            repeated = terms.value_counts()
            repeated = repeated[repeated >= 2]
            if repeated.empty:
                st.caption("Ainda não há descrições repetidas suficientes para apontar recorrências. Registre sintomas e observações de forma padronizada.")
            else:
                st.dataframe(repeated.rename_axis("Descrição").reset_index(name="Ocorrências"),
                             use_container_width=True, hide_index=True)


if page == "Fornecedores":
    st.subheader("Fornecedores e contatos")
    st.caption("Cadastre novos fornecedores ou atualize os dados dos contatos já importados.")

    supplier_rows = df("SELECT id, company, address, phone, service, service_done FROM suppliers ORDER BY company")
    supplier_mode = st.radio(
        "Ação", ["Adicionar fornecedor", "Editar fornecedor"],
        horizontal=True, key="supplier_mode"
    )

    selected_supplier = None
    if supplier_mode == "Editar fornecedor":
        if supplier_rows.empty:
            st.info("Não há fornecedores cadastrados para editar.")
        else:
            supplier_options = {
                f'{r["company"] or "Sem nome"} — ID {int(r["id"])}': int(r["id"])
                for _, r in supplier_rows.iterrows()
            }
            chosen_supplier = st.selectbox(
                "Selecione o fornecedor", list(supplier_options),
                key="supplier_edit_select"
            )
            selected_supplier = supplier_rows[
                supplier_rows["id"] == supplier_options[chosen_supplier]
            ].iloc[0]

    if supplier_mode == "Adicionar fornecedor" or selected_supplier is not None:
        with st.form("supplier_editor_form", clear_on_submit=(supplier_mode == "Adicionar fornecedor")):
            company = st.text_input(
                "Empresa", value="" if selected_supplier is None else str(selected_supplier["company"] or "")
            )
            address = st.text_input(
                "Endereço", value="" if selected_supplier is None else str(selected_supplier["address"] or "")
            )
            phone = st.text_input(
                "Telefone", value="" if selected_supplier is None else str(selected_supplier["phone"] or "")
            )
            service = st.text_area(
                "Serviços oferecidos", value="" if selected_supplier is None else str(selected_supplier["service"] or "")
            )
            service_done = st.text_area(
                "Serviço realizado", value="" if selected_supplier is None else str(selected_supplier["service_done"] or "")
            )
            supplier_submit = st.form_submit_button(
                "Salvar fornecedor" if supplier_mode == "Adicionar fornecedor" else "Atualizar fornecedor",
                use_container_width=True
            )
        if supplier_submit:
            if not company.strip():
                st.error("Informe o nome da empresa.")
            else:
                with conn() as c:
                    if supplier_mode == "Adicionar fornecedor":
                        c.execute(
                            "INSERT INTO suppliers(company, address, phone, service, service_done) VALUES (?, ?, ?, ?, ?)",
                            (company.strip(), address.strip(), phone.strip(), service.strip(), service_done.strip())
                        )
                        st.success("Fornecedor cadastrado com sucesso.")
                    else:
                        c.execute(
                            "UPDATE suppliers SET company=?, address=?, phone=?, service=?, service_done=? WHERE id=?",
                            (company.strip(), address.strip(), phone.strip(), service.strip(), service_done.strip(),
                             int(selected_supplier["id"]))
                        )
                        st.success("Fornecedor atualizado com sucesso.")
                st.rerun()

    st.markdown("#### Fornecedores cadastrados")
    suppliers_view = df(
        "SELECT company AS Empresa, address AS Endereço, phone AS Telefone, "
        "service AS Serviços, service_done AS 'Serviço realizado' FROM suppliers ORDER BY company"
    )
    if suppliers_view.empty:
        st.info("Nenhum fornecedor cadastrado.")
    else:
        st.dataframe(suppliers_view, use_container_width=True, hide_index=True)
        st.download_button(
            "Exportar contatos (CSV)", suppliers_view.to_csv(index=False).encode("utf-8-sig"),
            file_name="fornecedores_rl_connect.csv", mime="text/csv"
        )

if page == "Configurações":
    st.subheader("Configurações do sistema")
    st.caption("Parâmetros e informações gerais da aplicação.")
    st.markdown("#### Catálogo de manutenções")
    st.caption("Gerencie os serviços, unidades, intervalos e recomendações usados no controle preventivo.")

    catalog_rows = df(
        "SELECT id, name, unit, interval_value, recommendation FROM catalog ORDER BY name"
    )
    catalog_mode = st.radio(
        "Ação do catálogo", ["Adicionar serviço", "Editar serviço"],
        horizontal=True, key="catalog_mode"
    )

    selected_catalog = None
    if catalog_mode == "Editar serviço":
        if catalog_rows.empty:
            st.info("O catálogo ainda não possui serviços para editar.")
        else:
            catalog_options = {
                f'{r["name"]} — ID {int(r["id"])}': int(r["id"])
                for _, r in catalog_rows.iterrows()
            }
            chosen_catalog = st.selectbox(
                "Selecione o serviço", list(catalog_options),
                key="catalog_edit_select"
            )
            selected_catalog = catalog_rows[
                catalog_rows["id"] == catalog_options[chosen_catalog]
            ].iloc[0]

    if catalog_mode == "Adicionar serviço" or selected_catalog is not None:
        with st.form("catalog_editor_form", clear_on_submit=(catalog_mode == "Adicionar serviço")):
            catalog_name = st.text_input(
                "Nome do serviço", value="" if selected_catalog is None else str(selected_catalog["name"])
            )
            catalog_unit = st.selectbox(
                "Unidade do intervalo", ["km", "dias", "meses", "indeterminado"],
                index=0 if selected_catalog is None else (
                    ["km", "dias", "meses", "indeterminado"].index(selected_catalog["unit"])
                    if selected_catalog["unit"] in ["km", "dias", "meses", "indeterminado"] else 0
                )
            )
            # Exibe o intervalo com separador de milhar (ponto), mantendo o valor
            # numérico puro no banco de dados.
            current_interval = (
                0 if selected_catalog is None or pd.isna(selected_catalog["interval_value"])
                else int(selected_catalog["interval_value"])
            )
            catalog_interval_text = st.text_input(
                "Intervalo",
                value=fmt_km(current_interval),
                help="Informe apenas números. Ex.: 10.000. Use 0 quando não houver intervalo definido."
            )
            catalog_recommendation = st.text_area(
                "Recomendação",
                value="" if selected_catalog is None else str(selected_catalog["recommendation"] or "")
            )
            catalog_submit = st.form_submit_button(
                "Salvar serviço" if catalog_mode == "Adicionar serviço" else "Atualizar serviço",
                use_container_width=True
            )
        if catalog_submit:
            if not catalog_name.strip():
                st.error("Informe o nome do serviço.")
            else:
                # Aceita valores digitados como 10000, 10.000 ou 10 000.
                interval_digits = re.sub(r"[^0-9]", "", catalog_interval_text or "")
                if not interval_digits:
                    st.error("Informe um intervalo numérico válido.")
                    st.stop()
                catalog_interval = int(interval_digits)
                interval_value = None if catalog_interval == 0 or catalog_unit == "indeterminado" else catalog_interval
                try:
                    with conn() as c:
                        if catalog_mode == "Adicionar serviço":
                            c.execute(
                                "INSERT INTO catalog(name, unit, interval_value, recommendation) VALUES (?, ?, ?, ?)",
                                (catalog_name.strip(), catalog_unit, interval_value, catalog_recommendation.strip())
                            )
                            st.success("Serviço adicionado ao catálogo.")
                        else:
                            c.execute(
                                "UPDATE catalog SET name=?, unit=?, interval_value=?, recommendation=? WHERE id=?",
                                (catalog_name.strip(), catalog_unit, interval_value, catalog_recommendation.strip(),
                                 int(selected_catalog["id"]))
                            )
                            st.success("Serviço atualizado.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Já existe um serviço com esse nome no catálogo.")

    st.markdown("#### Serviços cadastrados")
    catalog_view = df(
        "SELECT name AS Serviço, unit AS Unidade, interval_value AS Intervalo, "
        "recommendation AS Recomendação FROM catalog ORDER BY name"
    )
    if not catalog_view.empty:
        catalog_view["Intervalo"] = catalog_view["Intervalo"].apply(
            lambda value: "" if pd.isna(value) else fmt_km(value)
        )
    if catalog_view.empty:
        st.info("O catálogo ainda não possui serviços cadastrados.")
    else:
        st.dataframe(catalog_view, use_container_width=True, hide_index=True)
        st.download_button(
            "Exportar catálogo (CSV)",
            catalog_view.to_csv(index=False).encode("utf-8-sig"),
            file_name="catalogo_manutencoes.csv",
            mime="text/csv",
        )
    st.markdown("#### Banco de dados")
    st.info(f"Banco local SQLite: `{DB.name}`. Faça cópias de segurança periódicas deste arquivo.")
    if st.button("Gerar backup completo (banco + documentos anexados)", key="make_backup"):
        st.session_state.backup_blob = (make_backup_zip(), datetime.now().strftime("%Y%m%d_%H%M"))
    if "backup_blob" in st.session_state:
        _blob, _stamp = st.session_state.backup_blob
        st.download_button("Baixar backup (.zip)", _blob, file_name=f"backup_frota_{_stamp}.zip",
                           mime="application/zip", key="download_backup")
    st.markdown("#### Sobre")
    st.write("Gestão de Manutenção de Frota — RL Connect")
    st.caption("Aplicação local desenvolvida em Python, Streamlit e SQLite.")

st.divider()
st.caption("Importante: os intervalos do catálogo são referências iniciais da planilha. Confirme sempre os prazos no manual do fabricante e nas condições reais de uso.")
