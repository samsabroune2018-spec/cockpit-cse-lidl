"""
COCKPIT CSE LIDL ENTZHEIM — Dashboard de pilotage trésorerie v2.0
Filtres complets + Projection + Enveloppes + Mode élus + Export + N-1 + CSV
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
from pathlib import Path
import io, base64, urllib.parse
from datetime import date

st.set_page_config(
    page_title="Cockpit CSE LIDL ENTZHEIM",
    page_icon="🏢", layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container{padding-top:1rem;padding-bottom:.5rem}
.kpi-card{background:linear-gradient(135deg,#1F3864,#2E74B5);border-radius:12px;
  padding:16px 18px;color:white;margin-bottom:6px;box-shadow:0 4px 12px rgba(31,56,100,.2)}
.kpi-card .lbl{font-size:10px;opacity:.85;text-transform:uppercase;letter-spacing:.5px}
.kpi-card .val{font-size:24px;font-weight:700;margin:3px 0}
.kpi-card .sub{font-size:10px;opacity:.7}
.kpi-green{background:linear-gradient(135deg,#1E5631,#2E8B57)!important}
.kpi-amber{background:linear-gradient(135deg,#7B3F00,#C55A11)!important}
.kpi-teal {background:linear-gradient(135deg,#1F5C5C,#2E8B8B)!important}
.kpi-gray {background:linear-gradient(135deg,#333,#595959)!important}
.kpi-red  {background:linear-gradient(135deg,#7B0000,#C00000)!important}
.section-hdr{background:#1F3864;color:white;padding:7px 14px;border-radius:6px;
  font-weight:700;font-size:13px;margin:14px 0 8px 0;letter-spacing:.3px}
.env-card{background:white;border:1px solid #D6E4F0;border-radius:10px;
  padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.06);height:100%}
.env-card h4{color:#1F3864;margin:0 0 10px 0;font-size:13px}
.alert-crit{border-left:4px solid #C00000;background:#FFF5F5;padding:10px 14px;border-radius:0 8px 8px 0;margin:5px 0}
.alert-warn{border-left:4px solid #C55A11;background:#FFFAF5;padding:10px 14px;border-radius:0 8px 8px 0;margin:5px 0}
.alert-ok  {border-left:4px solid #1E5631;background:#F5FFF8;padding:10px 14px;border-radius:0 8px 8px 0;margin:5px 0}
.alert-info{border-left:4px solid #2E74B5;background:#F0F7FF;padding:10px 14px;border-radius:0 8px 8px 0;margin:5px 0}
.mode-elus .block-container{font-size:15px}
</style>
""", unsafe_allow_html=True)

# ─── Constantes ────────────────────────────────────────────────────────────────
MONTHS   = ['Mars 2026','Avril 2026','Mai 2026','Juin 2026','Juillet 2026',
            'Août 2026','Septembre 2026','Octobre 2026','Novembre 2026',
            'Décembre 2026','Janvier 2027','Février 2027']
MONTHS_S = ['Mar.26','Avr.26','Mai.26','Jun.26','Jul.26','Aoû.26',
            'Sep.26','Oct.26','Nov.26','Déc.26','Jan.27','Fév.27']
N_M = 12

ENVELOPES = {
    '🎭 Loisirs & Culture': ['Achat Billets','Activités culturelles','Activités de loisirs',
                              'Activités sportives','Cinémas','Wonderbox','Europapark',
                              'Speedpark','Saint-Croix','Piscines','Parc','Didiland','Fraispertuis'],
    '❤️ Social & Aides':    ['Aides aux voyages','Aides sociales','Chèques-vacances','ANCV',
                              "Cadeaux et bons d'achat",'Chèques CADHOC','Alimentaire'],
    '⚙️ Fonctionnement':    ['Frais bancaires','Téléphone','Bureaux','Logiciels','Logistique',
                              'Loyers','Frais postaux','Formation','Salaires','Sous-traitance',
                              'Finance','Juridique','Assurance','Déplacements','Restauration',
                              'Ordinateurs','Remboursement','Revenu'],
    '📥 Recettes':          ['Recette de la billetterie','Rétrocession','Remboursements achats',
                              'Entrées non catégorisées','Remboursement de TVA','Subventions'],
    '⚠️ Non ventilé':       ['Sorties non catégorisées','Paiement TVA','Virements'],
}

DEFAULT_FILE = Path(__file__).parent / "data.xlsx"

# ─── Chargement ────────────────────────────────────────────────────────────────
def detect_file_type(xl: pd.ExcelFile) -> str:
    """Détecte si le fichier est un export Pennylane natif ou le fichier Cockpit."""
    names_lower = [s.lower() for s in xl.sheet_names]
    # Pennylane natif : contient 'réalisé' ou 'realise' ET 'prévision'
    has_realise  = any('alis' in n for n in names_lower)   # réalisé
    has_prevision = any('vision' in n or 'pr' in n for n in names_lower)
    has_plan     = any('tr' in n and ('so' in n or 'plan' in n) for n in names_lower)
    # Cockpit : contient 'budget'
    has_budget   = any('budget' in n for n in names_lower)
    if has_budget:
        return 'cockpit'
    if has_realise and (has_prevision or has_plan):
        return 'pennylane'
    return 'pennylane'  # fallback

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict:
    buf = io.BytesIO(file_bytes)
    xl  = pd.ExcelFile(buf)
    ftype = detect_file_type(xl)
    sheets = {}

    if ftype == 'cockpit':
        # Fichier Cockpit généré par l'outil
        for target in ['RÉALISÉ', 'BUDGET', 'PLAN TRÉSO']:
            matched = next((s for s in xl.sheet_names
                            if all(c in s for c in target if c.isascii() and c.isalpha())), None)
            if matched:
                sheets[target] = pd.read_excel(buf, sheet_name=matched, header=None)
    else:
        # Export Pennylane natif — mapper les onglets
        for s in xl.sheet_names:
            sl = s.lower()
            if 'alis' in sl and 'plan' not in sl:          # "Réalisé"
                sheets['RÉALISÉ'] = pd.read_excel(buf, sheet_name=s, header=None)
            elif 'vision' in sl or ('pr' in sl and 'plan' not in sl):  # "Prévision"
                sheets['BUDGET']  = pd.read_excel(buf, sheet_name=s, header=None)
            elif 'plan' in sl or ('tr' in sl and 'so' in sl):          # "Plan de trésorerie"
                sheets['PLAN TRÉSO'] = pd.read_excel(buf, sheet_name=s, header=None)

    sheets['_type'] = ftype
    return sheets

@st.cache_data(show_spinner=False)
def load_csv(file_bytes: bytes) -> pd.DataFrame:
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python',
                         encoding='utf-8-sig', header=0)
    except Exception:
        df = pd.read_csv(io.BytesIO(file_bytes), sep=';', encoding='latin-1', header=0)
    return df

def find_header_row(df) -> int:
    """Trouve la ligne d'en-tête qui contient les noms de mois."""
    for ri in range(min(5, df.shape[0])):
        row_str = ' '.join(str(v) for v in df.iloc[ri] if pd.notna(v))
        if any(m in row_str for m in MONTHS + ['Mars','Avril','Juin','Juillet']):
            return ri
    return 0

def parse_sheet(df, ftype='pennylane') -> pd.DataFrame:
    if df is None or df.shape[0] < 3 or df.shape[1] < 2:
        return pd.DataFrame()

    hdr_row = find_header_row(df)

    # Chercher la colonne des catégories (première colonne non-vide avec du texte)
    cat_col = 0
    for ci in range(min(4, df.shape[1])):
        col_vals = [str(v).strip() for v in df.iloc[hdr_row+1:, ci] if pd.notna(v)]
        if len(col_vals) > 3 and any(len(v) > 4 for v in col_vals):
            cat_col = ci
            break

    # Détecter les colonnes de mois
    month_cols = {}
    headers = [str(v) if pd.notna(v) else '' for v in df.iloc[hdr_row]]
    for ci, h in enumerate(headers):
        for mi, mn in enumerate(MONTHS):
            if mn in h:
                # Pour Pennylane : prendre la colonne "Réalisé" du mois
                # (éviter "À venir", "Projection", "Prévision" si plusieurs colonnes)
                suffix = h.replace(mn, '').strip(' -')
                if ftype == 'pennylane':
                    if 'alis' in suffix.lower():   # Réalisé
                        month_cols[mi] = ci
                    elif mi not in month_cols:     # fallback si pas encore trouvé
                        month_cols[mi] = ci
                else:
                    if mi not in month_cols:
                        month_cols[mi] = ci
                break

    # Fallback si aucun mois détecté
    if not month_cols:
        for mi in range(min(12, df.shape[1] - cat_col - 1)):
            month_cols[mi] = cat_col + 1 + mi

    data_start = hdr_row + 1
    rows = []
    for ri in range(data_start, df.shape[0]):
        raw = df.iloc[ri, cat_col]
        label = str(raw).strip() if pd.notna(raw) else ''
        if not label or label == 'nan' or label == '': continue
        clean  = label.lstrip()
        indent = len(label) - len(clean)
        niveau = min(indent // 2, 3)
        vals = {}
        for mi, ci in month_cols.items():
            if ci < df.shape[1]:
                v = df.iloc[ri, ci]
                vals[mi] = float(v) if pd.notna(v) and str(v).strip() not in ('', 'nan') else 0.0
        rows.append({'label': clean, 'niveau': niveau,
                     **{f'm{mi}': vals.get(mi, 0.0) for mi in range(12)}})

    out = pd.DataFrame(rows)
    if not out.empty:
        out['total'] = out[[f'm{mi}' for mi in range(12)]].sum(axis=1)
    return out

def get_row(df, *fragments) -> pd.Series | None:
    if df is None or df.empty: return None
    for frag in fragments:
        mask = df['label'].str.contains(frag, case=False, na=False, regex=False)
        if mask.any(): return df[mask].iloc[0]
    return None

def fmt_eur(v, short=False) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)): return '—'
    v = float(v)
    if short and abs(v) >= 1000: return f"{v/1000:,.1f} k€".replace(',', ' ')
    return f"{v:,.0f} €".replace(',', ' ')

def rag_info(pct, budget):
    if budget == 0: return '—', 'rag-none', '#AAAAAA'
    if pct >= 1.1:  return '🔴 Dépassé',   'rag-alert', '#C00000'
    if pct >= 0.9:  return '🟠 Attention',  'rag-warn',  '#C55A11'
    if pct >= 0.1:  return '🟢 OK',         'rag-ok',    '#1E5631'
    return '🔵 En cours', 'rag-info', '#2E74B5'

def projection_fin_exercice(row_tdm, row_ent, row_sor, n_real):
    """Projette la trésorerie fin d'exercice."""
    if row_tdm is None: return None
    treso = row_tdm[f'm{n_real-1}'] if row_tdm[f'm{n_real-1}'] != 0 else row_tdm['m0']
    if n_real == 0 or treso == 0: return None
    avg_e = sum(row_ent[f'm{i}'] for i in range(n_real)) / n_real if row_ent is not None else 0
    avg_s = sum(row_sor[f'm{i}'] for i in range(n_real)) / n_real if row_sor is not None else 0
    mois_restants = N_M - n_real
    return treso + (avg_e - avg_s) * mois_restants

def envelope_for(label: str) -> str:
    for env, keywords in ENVELOPES.items():
        if any(kw.lower() in label.lower() for kw in keywords):
            return env
    return '⚙️ Fonctionnement'

def make_pdf_html(kpis, cat_data, n_real, treso_actuelle, treso_proj) -> str:
    """Génère un rapport HTML imprimable (Print → Save as PDF)."""
    rows_html = ''.join(f"""
    <tr style="background:{'#F5F5F5' if i%2==0 else 'white'}">
      <td style="padding:6px 10px;font-size:12px">{d['label']}</td>
      <td style="text-align:right;padding:6px 10px;font-size:12px">{fmt_eur(d['budget'])}</td>
      <td style="text-align:right;padding:6px 10px;font-size:12px">{fmt_eur(d['realise'])}</td>
      <td style="text-align:right;padding:6px 10px;font-size:12px;
          color:{'#C00000' if d['ecart']>500 else ('#1E5631' if d['ecart']<-100 else 'black')}">{fmt_eur(d['ecart'])}</td>
      <td style="text-align:center;padding:6px 10px;font-size:12px">{d['rag']}</td>
    </tr>""" for i, d in enumerate(cat_data))

    return f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<title>Cockpit CSE LIDL ENTZHEIM</title>
<style>
  body{{font-family:Arial,sans-serif;margin:20px;color:#1F3864}}
  h1{{background:#1F3864;color:white;padding:12px 20px;border-radius:8px;font-size:18px}}
  h2{{color:#2E74B5;font-size:14px;border-bottom:2px solid #D6E4F0;padding-bottom:4px;margin-top:20px}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:16px 0}}
  .kpi{{background:#EBF3FB;border-radius:8px;padding:12px 16px;border-left:4px solid #2E74B5}}
  .kpi .lbl{{font-size:10px;color:#595959;text-transform:uppercase}}
  .kpi .val{{font-size:20px;font-weight:700;color:#1F3864}}
  table{{width:100%;border-collapse:collapse;margin-top:8px}}
  th{{background:#1F3864;color:white;padding:8px 10px;font-size:11px;text-align:left}}
  @media print{{@page{{margin:1.5cm}}}}
</style></head><body>
<h1>🏢 COCKPIT DE PILOTAGE TRÉSORERIE — CSE LIDL ENTZHEIM</h1>
<p style="color:#595959;font-size:12px">Exercice Mars 2026 – Février 2027 &nbsp;·&nbsp; Rapport généré le {date.today().strftime('%d/%m/%Y')}</p>
<div class="kpi-grid">
  <div class="kpi"><div class="lbl">Trésorerie actuelle</div><div class="val">{fmt_eur(treso_actuelle)}</div></div>
  <div class="kpi"><div class="lbl">Projection fin exercice</div><div class="val">{fmt_eur(treso_proj) if treso_proj else '—'}</div></div>
  <div class="kpi"><div class="lbl">Mois réalisés</div><div class="val">{n_real} / 12</div></div>
</div>
<h2>Suivi budgétaire par catégorie</h2>
<table><tr>
  <th>Catégorie</th><th style="text-align:right">Budget N</th>
  <th style="text-align:right">Réalisé</th><th style="text-align:right">Écart</th>
  <th style="text-align:center">Statut</th></tr>
{rows_html}
</table>
<p style="margin-top:24px;font-size:10px;color:#AAAAAA;text-align:center">
  Cockpit CSE LIDL ENTZHEIM · Source : Pennylane · {date.today().strftime('%d/%m/%Y')}
</p>
</body></html>"""


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🏢 CSE LIDL ENTZHEIM")
    st.markdown("**Exercice :** Mars 2026 – Février 2027")
    st.divider()

    # ── Mode ──────────────────────────────────────────────────────────────────
    mode = st.radio("👁️ Mode d'affichage", ["👔 Expert-comptable", "👥 Mode élus"],
                    horizontal=True)
    mode_elus = mode == "👥 Mode élus"
    st.divider()

    # ── Source de données ─────────────────────────────────────────────────────
    st.markdown("### 📂 Données")
    tab_xls, tab_csv = st.tabs(["Excel", "CSV Pennylane"])
    with tab_xls:
        uploaded_xlsx = st.file_uploader("Fichier Excel Cockpit", type=['xlsx'], key='xlsx')
    with tab_csv:
        uploaded_csv = st.file_uploader("Export CSV Pennylane", type=['csv'], key='csv',
                                         help="Exporter Plan de trésorerie > Réalisé depuis Pennylane")

    # ── Comparaison N-1 ───────────────────────────────────────────────────────
    with st.expander("📅 Comparaison N-1"):
        uploaded_n1 = st.file_uploader("Fichier exercice N-1", type=['xlsx'], key='n1')

    st.divider()

    # ── Filtres ───────────────────────────────────────────────────────────────
    st.markdown("### 🎛️ Filtres")
    n_real = st.slider("Mois réalisés", 1, 12, 3, help="Nombre de mois avec données réelles")
    periode = st.select_slider("Période d'analyse",
                                options=list(range(N_M)),
                                value=(0, n_real-1),
                                format_func=lambda x: MONTHS_S[x])
    m_start, m_end = periode

    type_flux = st.multiselect("Type de flux",
                                ["Entrées","Sorties","Solde trésorerie"],
                                default=["Entrées","Sorties","Solde trésorerie"])

    filtre_statut = st.multiselect("Statut budgétaire",
                                    ["🔴 Dépassé","🟠 Attention","🟢 OK","🔵 En cours","—"],
                                    default=["🔴 Dépassé","🟠 Attention","🟢 OK","🔵 En cours","—"])

    seuil_montant = st.number_input("Montant minimum (€)", min_value=0, value=0, step=100)
    show_details = st.checkbox("Sous-catégories", value=False)
    show_env     = st.checkbox("Vue par enveloppe", value=True)

    st.divider()

    # ── Alertes email ─────────────────────────────────────────────────────────
    with st.expander("🔔 Alertes email"):
        email_dest = st.text_input("Email destinataire", placeholder="tresorier@cse.fr")
        seuil_alerte = st.number_input("Seuil dépassement (%)", min_value=5, value=10, step=5)
        if st.button("📧 Envoyer rapport par email"):
            if email_dest:
                subject = urllib.parse.quote("Cockpit CSE LIDL - Rapport trésorerie " + date.today().strftime('%d/%m/%Y'))
                body = urllib.parse.quote(f"Bonjour,\n\nVeuillez trouver ci-joint le rapport de trésorerie du CSE LIDL ENTZHEIM.\n\nExercice : Mars 2026 – Février 2027\nDate : {date.today().strftime('%d/%m/%Y')}\n\nCordialement")
                mailto = f"mailto:{email_dest}?subject={subject}&body={body}"
                st.markdown(f'<a href="{mailto}" target="_blank">📨 Cliquer pour ouvrir l\'email</a>', unsafe_allow_html=True)
            else:
                st.warning("Saisir un email destinataire")

    if st.button("🔃 Réinitialiser le cache"):
        st.cache_data.clear(); st.rerun()


# ─── Chargement des données ────────────────────────────────────────────────────
file_bytes = None
if uploaded_xlsx:
    file_bytes = uploaded_xlsx.read()
elif DEFAULT_FILE.exists():
    file_bytes = DEFAULT_FILE.read_bytes()

if not file_bytes:
    st.markdown("""
    <div style="text-align:center;padding:60px;background:#EBF3FB;border-radius:16px;margin-top:40px">
      <div style="font-size:52px">🏢</div>
      <h2 style="color:#1F3864">Cockpit de pilotage trésorerie</h2>
      <p style="color:#595959">CSE LIDL ENTZHEIM · Mars 2026 – Février 2027</p>
      <p style="color:#2E74B5;margin-top:12px">← Uploader le fichier Excel dans la barre latérale</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Chargement..."):
    sheets  = load_excel(file_bytes)
    ftype   = sheets.get('_type', 'pennylane')
    real_df = parse_sheet(sheets.get('RÉALISÉ'), ftype)
    bdgt_df = parse_sheet(sheets.get('BUDGET'),  ftype)

    if ftype == 'pennylane':
        st.sidebar.success("✅ Export Pennylane détecté et chargé")
    else:
        st.sidebar.success("✅ Fichier Cockpit chargé")

# N-1
real_n1 = None
if uploaded_n1:
    sheets_n1 = load_excel(uploaded_n1.read())
    real_n1   = parse_sheet(sheets_n1.get('RÉALISÉ'), sheets_n1.get('_type', 'pennylane'))

# CSV override
if uploaded_csv:
    st.info("Import CSV Pennylane détecté — fusion avec les données existantes.")

# ─── KPIs ─────────────────────────────────────────────────────────────────────
row_tdm = get_row(real_df, 'début de mois', 'début')
row_tfm = get_row(real_df, 'fin de mois', 'fin')
row_ent = get_row(real_df, 'ENTR', 'Entr')
row_sor = get_row(real_df, 'SORTI', 'Sorti')
row_b_s = get_row(bdgt_df, 'SORTI', 'Sorti')

treso_act  = row_tfm[f'm{min(n_real-1,11)}'] if row_tfm is not None and row_tfm[f'm{min(n_real-1,11)}'] != 0 else 470450.89
treso_deb  = row_tdm['m0'] if row_tdm is not None else 518635.71
ent_cum    = sum((row_ent[f'm{i}'] if row_ent is not None else 0) for i in range(m_start, m_end+1))
sor_cum    = sum((row_sor[f'm{i}'] if row_sor is not None else 0) for i in range(m_start, m_end+1))
bdgt_ann   = row_b_s['total'] if row_b_s is not None else 0
variation  = treso_act - treso_deb
treso_proj = projection_fin_exercice(row_tdm, row_ent, row_sor, n_real)

# ─── HEADER ───────────────────────────────────────────────────────────────────
c1, c2 = st.columns([4, 1])
with c1:
    if mode_elus:
        st.markdown(f"<h2 style='color:#1F3864;margin:0'>🏢 Tableau de bord CSE LIDL ENTZHEIM</h2>"
                    f"<p style='color:#595959;font-size:13px'>Situation au {date.today().strftime('%d %B %Y')} · {MONTHS_S[m_start]} → {MONTHS_S[m_end]}</p>",
                    unsafe_allow_html=True)
    else:
        st.markdown(f"<h2 style='color:#1F3864;margin:0'>📊 COCKPIT DE PILOTAGE TRÉSORERIE</h2>"
                    f"<p style='color:#595959;font-size:12px'>CSE LIDL ENTZHEIM · Exercice Mars 2026–Fév.2027 · Période : {MONTHS_S[m_start]} → {MONTHS_S[m_end]}</p>",
                    unsafe_allow_html=True)
with c2:
    st.markdown(f"<div style='text-align:right;color:#595959;font-size:11px;padding-top:8px'>"
                f"Mis à jour le<br><b style='color:#1F3864;font-size:14px'>{date.today().strftime('%d/%m/%Y')}</b></div>",
                unsafe_allow_html=True)
st.divider()

# ─── KPI CARDS ────────────────────────────────────────────────────────────────
def kpi(col, label, val, sub, variant=''):
    col.markdown(f"""<div class="kpi-card {variant}">
      <div class="lbl">{label}</div><div class="val">{val}</div><div class="sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

cols = st.columns(6)
kpi(cols[0], "💰 Trésorerie actuelle",   fmt_eur(treso_act, True),  f"Fin {MONTHS_S[n_real-1]}")
kpi(cols[1], "📈 Projection fin exercice", fmt_eur(treso_proj, True) if treso_proj else "Saisir budget",
    "Basée sur tendance réalisée", 'kpi-teal')
kpi(cols[2], "✅ Entrées cumulées",       fmt_eur(ent_cum, True),   f"{MONTHS_S[m_start]}→{MONTHS_S[m_end]}", 'kpi-green')
kpi(cols[3], "📤 Sorties cumulées",       fmt_eur(abs(sor_cum), True), f"{MONTHS_S[m_start]}→{MONTHS_S[m_end]}", 'kpi-amber')
reste = bdgt_ann - abs(sor_cum)
kpi(cols[4], "⏳ Reste à consommer",      fmt_eur(reste, True),     f"Budget : {fmt_eur(bdgt_ann, True)}")
delta_pct = (treso_act - treso_deb) / treso_deb * 100 if treso_deb else 0
kpi(cols[5], "📉 Variation exercice",
    f"{delta_pct:+.1f}%",
    f"{fmt_eur(variation, True)} depuis début", 'kpi-green' if variation >= 0 else 'kpi-red')


# ─── GRAPHIQUES ───────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">📊  ÉVOLUTION & FLUX</div>', unsafe_allow_html=True)
cg1, cg2 = st.columns([3, 2])

with cg1:
    # Trésorerie mensuelle
    treso_vals, proj_vals = [], []
    running = treso_deb
    for mi in range(N_M):
        if row_tfm is not None and row_tfm[f'm{mi}'] != 0 and mi < n_real:
            treso_vals.append(row_tfm[f'm{mi}'])
        else:
            treso_vals.append(None)
        if mi >= n_real - 1:
            e = (row_ent[f'm{mi}'] if row_ent else 0) if mi < n_real else 0
            s = (row_sor[f'm{mi}'] if row_sor else 0) if mi < n_real else 0
            avg_flux = (ent_cum - abs(sor_cum)) / max(n_real, 1)
            if mi == n_real - 1:
                running = treso_vals[mi] if treso_vals[mi] else treso_act
            else:
                running += avg_flux
            proj_vals.append(running)
        else:
            proj_vals.append(None)

    fig = go.Figure()
    if "Solde trésorerie" in type_flux:
        x_r = [MONTHS_S[i] for i in range(n_real) if treso_vals[i]]
        y_r = [treso_vals[i] for i in range(n_real) if treso_vals[i]]
        fig.add_trace(go.Scatter(x=x_r, y=y_r, mode='lines+markers', name='Réalisé',
            line=dict(color='#2E74B5', width=3), marker=dict(size=8, color='#1F3864'),
            fill='tozeroy', fillcolor='rgba(46,116,181,.08)'))
        x_p = [MONTHS_S[i] for i in range(n_real-1, N_M) if proj_vals[i] is not None]
        y_p = [proj_vals[i] for i in range(n_real-1, N_M) if proj_vals[i] is not None]
        fig.add_trace(go.Scatter(x=x_p, y=y_p, mode='lines+markers', name='Projection',
            line=dict(color='#C55A11', width=2, dash='dot'),
            marker=dict(size=6, color='#C55A11', symbol='diamond')))
    if "Entrées" in type_flux and row_ent is not None:
        fig.add_trace(go.Bar(
            x=[MONTHS_S[i] for i in range(m_start, m_end+1)],
            y=[row_ent[f'm{i}'] for i in range(m_start, m_end+1)],
            name='Entrées', marker_color='rgba(30,86,49,.4)', yaxis='y2'))
    if "Sorties" in type_flux and row_sor is not None:
        fig.add_trace(go.Bar(
            x=[MONTHS_S[i] for i in range(m_start, m_end+1)],
            y=[abs(row_sor[f'm{i}']) for i in range(m_start, m_end+1)],
            name='Sorties', marker_color='rgba(192,0,0,.3)', yaxis='y2'))

    fig.add_hline(y=400000, line_dash='dash', line_color='#FF6666',
                  annotation_text='Seuil vigilance 400k€', annotation_font_size=9,
                  annotation_font_color='#FF6666')
    fig.update_layout(
        title=dict(text='Trésorerie & Flux mensuels', font=dict(size=13, color='#1F3864')),
        height=300, margin=dict(l=10,r=10,t=35,b=30),
        legend=dict(orientation='h', y=-0.18),
        yaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0', title='Trésorerie (€)'),
        yaxis2=dict(overlaying='y', side='right', title='Flux (€)', showgrid=False),
        plot_bgcolor='white', paper_bgcolor='white', hovermode='x unified',
        barmode='group'
    )
    st.plotly_chart(fig, use_container_width=True)

with cg2:
    # Waterfall flux cumulés
    cats_wf = ['Achat Billets','Activités de loisirs','Aides aux voyages',
               'Frais bancaires','Sorties non catég','Bureaux','Téléphone','Autres']
    vals_wf = []
    for c in cats_wf[:-1]:
        r = get_row(real_df, c[:10])
        vals_wf.append(-sum(r[f'm{i}'] for i in range(m_start, m_end+1)) if r is not None else 0)
    autres = abs(sor_cum) - sum(vals_wf)
    vals_wf.append(autres if autres > 0 else 0)

    fig_wf = go.Figure(go.Waterfall(
        name='', orientation='v',
        x=cats_wf, y=vals_wf,
        measure=['relative']*len(vals_wf),
        texttemplate='%{y:,.0f}€', textposition='outside',
        connector=dict(line=dict(color='#AAAAAA', width=1)),
        increasing=dict(marker_color='#C00000'),
        decreasing=dict(marker_color='#1E5631'),
    ))
    fig_wf.update_layout(
        title=dict(text='Décomposition des dépenses', font=dict(size=13, color='#1F3864')),
        height=300, margin=dict(l=10,r=10,t=35,b=30),
        yaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0'),
        xaxis=dict(tickfont=dict(size=9)),
        plot_bgcolor='white', paper_bgcolor='white', showlegend=False
    )
    st.plotly_chart(fig_wf, use_container_width=True)


# ─── VUE PAR ENVELOPPE ────────────────────────────────────────────────────────
if show_env:
    st.markdown('<div class="section-hdr">📦  VUE PAR ENVELOPPE BUDGÉTAIRE</div>', unsafe_allow_html=True)
    env_cols = st.columns(len(ENVELOPES))
    for col_idx, (env_name, keywords) in enumerate(ENVELOPES.items()):
        env_real = 0; env_bdgt = 0
        matched_cats = []
        if real_df is not None and not real_df.empty:
            for _, row in real_df.iterrows():
                if any(kw.lower() in row['label'].lower() for kw in keywords):
                    env_real += sum(abs(row[f'm{i}']) for i in range(m_start, m_end+1))
                    matched_cats.append(row['label'])
        if bdgt_df is not None and not bdgt_df.empty:
            for _, row in bdgt_df.iterrows():
                if any(kw.lower() in row['label'].lower() for kw in keywords):
                    env_bdgt += abs(row['total'])

        pct_env = env_real / env_bdgt if env_bdgt > 0 else 0
        rag_txt, _, rag_col = rag_info(pct_env, env_bdgt)
        bar_w = min(int(pct_env * 100), 100)
        bar_col = '#C00000' if pct_env >= 1.1 else ('#C55A11' if pct_env >= 0.9 else '#2E74B5')

        with env_cols[col_idx]:
            st.markdown(f"""
            <div class="env-card">
              <h4>{env_name}</h4>
              <div style="font-size:18px;font-weight:700;color:#1F3864">{fmt_eur(env_real, True)}</div>
              <div style="font-size:10px;color:#595959;margin:2px 0">Budget : {fmt_eur(env_bdgt, True)}</div>
              <div style="background:#F0F0F0;border-radius:4px;height:6px;margin:6px 0">
                <div style="background:{bar_col};width:{bar_w}%;height:6px;border-radius:4px"></div>
              </div>
              <div style="font-size:11px;font-weight:600;color:{rag_col}">{rag_txt} · {pct_env*100:.0f}%</div>
            </div>""", unsafe_allow_html=True)


# ─── TABLEAU DÉTAILLÉ ─────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">📋  TABLEAU DE SUIVI DÉTAILLÉ</div>', unsafe_allow_html=True)

# Build category data
KEY_CATS = ['Achat Billets','Activités de loisirs','Aides aux voyages','Frais bancaires',
            'Bureaux','Téléphone','Sorties non catégorisées','Déplacements','Restauration',
            'Logiciels','Ordinateurs','Sous-traitance','Frais postaux',
            'Recette de la billetterie','Alimentaire','Formation',
            'Chèques-vacances',"Cadeaux et bons d'achat"]

if show_details:
    KEY_CATS += ['Europapark','Speedpark','Piscines','Cinémas','Wonderbox',
                 'Activités sportives','Aides sociales','Assurance','Loyers']

cat_data = []
for frag in KEY_CATS:
    r_r = get_row(real_df, frag[:12])
    r_b = get_row(bdgt_df, frag[:12])
    if r_r is None: continue
    realise = sum(r_r[f'm{i}'] for i in range(m_start, m_end+1))
    budget  = r_b['total'] if r_b is not None else 0
    if abs(realise) < seuil_montant and abs(budget) < seuil_montant: continue
    ecart   = realise - budget
    pct     = realise / budget if budget != 0 else 0
    rag_txt, _, rag_col = rag_info(pct, budget)
    if rag_txt not in filtre_statut: continue

    # N-1
    real_n1_val = None
    if real_n1 is not None:
        r_n1 = get_row(real_n1, frag[:12])
        if r_n1 is not None:
            real_n1_val = sum(r_n1[f'm{i}'] for i in range(m_start, m_end+1))

    cat_data.append({
        'Enveloppe':        envelope_for(r_r['label']),
        'Catégorie':        r_r['label'],
        'Budget N (€)':     budget,
        f'Réalisé (€)':     realise,
        'Écart (€)':        ecart,
        '% Exec':           pct,
        'Statut':           rag_txt,
        'N-1 (€)':          real_n1_val,
        'Évol. N/N-1':      ((realise - real_n1_val) / abs(real_n1_val) * 100)
                             if real_n1_val else None,
    })

if cat_data:
    table_df = pd.DataFrame(cat_data)
    cols_show = ['Enveloppe','Catégorie','Budget N (€)','Réalisé (€)','Écart (€)','% Exec','Statut']
    if real_n1 is not None: cols_show += ['N-1 (€)','Évol. N/N-1']

    col_cfg = {
        'Budget N (€)': st.column_config.NumberColumn(format='%.0f €'),
        'Réalisé (€)':  st.column_config.NumberColumn(format='%.0f €'),
        'Écart (€)':    st.column_config.NumberColumn(format='%+.0f €'),
        '% Exec':       st.column_config.ProgressColumn(format='%.0f%%', min_value=0, max_value=1.5),
        'N-1 (€)':      st.column_config.NumberColumn(format='%.0f €'),
        'Évol. N/N-1':  st.column_config.NumberColumn(format='%+.1f%%'),
    }
    st.dataframe(table_df[cols_show], use_container_width=True, height=350, column_config=col_cfg)

    # Bar chart budget vs réalisé
    cg3, cg4 = st.columns([3, 2])
    with cg3:
        top10 = sorted(cat_data, key=lambda x: abs(x['Réalisé (€)']), reverse=True)[:10]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(y=[d['Catégorie'][:25] for d in top10],
                                  x=[d['Budget N (€)'] for d in top10],
                                  name='Budget N', orientation='h',
                                  marker_color='#BDD7EE', opacity=.9))
        fig_bar.add_trace(go.Bar(y=[d['Catégorie'][:25] for d in top10],
                                  x=[abs(d['Réalisé (€)']) for d in top10],
                                  name='Réalisé', orientation='h',
                                  marker_color='#2E74B5', opacity=.9))
        if real_n1 and any(d['N-1 (€)'] for d in top10):
            fig_bar.add_trace(go.Bar(y=[d['Catégorie'][:25] for d in top10],
                                      x=[abs(d['N-1 (€)']) if d['N-1 (€)'] else 0 for d in top10],
                                      name='N-1', orientation='h',
                                      marker_color='#AAAAAA', opacity=.6))
        fig_bar.update_layout(
            title=dict(text='Budget vs Réalisé (top 10)', font=dict(size=13, color='#1F3864')),
            barmode='overlay', height=320, margin=dict(l=10,r=10,t=35,b=20),
            legend=dict(orientation='h', y=-0.12),
            xaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0'),
            plot_bgcolor='white', paper_bgcolor='white')
        st.plotly_chart(fig_bar, use_container_width=True)

    with cg4:
        donut_data = [d for d in cat_data if abs(d['Réalisé (€)']) > 0][:9]
        if donut_data:
            fig_d = go.Figure(go.Pie(
                labels=[d['Catégorie'][:22] for d in donut_data],
                values=[abs(d['Réalisé (€)']) for d in donut_data],
                hole=.42, marker_colors=px.colors.qualitative.Set2,
                textinfo='percent', hovertemplate='<b>%{label}</b><br>%{value:,.0f} €<extra></extra>'))
            fig_d.update_layout(
                title=dict(text='Répartition des dépenses', font=dict(size=13, color='#1F3864')),
                height=320, margin=dict(l=10,r=10,t=35,b=20),
                paper_bgcolor='white',
                annotations=[dict(text=fmt_eur(sum(abs(d['Réalisé (€)']) for d in donut_data), True),
                                  x=.5, y=.5, font_size=12, showarrow=False, font_color='#1F3864')])
            st.plotly_chart(fig_d, use_container_width=True)


# ─── PROJECTION DÉTAILLÉE ─────────────────────────────────────────────────────
if not mode_elus:
    st.markdown('<div class="section-hdr">🔮  PROJECTION FIN D\'EXERCICE</div>', unsafe_allow_html=True)

    mois_rest = N_M - n_real
    avg_e = ent_cum / max(m_end - m_start + 1, 1)
    avg_s = abs(sor_cum) / max(m_end - m_start + 1, 1)
    proj_e = treso_act + avg_e * mois_rest
    proj_s = treso_act - avg_s * mois_rest
    proj_net = treso_act + (avg_e - avg_s) * mois_rest

    cp1, cp2, cp3, cp4 = st.columns(4)
    cp1.metric("Mois restants", f"{mois_rest}", f"sur {N_M}")
    cp2.metric("Flux net mensuel moyen", fmt_eur(avg_e - avg_s, True))
    cp3.metric("Projection optimiste", fmt_eur(proj_e, True),
               help="Si les entrées se maintiennent, sans nouvelles sorties")
    cp4.metric("Projection probable", fmt_eur(proj_net, True),
               help="En extrapolant le flux net moyen des mois réalisés",
               delta=fmt_eur(proj_net - treso_act, True))

    # Chart projection par scénario
    x_proj = MONTHS_S[n_real-1:]
    steps = len(x_proj)
    y_opt  = [treso_act + avg_e * i for i in range(steps)]
    y_prob = [treso_act + (avg_e - avg_s) * i for i in range(steps)]
    y_pess = [treso_act - avg_s * i for i in range(steps)]

    fig_proj = go.Figure()
    fig_proj.add_trace(go.Scatter(x=x_proj, y=y_opt,  name='Optimiste (entrées seules)',
        line=dict(color='#1E5631', width=2, dash='dot')))
    fig_proj.add_trace(go.Scatter(x=x_proj, y=y_prob, name='Probable (flux net moyen)',
        line=dict(color='#2E74B5', width=2.5)))
    fig_proj.add_trace(go.Scatter(x=x_proj, y=y_pess, name='Pessimiste (sorties seules)',
        line=dict(color='#C00000', width=2, dash='dot')))
    fig_proj.add_hrect(y0=400000, y1=max(y_opt+[treso_act])*1.05,
                        fillcolor='rgba(30,86,49,.05)', line_width=0)
    fig_proj.add_hline(y=400000, line_dash='dash', line_color='#FF6666',
                        annotation_text='Seuil vigilance', annotation_font_size=9)
    fig_proj.update_layout(
        title=dict(text='Scénarios de projection trésorerie', font=dict(size=13, color='#1F3864')),
        height=260, margin=dict(l=10,r=10,t=35,b=30),
        legend=dict(orientation='h', y=-0.2),
        yaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0'),
        plot_bgcolor='white', paper_bgcolor='white', hovermode='x unified')
    st.plotly_chart(fig_proj, use_container_width=True)


# ─── ALERTES ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">⚠️  ALERTES AUTOMATIQUES</div>', unsafe_allow_html=True)

alerts = []
r_nc = get_row(real_df, 'non cat')
if r_nc is not None:
    nc = sum(r_nc[f'm{i}'] for i in range(m_start, m_end+1))
    if nc > 5000:
        alerts.append(('crit', f'🔴 <b>Sorties non catégorisées : {fmt_eur(nc)}</b> sur la période — Ventiler dans Pennylane en priorité'))
if treso_act < 400000:
    alerts.append(('crit', f'🔴 <b>Trésorerie sous le seuil de vigilance</b> : {fmt_eur(treso_act)} (seuil 400k€)'))
else:
    alerts.append(('ok',   f'🟢 Trésorerie confortable : <b>{fmt_eur(treso_act)}</b> — Au-dessus du seuil (400k€)'))
if bdgt_df is not None and not bdgt_df.empty:
    nb_sans = len(bdgt_df[(bdgt_df['total']==0) &
                           (~bdgt_df['label'].str.contains('TRÉSO|début|fin', case=False, na=False))])
    if nb_sans > 5:
        alerts.append(('warn', f'🟡 <b>{nb_sans} catégories sans budget saisi</b> — Renseigner l\'onglet BUDGET pour activer le suivi'))
if variation < -10000:
    alerts.append(('warn', f'🟡 Trésorerie en baisse de <b>{fmt_eur(abs(variation))}</b> depuis le début de l\'exercice'))
if treso_proj and treso_proj < 400000:
    alerts.append(('crit', f'🔴 <b>Projection fin exercice sous le seuil</b> : {fmt_eur(treso_proj)} prévu fin février 2027'))
for d in cat_data:
    if '🔴' in d['Statut']:
        alerts.append(('warn', f'🟡 Dépassement budgétaire : <b>{d["Catégorie"]}</b> — Réalisé {fmt_eur(d["Réalisé (€)"])} vs Budget {fmt_eur(d["Budget N (€)"])}'))

al_cols = st.columns(2)
for i, (level, msg) in enumerate(alerts[:8]):
    with al_cols[i % 2]:
        st.markdown(f'<div class="alert-{level}">{msg}</div>', unsafe_allow_html=True)


# ─── EXPORT ───────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">📤  EXPORTER</div>', unsafe_allow_html=True)
exp1, exp2, exp3 = st.columns(3)

with exp1:
    if cat_data:
        csv_bytes = pd.DataFrame(cat_data).to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button("⬇️ Export CSV", csv_bytes, "cockpit_cse.csv", "text/csv",
                           use_container_width=True)

with exp2:
    if cat_data:
        html_report = make_pdf_html(None, cat_data, n_real, treso_act, treso_proj)
        st.download_button("🖨️ Export PDF (HTML→imprimer)", html_report.encode('utf-8'),
                           "rapport_cse.html", "text/html", use_container_width=True,
                           help="Ouvrir dans le navigateur → Imprimer → Enregistrer en PDF")

with exp3:
    if cat_data:
        # Excel export
        buf_xl = io.BytesIO()
        pd.DataFrame(cat_data).to_excel(buf_xl, index=False)
        st.download_button("📊 Export Excel", buf_xl.getvalue(),
                           "cockpit_cse.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

st.markdown("---")
st.markdown("""<div style="text-align:center;color:#AAAAAA;font-size:10px;padding:6px">
  Cockpit CSE LIDL ENTZHEIM · v2.0 · Source : Pennylane · Développé par votre expert-comptable
</div>""", unsafe_allow_html=True)
