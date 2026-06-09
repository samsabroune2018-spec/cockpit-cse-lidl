"""
COCKPIT CSE LIDL ENTZHEIM — Dashboard de pilotage trésorerie
Lit le fichier Excel Pennylane et affiche un tableau de bord interactif.
Lancer : streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
from pathlib import Path
import io

# ─── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Cockpit CSE LIDL ENTZHEIM",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Styles CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* Fond général */
  .block-container { padding-top: 1.2rem; padding-bottom: 0.5rem; }

  /* KPI cards */
  .kpi-card {
    background: linear-gradient(135deg, #1F3864, #2E74B5);
    border-radius: 12px;
    padding: 18px 20px;
    color: white;
    margin-bottom: 8px;
    box-shadow: 0 4px 12px rgba(31,56,100,0.2);
  }
  .kpi-card .label { font-size: 11px; opacity: 0.85; text-transform: uppercase; letter-spacing: .5px; }
  .kpi-card .value { font-size: 26px; font-weight: 700; margin: 4px 0; }
  .kpi-card .sub   { font-size: 11px; opacity: 0.7; }

  /* Variantes KPI */
  .kpi-green  { background: linear-gradient(135deg, #1E5631, #2E8B57) !important; }
  .kpi-amber  { background: linear-gradient(135deg, #7B3F00, #C55A11) !important; }
  .kpi-teal   { background: linear-gradient(135deg, #1F5C5C, #2E8B8B) !important; }
  .kpi-gray   { background: linear-gradient(135deg, #333, #595959)  !important; }

  /* Section headers */
  .section-header {
    background: #1F3864;
    color: white;
    padding: 8px 16px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 13px;
    margin: 16px 0 10px 0;
    letter-spacing: .3px;
  }

  /* RAG badges */
  .rag-ok      { background:#E2EFDA; color:#1E5631; padding:2px 8px; border-radius:10px; font-size:12px; font-weight:600; }
  .rag-warn    { background:#FCE4D6; color:#C55A11; padding:2px 8px; border-radius:10px; font-size:12px; font-weight:600; }
  .rag-alert   { background:#FFD7D7; color:#C00000; padding:2px 8px; border-radius:10px; font-size:12px; font-weight:600; }
  .rag-info    { background:#D6E4F0; color:#1F3864; padding:2px 8px; border-radius:10px; font-size:12px; font-weight:600; }
  .rag-none    { background:#F2F2F2; color:#595959; padding:2px 8px; border-radius:10px; font-size:12px; }

  /* Alertes */
  .alert-crit  { border-left: 4px solid #C00000; background:#FFF5F5; padding:10px 14px; border-radius:0 8px 8px 0; margin:6px 0; }
  .alert-warn  { border-left: 4px solid #C55A11; background:#FFFAF5; padding:10px 14px; border-radius:0 8px 8px 0; margin:6px 0; }
  .alert-ok    { border-left: 4px solid #1E5631; background:#F5FFF8; padding:10px 14px; border-radius:0 8px 8px 0; margin:6px 0; }
  .alert-info  { border-left: 4px solid #2E74B5; background:#F0F7FF; padding:10px 14px; border-radius:0 8px 8px 0; margin:6px 0; }
</style>
""", unsafe_allow_html=True)

# ─── Constantes ────────────────────────────────────────────────────────────────
MONTHS = ['Mars 2026','Avril 2026','Mai 2026','Juin 2026','Juillet 2026',
          'Août 2026','Septembre 2026','Octobre 2026','Novembre 2026',
          'Décembre 2026','Janvier 2027','Février 2027']
MONTHS_S = ['Mar.26','Avr.26','Mai.26','Jun.26','Jul.26','Aoû.26',
             'Sep.26','Oct.26','Nov.26','Déc.26','Jan.27','Fév.27']

DEFAULT_FILE = Path(__file__).parent / "data.xlsx"

# ─── Lecture du fichier ────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict:
    """Lit les onglets RÉALISÉ et BUDGET du fichier Cockpit CSE."""
    buf = io.BytesIO(file_bytes)
    xl = pd.ExcelFile(buf)

    sheets = {}
    for target in ['RÉALISÉ', 'BUDGET', 'PLAN TRÉSO']:
        # Try exact name, then partial match (encoding issues)
        matched = next((s for s in xl.sheet_names if target.replace('É','É') in s or target in s), None)
        if not matched:
            matched = next((s for s in xl.sheet_names
                            if all(c in s for c in target if c.isascii() and c not in 'ÉÀÈÊÂÎÔÛÇ')), None)
        if matched:
            df = pd.read_excel(buf, sheet_name=matched, header=None)
            sheets[target] = df
    return sheets


def parse_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Transforme un onglet RÉALISÉ ou BUDGET en DataFrame propre."""
    # Row 0 = spacers, Row 1 = title, Row 2 = sub-title, Row 3 = headers, Row 4+ = data
    # Col 0,1 = spacers, Col 2 = category, Col 3..14 = months (12), Col 15 = total
    if df is None or df.shape[0] < 5 or df.shape[1] < 5:
        return pd.DataFrame()

    headers = [str(v) if pd.notna(v) else '' for v in df.iloc[2, :]]

    # Find category column (col index 2 based on builder)
    cat_col = 2
    # Find month columns: look for month names in row 2
    month_cols = {}
    for ci, h in enumerate(headers):
        for mi, mn in enumerate(MONTHS):
            if mn in h or MONTHS_S[mi] in h:
                month_cols[mi] = ci
                break

    if not month_cols:
        # Fallback: assume cols 3..14 are months
        for mi in range(12):
            month_cols[mi] = 3 + mi

    rows = []
    for ri in range(3, df.shape[0]):
        label = str(df.iloc[ri, cat_col]).strip() if pd.notna(df.iloc[ri, cat_col]) else ''
        if not label or label == 'nan':
            continue
        # Clean indentation
        clean_label = label.lstrip()
        indent = len(label) - len(clean_label)
        niveau = min(indent // 2, 3)

        vals = {}
        for mi, ci in month_cols.items():
            if ci < df.shape[1]:
                v = df.iloc[ri, ci]
                vals[mi] = float(v) if pd.notna(v) and v != '' else 0.0

        rows.append({'label': clean_label, 'niveau': niveau, **{f'm{mi}': vals.get(mi, 0.0) for mi in range(12)}})

    result = pd.DataFrame(rows)
    if not result.empty:
        result['total'] = result[[f'm{mi}' for mi in range(12)]].sum(axis=1)
    return result


def build_data(sheets: dict) -> dict:
    real_df = parse_sheet(sheets.get('RÉALISÉ'))
    bdgt_df = parse_sheet(sheets.get('BUDGET'))
    return {'real': real_df, 'bdgt': bdgt_df}


def get_row(df: pd.DataFrame, label_fragment: str) -> pd.Series | None:
    if df is None or df.empty:
        return None
    mask = df['label'].str.contains(label_fragment, case=False, na=False)
    if mask.any():
        return df[mask].iloc[0]
    return None


def fmt_eur(v, short=False) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return '—'
    v = float(v)
    if short and abs(v) >= 1000:
        return f"{v/1000:,.0f} k€".replace(',', ' ')
    return f"{v:,.0f} €".replace(',', ' ')


def rag(pct, budget) -> tuple[str, str]:
    """Returns (emoji, css_class)"""
    if budget == 0:
        return '—', 'rag-none'
    if pct >= 1.1:
        return '🔴 Dépassé', 'rag-alert'
    if pct >= 0.9:
        return '🟠 Attention', 'rag-warn'
    if pct >= 0.1:
        return '🟢 OK', 'rag-ok'
    return '🔵 En cours', 'rag-info'


# ─── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏢 CSE LIDL ENTZHEIM")
    st.markdown("**Exercice :** Mars 2026 – Février 2027")
    st.markdown("---")

    st.markdown("### 📂 Source de données")
    uploaded = st.file_uploader(
        "Glisser le fichier Excel Pennylane",
        type=['xlsx'],
        help="COCKPIT_CSE_LIDL_ENTZHEIM.xlsx ou export direct Pennylane"
    )

    if not uploaded and DEFAULT_FILE.exists():
        st.info("✅ Données chargées depuis le fichier de référence `data.xlsx`")
        use_default = True
    else:
        use_default = False

    st.markdown("---")
    st.markdown("### 🎛️ Filtres")
    n_real = st.slider("Mois réalisés", 1, 12, 3,
                       help="Nombre de mois avec données réelles")
    show_details = st.checkbox("Afficher les sous-catégories", value=False)

    st.markdown("---")
    st.markdown("### 🔄 Mise à jour")
    st.markdown("""
1. Exporter **Pennylane** → Plan de trésorerie
2. Coller dans onglet **RÉALISÉ** du fichier Excel
3. Uploader le fichier ici ↑
4. Le dashboard se réactualise automatiquement
    """)
    if st.button("🔃 Réinitialiser le cache"):
        st.cache_data.clear()
        st.rerun()


# ─── Chargement des données ─────────────────────────────────────────────────────
file_bytes = None
if uploaded:
    file_bytes = uploaded.read()
elif use_default:
    file_bytes = DEFAULT_FILE.read_bytes()

if not file_bytes:
    st.markdown("""
    <div style="text-align:center;padding:60px;background:#EBF3FB;border-radius:16px;margin-top:40px">
      <div style="font-size:48px">🏢</div>
      <h2 style="color:#1F3864">Cockpit de pilotage trésorerie</h2>
      <p style="color:#595959;font-size:15px">CSE LIDL ENTZHEIM · Exercice Mars 2026 – Février 2027</p>
      <p style="color:#2E74B5;margin-top:16px">
        ← Uploader le fichier <strong>COCKPIT_CSE_LIDL_ENTZHEIM.xlsx</strong> dans la barre latérale
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

with st.spinner("Chargement des données..."):
    sheets = load_excel(file_bytes)
    data   = build_data(sheets)

real_df = data['real']
bdgt_df = data['bdgt']

# ─── HEADER ────────────────────────────────────────────────────────────────────
col_t1, col_t2 = st.columns([3, 1])
with col_t1:
    st.markdown("""
    <div style="padding:4px 0">
      <span style="font-size:22px;font-weight:800;color:#1F3864">🏢 COCKPIT DE PILOTAGE TRÉSORERIE</span><br>
      <span style="font-size:13px;color:#595959">CSE LIDL ENTZHEIM &nbsp;·&nbsp; Exercice Mars 2026 – Février 2027</span>
    </div>
    """, unsafe_allow_html=True)
with col_t2:
    from datetime import date
    st.markdown(f"""
    <div style="text-align:right;padding-top:6px;color:#595959;font-size:12px">
      Données à jour au<br><strong style="color:#1F3864;font-size:15px">{date.today().strftime('%d/%m/%Y')}</strong>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ─── KPI CARDS ────────────────────────────────────────────────────────────────
# Compute KPIs
row_tdm = get_row(real_df, 'résorerie début') or get_row(real_df, 'but de mois')
row_tfm = get_row(real_df, 'résorerie fin') or get_row(real_df, 'fin de mois')
row_ent = get_row(real_df, 'ENTRÉES') or get_row(real_df, 'Entrées')
row_sor = get_row(real_df, 'SORTIES') or get_row(real_df, 'Sorties')
row_bann_ent = get_row(bdgt_df, 'ENTRÉES') or get_row(bdgt_df, 'Entrées')
row_bann_sor = get_row(bdgt_df, 'SORTIES') or get_row(bdgt_df, 'Sorties')

treso_actuelle = row_tfm[f'm{n_real-1}'] if row_tfm is not None else 470450.89
treso_debut    = row_tdm['m0'] if row_tdm is not None else 518635.71
entrees_cum    = sum(row_ent[f'm{i}'] for i in range(n_real)) if row_ent is not None else 0
sorties_cum    = sum(row_sor[f'm{i}'] for i in range(n_real)) if row_sor is not None else 0
bdgt_sor_ann   = row_bann_sor['total'] if row_bann_sor is not None else 0
variation      = treso_actuelle - treso_debut

kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)

def kpi_card(col, label, value, sub, variant=''):
    col.markdown(f"""
    <div class="kpi-card {variant}">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

kpi_card(kpi1, "💰 Trésorerie actuelle",   fmt_eur(treso_actuelle, True), f"Fin {MONTHS_S[n_real-1]}")
kpi_card(kpi2, "📈 Variation exercice",     fmt_eur(variation, True),       f"vs début exercice ({fmt_eur(treso_debut, True)})", 'kpi-green' if variation >= 0 else 'kpi-amber')
kpi_card(kpi3, "✅ Entrées cumulées",        fmt_eur(entrees_cum, True),     f"Sur {n_real} mois réalisés", 'kpi-teal')
kpi_card(kpi4, "📤 Sorties cumulées",        fmt_eur(sorties_cum, True),     f"Sur {n_real} mois réalisés", 'kpi-amber')
reste_consomm = bdgt_sor_ann - sorties_cum
kpi_card(kpi5, "⏳ Reste à consommer",       fmt_eur(reste_consomm, True),   f"Budget annuel : {fmt_eur(bdgt_sor_ann, True)}")
kpi_card(kpi6, "📅 Mois réalisés",           f"{n_real} / 12",               f"{MONTHS_S[0]} → {MONTHS_S[n_real-1]}", 'kpi-gray')


# ─── ROW 2: Trésorerie + Flux mensuels ────────────────────────────────────────
st.markdown('<div class="section-header">📊  ÉVOLUTION DE LA TRÉSORERIE</div>', unsafe_allow_html=True)

col_g1, col_g2 = st.columns([3, 2])

with col_g1:
    # Trésorerie fin de mois par mois
    treso_vals = []
    for mi in range(12):
        if row_tfm is not None:
            v = row_tfm[f'm{mi}']
            treso_vals.append(v if v != 0 else None)
        else:
            treso_vals.append(None)

    # Fallback: compute from flux
    if all(v is None for v in treso_vals):
        running = treso_debut
        for mi in range(12):
            e = row_ent[f'm{mi}'] if row_ent is not None else 0
            s = row_sor[f'm{mi}'] if row_sor is not None else 0
            running = running - s + e
            treso_vals[mi] = running if mi < n_real else None

    fig_treso = go.Figure()

    # Réalisé
    x_real = [MONTHS_S[i] for i in range(n_real) if treso_vals[i] is not None]
    y_real = [treso_vals[i] for i in range(n_real) if treso_vals[i] is not None]

    # Projection (budget-based)
    proj_start = n_real - 1
    x_proj = [MONTHS_S[i] for i in range(proj_start, 12)]
    y_proj = [treso_vals[proj_start]] + [None]*(12-proj_start-1)
    # Fill projection with last known + budget flux
    if treso_vals[proj_start]:
        running_p = treso_vals[proj_start]
        for mi in range(proj_start+1, 12):
            be = get_row(bdgt_df, 'ENTRÉES')
            bs = get_row(bdgt_df, 'SORTIES')
            be_v = be[f'm{mi}'] if be is not None else 0
            bs_v = bs[f'm{mi}'] if bs is not None else 0
            running_p = running_p - bs_v + be_v
            y_proj[mi - proj_start] = running_p

    fig_treso.add_trace(go.Scatter(
        x=x_real, y=y_real, mode='lines+markers',
        name='Réalisé',
        line=dict(color='#2E74B5', width=3),
        marker=dict(size=8, color='#1F3864'),
        fill='tozeroy', fillcolor='rgba(46,116,181,0.08)'
    ))

    fig_treso.add_trace(go.Scatter(
        x=x_proj, y=y_proj, mode='lines+markers',
        name='Projection',
        line=dict(color='#C55A11', width=2, dash='dot'),
        marker=dict(size=6, color='#C55A11', symbol='diamond')
    ))

    # Add threshold line
    fig_treso.add_hline(y=400000, line_dash='dash', line_color='red',
                         annotation_text='Seuil de vigilance (400k€)',
                         annotation_position='bottom right',
                         annotation_font_color='red', annotation_font_size=10)

    fig_treso.update_layout(
        title=dict(text='Trésorerie fin de mois (€)', font=dict(size=13, color='#1F3864')),
        height=280, margin=dict(l=10, r=10, t=35, b=30),
        legend=dict(orientation='h', y=-0.15, x=0),
        yaxis=dict(tickformat=',.0f', ticksuffix=' €', gridcolor='#F0F0F0'),
        xaxis=dict(gridcolor='#F0F0F0'),
        plot_bgcolor='white', paper_bgcolor='white',
        hovermode='x unified'
    )
    st.plotly_chart(fig_treso, use_container_width=True)


with col_g2:
    # Flux mensuels (entrées / sorties bar chart)
    ent_vals = [row_ent[f'm{i}'] if row_ent is not None else 0 for i in range(n_real)]
    sor_vals = [abs(row_sor[f'm{i}']) if row_sor is not None else 0 for i in range(n_real)]

    fig_flux = go.Figure()
    fig_flux.add_trace(go.Bar(
        name='Entrées', x=[MONTHS_S[i] for i in range(n_real)], y=ent_vals,
        marker_color='#1E5631', opacity=0.85
    ))
    fig_flux.add_trace(go.Bar(
        name='Sorties', x=[MONTHS_S[i] for i in range(n_real)], y=sor_vals,
        marker_color='#C00000', opacity=0.85
    ))
    fig_flux.update_layout(
        title=dict(text='Flux mensuels réalisés (€)', font=dict(size=13, color='#1F3864')),
        barmode='group', height=280, margin=dict(l=10, r=10, t=35, b=30),
        legend=dict(orientation='h', y=-0.15, x=0),
        yaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0'),
        xaxis=dict(gridcolor='#F0F0F0'),
        plot_bgcolor='white', paper_bgcolor='white',
    )
    st.plotly_chart(fig_flux, use_container_width=True)


# ─── ROW 3: Budget vs Réalisé par catégorie ────────────────────────────────────
st.markdown('<div class="section-header">💰  BUDGET VS RÉALISÉ PAR CATÉGORIE</div>', unsafe_allow_html=True)

col_g3, col_g4 = st.columns([3, 2])

# Build category comparison
KEY_CATS = [
    'Achat Billets', 'Activités de loisirs', 'Aides aux voyages',
    'Frais bancaires', 'Bureaux', 'Téléphone', 'Sorties non',
    'Déplacements', 'Restauration', 'Logiciels', 'Ordinateurs',
    'Sous-traitance', 'Frais postaux', 'Recette de la billetterie',
]

cat_data = []
for frag in KEY_CATS:
    r_r = get_row(real_df, frag)
    r_b = get_row(bdgt_df, frag)
    if r_r is None:
        continue
    realise  = sum(r_r[f'm{i}'] for i in range(n_real))
    budget_n = r_b['total'] if r_b is not None else 0
    ecart    = realise - budget_n
    pct_v    = realise / budget_n if budget_n != 0 else 0
    rag_txt, rag_cls = rag(pct_v, budget_n)
    cat_data.append({
        'label':   r_r['label'],
        'realise': abs(realise),
        'budget':  abs(budget_n),
        'ecart':   ecart,
        'pct':     pct_v,
        'rag':     rag_txt,
        'rag_cls': rag_cls,
    })

cat_data.sort(key=lambda x: x['realise'], reverse=True)

with col_g3:
    if cat_data:
        labels  = [d['label'][:28] for d in cat_data[:12]]
        b_vals  = [d['budget'] for d in cat_data[:12]]
        r_vals  = [d['realise'] for d in cat_data[:12]]

        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            name='Budget N', y=labels, x=b_vals,
            orientation='h', marker_color='#BDD7EE', opacity=0.9
        ))
        fig_bar.add_trace(go.Bar(
            name='Réalisé cumulé', y=labels, x=r_vals,
            orientation='h', marker_color='#2E74B5', opacity=0.9
        ))
        fig_bar.update_layout(
            title=dict(text='Top catégories : Budget annuel vs Réalisé cumulé', font=dict(size=13, color='#1F3864')),
            barmode='overlay', height=380, margin=dict(l=10, r=10, t=35, b=20),
            legend=dict(orientation='h', y=-0.1),
            xaxis=dict(tickformat=',.0f', gridcolor='#F0F0F0'),
            plot_bgcolor='white', paper_bgcolor='white',
        )
        st.plotly_chart(fig_bar, use_container_width=True)

with col_g4:
    # Donut chart: répartition des sorties réalisées
    donut_cats = [d for d in cat_data if d['realise'] > 0][:8]
    if donut_cats:
        fig_donut = go.Figure(go.Pie(
            labels=[d['label'][:20] for d in donut_cats],
            values=[d['realise'] for d in donut_cats],
            hole=0.45,
            marker_colors=px.colors.qualitative.Set2,
            textinfo='percent',
            hovertemplate='<b>%{label}</b><br>%{value:,.0f} €<br>%{percent}<extra></extra>'
        ))
        fig_donut.update_layout(
            title=dict(text='Répartition des dépenses réalisées', font=dict(size=13, color='#1F3864')),
            height=380, margin=dict(l=10, r=10, t=35, b=20),
            legend=dict(font=dict(size=10)),
            paper_bgcolor='white',
            annotations=[dict(text=fmt_eur(sum(d['realise'] for d in donut_cats), True),
                               x=0.5, y=0.5, font_size=13, showarrow=False,
                               font_color='#1F3864')]
        )
        st.plotly_chart(fig_donut, use_container_width=True)


# ─── ROW 4: Tableau détaillé ─────────────────────────────────────────────────
st.markdown('<div class="section-header">📋  TABLEAU DE SUIVI DÉTAILLÉ</div>', unsafe_allow_html=True)

if cat_data:
    # Build display table
    table_rows = []
    for d in cat_data:
        pct_str = f"{d['pct']*100:.1f}%" if d['budget'] > 0 else '—'
        table_rows.append({
            'Catégorie':          d['label'],
            'Budget N (€)':       d['budget'],
            f'Réalisé {n_real}m (€)': d['realise'],
            'Écart (€)':          d['ecart'],
            '% Exec':             pct_str,
            'Statut':             d['rag'],
        })

    table_df = pd.DataFrame(table_rows)

    # Style the dataframe
    def style_row(row):
        styles = [''] * len(row)
        if '🔴' in str(row.get('Statut', '')):
            return [f'background-color: #FFD7D7; color: #C00000'] * len(row)
        if '🟠' in str(row.get('Statut', '')):
            return [f'background-color: #FCE4D6; color: #C55A11'] * len(row)
        if '🟢' in str(row.get('Statut', '')):
            return [f'background-color: #E2EFDA; color: #1E5631'] * len(row)
        return styles

    styled = (table_df.style
              .apply(style_row, axis=1)
              .format({
                  'Budget N (€)':              lambda x: f"{x:,.0f} €" if x > 0 else "—",
                  f'Réalisé {n_real}m (€)':    lambda x: f"{x:,.0f} €".replace(',', ' '),
                  'Écart (€)':                 lambda x: f"{x:+,.0f} €".replace(',', ' '),
              })
              .set_table_styles([
                  {'selector': 'th', 'props': [('background-color', '#1F3864'),
                                                ('color', 'white'),
                                                ('font-weight', 'bold'),
                                                ('font-size', '12px'),
                                                ('padding', '8px')]},
                  {'selector': 'td', 'props': [('font-size', '12px'),
                                                ('padding', '6px 10px')]},
              ]))

    st.dataframe(table_df, use_container_width=True, height=350,
                 column_config={
                     'Budget N (€)':              st.column_config.NumberColumn(format='%.0f €'),
                     f'Réalisé {n_real}m (€)':    st.column_config.NumberColumn(format='%.0f €'),
                     'Écart (€)':                 st.column_config.NumberColumn(format='%+.0f €'),
                 })


# ─── ROW 5: Détail mensuel ─────────────────────────────────────────────────────
if show_details and not real_df.empty:
    st.markdown('<div class="section-header">📅  DÉTAIL MENSUEL PAR CATÉGORIE</div>', unsafe_allow_html=True)

    # Monthly heatmap for top categories
    top_labels = [d['label'] for d in cat_data[:10]]
    heatmap_data = []
    for lbl in top_labels:
        r = get_row(real_df, lbl)
        if r is not None:
            heatmap_data.append([r[f'm{i}'] for i in range(n_real)])

    if heatmap_data:
        fig_heat = go.Figure(go.Heatmap(
            z=heatmap_data,
            x=[MONTHS_S[i] for i in range(n_real)],
            y=top_labels,
            colorscale='Blues',
            hoverongaps=False,
            hovertemplate='<b>%{y}</b><br>%{x}: %{z:,.0f} €<extra></extra>',
            text=[[f"{v:,.0f} €" for v in row] for row in heatmap_data],
            texttemplate='%{text}', textfont_size=9
        ))
        fig_heat.update_layout(
            title=dict(text='Heatmap des dépenses réalisées (€)', font=dict(size=13, color='#1F3864')),
            height=350, margin=dict(l=10, r=10, t=35, b=20),
            xaxis=dict(side='top'),
            paper_bgcolor='white'
        )
        st.plotly_chart(fig_heat, use_container_width=True)


# ─── ROW 6: Alertes ───────────────────────────────────────────────────────────
st.markdown('<div class="section-header">⚠️  ALERTES AUTOMATIQUES</div>', unsafe_allow_html=True)

alerts = []

# Alert 1: Non catégorisé élevé
r_nc = get_row(real_df, 'non cat')
if r_nc is not None:
    nc_cum = sum(r_nc[f'm{i}'] for i in range(n_real))
    if nc_cum > 5000:
        alerts.append(('crit', f'🔴 Sorties non catégorisées : <strong>{fmt_eur(nc_cum)}</strong> cumulés sur {n_real} mois — Ventiler les écritures dans Pennylane'))

# Alert 2: Budget manquant
if bdgt_df is not None and not bdgt_df.empty:
    cats_sans_budget = bdgt_df[(bdgt_df['total'] == 0) & (~bdgt_df['label'].str.contains('TRÉSO|début|fin', case=False, na=False))]
    if len(cats_sans_budget) > 5:
        alerts.append(('warn', f'🟡 {len(cats_sans_budget)} catégories sans budget saisi — Le suivi budgétaire sera incomplet'))

# Alert 3: Trésorerie OK
if treso_actuelle > 400000:
    alerts.append(('ok', f'🟢 Trésorerie confortable : <strong>{fmt_eur(treso_actuelle)}</strong> — Au-dessus du seuil de vigilance (400k€)'))
else:
    alerts.append(('crit', f'🔴 Trésorerie basse : <strong>{fmt_eur(treso_actuelle)}</strong> — En-dessous du seuil de vigilance (400k€)'))

# Alert 4: Achat Billets
r_ab = get_row(real_df, 'Achat Billets')
if r_ab is not None:
    ab_cum = sum(r_ab[f'm{i}'] for i in range(n_real))
    if ab_cum < 0:
        alerts.append(('info', f'🔵 Achat Billets solde net : <strong>{fmt_eur(ab_cum)}</strong> — Remboursements supérieurs aux achats (normal si avances)'))

# Alert 5: Variation trésorerie
if variation < -10000:
    alerts.append(('warn', f'🟡 Trésorerie en baisse de <strong>{fmt_eur(abs(variation))}</strong> depuis le début de l\'exercice'))

# Display alerts
al_cols = st.columns(2)
for i, (level, msg) in enumerate(alerts):
    with al_cols[i % 2]:
        st.markdown(f'<div class="alert-{level}">{msg}</div>', unsafe_allow_html=True)


# ─── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:#AAAAAA;font-size:11px;padding:8px">
  Cockpit CSE LIDL ENTZHEIM · Données source : Pennylane ·
  Outil développé par votre expert-comptable
</div>
""", unsafe_allow_html=True)
