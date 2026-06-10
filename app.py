"""
COCKPIT CSE — v3.0 | Dashboard de pilotage trésorerie
Lecture directe du fichier Pennylane · Design professionnel
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from pathlib import Path
import io
from datetime import date

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Cockpit CSE LIDL ENTZHEIM",
                   page_icon="🏢", layout="wide",
                   initial_sidebar_state="collapsed")

MONTHS   = ["Mars 2026","Avril 2026","Mai 2026","Juin 2026","Juillet 2026",
            "Août 2026","Septembre 2026","Octobre 2026","Novembre 2026",
            "Décembre 2026","Janvier 2027","Février 2027"]
MONTHS_S = ["Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc","Jan","Fév"]
DEFAULT_FILE = Path(__file__).parent / "data.xlsx"

PALETTE = {
    "navy":   "#1a365d", "blue":   "#2b6cb0", "teal":   "#319795",
    "green":  "#276749", "amber":  "#c05621", "red":    "#c53030",
    "purple": "#553c9a", "gray":   "#4a5568", "light":  "#edf2f7",
}

# ══════════════════════════════════════════════════════════════════════════════
#  CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  .block-container {{ padding: 1.5rem 2rem 1rem; max-width: 1400px; }}
  div[data-testid="stSidebar"] {{ background: {PALETTE['navy']}; }}

  /* Hero */
  .hero {{
    background: linear-gradient(135deg, {PALETTE['navy']} 0%, {PALETTE['blue']} 60%, {PALETTE['teal']} 100%);
    border-radius: 16px; padding: 28px 36px; color: white; margin-bottom: 24px;
    box-shadow: 0 8px 32px rgba(26,54,93,.35);
  }}
  .hero h1 {{ margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -.3px; }}
  .hero p  {{ margin: 4px 0 0; opacity: .8; font-size: 13px; }}

  /* KPI */
  .kpi {{
    background: white; border-radius: 14px; padding: 20px 22px;
    box-shadow: 0 2px 12px rgba(0,0,0,.07); border-top: 4px solid {PALETTE['blue']};
    position: relative; overflow: hidden; height: 110px;
  }}
  .kpi.green  {{ border-top-color: {PALETTE['green']}; }}
  .kpi.amber  {{ border-top-color: {PALETTE['amber']}; }}
  .kpi.teal   {{ border-top-color: {PALETTE['teal']}; }}
  .kpi.red    {{ border-top-color: {PALETTE['red']}; }}
  .kpi.purple {{ border-top-color: {PALETTE['purple']}; }}
  .kpi-icon {{ position:absolute; right:16px; top:16px; font-size:28px; opacity:.15; }}
  .kpi-lbl  {{ font-size: 11px; font-weight: 600; color: {PALETTE['gray']};
               text-transform: uppercase; letter-spacing: .6px; margin-bottom: 6px; }}
  .kpi-val  {{ font-size: 28px; font-weight: 800; color: {PALETTE['navy']}; line-height: 1; }}
  .kpi-sub  {{ font-size: 11px; color: {PALETTE['gray']}; margin-top: 4px; }}

  /* Section */
  .sec-title {{
    font-size: 13px; font-weight: 700; color: {PALETTE['navy']};
    text-transform: uppercase; letter-spacing: .8px;
    border-left: 4px solid {PALETTE['blue']}; padding-left: 10px;
    margin: 24px 0 14px;
  }}

  /* Badges RAG */
  .rag {{ display:inline-block; padding:3px 10px; border-radius:20px;
          font-size:11px; font-weight:600; }}
  .rag-green  {{ background:#c6f6d5; color:#276749; }}
  .rag-amber  {{ background:#feebc8; color:#c05621; }}
  .rag-red    {{ background:#fed7d7; color:#c53030; }}
  .rag-blue   {{ background:#bee3f8; color:#2b6cb0; }}
  .rag-gray   {{ background:#e2e8f0; color:#4a5568; }}

  /* Alertes */
  .alert {{ border-radius:10px; padding:12px 16px; margin:6px 0;
            border-left:5px solid; font-size:13px; }}
  .alert-red    {{ background:#fff5f5; border-color:#c53030; color:#742a2a; }}
  .alert-amber  {{ background:#fffaf0; border-color:#c05621; color:#7b341e; }}
  .alert-green  {{ background:#f0fff4; border-color:#276749; color:#1c4532; }}
  .alert-blue   {{ background:#ebf8ff; border-color:#2b6cb0; color:#1a365d; }}

  /* Table */
  .stDataFrame {{ border-radius:12px; overflow:hidden; }}

  /* Divider */
  hr {{ border:none; border-top:1px solid #e2e8f0; margin:20px 0; }}

  /* Upload zone */
  [data-testid="stFileUploader"] {{ background:#f7fafc; border-radius:12px; padding:8px; }}

  /* Progress bar */
  .prog-bg {{ background:#e2e8f0; border-radius:4px; height:8px; margin-top:4px; }}
  .prog-fill {{ height:8px; border-radius:4px; transition:width .5s; }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PARSING ROBUSTE
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_pennylane(file_bytes: bytes) -> dict:
    """
    Lit un fichier Pennylane (Plan de trésorerie ou Cockpit CSE).
    Retourne un dict avec les DataFrames réalisé, budget et métadonnées.
    """
    buf = io.BytesIO(file_bytes)
    xl  = pd.ExcelFile(buf)

    result = {"realise": {}, "budget": {}, "treso": {}, "mois_realises": []}

    # ── Trouver les bons onglets ──────────────────────────────────────────────
    sheet_map = {}
    for s in xl.sheet_names:
        sl = s.lower().replace("é","e").replace("è","e").replace("ê","e")
        if "realise" in sl or ("alis" in sl and "plan" not in sl):
            sheet_map["realise"] = s
        elif "prevision" in sl or "vision" in sl or ("budget" in sl):
            sheet_map["budget"] = s
        elif "plan" in sl and ("treso" in sl or "tresor" in sl):
            sheet_map["plan"] = s

    # ── Lire l'onglet Réalisé ─────────────────────────────────────────────────
    def read_and_parse(sheet_name, data_type="realise"):
        if sheet_name not in xl.sheet_names:
            return
        df = pd.read_excel(buf, sheet_name=sheet_name, header=None)

        # Trouver la ligne d'en-tête (contient des noms de mois)
        hdr_row = 0
        for ri in range(min(6, len(df))):
            row_str = " ".join(str(v) for v in df.iloc[ri] if pd.notna(v))
            if any(m[:4] in row_str for m in MONTHS):
                hdr_row = ri
                break

        # Trouver la colonne des catégories
        cat_col = 0
        for ci in range(min(5, df.shape[1])):
            vals = [v for v in df.iloc[hdr_row+1:hdr_row+10, ci] if pd.notna(v)]
            if len(vals) >= 3 and any(isinstance(v, str) and len(v) > 3 for v in vals):
                cat_col = ci
                break

        # Mapper colonnes → mois
        headers = list(df.iloc[hdr_row])
        month_cols = {}  # month_index -> col_index
        for ci, h in enumerate(headers):
            h_str = str(h) if pd.notna(h) else ""
            for mi, mn in enumerate(MONTHS):
                if mn in h_str:
                    h_lower = h_str.lower()
                    # Priorité : Réalisé pour data_type=realise, Prévision pour budget
                    if data_type == "realise" and "alis" in h_lower:
                        month_cols[mi] = ci
                    elif data_type == "budget" and ("vision" in h_lower or "vision" in h_lower):
                        month_cols[mi] = ci
                    elif mi not in month_cols:
                        month_cols[mi] = ci

        if not month_cols:
            # Fallback : colonnes après cat_col
            for mi in range(min(12, df.shape[1] - cat_col - 1)):
                month_cols[mi] = cat_col + 1 + mi

        # Extraire les données
        for ri in range(hdr_row + 1, len(df)):
            raw = df.iloc[ri, cat_col]
            if not pd.notna(raw) or str(raw).strip() in ("", "nan"):
                continue
            label = str(raw).strip().lstrip()
            vals = {}
            for mi, ci in month_cols.items():
                if ci < df.shape[1]:
                    v = df.iloc[ri, ci]
                    vals[mi] = float(v) if pd.notna(v) and str(v).strip() not in ("", "nan") else 0.0

            # Stocker dans le bon dictionnaire
            if data_type == "realise":
                result["realise"][label] = vals
            else:
                result["budget"][label] = vals

        # Mois réalisés = colonnes avec des données
        if data_type == "realise" and not result["mois_realises"]:
            result["mois_realises"] = sorted(month_cols.keys())

    # Lire chaque onglet disponible
    if "realise" in sheet_map:
        read_and_parse(sheet_map["realise"], "realise")
    if "budget" in sheet_map:
        read_and_parse(sheet_map["budget"], "budget")

    # Si onglet Plan de trésorerie disponible et réalisé vide, l'utiliser
    if not result["realise"] and "plan" in sheet_map:
        read_and_parse(sheet_map["plan"], "realise")

    return result


def val(data_dict: dict, label_fragment: str, month_idx: int,
        data_type: str = "realise") -> float:
    """Cherche une valeur dans le dict par fragment de label."""
    src = data_dict.get(data_type, {})
    for lbl, vals in src.items():
        if label_fragment.lower() in lbl.lower():
            return vals.get(month_idx, 0.0)
    return 0.0

def val_cum(data_dict: dict, label_fragment: str, months: list,
            data_type: str = "realise") -> float:
    return sum(val(data_dict, label_fragment, m, data_type) for m in months)

def fmt(v: float, short=True) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    if short and abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f} M€"
    if short and abs(v) >= 1_000:
        return f"{v/1_000:,.0f} k€".replace(",", " ")
    return f"{v:,.0f} €".replace(",", " ")

def rag_badge(pct: float, budget: float) -> str:
    if budget == 0:
        return '<span class="rag rag-gray">Pas de budget</span>'
    if pct >= 1.15:
        return f'<span class="rag rag-red">🔴 Dépassé ({pct*100:.0f}%)</span>'
    if pct >= 0.90:
        return f'<span class="rag rag-amber">🟠 Attention ({pct*100:.0f}%)</span>'
    if pct >= 0.05:
        return f'<span class="rag rag-green">🟢 OK ({pct*100:.0f}%)</span>'
    return f'<span class="rag rag-blue">🔵 En cours ({pct*100:.0f}%)</span>'


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="padding:20px 16px 10px;text-align:center">
      <div style="font-size:36px">🏢</div>
      <div style="color:white;font-weight:800;font-size:16px;margin-top:8px">CSE LIDL ENTZHEIM</div>
      <div style="color:#90cdf4;font-size:12px;margin-top:4px">Mars 2026 – Février 2027</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    uploaded = st.file_uploader("📂 Fichier Pennylane (.xlsx)",
                                type=["xlsx"],
                                help="Uploader votre export Pennylane ou le fichier Cockpit CSE")

    st.divider()
    n_real = st.slider("📅 Mois réalisés", 1, 12, 3,
                       help="Nombre de mois avec données réelles disponibles")
    mode_elus = st.toggle("👥 Mode présentation élus", value=False)
    st.divider()
    if st.button("🔃 Actualiser"):
        st.cache_data.clear(); st.rerun()

    st.markdown("""
    <div style="color:#718096;font-size:11px;padding:16px 0 0;text-align:center">
      Workflow mensuel :<br>
      Pennylane → Export Excel → Upload ici
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
file_bytes = None
if uploaded:
    file_bytes = uploaded.read()
elif DEFAULT_FILE.exists():
    file_bytes = DEFAULT_FILE.read_bytes()

if not file_bytes:
    st.markdown("""
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
                padding:80px 20px;background:white;border-radius:20px;margin-top:40px;
                box-shadow:0 4px 20px rgba(0,0,0,.08)">
      <div style="font-size:64px">🏢</div>
      <h2 style="color:#1a365d;margin:16px 0 8px;font-weight:800">Cockpit de pilotage trésorerie</h2>
      <p style="color:#718096;font-size:15px;margin:0">CSE LIDL ENTZHEIM · Mars 2026 – Février 2027</p>
      <div style="background:#ebf8ff;border:2px dashed #90cdf4;border-radius:12px;
                  padding:20px 40px;margin-top:24px;text-align:center">
        <p style="color:#2b6cb0;margin:0;font-weight:600">
          ← Ouvrir le menu latéral et uploader votre fichier Pennylane Excel
        </p>
      </div>
    </div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Chargement des données Pennylane..."):
    data = parse_pennylane(file_bytes)

real = data["realise"]
bdgt = data["budget"]
m_range = list(range(n_real))  # mois réalisés

# ── Calculs clés ──────────────────────────────────────────────────────────────
# Trésorerie
treso_debut = 518_635.71  # Mars 2026 début
treso_fin_vals = {}
for lbl, vals in real.items():
    if "fin" in lbl.lower() and "mois" in lbl.lower():
        treso_fin_vals = vals
        break
if not treso_fin_vals:
    for lbl, vals in real.items():
        if "d" in lbl.lower() and "but" in lbl.lower():
            treso_debut = vals.get(0, treso_debut)
            break

treso_act = treso_fin_vals.get(n_real - 1, 470_450.89) if treso_fin_vals else 470_450.89
treso_list = [treso_fin_vals.get(m, 0) for m in range(n_real)]

# Entrées / Sorties
entrees_cum = 0
sorties_cum = 0
for lbl, vals in real.items():
    ll = lbl.lower()
    if ll in ("entrées", "entrees", "entrées"):
        entrees_cum = sum(vals.get(m, 0) for m in m_range)
    if ll in ("sorties",):
        sorties_cum = sum(abs(vals.get(m, 0)) for m in m_range)

# Fallback si totaux non trouvés
if entrees_cum == 0:
    for lbl, vals in real.items():
        if lbl.lower().startswith("entr") and "non" not in lbl.lower() and "remb" not in lbl.lower():
            entrees_cum += sum(vals.get(m, 0) for m in m_range)
            break

if sorties_cum == 0:
    EXCL = {"trésorerie", "tresorerie", "entrées", "entrees", "sorties"}
    for lbl, vals in real.items():
        if lbl.lower() not in EXCL and not lbl.lower().startswith("entr"):
            v = sum(abs(vals.get(m, 0)) for m in m_range)
            sorties_cum += v

bdgt_sorties = 0
for lbl, vals in bdgt.items():
    if "sorti" in lbl.lower():
        bdgt_sorties = sum(abs(vals.get(m, 0)) for m in range(12))
        break

# Projection fin d'exercice
mois_rest = 12 - n_real
flux_net_moyen = (entrees_cum - sorties_cum) / max(n_real, 1)
treso_proj = treso_act + flux_net_moyen * mois_rest


# ══════════════════════════════════════════════════════════════════════════════
#  HERO
# ══════════════════════════════════════════════════════════════════════════════
mode_lbl = "Mode Élus" if mode_elus else "Expert-Comptable"
st.markdown(f"""
<div class="hero">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div>
      <h1>🏢 Cockpit de Pilotage Trésorerie</h1>
      <p>CSE LIDL ENTZHEIM &nbsp;·&nbsp; Exercice Mars 2026 – Février 2027 &nbsp;·&nbsp;
         {n_real} mois réalisés sur 12 &nbsp;·&nbsp; {mode_lbl}</p>
    </div>
    <div style="text-align:right;opacity:.85">
      <div style="font-size:22px;font-weight:800">{date.today().strftime('%d %b %Y')}</div>
      <div style="font-size:12px;margin-top:2px">Dernière mise à jour</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  KPIs
# ══════════════════════════════════════════════════════════════════════════════
variation = treso_act - treso_debut
reste = bdgt_sorties - sorties_cum if bdgt_sorties > 0 else None

k1, k2, k3, k4, k5 = st.columns(5)

def kpi_html(icon, label, value, sub="", variant=""):
    return f"""
    <div class="kpi {variant}">
      <div class="kpi-icon">{icon}</div>
      <div class="kpi-lbl">{label}</div>
      <div class="kpi-val">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>"""

k1.markdown(kpi_html("💰", "Trésorerie actuelle",
    fmt(treso_act), f"Fin {MONTHS_S[n_real-1]} 2026"), unsafe_allow_html=True)

k2.markdown(kpi_html("📈", "Projection fév. 2027",
    fmt(treso_proj),
    f"Flux moyen {fmt(flux_net_moyen)}/mois",
    "teal" if treso_proj > 400_000 else "red"), unsafe_allow_html=True)

k3.markdown(kpi_html("📤", "Dépenses réalisées",
    fmt(sorties_cum), f"Sur {n_real} mois", "amber"), unsafe_allow_html=True)

k4.markdown(kpi_html("📥", "Recettes encaissées",
    fmt(entrees_cum), f"Sur {n_real} mois", "green"), unsafe_allow_html=True)

delta_sign = "▲" if variation >= 0 else "▼"
k5.markdown(kpi_html("📉", "Variation trésorerie",
    f"{delta_sign} {fmt(abs(variation))}",
    f"Depuis début exercice",
    "green" if variation >= 0 else "red"), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  GRAPHIQUES PRINCIPAUX
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-title">📊  Évolution de la trésorerie</div>',
            unsafe_allow_html=True)

cg1, cg2 = st.columns([3, 2])

with cg1:
    # Line chart trésorerie
    x_real = [MONTHS_S[m] for m in range(n_real) if treso_list[m] != 0]
    y_real = [treso_list[m] for m in range(n_real) if treso_list[m] != 0]

    # Projection
    proj_pts = [treso_act]
    running  = treso_act
    for _ in range(mois_rest):
        running += flux_net_moyen
        proj_pts.append(running)
    x_proj = [MONTHS_S[m] for m in range(n_real - 1, 12)]
    y_proj = proj_pts[:len(x_proj)]

    fig = go.Figure()

    # Zone réalisé
    fig.add_trace(go.Scatter(
        x=x_real, y=y_real, mode="lines+markers", name="Réalisé",
        line=dict(color=PALETTE["blue"], width=3),
        marker=dict(size=9, color=PALETTE["navy"], line=dict(color="white", width=2)),
        fill="tozeroy",
        fillcolor=f"rgba(43,108,176,0.07)",
        hovertemplate="<b>%{x}</b><br>Trésorerie : %{y:,.0f} €<extra></extra>"
    ))

    # Projection
    fig.add_trace(go.Scatter(
        x=x_proj, y=y_proj, mode="lines+markers", name="Projection",
        line=dict(color=PALETTE["teal"], width=2.5, dash="dot"),
        marker=dict(size=7, color=PALETTE["teal"], symbol="diamond"),
        hovertemplate="<b>%{x}</b><br>Projection : %{y:,.0f} €<extra></extra>"
    ))

    # Seuil de vigilance
    fig.add_hrect(y0=0, y1=400_000, fillcolor="rgba(197,48,48,.04)",
                  line_width=0, annotation_text="Zone de vigilance",
                  annotation_position="bottom right",
                  annotation_font=dict(size=10, color=PALETTE["red"]))
    fig.add_hline(y=400_000, line=dict(color=PALETTE["red"], dash="dash", width=1.5),
                  annotation_text="Seuil 400k€",
                  annotation_font=dict(size=10, color=PALETTE["red"]),
                  annotation_position="right")

    fig.update_layout(
        height=300, margin=dict(l=0, r=0, t=10, b=0),
        legend=dict(orientation="h", y=-0.2, x=0,
                    font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
        yaxis=dict(tickformat=",.0f", gridcolor="#f0f4f8",
                   title="", ticksuffix=" €", tickfont=dict(size=10)),
        xaxis=dict(gridcolor="#f0f4f8", tickfont=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified", showlegend=True
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

with cg2:
    # Donut répartition dépenses
    CAT_LABELS = {
        "Achat Billets": "🎟 Billetterie",
        "Activités de loisirs": "🎡 Loisirs",
        "Sorties non catégorisées": "❓ Non ventilé",
        "Aides aux voyages": "✈️ Aides voyages",
        "Frais bancaires": "🏦 Frais bancaires",
        "Bureaux": "🖨 Bureaux",
        "Téléphone & Internet": "📱 Téléphone",
        "Restauration": "🍽 Restauration",
        "Déplacements": "🚗 Déplacements",
        "Logiciels & Services Web": "💻 Logiciels",
    }

    donut_vals, donut_lbls = [], []
    for cat_key, cat_label in CAT_LABELS.items():
        v = sum(abs(real.get(k, {}).get(m, 0))
                for k in real
                for m in m_range
                if cat_key.lower()[:8] in k.lower()[:8])
        # simpler: direct lookup
        v2 = 0
        for k, vals in real.items():
            if cat_key.lower()[:10] in k.lower():
                v2 = sum(abs(vals.get(m, 0)) for m in m_range)
                break
        if v2 > 0:
            donut_vals.append(v2)
            donut_lbls.append(cat_label)

    if donut_vals:
        total_d = sum(donut_vals)
        fig_d = go.Figure(go.Pie(
            labels=donut_lbls, values=donut_vals, hole=0.5,
            marker=dict(colors=px.colors.qualitative.Set2,
                        line=dict(color="white", width=2)),
            textinfo="percent", textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>%{value:,.0f} €<br>%{percent}<extra></extra>",
            sort=True
        ))
        fig_d.update_layout(
            height=300, margin=dict(l=0, r=0, t=10, b=0),
            paper_bgcolor="white", showlegend=True,
            legend=dict(font=dict(size=10), orientation="v", x=1, y=0.5),
            annotations=[dict(text=f"<b>{fmt(total_d)}</b>",
                               x=0.5, y=0.5, font=dict(size=13, color=PALETTE["navy"]),
                               showarrow=False)]
        )
        st.plotly_chart(fig_d, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info("Les données de répartition seront affichées après chargement du fichier Pennylane.")


# ══════════════════════════════════════════════════════════════════════════════
#  TABLEAU BUDGET VS RÉALISÉ
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-title">💰  Budget vs Réalisé par catégorie</div>',
            unsafe_allow_html=True)

TRACKED = [
    ("Achat Billets",               "🎟 Billetterie"),
    ("Activités de loisirs",        "🎡 Activités loisirs"),
    ("Activités sportives",         "🏃 Activités sportives"),
    ("Activités culturelles",       "🎭 Activités culturelles"),
    ("Aides aux voyages",           "✈️ Aides aux voyages"),
    ("Aides sociales",              "❤️ Aides sociales"),
    ("Chèques-vacances",            "🏖 Chèques-vacances"),
    ("Cadeaux et bons",             "🎁 Cadeaux & bons"),
    ("Restauration",                "🍽 Restauration"),
    ("Frais bancaires",             "🏦 Frais bancaires"),
    ("Téléphone",                   "📱 Téléphone & Internet"),
    ("Bureaux",                     "🖨 Bureaux & fournitures"),
    ("Déplacements",                "🚗 Déplacements"),
    ("Logiciels",                   "💻 Logiciels & web"),
    ("Formation",                   "📚 Formation des élus"),
    ("Sous-traitance",              "🔧 Sous-traitance"),
    ("Frais postaux",               "📮 Frais postaux"),
    ("Sorties non catégorisées",    "❓ Non catégorisé"),
]

rows = []
for key, label in TRACKED:
    r_val = 0
    b_val = 0
    for k, vals in real.items():
        if key.lower()[:10] in k.lower():
            r_val = sum(abs(vals.get(m, 0)) for m in m_range)
            break
    for k, vals in bdgt.items():
        if key.lower()[:10] in k.lower():
            b_val = sum(abs(vals.get(m, 0)) for m in range(12))
            break
    if r_val == 0 and b_val == 0:
        continue
    pct = r_val / b_val if b_val > 0 else 0
    rows.append({"label": label, "key": key,
                 "budget": b_val, "realise": r_val, "ecart": r_val - b_val, "pct": pct})

rows.sort(key=lambda x: x["realise"], reverse=True)

if rows:
    # Graphique barres horizontales
    cb1, cb2 = st.columns([2, 3])

    with cb1:
        # Tableau avec progress bars
        for r in rows:
            pct_display = min(r["pct"], 1.5)
            bar_color = (PALETTE["red"] if r["pct"] >= 1.15 else
                         PALETTE["amber"] if r["pct"] >= 0.90 else PALETTE["green"])
            bar_pct = int(pct_display / 1.5 * 100)
            rag_txt = ("🔴" if r["pct"] >= 1.15 else "🟠" if r["pct"] >= 0.90 else
                       "🟢" if r["pct"] >= 0.05 else "🔵")

            budget_str = f" / {fmt(r['budget'])}" if r['budget'] > 0 else ""
            st.markdown(f"""
            <div style="margin-bottom:10px">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:2px">
                <span style="font-size:12px;font-weight:600;color:{PALETTE['navy']}">{r['label']}</span>
                <span style="font-size:12px;color:{PALETTE['gray']}">{rag_txt} {fmt(r['realise'])}{budget_str}</span>
              </div>
              <div class="prog-bg">
                <div class="prog-fill" style="width:{bar_pct}%;background:{bar_color}"></div>
              </div>
            </div>""", unsafe_allow_html=True)

    with cb2:
        # Bar chart Budget vs Réalisé
        top = rows[:12]
        fig_b = go.Figure()
        fig_b.add_trace(go.Bar(
            y=[r["label"] for r in top], x=[r["budget"] for r in top],
            name="Budget annuel", orientation="h",
            marker=dict(color=f"rgba(43,108,176,0.25)",
                        line=dict(color=PALETTE["blue"], width=1.5)),
            hovertemplate="%{y}<br>Budget : %{x:,.0f} €<extra></extra>"
        ))
        fig_b.add_trace(go.Bar(
            y=[r["label"] for r in top], x=[r["realise"] for r in top],
            name=f"Réalisé ({n_real} mois)", orientation="h",
            marker=dict(color=[
                PALETTE["red"] if r["pct"] >= 1.15 else
                PALETTE["amber"] if r["pct"] >= 0.90 else
                PALETTE["blue"]
                for r in top]),
            hovertemplate="%{y}<br>Réalisé : %{x:,.0f} €<extra></extra>"
        ))
        fig_b.update_layout(
            barmode="overlay", height=380,
            margin=dict(l=0, r=0, t=10, b=0),
            legend=dict(orientation="h", y=-0.12, font=dict(size=11),
                        bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(tickformat=",.0f", gridcolor="#f0f4f8",
                       ticksuffix=" €", tickfont=dict(size=10)),
            yaxis=dict(tickfont=dict(size=11)),
            plot_bgcolor="white", paper_bgcolor="white"
        )
        st.plotly_chart(fig_b, use_container_width=True,
                        config={"displayModeBar": False})


# ══════════════════════════════════════════════════════════════════════════════
#  TABLEAU MENSUEL
# ══════════════════════════════════════════════════════════════════════════════
if not mode_elus:
    st.markdown('<div class="sec-title">📅  Détail mensuel des dépenses</div>',
                unsafe_allow_html=True)

    monthly_rows = []
    for key, label in TRACKED[:12]:
        row_d = {"Catégorie": label}
        for m in range(n_real):
            v = 0
            for k, vals in real.items():
                if key.lower()[:10] in k.lower():
                    v = abs(vals.get(m, 0))
                    break
            row_d[MONTHS_S[m]] = v if v > 0 else None
        total = sum(row_d.get(MONTHS_S[m], 0) or 0 for m in range(n_real))
        row_d["Total"] = total
        if total > 0:
            monthly_rows.append(row_d)

    if monthly_rows:
        df_monthly = pd.DataFrame(monthly_rows)
        df_monthly = df_monthly.sort_values("Total", ascending=False)

        fmt_cols = {c: st.column_config.NumberColumn(c, format="%.0f €")
                    for c in df_monthly.columns if c != "Catégorie"}
        st.dataframe(df_monthly, use_container_width=True, height=340,
                     column_config=fmt_cols, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PROJECTION
# ══════════════════════════════════════════════════════════════════════════════
if not mode_elus:
    st.markdown('<div class="sec-title">🔮  Projection fin d\'exercice (février 2027)</div>',
                unsafe_allow_html=True)

    cp1, cp2, cp3 = st.columns(3)
    with cp1:
        st.metric("Trésorerie actuelle", fmt(treso_act),
                  delta=f"{flux_net_moyen:+,.0f} €/mois moyen")
    with cp2:
        st.metric("Projection probable", fmt(treso_proj),
                  delta=f"{mois_rest} mois restants",
                  delta_color="off")
    with cp3:
        pct_conso = sorties_cum / bdgt_sorties * 100 if bdgt_sorties > 0 else 0
        st.metric("Budget consommé", f"{pct_conso:.1f}%" if bdgt_sorties > 0 else "—",
                  delta="Budget à saisir" if bdgt_sorties == 0 else f"Sur {fmt(bdgt_sorties)}")

    # Graphique scénarios
    x_sc = [MONTHS_S[m] for m in range(n_real - 1, 12)]
    n_pts = len(x_sc)
    avg_s_m = sorties_cum / n_real if n_real > 0 else 0
    avg_e_m = entrees_cum / n_real if n_real > 0 else 0

    y_opt  = [treso_act + avg_e_m * i        for i in range(n_pts)]
    y_prob = [treso_act + flux_net_moyen * i  for i in range(n_pts)]
    y_pess = [treso_act - avg_s_m * i        for i in range(n_pts)]

    fig_sc = go.Figure()
    fig_sc.add_trace(go.Scatter(x=x_sc, y=y_opt, name="Optimiste",
        line=dict(color=PALETTE["green"], width=2, dash="dot"),
        hovertemplate="%{x} : %{y:,.0f} €<extra>Optimiste</extra>"))
    fig_sc.add_trace(go.Scatter(x=x_sc, y=y_prob, name="Probable",
        line=dict(color=PALETTE["blue"], width=3),
        fill="tonexty", fillcolor="rgba(43,108,176,.06)",
        hovertemplate="%{x} : %{y:,.0f} €<extra>Probable</extra>"))
    fig_sc.add_trace(go.Scatter(x=x_sc, y=y_pess, name="Pessimiste",
        line=dict(color=PALETTE["red"], width=2, dash="dot"),
        fill="tonexty", fillcolor="rgba(197,48,48,.05)",
        hovertemplate="%{x} : %{y:,.0f} €<extra>Pessimiste</extra>"))
    fig_sc.add_hline(y=400_000, line=dict(color=PALETTE["red"], dash="dash", width=1),
                     annotation_text="Seuil vigilance 400k€",
                     annotation_font=dict(size=10, color=PALETTE["red"]))
    fig_sc.update_layout(
        height=240, margin=dict(l=0, r=0, t=10, b=0),
        legend=dict(orientation="h", y=-0.25, font=dict(size=11),
                    bgcolor="rgba(0,0,0,0)"),
        yaxis=dict(tickformat=",.0f", gridcolor="#f0f4f8",
                   ticksuffix=" €", tickfont=dict(size=10)),
        xaxis=dict(tickfont=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white", hovermode="x unified"
    )
    st.plotly_chart(fig_sc, use_container_width=True, config={"displayModeBar": False})


# ══════════════════════════════════════════════════════════════════════════════
#  ALERTES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-title">⚠️  Alertes automatiques</div>',
            unsafe_allow_html=True)

alerts = []

# Trésorerie
if treso_act >= 400_000:
    alerts.append(("green", "🟢",
        f"<b>Trésorerie saine :</b> {fmt(treso_act)} — Au-dessus du seuil de vigilance (400k€)"))
else:
    alerts.append(("red", "🔴",
        f"<b>ALERTE Trésorerie :</b> {fmt(treso_act)} — Sous le seuil de vigilance (400k€) !"))

# Projection
if treso_proj < 400_000:
    alerts.append(("red", "🔴",
        f"<b>Projection à risque :</b> {fmt(treso_proj)} prévu fin février 2027"))
elif treso_proj < 450_000:
    alerts.append(("amber", "🟠",
        f"<b>Projection à surveiller :</b> {fmt(treso_proj)} prévu fin février 2027"))

# Non catégorisé
nc_val = 0
for k, vals in real.items():
    if "non cat" in k.lower():
        nc_val = sum(abs(vals.get(m, 0)) for m in m_range)
        break
if nc_val > 5_000:
    alerts.append(("red", "🔴",
        f"<b>Sorties non catégorisées élevées :</b> {fmt(nc_val)} sur {n_real} mois"
        " — Ventiler les écritures dans Pennylane pour une analyse fiable"))

# Budget
if bdgt_sorties == 0:
    alerts.append(("amber", "🟠",
        "<b>Budget non saisi :</b> Le suivi budgétaire est incomplet — "
        "Renseigner les prévisions dans Pennylane ou dans l'onglet Budget"))
else:
    pct_bgt = sorties_cum / bdgt_sorties
    if pct_bgt > 1.1:
        alerts.append(("red", "🔴",
            f"<b>Dépassement budgétaire global :</b> {pct_bgt*100:.1f}% du budget annuel "
            f"consommé en {n_real} mois ({fmt(sorties_cum)} / {fmt(bdgt_sorties)})"))

# Dépassements par catégorie
for r in rows:
    if r["pct"] >= 1.15 and r["realise"] > 500:
        alerts.append(("amber", "🟠",
            f"<b>{r['label']} :</b> {fmt(r['realise'])} réalisé "
            f"vs {fmt(r['budget'])} budgeté ({r['pct']*100:.0f}%)"))

# Affichage
al1, al2 = st.columns(2)
for i, (level, icon, msg) in enumerate(alerts[:8]):
    with (al1 if i % 2 == 0 else al2):
        st.markdown(f'<div class="alert alert-{level}">{msg}</div>',
                    unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
ex1, ex2, ex3 = st.columns([1, 1, 2])

if rows:
    with ex1:
        csv = pd.DataFrame(rows)[["label","budget","realise","ecart","pct"]] \
                .rename(columns={"label":"Catégorie","budget":"Budget N (€)",
                                  "realise":"Réalisé (€)","ecart":"Écart (€)",
                                  "pct":"% Exec"})
        csv["% Exec"] = (csv["% Exec"] * 100).round(1)
        st.download_button("⬇️ Export CSV",
            csv.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            "cockpit_cse.csv", "text/csv", use_container_width=True)

    with ex2:
        # Rapport HTML
        rows_html = "".join(
            f"<tr style='background:{'#f7fafc' if i%2==0 else 'white'}'>"
            f"<td style='padding:6px 12px;font-size:12px'>{r['label']}</td>"
            f"<td style='text-align:right;padding:6px 12px;font-size:12px'>{fmt(r['budget'], False)}</td>"
            f"<td style='text-align:right;padding:6px 12px;font-size:12px'>{fmt(r['realise'], False)}</td>"
            f"<td style='text-align:right;padding:6px 12px;font-size:12px;"
            f"color:{'#c53030' if r['ecart']>500 else ('green' if r['ecart']<-200 else 'black')}'>"
            f"{fmt(r['ecart'], False)}</td>"
            f"<td style='text-align:center;padding:6px 12px;font-size:12px'>"
            f"{'🔴' if r['pct']>=1.15 else '🟠' if r['pct']>=0.9 else '🟢'}</td></tr>"
            for i, r in enumerate(rows))
        html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<title>Cockpit CSE LIDL ENTZHEIM</title>
<style>body{{font-family:Arial,sans-serif;margin:30px;color:#1a365d}}
h1{{background:#1a365d;color:white;padding:16px 24px;border-radius:8px;font-size:18px;margin:0 0 16px}}
.kpis{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:16px 0}}
.kpi{{background:#ebf8ff;border-radius:8px;padding:12px 16px;border-left:4px solid #2b6cb0}}
.kpi b{{display:block;font-size:18px;color:#1a365d}}
table{{width:100%;border-collapse:collapse;margin-top:16px}}
th{{background:#1a365d;color:white;padding:8px 12px;font-size:11px;text-align:left}}
@media print{{@page{{margin:1.5cm}}}}</style></head><body>
<h1>🏢 Cockpit CSE LIDL ENTZHEIM — Rapport du {date.today().strftime('%d/%m/%Y')}</h1>
<div class="kpis">
  <div class="kpi"><span>Trésorerie actuelle</span><b>{fmt(treso_act, False)}</b></div>
  <div class="kpi"><span>Dépenses réalisées ({n_real} mois)</span><b>{fmt(sorties_cum, False)}</b></div>
  <div class="kpi"><span>Projection fév. 2027</span><b>{fmt(treso_proj, False)}</b></div>
</div>
<table><tr><th>Catégorie</th><th style='text-align:right'>Budget N</th>
<th style='text-align:right'>Réalisé</th><th style='text-align:right'>Écart</th>
<th style='text-align:center'>Statut</th></tr>{rows_html}</table>
<p style='margin-top:24px;font-size:10px;color:#aaa;text-align:center'>
  CSE LIDL ENTZHEIM · Source Pennylane · {date.today().strftime('%d/%m/%Y')}
</p></body></html>"""
        st.download_button("🖨️ Export PDF (HTML)",
            html.encode("utf-8"), "rapport_cse.html",
            "text/html", use_container_width=True,
            help="Ouvrir → Ctrl+P → Enregistrer en PDF")

st.markdown(f"""
<div style="text-align:center;color:#a0aec0;font-size:11px;padding:16px 0 4px">
  Cockpit CSE LIDL ENTZHEIM · v3.0 · Source : Pennylane ·
  Exercice Mars 2026–Fév.2027 · {date.today().strftime('%d/%m/%Y')}
</div>""", unsafe_allow_html=True)
