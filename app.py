"""
app.py — Application RTE Gestion BDD Matériel
Pages : Dashboard · Base de données · Demandes · Import MADU · Administration
"""

import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
import os

from utils import (
    load_data, save_data,
    load_requests, save_requests, add_request,
    apply_modification, apply_addition,
    validate_madu_file, compare_madu, get_madu_summary,
    ALL_COLUMNS, COLUMN_OPTIONS,
)

# ── Charte graphique RTE ──────────────────────────────────────────────────────
RTE_BLUE   = "#009FE3"
RTE_DARK   = "#003F6B"
RTE_LIGHT  = "#E6F5FC"
RTE_GREY   = "#F4F6F8"
RTE_TEXT   = "#1A1A2E"
RTE_WHITE  = "#FFFFFF"
RTE_GREEN  = "#22C55E"
RTE_ORANGE = "#F5A623"
RTE_RED    = "#E8453C"
RTE_TEAL   = "#00C4B4"

PALETTE = [
    "#009FE3", "#003F6B", "#00C4B4", "#F5A623",
    "#7ED321", "#BD10E0", "#E8453C", "#4A90D9", "#50C878", "#FF6B35",
]

st.set_page_config(
    page_title="RTE – Gestion BDD Matériel",
    page_icon="data/RTE_logo.png",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ══════════════════════════════════════════════════════════════════════════════
# CSS GLOBAL
# ══════════════════════════════════════════════════════════════════════════════
def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html,body,[class*="css"]{{font-family:'Inter',sans-serif;color:{RTE_TEXT};}}
    .stApp{{background-color:{RTE_GREY};}}
    .block-container{{padding-top:3rem!important;padding-bottom:2rem!important;}}

    /* ── Sidebar ── */
    [data-testid="stSidebar"]{{
        background:linear-gradient(180deg,{RTE_DARK} 0%,#005A96 100%)!important;
        border-right:none!important;
    }}
    [data-testid="stSidebar"] *{{color:{RTE_WHITE}!important;}}
    [data-testid="stSidebar"] hr{{border-color:rgba(255,255,255,.2)!important;}}
    [data-testid="stSidebar"] .stButton>button{{
        background:transparent!important;
        border:1px solid rgba(255,255,255,.25)!important;
        color:{RTE_WHITE}!important;border-radius:8px!important;
        font-weight:500!important;transition:all .2s ease!important;
        text-align:left!important;
    }}
    [data-testid="stSidebar"] .stButton>button:hover{{
        background:rgba(255,255,255,.15)!important;
        border-color:rgba(255,255,255,.5)!important;
    }}
    [data-testid="stSidebar"] .stButton>button[kind="primary"]{{
        background:{RTE_BLUE}!important;border-color:{RTE_BLUE}!important;font-weight:700!important;
    }}

    /* ── En-tête de page ── */
    .rte-header{{
        background:linear-gradient(135deg,{RTE_DARK} 0%,{RTE_BLUE} 100%);
        padding:1.3rem 2rem;border-radius:12px;margin-bottom:1.3rem;
        display:flex;align-items:center;gap:1.2rem;
        box-shadow:0 4px 15px rgba(0,159,227,.25);
    }}
    .rte-header h1{{color:white!important;font-size:1.5rem!important;font-weight:700!important;margin:0!important;padding:0!important;}}
    .rte-header p{{color:rgba(255,255,255,.82)!important;margin:.2rem 0 0!important;font-size:.87rem!important;}}

    /* ── Cartes KPI ── */
    .kpi-card{{
        background:white;border-radius:12px;padding:1.1rem 1.3rem;
        box-shadow:0 2px 10px rgba(0,0,0,.07);border-left:4px solid {RTE_BLUE};
        transition:transform .2s ease,box-shadow .2s ease;height:100%;
    }}
    .kpi-card:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,159,227,.18);}}
    .kpi-label{{font-size:.74rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B7280;margin-bottom:.3rem;}}
    .kpi-value{{font-size:1.85rem;font-weight:700;color:{RTE_DARK};line-height:1.1;}}
    .kpi-sub{{font-size:.74rem;color:#9CA3AF;margin-top:.15rem;}}

    /* ── Titre de section ── */
    .section-title{{
        font-size:1rem;font-weight:600;color:{RTE_DARK};
        padding-bottom:.35rem;border-bottom:2px solid {RTE_BLUE};margin-bottom:.75rem;
    }}

    /* ── Barre de filtres ── */
    .filter-bar{{
        background:white;border-radius:10px;padding:.9rem 1.2rem;
        box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:1rem;
        border-left:4px solid {RTE_TEAL};
    }}

    /* ── Boutons principaux ── */
    .stButton>button[kind="primary"]{{
        background:{RTE_BLUE}!important;border:none!important;border-radius:8px!important;
        font-weight:600!important;box-shadow:0 2px 8px rgba(0,159,227,.35)!important;
        transition:all .2s ease!important;
    }}
    .stButton>button[kind="primary"]:hover{{background:{RTE_DARK}!important;}}
    .stDownloadButton>button{{
        background:white!important;color:{RTE_BLUE}!important;
        border:2px solid {RTE_BLUE}!important;border-radius:8px!important;
        font-weight:600!important;transition:all .2s ease!important;
    }}
    .stDownloadButton>button:hover{{background:{RTE_BLUE}!important;color:white!important;}}

    /* ── Cartes de type de demande ── */
    .type-card{{
        border:2px solid #E5E7EB;border-radius:12px;padding:1.1rem 1.3rem;
        cursor:pointer;transition:all .2s ease;background:white;
        text-align:center;height:100%;
    }}
    .type-card:hover{{border-color:{RTE_BLUE};box-shadow:0 4px 12px rgba(0,159,227,.15);}}
    .type-card.selected{{border-color:{RTE_BLUE};background:{RTE_LIGHT};}}

    /* ── Badges de statut ── */
    .badge{{display:inline-block;padding:.2rem .65rem;border-radius:999px;font-size:.74rem;font-weight:600;}}
    .badge-wait{{background:#FEF3C7;color:#92400E;}}
    .badge-ok{{background:#D1FAE5;color:#065F46;}}
    .badge-ko{{background:#FEE2E2;color:#991B1B;}}

    /* ── Indicateur d'étapes ── */
    .steps{{display:flex;gap:.5rem;margin-bottom:1.2rem;align-items:center;}}
    .step{{display:flex;align-items:center;gap:.4rem;}}
    .step-num{{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.78rem;font-weight:700;}}
    .step-num.active{{background:{RTE_BLUE};color:white;}}
    .step-num.done{{background:{RTE_GREEN};color:white;}}
    .step-num.idle{{background:#E5E7EB;color:#6B7280;}}
    .step-label{{font-size:.82rem;font-weight:500;}}
    .step-label.active{{color:{RTE_BLUE};font-weight:700;}}
    .step-label.idle{{color:#9CA3AF;}}
    .step-sep{{flex:1;height:2px;background:#E5E7EB;}}
    .step-sep.done{{background:{RTE_GREEN};}}

    /* ── Widgets Streamlit ── */
    [data-testid="stExpander"]{{border:1px solid #E5E7EB!important;border-radius:10px!important;overflow:hidden;}}
    [data-testid="stExpander"] summary{{background:{RTE_GREY}!important;font-weight:600!important;color:{RTE_DARK}!important;}}
    [data-testid="stMetric"]{{background:white;border-radius:10px;padding:.8rem 1rem;box-shadow:0 2px 8px rgba(0,0,0,.06);border-left:3px solid {RTE_BLUE};}}
    [data-testid="stMetricLabel"]{{color:#6B7280!important;font-size:.82rem!important;}}
    [data-testid="stMetricValue"]{{color:{RTE_DARK}!important;font-weight:700!important;}}
    [data-testid="stDataFrame"]{{border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06);}}

    /* ── Zone d'upload MADU ── */
    .madu-upload-zone{{
        background:white;border:2px dashed {RTE_BLUE};border-radius:14px;
        padding:2rem;text-align:center;
        box-shadow:0 2px 10px rgba(0,159,227,.08);
    }}

    /* ── Tableau de comparaison MADU ── */
    .madu-avancee{{background:#D1FAE5;color:#065F46;padding:.15rem .5rem;border-radius:6px;font-weight:600;font-size:.82rem;}}
    .madu-retardee{{background:#FEE2E2;color:#991B1B;padding:.15rem .5rem;border-radius:6px;font-weight:600;font-size:.82rem;}}
    .madu-inchangee{{background:#F3F4F6;color:#6B7280;padding:.15rem .5rem;border-radius:6px;font-weight:600;font-size:.82rem;}}
    .madu-notfound{{background:#FEF3C7;color:#92400E;padding:.15rem .5rem;border-radius:6px;font-weight:600;font-size:.82rem;}}
    </style>
    """, unsafe_allow_html=True)


inject_css()


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS UI
# ══════════════════════════════════════════════════════════════════════════════

def get_logo_b64() -> str | None:
    p = "data/RTE_logo.png"
    if os.path.exists(p):
        with open(p, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


LOGO_B64 = get_logo_b64()


def page_header(title: str, subtitle: str = ""):
    logo = (
        f'<img src="data:image/png;base64,{LOGO_B64}" style="height:50px;border-radius:8px;"/>'
        if LOGO_B64 else ""
    )
    st.markdown(f"""
    <div class="rte-header">{logo}
      <div>
        <h1>{title}</h1>
        {"<p>" + subtitle + "</p>" if subtitle else ""}
      </div>
    </div>""", unsafe_allow_html=True)


def kpi_card(label: str, value, sub: str = "", accent: str = RTE_BLUE):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{accent}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
    </div>""", unsafe_allow_html=True)


def section_title(text: str):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def alert_box(text: str, color: str = RTE_LIGHT, border: str = RTE_BLUE):
    st.markdown(
        f'<div style="background:{color};border-left:4px solid {border};'
        f'border-radius:6px;padding:.65rem 1rem;margin-bottom:.8rem;">{text}</div>',
        unsafe_allow_html=True,
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    if LOGO_B64:
        st.markdown(
            f'<div style="text-align:center;padding:1rem 0 .5rem;">'
            f'<img src="data:image/png;base64,{LOGO_B64}" style="width:76px;border-radius:50%;"/>'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div style="text-align:center;font-size:.93rem;font-weight:700;'
        'letter-spacing:.08em;padding-bottom:.2rem;">BDD MATÉRIEL</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<hr style="margin:.4rem 0 1rem;"/>', unsafe_allow_html=True)

    PAGES = {
        "📊  Dashboard":            "dashboard",
        "🗄️  Base de données":      "bdd",
        "📝  Demandes":             "demandes",
        "📅  Import MADU":          "madu",
        "🔐  Administration":       "admin",
    }
    if "page" not in st.session_state:
        st.session_state.page = "dashboard"

    for label, key in PAGES.items():
        if st.button(
            label,
            use_container_width=True,
            type="primary" if st.session_state.page == key else "secondary",
            key=f"nav_{key}",
        ):
            st.session_state.page = key
            st.rerun()

    st.markdown('<hr style="margin:1rem 0 .5rem;"/>', unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.69rem;color:rgba(255,255,255,.5);text-align:center;">'
        "Réseau de Transport d'Électricité<br/>Gestion des équipements réseau</div>",
        unsafe_allow_html=True,
    )


# ── Paramètres des graphiques ─────────────────────────────────────────────────
CHART = dict(
    plot_bgcolor="white",
    paper_bgcolor="white",
    font=dict(family="Inter, sans-serif", color=RTE_TEXT),
    margin=dict(t=30, b=8, l=8, r=8),
)


@st.cache_data(ttl=30)
def get_data() -> pd.DataFrame:
    return load_data()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : DASHBOARD INTERACTIF
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "dashboard":
    df_all = get_data()
    page_header(
        "Tableau de bord",
        "Vue d'ensemble interactive — cliquez sur les graphiques pour filtrer",
    )

    # ── Filtres globaux ───────────────────────────────────────────────────────
    with st.container():
        st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)
        with fc1:
            f_crpt = st.selectbox(
                "🏙️ Site (CRPT)",
                ["Tous"] + sorted(df_all["CRPT"].dropna().unique().tolist()),
                key="db_crpt",
            )
        with fc2:
            f_prog = st.selectbox(
                "🏭 Programme",
                ["Tous"] + sorted(df_all["Programme industrielle"].dropna().unique().tolist()),
                key="db_prog",
            )
        with fc3:
            f_bis = st.selectbox(
                "🔩 Type BIS",
                ["Tous"] + sorted(df_all["Type de BIS"].dropna().unique().tolist()),
                key="db_bis",
            )
        with fc4:
            f_sp = st.selectbox(
                "📋 Sous-politique",
                ["Tous"] + sorted(df_all["Sous-politique"].dropna().unique().tolist()),
                key="db_sp",
            )
        with fc5:
            f_plqf = st.selectbox(
                "🏢 PLQF",
                ["Tous"] + sorted(df_all["Affectation PLQF"].dropna().unique().tolist()),
                key="db_plqf",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Application des filtres ───────────────────────────────────────────────
    df = df_all.copy()
    if f_crpt != "Tous":  df = df[df["CRPT"] == f_crpt]
    if f_prog != "Tous":  df = df[df["Programme industrielle"] == f_prog]
    if f_bis  != "Tous":  df = df[df["Type de BIS"] == f_bis]
    if f_sp   != "Tous":  df = df[df["Sous-politique"] == f_sp]
    if f_plqf != "Tous":  df = df[df["Affectation PLQF"] == f_plqf]

    n_total = len(df)
    n_all   = len(df_all)
    is_filt = n_total < n_all

    # ── KPIs principaux ───────────────────────────────────────────────────────
    assigned = df["Affectation PLQF"].notna().sum()
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card(
            "Demandes affichées", n_total,
            sub=f"sur {n_all} au total" if is_filt else "dans la base",
            accent=RTE_BLUE,
        )
    with k2:
        kpi_card("Sites CRPT", df["CRPT"].nunique(), accent=RTE_DARK)
    with k3:
        kpi_card("Types de BIS", df["Type de BIS"].nunique(), accent=RTE_TEAL)
    with k4:
        kpi_card(
            "Affectées PLQF", assigned,
            sub=f"{int(100 * assigned / n_total) if n_total else 0} % du total",
            accent=RTE_ORANGE,
        )

    if is_filt:
        alert_box(
            f"🔍 Filtre actif — <strong>{n_total}</strong> demande(s) sur "
            f"<strong>{n_all}</strong> affichées.",
            color=RTE_LIGHT, border=RTE_BLUE,
        )

    # ── Bloc MADU : affiché uniquement si une comparaison a été chargée ───────
    madu_compare = st.session_state.get("madu_compare_result")
    if madu_compare is not None:
        st.markdown("<br/>", unsafe_allow_html=True)
        section_title("📅 Analyse des écarts de dates MADU")

        summary = get_madu_summary(madu_compare)

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            kpi_card("Projets comparés",   summary["trouves"],
                     sub=f"{summary['non_trouves']} non trouvé(s)", accent=RTE_BLUE)
        with m2:
            kpi_card("Dates avancées ✅",  summary["avancees"],
                     sub="date rapprochée", accent=RTE_GREEN)
        with m3:
            kpi_card("Dates retardées ⚠️", summary["retardees"],
                     sub="date repoussée",  accent=RTE_RED)
        with m4:
            kpi_card("Inchangées",          summary["inchangees"],
                     sub="aucun écart",      accent=RTE_TEAL)
        with m5:
            kpi_card("Écart moyen",
                     f"{summary['ecart_moyen_jours']:+.0f} j",
                     sub=f"min {summary['ecart_min_jours']:+d} j  /  max {summary['ecart_max_jours']:+d} j",
                     accent=RTE_ORANGE)

        st.markdown("<br/>", unsafe_allow_html=True)
        col_chart1, col_chart2 = st.columns([2, 1])

        with col_chart1:
            section_title("📊 Distribution des écarts de dates MADU (jours)")
            df_ecarts = madu_compare[madu_compare["Statut"].isin(["Avancée","Retardée","Inchangée"])].copy()
            if not df_ecarts.empty:
                color_map = {
                    "Avancée":   RTE_GREEN,
                    "Retardée":  RTE_RED,
                    "Inchangée": "#9CA3AF",
                }
                fig_ecart = px.bar(
                    df_ecarts.sort_values("Écart (jours)"),
                    x="EOTP2",
                    y="Écart (jours)",
                    color="Statut",
                    color_discrete_map=color_map,
                    hover_data=["Projet (import)", "Ancienne date MADU", "Nouvelle date MADU"],
                    text="Écart (jours)",
                )
                fig_ecart.add_hline(y=0, line_dash="dash", line_color="#6B7280", line_width=1)
                fig_ecart.update_layout(
                    height=330, **CHART,
                    xaxis_title="EOTP2", yaxis_title="Écart (jours)",
                    xaxis_tickangle=-30, showlegend=True,
                )
                fig_ecart.update_traces(textposition="outside", marker_line_width=0)
                st.plotly_chart(fig_ecart, use_container_width=True)
            else:
                st.info("Aucun projet avec écart à afficher.")

        with col_chart2:
            section_title("🥧 Répartition des statuts")
            df_pie = madu_compare["Statut"].value_counts().reset_index()
            df_pie.columns = ["Statut", "Nombre"]
            color_pie = {
                "Avancée":               RTE_GREEN,
                "Retardée":              RTE_RED,
                "Inchangée":             "#9CA3AF",
                "Non trouvé dans la BDD": RTE_ORANGE,
                "Date manquante":        "#E5E7EB",
            }
            fig_pie = px.pie(
                df_pie, names="Statut", values="Nombre",
                color="Statut", color_discrete_map=color_pie,
                hole=0.5,
            )
            fig_pie.update_layout(height=330, **CHART)
            fig_pie.update_traces(
                textinfo="percent+label",
                marker=dict(line=dict(color="white", width=2)),
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        # ── Tableau détaillé MADU ─────────────────────────────────────────────
        section_title("📋 Tableau détaillé des écarts MADU")

        tab_f1, tab_f2 = st.columns([3, 1])
        with tab_f1:
            madu_filter = st.selectbox(
                "Filtrer par statut",
                ["Tous", "Avancée", "Retardée", "Inchangée", "Non trouvé dans la BDD"],
                key="madu_dash_filter",
            )
        with tab_f2:
            show_mod_only = st.checkbox("Modifiés uniquement", value=False, key="madu_mod_only")

        df_show = madu_compare.copy()
        if madu_filter != "Tous":
            df_show = df_show[df_show["Statut"] == madu_filter]
        if show_mod_only:
            df_show = df_show[df_show["Statut"].isin(["Avancée", "Retardée"])]

        # Mise en forme des colonnes de dates
        for dc in ["Ancienne date MADU", "Nouvelle date MADU"]:
            df_show[dc] = pd.to_datetime(df_show[dc], errors="coerce").dt.strftime("%d/%m/%Y")

        # Coloration ligne par ligne via style
        def _color_row(row):
            colors = {
                "Avancée":               "background-color:#D1FAE5",
                "Retardée":              "background-color:#FEE2E2",
                "Inchangée":             "",
                "Non trouvé dans la BDD": "background-color:#FEF3C7",
            }
            c = colors.get(row["Statut"], "")
            return [c] * len(row)

        styled = df_show.reset_index(drop=True).style.apply(_color_row, axis=1)
        st.dataframe(styled, use_container_width=True, height=380)

        # Export du tableau MADU
        dl1, dl2, _ = st.columns([1, 1, 3])
        with dl1:
            st.download_button(
                "⬇️ CSV",
                data=df_show.to_csv(index=False, sep=";", encoding="utf-8-sig"),
                file_name="analyse_madu.csv", mime="text/csv",
                use_container_width=True, key="dl_madu_csv_dash",
            )
        with dl2:
            buf = io.BytesIO()
            df_show.to_excel(buf, index=False)
            st.download_button(
                "⬇️ Excel", data=buf.getvalue(),
                file_name="analyse_madu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_madu_xlsx_dash",
            )

        alert_box(
            '💡 Ces données proviennent du dernier fichier importé via la page '
            '<strong>📅 Import MADU</strong>. Pour mettre à jour, importez un nouveau fichier.',
            color="#F0FDF4", border=RTE_GREEN,
        )

    st.markdown("<br/>", unsafe_allow_html=True)

    # ── Graphiques standards ──────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        section_title("📍 Demandes par site (CRPT)")
        c = df["CRPT"].value_counts().reset_index()
        c.columns = ["CRPT", "Nombre"]
        fig = px.bar(
            c, x="Nombre", y="CRPT", orientation="h",
            color="CRPT", color_discrete_sequence=PALETTE, text="Nombre",
        )
        fig.update_layout(
            showlegend=False, height=310, **CHART,
            yaxis_title="", xaxis_title="Nombre de demandes",
            clickmode="event+select",
        )
        fig.update_traces(
            textposition="outside", marker_line_width=0,
            selected=dict(marker=dict(opacity=1)),
            unselected=dict(marker=dict(opacity=0.4)),
        )
        ev1 = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="chart_crpt")
        if ev1 and ev1.get("selection") and ev1["selection"].get("points"):
            val = ev1["selection"]["points"][0]["label"]
            if f_crpt != val:
                st.session_state["db_crpt"] = val
                st.rerun()

    with col2:
        section_title("🏭 Répartition par programme industriel")
        p = df["Programme industrielle"].value_counts().reset_index()
        p.columns = ["Programme", "Nombre"]
        fig2 = px.pie(p, names="Programme", values="Nombre",
                      color_discrete_sequence=PALETTE, hole=0.45)
        fig2.update_layout(height=310, **CHART)
        fig2.update_traces(
            textinfo="percent+label", textfont_size=12,
            marker=dict(line=dict(color="white", width=2)),
            pull=[0.03] * len(p),
        )
        ev2 = st.plotly_chart(fig2, use_container_width=True, on_select="rerun", key="chart_prog")
        if ev2 and ev2.get("selection") and ev2["selection"].get("points"):
            val = ev2["selection"]["points"][0]["label"]
            if f_prog != val:
                st.session_state["db_prog"] = val
                st.rerun()

    col3, col4 = st.columns(2)

    with col3:
        section_title("🔩 Demandes par type de BIS")
        b = df["Type de BIS"].value_counts().reset_index()
        b.columns = ["Type BIS", "Nombre"]
        fig3 = px.bar(b, x="Type BIS", y="Nombre",
                      color="Type BIS", color_discrete_sequence=PALETTE, text="Nombre")
        fig3.update_layout(
            showlegend=False, height=310, **CHART,
            xaxis_title="", yaxis_title="Nombre", xaxis_tickangle=-30,
            clickmode="event+select",
        )
        fig3.update_traces(
            textposition="outside", marker_line_width=0,
            selected=dict(marker=dict(opacity=1)),
            unselected=dict(marker=dict(opacity=0.4)),
        )
        ev3 = st.plotly_chart(fig3, use_container_width=True, on_select="rerun", key="chart_bis")
        if ev3 and ev3.get("selection") and ev3["selection"].get("points"):
            val = ev3["selection"]["points"][0]["x"]
            if f_bis != val:
                st.session_state["db_bis"] = val
                st.rerun()

    with col4:
        section_title("📋 Répartition par sous-politique")
        s = df["Sous-politique"].value_counts().reset_index()
        s.columns = ["Sous-politique", "Nombre"]
        fig4 = px.bar(s, x="Sous-politique", y="Nombre",
                      color="Sous-politique", color_discrete_sequence=PALETTE, text="Nombre")
        fig4.update_layout(
            showlegend=False, height=310, **CHART,
            xaxis_title="", yaxis_title="Nombre", xaxis_tickangle=-30,
            clickmode="event+select",
        )
        fig4.update_traces(
            textposition="outside", marker_line_width=0,
            selected=dict(marker=dict(opacity=1)),
            unselected=dict(marker=dict(opacity=0.4)),
        )
        ev4 = st.plotly_chart(fig4, use_container_width=True, on_select="rerun", key="chart_sp")
        if ev4 and ev4.get("selection") and ev4["selection"].get("points"):
            val = ev4["selection"]["points"][0]["x"]
            if f_sp != val:
                st.session_state["db_sp"] = val
                st.rerun()

    col5, col6 = st.columns([2, 1])

    with col5:
        section_title("📈 Évolution des mises en service (MES PFM1)")
        df_mes = df.dropna(subset=["MES PFM1"]).copy()
        if len(df_mes):
            df_mes["Année-Mois"] = df_mes["MES PFM1"].dt.to_period("M").astype(str)
            ts = (df_mes.groupby("Année-Mois")
                        .size()
                        .reset_index(name="Nombre")
                        .sort_values("Année-Mois"))
            fig5 = go.Figure()
            fig5.add_trace(go.Scatter(
                x=ts["Année-Mois"], y=ts["Nombre"],
                mode="lines+markers",
                line=dict(color=RTE_BLUE, width=2.5),
                marker=dict(color=RTE_BLUE, size=7, line=dict(color="white", width=1.5)),
                fill="tozeroy", fillcolor="rgba(0,159,227,0.12)",
                hovertemplate="<b>%{x}</b><br>%{y} MES<extra></extra>",
            ))
            fig5.update_layout(
                height=310, **CHART,
                xaxis_title="", yaxis_title="Mises en service",
                xaxis=dict(tickangle=-45), hovermode="x unified",
            )
            st.plotly_chart(fig5, use_container_width=True)
        else:
            st.info("Aucune donnée MES PFM1 pour la sélection.")

    with col6:
        section_title("💧 Huile / Air")
        dh = df.copy()
        dh["Huile/Air"] = dh["Huile/Air"].str.lower().str.strip()
        ha = dh["Huile/Air"].value_counts().reset_index()
        ha.columns = ["Type", "Nombre"]
        fig6 = px.pie(ha, names="Type", values="Nombre",
                      color_discrete_sequence=[RTE_BLUE, RTE_DARK], hole=0.52)
        fig6.update_layout(height=310, **CHART)
        fig6.update_traces(
            textinfo="percent+label",
            marker=dict(line=dict(color="white", width=2)),
        )
        st.plotly_chart(fig6, use_container_width=True)

    section_title("🏢 Affectation PLQF — répartition par fournisseur")
    plqf = df["Affectation PLQF"].fillna("Non affecté").value_counts().reset_index()
    plqf.columns = ["Fournisseur", "Nombre"]
    fig7 = px.bar(plqf, x="Fournisseur", y="Nombre",
                  color="Fournisseur", color_discrete_sequence=PALETTE, text="Nombre")
    fig7.update_layout(
        showlegend=False, height=320, **CHART,
        xaxis_title="", yaxis_title="Nombre de demandes",
        xaxis_tickangle=-30, clickmode="event+select",
    )
    fig7.update_traces(
        textposition="outside", marker_line_width=0,
        selected=dict(marker=dict(opacity=1)),
        unselected=dict(marker=dict(opacity=0.4)),
    )
    ev7 = st.plotly_chart(fig7, use_container_width=True, on_select="rerun", key="chart_plqf")
    if ev7 and ev7.get("selection") and ev7["selection"].get("points"):
        val = ev7["selection"]["points"][0]["x"]
        if val != "Non affecté" and f_plqf != val:
            st.session_state["db_plqf"] = val
            st.rerun()

    if is_filt:
        st.markdown("<br/>", unsafe_allow_html=True)
        if st.button("🔄 Réinitialiser tous les filtres"):
            for k in ["db_crpt", "db_prog", "db_bis", "db_sp", "db_plqf"]:
                st.session_state[k] = "Tous"
            st.rerun()

    with st.expander(f"📋 Voir les {n_total} entrée(s) correspondantes", expanded=False):
        st.dataframe(
            df.drop(columns=["__id__"]).reset_index(drop=True),
            use_container_width=True, height=380,
        )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : BASE DE DONNÉES
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "bdd":
    df = get_data()
    page_header("Base de données", f"{len(df)} entrées · Filtrage avancé et export")

    with st.expander("🔍 Filtres avancés", expanded=True):
        f1, f2, f3 = st.columns(3)
        with f1:
            sel_crpt = st.selectbox("Site (CRPT)",
                ["Tous"] + sorted(df["CRPT"].dropna().unique().tolist()))
        with f2:
            sel_prog = st.selectbox("Programme industriel",
                ["Tous"] + sorted(df["Programme industrielle"].dropna().unique().tolist()))
        with f3:
            sel_bis = st.selectbox("Type de BIS",
                ["Tous"] + sorted(df["Type de BIS"].dropna().unique().tolist()))

        f4, f5, f6 = st.columns(3)
        with f4:
            sel_sp = st.selectbox("Sous-politique",
                ["Tous"] + sorted(df["Sous-politique"].dropna().unique().tolist()))
        with f5:
            sel_ha = st.selectbox("Huile/Air",
                ["Tous"] + sorted(df["Huile/Air"].dropna().str.lower().unique().tolist()))
        with f6:
            sel_plqf = st.selectbox("Affectation PLQF",
                ["Tous"] + sorted(df["Affectation PLQF"].dropna().unique().tolist()))

        search = st.text_input("🔎 Recherche libre (nom de projet, poste, RUO…)")
        date_range = None
        mes_min = df["MES PFM1"].dropna().min()
        mes_max = df["MES PFM1"].dropna().max()
        if pd.notna(mes_min) and pd.notna(mes_max):
            date_range = st.date_input(
                "Plage de dates MES PFM1",
                value=(mes_min.date(), mes_max.date()),
                min_value=mes_min.date(), max_value=mes_max.date(),
            )

    mask = pd.Series([True] * len(df), index=df.index)
    if sel_crpt  != "Tous": mask &= df["CRPT"] == sel_crpt
    if sel_prog  != "Tous": mask &= df["Programme industrielle"] == sel_prog
    if sel_bis   != "Tous": mask &= df["Type de BIS"] == sel_bis
    if sel_sp    != "Tous": mask &= df["Sous-politique"] == sel_sp
    if sel_ha    != "Tous": mask &= df["Huile/Air"].str.lower() == sel_ha
    if sel_plqf  != "Tous": mask &= df["Affectation PLQF"] == sel_plqf
    if search:
        mask &= df.apply(lambda r: search.lower() in " ".join(r.astype(str)).lower(), axis=1)
    if date_range and len(date_range) == 2:
        d0, d1 = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
        mask &= df["MES PFM1"].isna() | ((df["MES PFM1"] >= d0) & (df["MES PFM1"] <= d1))

    df_filtered = df[mask].drop(columns=["__id__"])
    st.markdown(
        f'<div style="background:{RTE_LIGHT};border-left:4px solid {RTE_BLUE};'
        f'border-radius:6px;padding:.55rem 1rem;margin-bottom:.8rem;font-weight:600;">'
        f'🔎 {len(df_filtered)} ligne(s) correspondante(s)</div>',
        unsafe_allow_html=True,
    )
    st.dataframe(df_filtered.reset_index(drop=True), use_container_width=True, height=480)

    st.markdown("---")
    section_title("📥 Exporter les données filtrées")
    ex1, ex2 = st.columns(2)
    with ex1:
        st.download_button(
            "⬇️ CSV",
            data=df_filtered.to_csv(index=False, sep=";", encoding="utf-8-sig"),
            file_name="export_bdd.csv", mime="text/csv",
            use_container_width=True,
        )
    with ex2:
        buf = io.BytesIO()
        df_filtered.to_excel(buf, index=False)
        st.download_button(
            "⬇️ Excel", data=buf.getvalue(),
            file_name="export_bdd.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : DEMANDES
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "demandes":
    df = get_data()
    page_header(
        "Demandes",
        "Modifier, supprimer ou ajouter une entrée · Validation admin requise",
    )

    # ── Initialisation de l'état du formulaire ────────────────────────────────
    for k, v in [("dem_step",1),("dem_type",None),("dem_id",None),
                 ("dem_changes",{}),("dem_raison",""),("dem_new_vals",{})]:
        if k not in st.session_state:
            st.session_state[k] = v

    step  = st.session_state.dem_step
    dtype = st.session_state.dem_type

    # ── Indicateur d'étapes ───────────────────────────────────────────────────
    labels = ["1 · Type", "2 · Sélection", "3 · Détails", "4 · Confirmation"]
    html = '<div class="steps">'
    for i, lbl in enumerate(labels, 1):
        if   i < step:  nc, lc = "done",   "active"
        elif i == step: nc, lc = "active",  "active"
        else:           nc, lc = "idle",    "idle"
        icon = "✓" if i < step else str(i)
        html += (f'<div class="step"><div class="step-num {nc}">{icon}</div>'
                 f'<span class="step-label {lc}">{lbl}</span></div>')
        if i < len(labels):
            html += f'<div class="step-sep {"done" if i < step else ""}"></div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

    # ── Étape 1 : Choisir le type ─────────────────────────────────────────────
    if step == 1:
        section_title("Quelle action souhaitez-vous effectuer ?")
        st.markdown("<br/>", unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        for col, typ, icon, lbl, desc in [
            (t1, "Modification", "✏️", "Modifier une ligne",
             "Corriger ou mettre à jour les informations d'une entrée existante"),
            (t2, "Suppression",  "🗑️", "Supprimer une ligne",
             "Demander la suppression d'une entrée de la base de données"),
            (t3, "Ajout",        "➕", "Ajouter une ligne",
             "Créer une nouvelle entrée dans la base de données"),
        ]:
            with col:
                if st.button(f"{icon}  {lbl}", use_container_width=True,
                             type="primary" if dtype == typ else "secondary",
                             key=f"type_{typ}"):
                    st.session_state.dem_type     = typ
                    st.session_state.dem_step     = 2
                    st.session_state.dem_changes  = {}
                    st.session_state.dem_new_vals = {}
                    st.rerun()
                st.markdown(
                    f'<div style="text-align:center;font-size:.78rem;color:#6B7280;'
                    f'margin-top:.3rem;">{desc}</div>',
                    unsafe_allow_html=True,
                )

    # ── Étape 2 : Sélectionner la ligne ──────────────────────────────────────
    elif step == 2:
        section_title(
            f"{'✏️ Modifier' if dtype=='Modification' else '🗑️ Supprimer' if dtype=='Suppression' else '➕ Ajouter'}"
            " — Sélection"
        )

        if dtype in ("Modification", "Suppression"):
            with st.container():
                st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
                sc1, sc2, sc3 = st.columns(3)
                with sc1:
                    s_crpt = st.selectbox(
                        "Filtrer par site",
                        ["Tous"] + sorted(df["CRPT"].dropna().unique().tolist()),
                        key="sel_crpt",
                    )
                with sc2:
                    s_search = st.text_input("Recherche (poste, projet, RUO…)", key="sel_search")
                with sc3:
                    s_prog = st.selectbox(
                        "Filtrer par programme",
                        ["Tous"] + sorted(df["Programme industrielle"].dropna().unique().tolist()),
                        key="sel_prog",
                    )
                st.markdown("</div>", unsafe_allow_html=True)

            df_sel = df.copy()
            if s_crpt != "Tous":
                df_sel = df_sel[df_sel["CRPT"] == s_crpt]
            if s_prog != "Tous":
                df_sel = df_sel[df_sel["Programme industrielle"] == s_prog]
            if s_search:
                df_sel = df_sel[df_sel.apply(
                    lambda r: s_search.lower() in " ".join(r.astype(str)).lower(), axis=1
                )]

            df_sel["__label__"] = df_sel.apply(
                lambda r: f"[#{r['__id__']}]  {r['CRPT']}  ·  {r['Poste']}  ·  {r['Nom de projet']}",
                axis=1,
            )

            if df_sel.empty:
                st.warning("Aucune ligne ne correspond aux critères de recherche.")
            else:
                sel_label = st.selectbox(
                    f"Ligne à {'modifier' if dtype=='Modification' else 'supprimer'}",
                    df_sel["__label__"].tolist(), key="sel_ligne",
                )
                sel_id = int(df_sel.loc[df_sel["__label__"] == sel_label, "__id__"].values[0])
                sel_row = df[df["__id__"] == sel_id].drop(columns=["__id__"]).T
                sel_row.columns = ["Valeur actuelle"]
                with st.expander("👁️ Aperçu de la ligne sélectionnée", expanded=True):
                    st.dataframe(sel_row, use_container_width=True, height=320)

                bp, bn = st.columns([1, 4])
                with bp:
                    if st.button("← Retour", key="back2"):
                        st.session_state.dem_step = 1; st.rerun()
                with bn:
                    if st.button("Continuer →", type="primary", key="next2"):
                        st.session_state.dem_id   = sel_id
                        st.session_state.dem_step = 3; st.rerun()
        else:
            alert_box(
                "➕ Vous allez créer une <strong>nouvelle entrée</strong>. "
                "Renseignez les champs à l'étape suivante.",
                RTE_LIGHT, RTE_BLUE,
            )
            bp, bn = st.columns([1, 4])
            with bp:
                if st.button("← Retour", key="back2_add"):
                    st.session_state.dem_step = 1; st.rerun()
            with bn:
                if st.button("Continuer →", type="primary", key="next2_add"):
                    st.session_state.dem_step = 3; st.rerun()

    # ── Étape 3 : Remplir les détails ─────────────────────────────────────────
    elif step == 3:
        sel_id = st.session_state.dem_id

        if dtype == "Modification":
            section_title("✏️ Modifier — Renseignez les champs à changer")
            alert_box(
                "Seuls les champs que vous modifiez seront soumis. "
                "Les autres resteront inchangés.",
                RTE_LIGHT, RTE_BLUE,
            )
            row_data = df[df["__id__"] == sel_id].drop(columns=["__id__"]).iloc[0]
            changes  = st.session_state.dem_changes.copy()

            for col_name in ALL_COLUMNS:
                cur_val = str(row_data.get(col_name, "")) if pd.notna(row_data.get(col_name)) else ""
                is_mod  = col_name in changes
                bg      = RTE_LIGHT if is_mod else "white"
                brd     = RTE_BLUE  if is_mod else "#E5E7EB"
                st.markdown(
                    f'<div style="background:{bg};border-left:3px solid {brd};'
                    f'border-radius:8px;padding:.6rem .9rem;margin-bottom:.4rem;">'
                    f'<span style="font-size:.72rem;font-weight:600;text-transform:uppercase;'
                    f'letter-spacing:.04em;color:#6B7280;">{col_name}{"  ✏️" if is_mod else ""}'
                    f'</span></div>',
                    unsafe_allow_html=True,
                )
                if col_name in COLUMN_OPTIONS:
                    nv = st.selectbox(
                        f"_{col_name}_", ["(inchangé)"] + COLUMN_OPTIONS[col_name],
                        index=0, key=f"mod_{col_name}", label_visibility="collapsed",
                    )
                    if nv != "(inchangé)": changes[col_name] = nv
                    elif col_name in changes: del changes[col_name]
                else:
                    nv = st.text_input(
                        f"_{col_name}_", value=changes.get(col_name, ""),
                        placeholder=f"Valeur actuelle : {cur_val}",
                        key=f"mod_{col_name}", label_visibility="collapsed",
                    )
                    if nv.strip():   changes[col_name] = nv.strip()
                    elif col_name in changes: del changes[col_name]

            st.session_state.dem_changes = changes
            if changes:
                alert_box(
                    f"✏️ <strong>{len(changes)} champ(s) modifié(s)</strong> : "
                    f"{', '.join(changes.keys())}",
                    RTE_LIGHT, RTE_BLUE,
                )

        elif dtype == "Suppression":
            section_title("🗑️ Supprimer — Motif de la demande")
            row_data = df[df["__id__"] == sel_id].drop(columns=["__id__"]).iloc[0]
            alert_box(
                f"⚠️ Suppression de : <strong>{row_data.get('CRPT','')} · "
                f"{row_data.get('Poste','')} · {row_data.get('Nom de projet','')}</strong>",
                "#FEF3C7", RTE_ORANGE,
            )
            with st.expander("👁️ Ligne concernée", expanded=False):
                st.dataframe(row_data.to_frame("Valeur actuelle"), use_container_width=True)
            raison = st.text_area(
                "Raison de la suppression *",
                value=st.session_state.dem_raison,
                placeholder="Expliquez pourquoi cette ligne doit être supprimée…",
                height=120, key="raison_supp",
            )
            st.session_state.dem_raison = raison

        else:  # Ajout
            section_title("➕ Ajouter — Renseignez les informations de la nouvelle entrée")
            alert_box(
                "Les champs marqués <strong>*</strong> sont obligatoires.",
                RTE_LIGHT, RTE_BLUE,
            )
            new_vals = st.session_state.dem_new_vals.copy()
            REQUIRED = ["CRPT", "Poste", "Nom de projet", "RUO"]
            for col_name in ALL_COLUMNS:
                req = " *" if col_name in REQUIRED else ""
                if col_name in COLUMN_OPTIONS:
                    v = st.selectbox(
                        f"{col_name}{req}",
                        ["— Sélectionner —"] + COLUMN_OPTIONS[col_name],
                        key=f"new_{col_name}",
                    )
                    new_vals[col_name] = v if v != "— Sélectionner —" else ""
                else:
                    v = st.text_input(
                        f"{col_name}{req}", value=new_vals.get(col_name, ""),
                        key=f"new_{col_name}",
                    )
                    new_vals[col_name] = v.strip()
            st.session_state.dem_new_vals = new_vals

        # ── Navigation étape 3 ────────────────────────────────────────────────
        st.markdown("---")
        bn1, bn2 = st.columns([1, 4])
        with bn1:
            if st.button("← Retour", key="back3"):
                st.session_state.dem_step = 2; st.rerun()
        with bn2:
            if dtype == "Modification" and not st.session_state.dem_changes:
                st.warning("Veuillez modifier au moins un champ avant de continuer.")
            elif dtype == "Suppression" and not st.session_state.dem_raison.strip():
                st.warning("Veuillez indiquer une raison avant de continuer.")
            elif dtype == "Ajout":
                missing = [c for c in ["CRPT","Poste","Nom de projet","RUO"]
                           if not st.session_state.dem_new_vals.get(c,"").strip()]
                if missing:
                    st.warning(f"Champs obligatoires manquants : {', '.join(missing)}")
                elif st.button("Prévisualiser →", type="primary", key="next3_add"):
                    st.session_state.dem_step = 4; st.rerun()
            else:
                if st.button("Prévisualiser →", type="primary", key="next3"):
                    st.session_state.dem_step = 4; st.rerun()

    # ── Étape 4 : Confirmation ────────────────────────────────────────────────
    elif step == 4:
        section_title("✅ Confirmation — Récapitulatif de votre demande")
        sel_id = st.session_state.dem_id

        if dtype == "Modification":
            changes  = st.session_state.dem_changes
            row_data = df[df["__id__"] == sel_id].drop(columns=["__id__"]).iloc[0]
            alert_box(
                f"✏️ <strong>Modification</strong> — ligne #{sel_id} : "
                f"<strong>{row_data.get('CRPT','')} · {row_data.get('Poste','')} · "
                f"{row_data.get('Nom de projet','')}</strong>",
                RTE_LIGHT, RTE_BLUE,
            )
            recap = [{"Champ": c,
                      "Valeur actuelle": str(row_data.get(c,"")) if pd.notna(row_data.get(c)) else "—",
                      "Nouvelle valeur": v}
                     for c, v in changes.items()]
            st.dataframe(pd.DataFrame(recap), use_container_width=True, hide_index=True)
            desc_str    = "\n".join(f"{k}: {v}" for k, v in changes.items())
            details_str = " | ".join(f"{k}: {v}" for k, v in row_data.items()
                                     if pd.notna(v))[:500]

        elif dtype == "Suppression":
            row_data = df[df["__id__"] == sel_id].drop(columns=["__id__"]).iloc[0]
            alert_box(
                f"🗑️ <strong>Suppression</strong> — ligne #{sel_id} : "
                f"<strong>{row_data.get('CRPT','')} · {row_data.get('Poste','')} · "
                f"{row_data.get('Nom de projet','')}</strong>",
                "#FEF3C7", RTE_ORANGE,
            )
            st.markdown(f"**Motif :** {st.session_state.dem_raison}")
            desc_str    = st.session_state.dem_raison
            details_str = " | ".join(f"{k}: {v}" for k, v in row_data.items()
                                     if pd.notna(v))[:500]

        else:  # Ajout
            new_vals = st.session_state.dem_new_vals
            alert_box("➕ <strong>Ajout</strong> d'une nouvelle entrée", RTE_LIGHT, RTE_TEAL)
            rdata = {k: v for k, v in new_vals.items() if v}
            st.dataframe(
                pd.DataFrame(list(rdata.items()), columns=["Champ","Valeur"]),
                use_container_width=True, hide_index=True,
            )
            desc_str    = "\n".join(f"{k}: {v}" for k, v in rdata.items())
            details_str = "Nouvelle ligne"
            sel_id      = "NOUVEAU"

        alert_box(
            "📌 Cette demande sera <strong>soumise à validation</strong> "
            "par un administrateur avant toute modification de la base.",
            "#FEF3C7", RTE_ORANGE,
        )
        bn1, bn2, _ = st.columns([1, 1, 3])
        with bn1:
            if st.button("← Modifier", key="back4"):
                st.session_state.dem_step = 3; st.rerun()
        with bn2:
            if st.button("📨 Envoyer la demande", type="primary", key="send"):
                add_request(dtype, sel_id, details_str, desc_str)
                for k in ["dem_step","dem_type","dem_id","dem_changes","dem_raison","dem_new_vals"]:
                    st.session_state.pop(k, None)
                st.success("✅ Demande envoyée ! En attente de validation.")
                st.balloons()

    # ── Historique ────────────────────────────────────────────────────────────
    st.markdown("---")
    section_title("📋 Historique de toutes les demandes")
    req_df = load_requests()
    if req_df.empty:
        st.info("Aucune demande enregistrée pour le moment.")
    else:
        hf_col, dl1, dl2 = st.columns([3, 1, 1])
        with hf_col:
            hf = st.selectbox(
                "Filtrer",
                ["Tous","En attente","Acceptée","Refusée","Modification","Suppression","Ajout"],
                key="hist_filter",
            )
        if hf in ["En attente","Acceptée","Refusée"]:
            disp = req_df[req_df["statut"] == hf]
        elif hf in ["Modification","Suppression","Ajout"]:
            disp = req_df[req_df["type"] == hf]
        else:
            disp = req_df

        with dl1:
            st.markdown("<br/>", unsafe_allow_html=True)
            st.download_button(
                "⬇️ CSV",
                data=disp.to_csv(index=False, sep=";", encoding="utf-8-sig"),
                file_name="historique_demandes.csv", mime="text/csv",
                use_container_width=True, key="dl_hist_csv_dem",
            )
        with dl2:
            st.markdown("<br/>", unsafe_allow_html=True)
            _buf = io.BytesIO(); disp.to_excel(_buf, index=False)
            st.download_button(
                "⬇️ Excel", data=_buf.getvalue(),
                file_name="historique_demandes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_hist_xlsx_dem",
            )
        st.markdown(
            f'<div style="font-size:.78rem;color:#6B7280;margin-bottom:.4rem;">'
            f'{len(disp)} demande(s) affichée(s)</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(disp.reset_index(drop=True), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : IMPORT MADU
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "madu":
    df_bdd = get_data()
    page_header(
        "Importation des dates de MADU",
        "Chargez un fichier Excel pour comparer les dates MADU avec la base existante",
    )

    # ── Instructions ──────────────────────────────────────────────────────────
    with st.expander("📖 Format attendu du fichier d'import", expanded=False):
        st.markdown("""
Le fichier Excel doit contenir **exactement ces 4 colonnes** (noms insensibles à la casse) :

| Colonne | Description | Exemple |
|---|---|---|
| `nom de projet` | Intitulé du projet | GIEN Instal Self 80 MVAr |
| `EOTP2` | Identifiant unique du projet (**clé de correspondance avec le champ RUO de la BDD**) | O20-112 |
| `date de MADU` | Nouvelle date de MADU au format JJ/MM/AAAA | 15/03/2025 |
| `matériel` | Type de matériel concerné | Self 80 MVAr |

> 💡 La correspondance est faite via **EOTP2** (fichier import) ↔ **RUO** (BDD principale).
        """)
        # Fichier exemple à télécharger
        ex_data = {
            "nom de projet":  ["Exemple projet A", "Exemple projet B"],
            "EOTP2":          ["O20-112", "O21-040"],
            "date de MADU":   ["15/03/2025", "01/06/2025"],
            "matériel":       ["Self 80 MVAr", "Self 80 MVAr"],
        }
        ex_buf = io.BytesIO()
        pd.DataFrame(ex_data).to_excel(ex_buf, index=False)
        st.download_button(
            "⬇️ Télécharger un fichier exemple",
            data=ex_buf.getvalue(),
            file_name="modele_import_madu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown("<br/>", unsafe_allow_html=True)
    section_title("📂 Charger le fichier d'import")

    uploaded = st.file_uploader(
        "Glissez-déposez votre fichier Excel ici ou cliquez pour parcourir",
        type=["xlsx", "xls"],
        help="Fichier Excel (.xlsx) contenant les colonnes : nom de projet, EOTP2, date de MADU, matériel",
    )

    if uploaded is not None:
        # ── Validation du fichier ─────────────────────────────────────────────
        df_import, errors = validate_madu_file(uploaded)

        # Affichage des avertissements non bloquants
        for err in errors:
            if err.startswith("⚠️"):
                st.warning(err)
            else:
                st.error(err)

        if df_import is None:
            st.stop()

        # ── Aperçu du fichier importé ─────────────────────────────────────────
        st.markdown("<br/>", unsafe_allow_html=True)
        section_title(f"👁️ Aperçu du fichier importé — {len(df_import)} ligne(s) valide(s)")

        df_preview = df_import.copy()
        df_preview["date de MADU"] = df_preview["date de MADU"].dt.strftime("%d/%m/%Y")
        st.dataframe(df_preview, use_container_width=True, height=250)

        # ── Lancement de la comparaison ───────────────────────────────────────
        st.markdown("<br/>", unsafe_allow_html=True)
        section_title("🔄 Comparaison avec la base de données")

        alert_box(
            f"La comparaison porte sur <strong>{len(df_import)}</strong> projet(s) importé(s) "
            f"et <strong>{len(df_bdd)}</strong> entrée(s) dans la BDD (clé : EOTP2 ↔ RUO).",
            RTE_LIGHT, RTE_BLUE,
        )

        if st.button("🚀 Lancer la comparaison", type="primary"):
            with st.spinner("Comparaison en cours…"):
                df_compare = compare_madu(df_import, df_bdd)
                summary    = get_madu_summary(df_compare)
                # Stockage en session pour le dashboard
                st.session_state["madu_compare_result"] = df_compare
            st.success("✅ Comparaison terminée ! Les résultats sont aussi disponibles sur le Dashboard.")

        # ── Affichage des résultats (si comparaison déjà faite) ───────────────
        df_compare = st.session_state.get("madu_compare_result")
        if df_compare is not None:
            summary = get_madu_summary(df_compare)

            st.markdown("<br/>", unsafe_allow_html=True)
            section_title("📊 Résultats de la comparaison")

            # KPIs
            m1, m2, m3, m4, m5 = st.columns(5)
            with m1:
                kpi_card("Projets trouvés", summary["trouves"],
                         sub=f"{summary['non_trouves']} non trouvé(s)", accent=RTE_BLUE)
            with m2:
                kpi_card("Dates avancées ✅", summary["avancees"],
                         sub="rapprochée", accent=RTE_GREEN)
            with m3:
                kpi_card("Dates retardées ⚠️", summary["retardees"],
                         sub="repoussée", accent=RTE_RED)
            with m4:
                kpi_card("Inchangées", summary["inchangees"],
                         sub="aucun écart", accent=RTE_TEAL)
            with m5:
                kpi_card(
                    "Écart moyen",
                    f"{summary['ecart_moyen_jours']:+.0f} j",
                    sub=f"min {summary['ecart_min_jours']:+d} j  /  max {summary['ecart_max_jours']:+d} j",
                    accent=RTE_ORANGE,
                )

            st.markdown("<br/>", unsafe_allow_html=True)

            # ── Graphiques ────────────────────────────────────────────────────
            gc1, gc2 = st.columns([2, 1])

            with gc1:
                section_title("📊 Écarts de dates MADU par projet (jours)")
                df_ecarts = df_compare[
                    df_compare["Statut"].isin(["Avancée","Retardée","Inchangée"])
                ].copy()
                if not df_ecarts.empty:
                    color_map = {
                        "Avancée":   RTE_GREEN,
                        "Retardée":  RTE_RED,
                        "Inchangée": "#9CA3AF",
                    }
                    fig_bar = px.bar(
                        df_ecarts.sort_values("Écart (jours)"),
                        x="EOTP2", y="Écart (jours)",
                        color="Statut", color_discrete_map=color_map,
                        hover_data=["Projet (import)", "Ancienne date MADU", "Nouvelle date MADU"],
                        text="Écart (jours)",
                    )
                    fig_bar.add_hline(y=0, line_dash="dash", line_color="#6B7280", line_width=1.5)
                    fig_bar.update_layout(
                        height=340, **CHART,
                        xaxis_title="EOTP2", yaxis_title="Écart (jours)",
                        xaxis_tickangle=-30,
                    )
                    fig_bar.update_traces(textposition="outside", marker_line_width=0)
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("Aucun projet avec écart trouvé dans la BDD.")

            with gc2:
                section_title("🥧 Répartition par statut")
                df_pie = df_compare["Statut"].value_counts().reset_index()
                df_pie.columns = ["Statut", "Nombre"]
                color_pie = {
                    "Avancée":                RTE_GREEN,
                    "Retardée":               RTE_RED,
                    "Inchangée":              "#9CA3AF",
                    "Non trouvé dans la BDD": RTE_ORANGE,
                    "Date manquante":         "#E5E7EB",
                }
                fig_pie = px.pie(
                    df_pie, names="Statut", values="Nombre",
                    color="Statut", color_discrete_map=color_pie, hole=0.5,
                )
                fig_pie.update_layout(height=340, **CHART)
                fig_pie.update_traces(
                    textinfo="percent+label",
                    marker=dict(line=dict(color="white", width=2)),
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            # ── Histogramme de distribution des écarts ────────────────────────
            section_title("📉 Distribution des écarts (histogramme)")
            df_hist = df_compare[df_compare["Écart (jours)"].notna()].copy()
            if not df_hist.empty:
                fig_hist = px.histogram(
                    df_hist, x="Écart (jours)",
                    color_discrete_sequence=[RTE_BLUE],
                    nbins=20,
                )
                fig_hist.add_vline(x=0, line_dash="dash", line_color=RTE_RED, line_width=1.5,
                                   annotation_text="Pas d'écart", annotation_position="top right")
                fig_hist.update_layout(
                    height=280, **CHART,
                    xaxis_title="Écart en jours", yaxis_title="Nombre de projets",
                    bargap=0.1,
                )
                st.plotly_chart(fig_hist, use_container_width=True)

            # ── Tableau détaillé ──────────────────────────────────────────────
            section_title("📋 Tableau détaillé des écarts")
            tf1, tf2 = st.columns([3, 1])
            with tf1:
                t_filter = st.selectbox(
                    "Filtrer par statut",
                    ["Tous", "Avancée", "Retardée", "Inchangée", "Non trouvé dans la BDD"],
                    key="madu_table_filter",
                )
            with tf2:
                only_mod = st.checkbox("Modifiés uniquement", key="madu_only_mod")

            df_show = df_compare.copy()
            if t_filter != "Tous":
                df_show = df_show[df_show["Statut"] == t_filter]
            if only_mod:
                df_show = df_show[df_show["Statut"].isin(["Avancée","Retardée"])]

            # Formatage des dates pour l'affichage
            for dc in ["Ancienne date MADU", "Nouvelle date MADU"]:
                df_show[dc] = pd.to_datetime(df_show[dc], errors="coerce").dt.strftime("%d/%m/%Y")

            # Coloration par statut
            def _color_row(row):
                mapping = {
                    "Avancée":               "background-color:#D1FAE5",
                    "Retardée":              "background-color:#FEE2E2",
                    "Inchangée":             "",
                    "Non trouvé dans la BDD": "background-color:#FEF3C7",
                }
                return [mapping.get(row["Statut"], "")] * len(row)

            styled = df_show.reset_index(drop=True).style.apply(_color_row, axis=1)
            st.dataframe(styled, use_container_width=True, height=400)

            # ── Export ────────────────────────────────────────────────────────
            st.markdown("---")
            section_title("📥 Exporter les résultats")
            e1, e2, _ = st.columns([1, 1, 3])
            with e1:
                st.download_button(
                    "⬇️ CSV",
                    data=df_show.to_csv(index=False, sep=";", encoding="utf-8-sig"),
                    file_name="comparaison_madu.csv", mime="text/csv",
                    use_container_width=True, key="dl_madu_csv",
                )
            with e2:
                buf_xl = io.BytesIO()
                df_show.to_excel(buf_xl, index=False)
                st.download_button(
                    "⬇️ Excel", data=buf_xl.getvalue(),
                    file_name="comparaison_madu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="dl_madu_xlsx",
                )

            # ── Effacer les résultats ─────────────────────────────────────────
            st.markdown("<br/>", unsafe_allow_html=True)
            if st.button("🗑️ Effacer les résultats"):
                st.session_state.pop("madu_compare_result", None)
                st.rerun()

    else:
        # Aucun fichier uploadé : afficher un état vide
        st.markdown("<br/>", unsafe_allow_html=True)
        if st.session_state.get("madu_compare_result") is not None:
            alert_box(
                "✅ Une analyse MADU est déjà chargée et visible sur le <strong>Dashboard</strong>. "
                "Importez un nouveau fichier pour mettre à jour.",
                "#F0FDF4", RTE_GREEN,
            )
        else:
            alert_box(
                "Aucun fichier chargé. Glissez-déposez votre fichier Excel ci-dessus "
                "pour démarrer la comparaison.",
                RTE_LIGHT, RTE_BLUE,
            )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : ADMINISTRATION
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "admin":
    ADMIN_PASSWORD = "admin1234"
    page_header("Administration", "Validation des demandes · Modification / Suppression / Ajout")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        st.markdown("<br/>", unsafe_allow_html=True)
        _, c, _ = st.columns([1, 1.3, 1])
        with c:
            st.markdown(
                f'<div style="background:white;border-radius:14px;padding:2rem 2rem 1.5rem;'
                f'box-shadow:0 4px 20px rgba(0,0,0,.1);border-top:4px solid {RTE_BLUE};">'
                f'<div style="text-align:center;margin-bottom:1.2rem;">'
                f'<span style="font-size:2.5rem;">🔐</span><br/>'
                f'<strong style="font-size:1.1rem;color:{RTE_DARK};">Espace Administrateur</strong>'
                f'</div>',
                unsafe_allow_html=True,
            )
            pwd = st.text_input(
                "Mot de passe", type="password",
                label_visibility="collapsed",
                placeholder="Entrez le mot de passe admin…",
            )
            if st.button("Se connecter", type="primary", use_container_width=True):
                if pwd == ADMIN_PASSWORD:
                    st.session_state.admin_logged_in = True
                    st.rerun()
                else:
                    st.error("Mot de passe incorrect.")
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        ci, cl = st.columns([4, 1])
        with ci:
            st.success("✅ Connecté en tant qu'administrateur")
        with cl:
            if st.button("🚪 Déconnexion", use_container_width=True):
                st.session_state.admin_logged_in = False; st.rerun()

        st.markdown("---")
        req_df   = load_requests()
        pending  = req_df[req_df["statut"] == "En attente"] if not req_df.empty else pd.DataFrame()
        accepted = len(req_df[req_df["statut"] == "Acceptée"]) if not req_df.empty else 0
        refused  = len(req_df[req_df["statut"] == "Refusée"])  if not req_df.empty else 0

        b1, b2, b3, b4 = st.columns(4)
        with b1: kpi_card("Total demandes", len(req_df) if not req_df.empty else 0, accent=RTE_BLUE)
        with b2: kpi_card("En attente",     len(pending),   accent=RTE_ORANGE)
        with b3: kpi_card("Acceptées",      accepted,       accent=RTE_GREEN)
        with b4: kpi_card("Refusées",       refused,        accent=RTE_RED)

        st.markdown("<br/>", unsafe_allow_html=True)
        section_title("📋 Demandes en attente de validation")

        if pending.empty:
            st.info("✅ Aucune demande en attente. Tout est à jour.")
        else:
            for _, row in pending.iterrows():
                icons = {"Modification": "✏️", "Suppression": "🗑️", "Ajout": "➕"}
                icon  = icons.get(row["type"], "📝")
                with st.expander(
                    f"{icon}  [#{row['id_demande']}]  {row['type']}  —  "
                    f"Ligne {row['id_ligne']}  —  {row['date_demande']}"
                ):
                    i1, i2 = st.columns(2)
                    with i1:
                        st.markdown(f"**Type :** `{row['type']}`")
                        st.markdown(f"**Ligne :** `{row['id_ligne']}`")
                        st.markdown(f"**Date :** {row['date_demande']}")
                    with i2:
                        st.markdown(f"**Statut :** `{row['statut']}`")

                    if row["type"] != "Ajout":
                        st.markdown("**Ligne concernée :**")
                        st.code(row["details_ligne"], language=None)

                    st.markdown("**Contenu de la demande :**")
                    st.info(row["description"])

                    # Aperçu avant/après pour les modifications
                    if row["type"] == "Modification":
                        try:
                            db_prev = load_data()
                            id_l    = int(row["id_ligne"])
                            db_mod  = apply_modification(id_l, row["description"], db_prev.copy())
                            orig    = db_prev[db_prev["__id__"] == id_l].drop(columns=["__id__"]).iloc[0]
                            modif   = db_mod[db_mod["__id__"]   == id_l].drop(columns=["__id__"]).iloc[0]
                            diff    = [(c, str(orig[c]), str(modif[c]))
                                       for c in orig.index if str(orig[c]) != str(modif[c])]
                            if diff:
                                st.markdown("**Aperçu des changements :**")
                                st.dataframe(
                                    pd.DataFrame(diff, columns=["Champ","Avant","Après"]),
                                    use_container_width=True, hide_index=True,
                                )
                        except Exception:
                            pass

                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button(
                            f"✅ Accepter #{row['id_demande']}",
                            key=f"acc_{row['id_demande']}", type="primary",
                            use_container_width=True,
                        ):
                            db = load_data()
                            try:
                                id_l = int(row["id_ligne"])
                            except Exception:
                                id_l = None

                            if row["type"] == "Suppression" and id_l is not None:
                                db = db[db["__id__"] != id_l]
                            elif row["type"] == "Modification" and id_l is not None:
                                db = apply_modification(id_l, row["description"], db)
                            elif row["type"] == "Ajout":
                                db = apply_addition(row["description"], db)

                            try:
                                save_data(db)
                                req_df.loc[req_df["id_demande"] == row["id_demande"], "statut"] = "Acceptée"
                                save_requests(req_df)
                                get_data.clear()
                                st.success(f"Demande #{row['id_demande']} acceptée et appliquée.")
                                st.rerun()
                            except PermissionError as e:
                                st.error(
                                    f"⚠️ **Fichier verrouillé** — Fermez `BDD.xlsx` dans Excel "
                                    f"puis cliquez à nouveau sur Accepter.\n\n_{e}_"
                                )
                    with c2:
                        if st.button(
                            f"❌ Refuser #{row['id_demande']}",
                            key=f"ref_{row['id_demande']}",
                            use_container_width=True,
                        ):
                            req_df.loc[req_df["id_demande"] == row["id_demande"], "statut"] = "Refusée"
                            save_requests(req_df)
                            st.warning(f"Demande #{row['id_demande']} refusée.")
                            st.rerun()

        st.markdown("---")
        section_title("📂 Historique complet")
        if req_df.empty:
            st.info("Aucune demande enregistrée.")
        else:
            sf_col, adl1, adl2 = st.columns([3, 1, 1])
            with sf_col:
                sf = st.selectbox(
                    "Filtrer par statut",
                    ["Tous","En attente","Acceptée","Refusée","Modification","Suppression","Ajout"],
                )
            disp = req_df if sf == "Tous" else (
                req_df[req_df["statut"] == sf] if sf in ["En attente","Acceptée","Refusée"]
                else req_df[req_df["type"] == sf]
            )
            with adl1:
                st.markdown("<br/>", unsafe_allow_html=True)
                st.download_button(
                    "⬇️ CSV",
                    data=disp.to_csv(index=False, sep=";", encoding="utf-8-sig"),
                    file_name="historique_admin.csv", mime="text/csv",
                    use_container_width=True, key="dl_admin_csv",
                )
            with adl2:
                st.markdown("<br/>", unsafe_allow_html=True)
                _buf2 = io.BytesIO(); disp.to_excel(_buf2, index=False)
                st.download_button(
                    "⬇️ Excel", data=_buf2.getvalue(),
                    file_name="historique_admin.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="dl_admin_xlsx",
                )
            st.markdown(
                f'<div style="font-size:.78rem;color:#6B7280;margin-bottom:.4rem;">'
                f'{len(disp)} demande(s) affichée(s)</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(disp.reset_index(drop=True), use_container_width=True)
