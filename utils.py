"""
utils.py — Fonctions métier de l'application RTE BDD Matériel
══════════════════════════════════════════════════════════════════════════════

Ce fichier expose exactement la même interface publique que l'ancien utils.py
(load_data, save_data, load_requests, etc.) mais les données sont désormais
lues et écrites dans Google Sheets au lieu de fichiers Excel locaux.

L'application (app.py) n'a donc besoin d'aucune modification : seul ce fichier
change pour brancher un autre backend de stockage.

ARCHITECTURE
────────────
  app.py  →  utils.py  →  google_sheets_connector.py  →  Google Sheets API
                ↑
          Interface identique à l'ancien utils.py (compatibilité garantie)
"""

import pandas as pd
import streamlit as st
from datetime import datetime

from google_sheets_connector import (
    # Lecture
    read_sheet_as_dataframe,
    # Écriture
    append_row,
    update_row,
    delete_row,
    overwrite_sheet,
    # Constantes
    SHEET_BDD,
    SHEET_DEMANDES,
    ALL_COLUMNS,
    DEMANDES_COLUMNS,
    DATE_COLS,
)

# ── Valeurs connues par colonne (inchangé par rapport à l'ancien utils.py) ───
COLUMN_OPTIONS = {
    "CRPT": ["Nantes", "Marseille", "Toulouse", "Nancy", "Lyon", "Paris", "Lille"],
    "Programme industrielle": ["SEPT", "POCC", "S3R", "LALS", "stock", "racco", "DPMI"],
    "Sous-politique": ["DR1", "R5C", "DR3", "RU5", "DR6", "DR1-accel", "DR1 rég"],
    "Etude tension haute": ["V1", "V2", "V3", "V3 Compléments"],
    "Type de BIS": ["245-95", "21-64", "63-30", "63-15", "245-200 Regl",
                    "90-30", "400-126", "420-300 reg"],
    "Huile/Air": ["huile", "air"],
    "Insonorisée": ["0", "1"],
    "Affectation PLQF": ["Siemens", "BEST", "GE", "HITACHI PSEM",
                         "Hitachi C", "Hitachi L", "TRENCH"],
}

# ── Paramètres MADU (inchangés) ───────────────────────────────────────────────
MADU_IMPORT_COLS = ["nom de projet", "EOTP2", "date de MADU", "matériel"]
MADU_KEY_BDD     = "RUO"
MADU_KEY_IMPORT  = "EOTP2"
MADU_DATE_BDD    = "Première MADU projet"


def _get_sheet_id() -> str:
    """
    Récupère l'ID du Google Spreadsheet depuis les secrets Streamlit.

    Le fichier .streamlit/secrets.toml doit contenir :
        [google_sheets]
        spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"

    Raises:
        KeyError : si la clé n'est pas trouvée dans les secrets
    """
    try:
        return st.secrets["google_sheets"]["spreadsheet_id"]
    except KeyError:
        raise KeyError(
            "L'ID du Google Spreadsheet est manquant dans les secrets.\n"
            "Ajoutez dans .streamlit/secrets.toml :\n"
            "  [google_sheets]\n"
            "  spreadsheet_id = \"votre_id_ici\""
        )


# ══════════════════════════════════════════════════════════════════════════════
# LECTURE DE LA BDD PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=30)
def load_data() -> pd.DataFrame:
    """
    Charge la BDD principale depuis l'onglet 'BDD' du Google Spreadsheet.

    Remplace : pd.read_excel("data/BDD.xlsx")

    @st.cache_data(ttl=30) : les données sont mises en cache 30 secondes.
    Streamlit recharge depuis Google Sheets maximum 2 fois par minute,
    ce qui évite d'atteindre les quotas de l'API.

    Returns:
        DataFrame avec __id__ (index 0-based, recalculé à chaque lecture)
    """
    sheet_id = _get_sheet_id()
    return read_sheet_as_dataframe(
        spreadsheet_id=sheet_id,
        sheet_name=SHEET_BDD,
        expected_columns=ALL_COLUMNS,
        date_columns=DATE_COLS,
    )


# ══════════════════════════════════════════════════════════════════════════════
# SAUVEGARDE DE LA BDD PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

def save_data(df: pd.DataFrame):
    """
    Remplace tout le contenu de l'onglet 'BDD' par le DataFrame fourni.

    Remplace : df.to_excel("data/BDD.xlsx")

    Appelé après une modification validée par l'admin (modification, suppression,
    ajout). Invalide automatiquement le cache Streamlit.

    Args:
        df : DataFrame à sauvegarder (avec ou sans __id__, sera ignoré)
    """
    sheet_id = _get_sheet_id()
    overwrite_sheet(sheet_id, SHEET_BDD, df, ALL_COLUMNS)
    # Invalider le cache pour que la prochaine lecture recharge depuis GSheets
    load_data.clear()


# ══════════════════════════════════════════════════════════════════════════════
# LECTURE DES DEMANDES
# ══════════════════════════════════════════════════════════════════════════════

def load_requests() -> pd.DataFrame:
    """
    Charge les demandes depuis l'onglet 'Demandes' du Google Spreadsheet.

    Remplace : pd.read_excel("data/demandes.xlsx") avec création si absent.

    Note : pas de cache ici car les demandes doivent toujours être à jour
    (l'admin et les utilisateurs peuvent agir simultanément).

    Returns:
        DataFrame des demandes avec colonnes DEMANDES_COLUMNS
    """
    sheet_id = _get_sheet_id()
    return read_sheet_as_dataframe(
        spreadsheet_id=sheet_id,
        sheet_name=SHEET_DEMANDES,
        expected_columns=DEMANDES_COLUMNS,
    )


# ══════════════════════════════════════════════════════════════════════════════
# SAUVEGARDE DES DEMANDES
# ══════════════════════════════════════════════════════════════════════════════

def save_requests(df: pd.DataFrame):
    """
    Remplace tout le contenu de l'onglet 'Demandes'.

    Utilisé pour mettre à jour le statut d'une demande (acceptée/refusée).

    Args:
        df : DataFrame complet des demandes
    """
    sheet_id = _get_sheet_id()
    overwrite_sheet(sheet_id, SHEET_DEMANDES, df, DEMANDES_COLUMNS)


# ══════════════════════════════════════════════════════════════════════════════
# AJOUT D'UNE DEMANDE
# ══════════════════════════════════════════════════════════════════════════════

def add_request(type_demande: str, id_ligne, details_ligne: str, description: str):
    """
    Ajoute une nouvelle demande à la fin de l'onglet 'Demandes'.

    Remplace : df = pd.concat([df, new_row]) + df.to_excel(...)

    Utilise append_row() qui ajoute une seule ligne via l'API,
    sans recharger tout le sheet → bien plus rapide que overwrite_sheet().

    Args:
        type_demande  : "Modification", "Suppression" ou "Ajout"
        id_ligne      : __id__ de la ligne concernée (ou "NOUVEAU")
        details_ligne : Résumé de la ligne actuelle (avant modification)
        description   : Description de la demande saisie par l'utilisateur
    """
    sheet_id = _get_sheet_id()

    # Calculer le prochain id_demande
    df_existing = load_requests()
    new_id = str(len(df_existing) + 1)

    new_row = {
        "id_demande":    new_id,
        "type":          type_demande,
        "id_ligne":      str(id_ligne),
        "details_ligne": details_ligne[:500],  # tronquer pour éviter les cellules trop grandes
        "description":   description,
        "date_demande":  datetime.now().strftime("%Y-%m-%d %H:%M"),
        "statut":        "En attente",
    }

    # Ajouter uniquement cette ligne → 1 seule requête API
    append_row(sheet_id, SHEET_DEMANDES, new_row, DEMANDES_COLUMNS)


# ══════════════════════════════════════════════════════════════════════════════
# APPLICATION D'UNE MODIFICATION (admin)
# ══════════════════════════════════════════════════════════════════════════════

def apply_modification(id_ligne: int, description: str, df_data: pd.DataFrame) -> pd.DataFrame:
    """
    Applique les modifications décrites au format 'Colonne: valeur'
    sur la ligne identifiée par id_ligne dans df_data (en mémoire).

    La persistance dans Google Sheets est assurée par save_data() ensuite.

    Args:
        id_ligne    : __id__ de la ligne à modifier
        description : Texte au format "Colonne: valeur\\nColonne2: valeur2"
        df_data     : DataFrame courant de la BDD

    Returns:
        DataFrame avec la ligne modifiée
    """
    for line in description.split("\n"):
        if ":" in line:
            col, val = line.split(":", 1)
            col, val = col.strip(), val.strip()
            if col in df_data.columns:
                df_data.loc[df_data["__id__"] == id_ligne, col] = val
    return df_data


def apply_addition(description: str, df_data: pd.DataFrame) -> pd.DataFrame:
    """
    Parse les lignes 'Colonne: valeur' et ajoute une nouvelle ligne à df_data
    (en mémoire). La persistance est assurée par save_data() ensuite.

    Args:
        description : Texte au format "Colonne: valeur\\nColonne2: valeur2"
        df_data     : DataFrame courant de la BDD

    Returns:
        DataFrame avec la nouvelle ligne ajoutée
    """
    new_row = {col: None for col in ALL_COLUMNS}
    for line in description.split("\n"):
        if ":" in line:
            col, val = line.split(":", 1)
            col, val = col.strip(), val.strip()
            if col in new_row:
                new_row[col] = val
    new_row["__id__"] = int(df_data["__id__"].max()) + 1 if len(df_data) > 0 else 0
    df_data = pd.concat([df_data, pd.DataFrame([new_row])], ignore_index=True)
    return df_data


# ══════════════════════════════════════════════════════════════════════════════
# MODULE IMPORT & COMPARAISON MADU  (inchangé — pas de stockage local)
# ══════════════════════════════════════════════════════════════════════════════

def validate_madu_file(uploaded_file):
    """
    Lit et valide le fichier Excel d'import MADU.
    Identique à l'ancienne version : le fichier MADU est toujours un upload
    temporaire, pas stocké localement ni dans Google Sheets.
    """
    errors = []
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
    except Exception as e:
        return None, [f"Impossible de lire le fichier Excel : {e}"]

    df.columns = [c.strip().lower() for c in df.columns]
    expected   = [c.lower() for c in MADU_IMPORT_COLS]
    missing    = [c for c in expected if c not in df.columns]
    if missing:
        return None, [
            f"Colonnes manquantes : {', '.join(missing)}. "
            f"Attendues : {', '.join(MADU_IMPORT_COLS)}."
        ]

    rename_map = {c.lower(): c for c in MADU_IMPORT_COLS}
    df = df.rename(columns=rename_map)

    n_before = len(df)
    df = df.dropna(subset=["EOTP2"])
    n_dropped = n_before - len(df)
    if n_dropped:
        errors.append(f"⚠️ {n_dropped} ligne(s) ignorée(s) : colonne 'EOTP2' vide.")

    df["date de MADU"] = df["date de MADU"].astype(str).str.strip()
    df["date de MADU"] = pd.to_datetime(df["date de MADU"], dayfirst=True, errors="coerce")
    n_bad = df["date de MADU"].isna().sum()
    if n_bad:
        errors.append(f"⚠️ {n_bad} date(s) non reconnue(s) — lignes exclues.")
        df = df.dropna(subset=["date de MADU"])

    df["EOTP2"] = df["EOTP2"].astype(str).str.strip()
    if df.empty:
        return None, ["Aucune ligne valide après nettoyage."]
    return df, errors


def compare_madu(df_import: pd.DataFrame, df_bdd: pd.DataFrame) -> pd.DataFrame:
    """Compare les dates MADU importées avec la BDD. Identique à l'ancienne version."""
    bdd = df_bdd[["__id__", MADU_KEY_BDD, "Nom de projet", "CRPT", MADU_DATE_BDD]].copy()
    bdd[MADU_KEY_BDD] = bdd[MADU_KEY_BDD].astype(str).str.strip()

    merged = df_import.merge(bdd, left_on="EOTP2", right_on=MADU_KEY_BDD, how="left")
    merged["Ancienne date MADU"] = pd.to_datetime(merged[MADU_DATE_BDD], errors="coerce")
    merged["Nouvelle date MADU"] = pd.to_datetime(merged["date de MADU"], errors="coerce")

    def _ecart(row):
        if pd.isna(row["Ancienne date MADU"]) or pd.isna(row["Nouvelle date MADU"]):
            return None
        return (row["Nouvelle date MADU"] - row["Ancienne date MADU"]).days

    def _statut(row):
        if pd.isna(row[MADU_KEY_BDD]):         return "Non trouvé dans la BDD"
        if pd.isna(row["Écart (jours)"]): return "Date manquante"
        e = row["Écart (jours)"]
        if e < 0: return "Avancée"
        if e > 0: return "Retardée"
        return "Inchangée"

    merged["Écart (jours)"] = merged.apply(_ecart, axis=1)
    merged["Statut"]        = merged.apply(_statut, axis=1)

    return merged[[
        "EOTP2", "nom de projet", "Nom de projet", "CRPT", "matériel",
        "Ancienne date MADU", "Nouvelle date MADU", "Écart (jours)", "Statut",
    ]].rename(columns={
        "nom de projet": "Projet (import)",
        "Nom de projet": "Projet (BDD)",
        "matériel":      "Matériel",
    })


def get_madu_summary(df_compare: pd.DataFrame) -> dict:
    """Indicateurs de synthèse MADU. Identique à l'ancienne version."""
    mask_found = df_compare["Statut"] != "Non trouvé dans la BDD"
    ecarts     = df_compare[mask_found]["Écart (jours)"].dropna()
    return {
        "total":             len(df_compare),
        "trouves":           int(mask_found.sum()),
        "non_trouves":       int((~mask_found).sum()),
        "avancees":          int((df_compare["Statut"] == "Avancée").sum()),
        "retardees":         int((df_compare["Statut"] == "Retardée").sum()),
        "inchangees":        int((df_compare["Statut"] == "Inchangée").sum()),
        "ecart_moyen_jours": round(float(ecarts.mean()), 1) if len(ecarts) else 0,
        "ecart_max_jours":   int(ecarts.max())  if len(ecarts) else 0,
        "ecart_min_jours":   int(ecarts.min())  if len(ecarts) else 0,
    }
