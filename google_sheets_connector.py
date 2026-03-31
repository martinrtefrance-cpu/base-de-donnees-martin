"""
google_sheets_connector.py
══════════════════════════════════════════════════════════════════════════════
Couche d'abstraction entre l'application et l'API Google Sheets.

Ce module remplace entièrement la logique de lecture/écriture des fichiers
Excel locaux (BDD.xlsx, demandes.xlsx) par des appels à Google Sheets.

PRINCIPE DE FONCTIONNEMENT
──────────────────────────
Un Google Spreadsheet contient deux feuilles (onglets) :
  • "BDD"      → équivalent de data/BDD.xlsx
  • "Demandes" → équivalent de data/demandes.xlsx

Chaque feuille a une ligne d'en-tête (ligne 1) puis les données.
L'identifiant interne __id__ n'est PAS stocké dans le sheet ;
il est recalculé dynamiquement (= index dans la feuille, 0-based).

AUTHENTIFICATION
────────────────
L'application utilise un compte de service Google (Service Account).
Les credentials sont lus depuis :
  1. Un fichier local  credentials/service_account.json  (dev local)
  2. Les secrets Streamlit  st.secrets["gcp_service_account"]  (déploiement)

══════════════════════════════════════════════════════════════════════════════
"""

import time
import json
import logging
from functools import wraps
from typing import Optional

import pandas as pd
import streamlit as st

# gspread est la bibliothèque principale pour interagir avec Google Sheets
import gspread
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound
from google.oauth2.service_account import Credentials

# ── Logging ───────────────────────────────────────────────────────────────────
logger = logging.getLogger(__name__)

# ── Scopes OAuth2 nécessaires ─────────────────────────────────────────────────
# spreadsheets      → lecture/écriture des cellules
# drive.readonly    → lister les spreadsheets accessibles (optionnel)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ── Noms des onglets dans le Google Spreadsheet ───────────────────────────────
SHEET_BDD      = "BDD"
SHEET_DEMANDES = "Demandes"

# ── Colonnes attendues (même ordre que l'ancien Excel) ────────────────────────
ALL_COLUMNS = [
    "CRPT", "Poste", "Nom de projet", "RUO", "Programme industrielle",
    "Sous-politique", "Etude tension haute", "Type de BIS", "Huile/Air",
    "Insonorisée", "Première MADU projet", "Première MEC projet",
    "MES PFM1", "Confiance Besoin", "Confiance MADU", "Affectation PLQF",
]

DEMANDES_COLUMNS = [
    "id_demande", "type", "id_ligne", "details_ligne",
    "description", "date_demande", "statut",
]

# Colonnes à parser comme dates lors de la lecture
DATE_COLS = ["Première MADU projet", "MES PFM1"]


# ══════════════════════════════════════════════════════════════════════════════
# DÉCORATEUR DE RETRY  (gestion des quotas et erreurs réseau)
# ══════════════════════════════════════════════════════════════════════════════

def retry_on_quota(max_retries: int = 5, base_delay: float = 2.0):
    """
    Décorateur qui rejoue automatiquement une fonction en cas d'erreur 429
    (quota Google Sheets dépassé) ou d'erreur réseau temporaire.

    Utilise un backoff exponentiel : 2s, 4s, 8s, 16s, 32s
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = base_delay
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except APIError as e:
                    # 429 = Too Many Requests (quota dépassé)
                    if e.response.status_code == 429 and attempt < max_retries - 1:
                        logger.warning(
                            f"Quota Google Sheets dépassé. "
                            f"Nouvelle tentative dans {delay:.0f}s "
                            f"(essai {attempt + 1}/{max_retries})"
                        )
                        time.sleep(delay)
                        delay *= 2
                    else:
                        raise
                except Exception as e:
                    if attempt < max_retries - 1:
                        logger.warning(f"Erreur temporaire : {e}. Retry dans {delay:.0f}s")
                        time.sleep(delay)
                        delay *= 2
                    else:
                        raise
        return wrapper
    return decorator


# ══════════════════════════════════════════════════════════════════════════════
# CONNEXION À GOOGLE SHEETS
# ══════════════════════════════════════════════════════════════════════════════

def _load_credentials() -> Credentials:
    """
    Charge les credentials du compte de service depuis :
      1. st.secrets (prioritaire → pour Streamlit Cloud)
      2. Fichier local credentials/service_account.json (dev)

    Raises:
        FileNotFoundError : si aucune source de credentials n'est trouvée
        ValueError        : si les credentials sont malformées
    """
    # ── Option 1 : Streamlit secrets (Streamlit Cloud / production) ───────────
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    # ── Option 2 : Fichier local (développement) ──────────────────────────────
    import os
    creds_path = os.path.join("credentials", "service_account.json")
    if os.path.exists(creds_path):
        return Credentials.from_service_account_file(creds_path, scopes=SCOPES)

    raise FileNotFoundError(
        "Credentials Google introuvables.\n"
        "• En local : placez votre fichier JSON dans credentials/service_account.json\n"
        "• En production : ajoutez [gcp_service_account] dans .streamlit/secrets.toml"
    )


@st.cache_resource(show_spinner=False)
def get_gspread_client() -> gspread.Client:
    """
    Crée et met en cache le client gspread.

    @st.cache_resource : le client est créé une seule fois pour toute la session
    Streamlit et partagé entre tous les reruns. C'est l'équivalent d'un singleton.
    """
    creds = _load_credentials()
    client = gspread.authorize(creds)
    logger.info("Client Google Sheets initialisé avec succès.")
    return client


def get_spreadsheet(spreadsheet_id: str) -> gspread.Spreadsheet:
    """
    Ouvre le Google Spreadsheet par son ID.

    L'ID se trouve dans l'URL du sheet :
    https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit

    Args:
        spreadsheet_id: L'identifiant unique du Google Spreadsheet

    Returns:
        L'objet Spreadsheet gspread

    Raises:
        SpreadsheetNotFound : si l'ID est incorrect ou le sheet non partagé
    """
    client = get_gspread_client()
    try:
        return client.open_by_key(spreadsheet_id)
    except SpreadsheetNotFound:
        raise SpreadsheetNotFound(
            f"Spreadsheet '{spreadsheet_id}' introuvable.\n"
            f"Vérifiez que :\n"
            f"  • L'ID est correct (trouvé dans l'URL du sheet)\n"
            f"  • Le sheet est partagé avec l'email du compte de service\n"
            f"    (avec le rôle Éditeur)"
        )


def get_worksheet(spreadsheet_id: str, sheet_name: str) -> gspread.Worksheet:
    """
    Récupère un onglet (worksheet) du spreadsheet.

    Args:
        spreadsheet_id : ID du Google Spreadsheet
        sheet_name     : nom de l'onglet ("BDD" ou "Demandes")

    Returns:
        L'objet Worksheet gspread
    """
    spreadsheet = get_spreadsheet(spreadsheet_id)
    try:
        return spreadsheet.worksheet(sheet_name)
    except WorksheetNotFound:
        raise WorksheetNotFound(
            f"Onglet '{sheet_name}' introuvable dans le spreadsheet.\n"
            f"Créez un onglet nommé exactement '{sheet_name}'."
        )


# ══════════════════════════════════════════════════════════════════════════════
# LECTURE DES DONNÉES
# ══════════════════════════════════════════════════════════════════════════════

@retry_on_quota()
def read_sheet_as_dataframe(
    spreadsheet_id: str,
    sheet_name: str,
    expected_columns: Optional[list] = None,
    date_columns: Optional[list] = None,
) -> pd.DataFrame:
    """
    Lit un onglet Google Sheets et le retourne sous forme de DataFrame pandas.

    La première ligne est utilisée comme en-tête.
    Les cellules vides sont remplacées par NaN.

    Args:
        spreadsheet_id   : ID du Google Spreadsheet
        sheet_name       : Nom de l'onglet
        expected_columns : Si fourni, crée le DataFrame avec ces colonnes même si vide
        date_columns     : Colonnes à convertir en datetime

    Returns:
        DataFrame avec une colonne __id__ (index 0-based, non stockée dans le sheet)
    """
    ws = get_worksheet(spreadsheet_id, sheet_name)

    # get_all_records() : lit toutes les lignes et retourne une liste de dicts
    # Les colonnes sont définies par la première ligne du sheet
    records = ws.get_all_records(
        empty2zero=False,    # les cellules vides restent "" (pas 0)
        head=1,              # la ligne 1 est l'en-tête
        default_blank="",    # valeur par défaut pour les cellules vides
    )

    # ── Construction du DataFrame ─────────────────────────────────────────────
    if records:
        df = pd.DataFrame(records)
        # Remplacer les chaînes vides par NaN (cohérent avec pd.read_excel)
        df = df.replace("", pd.NA)
    else:
        # Sheet vide : créer un DataFrame avec les bonnes colonnes
        cols = expected_columns or []
        df = pd.DataFrame(columns=cols)

    # ── Conversion des dates ──────────────────────────────────────────────────
    if date_columns:
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    # ── Ajout de l'identifiant interne __id__ ─────────────────────────────────
    # __id__ = index 0-based dans le sheet (NON stocké dans Google Sheets)
    # Il correspond à row_number - 2 dans le sheet (ligne 1 = headers)
    df["__id__"] = range(len(df))

    return df


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE — AJOUT D'UNE LIGNE
# ══════════════════════════════════════════════════════════════════════════════

@retry_on_quota()
def append_row(spreadsheet_id: str, sheet_name: str, row_data: dict, columns: list):
    """
    Ajoute une nouvelle ligne à la fin du sheet.

    Args:
        spreadsheet_id : ID du Google Spreadsheet
        sheet_name     : Nom de l'onglet
        row_data       : Dictionnaire {colonne: valeur}
        columns        : Ordre des colonnes (pour construire la ligne dans le bon ordre)

    Exemple:
        append_row(SHEET_ID, "Demandes", {"id_demande": "42", ...}, DEMANDES_COLUMNS)
    """
    ws = get_worksheet(spreadsheet_id, sheet_name)

    # Construire la ligne dans le même ordre que les colonnes du sheet
    row = [str(row_data.get(col, "")) if pd.notna(row_data.get(col, "")) else ""
           for col in columns]

    # append_row() ajoute la ligne après la dernière ligne non vide
    ws.append_row(row, value_input_option="USER_ENTERED")
    # USER_ENTERED : Google Sheets interprète les valeurs comme si l'utilisateur
    # les tapait (parse les dates, nombres, etc.)


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE — MISE À JOUR D'UNE LIGNE EXISTANTE
# ══════════════════════════════════════════════════════════════════════════════

@retry_on_quota()
def update_row(
    spreadsheet_id: str,
    sheet_name: str,
    row_id: int,
    new_data: dict,
    columns: list,
):
    """
    Met à jour une ligne existante dans le sheet.

    Args:
        spreadsheet_id : ID du Google Spreadsheet
        sheet_name     : Nom de l'onglet
        row_id         : __id__ de la ligne (0-based) → converti en numéro de ligne GSheet
        new_data       : Dictionnaire des champs à modifier {colonne: nouvelle_valeur}
        columns        : Ordre des colonnes dans le sheet

    Note sur la correspondance des indices :
        __id__ = 0  →  ligne GSheet 2  (ligne 1 = headers)
        __id__ = N  →  ligne GSheet N+2
    """
    ws = get_worksheet(spreadsheet_id, sheet_name)

    # Numéro de ligne dans le sheet (1-indexed, +1 pour les headers, +1 pour 1-index)
    gsheet_row = row_id + 2

    # Lire d'abord la ligne actuelle pour ne modifier que les champs demandés
    current_row = ws.row_values(gsheet_row)

    # S'assurer que current_row a assez de colonnes
    while len(current_row) < len(columns):
        current_row.append("")

    # Appliquer les modifications
    for col_name, new_val in new_data.items():
        if col_name in columns:
            col_idx = columns.index(col_name)
            current_row[col_idx] = str(new_val) if pd.notna(new_val) else ""

    # Mettre à jour toute la ligne en une seule requête API
    # (plus efficace que de mettre à jour cellule par cellule)
    ws.update(
        f"A{gsheet_row}",
        [current_row[:len(columns)]],
        value_input_option="USER_ENTERED",
    )


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE — SUPPRESSION D'UNE LIGNE
# ══════════════════════════════════════════════════════════════════════════════

@retry_on_quota()
def delete_row(spreadsheet_id: str, sheet_name: str, row_id: int):
    """
    Supprime physiquement une ligne du sheet (les lignes suivantes remontent).

    Args:
        spreadsheet_id : ID du Google Spreadsheet
        sheet_name     : Nom de l'onglet
        row_id         : __id__ de la ligne (0-based)

    ⚠️  Attention : après une suppression, tous les __id__ des lignes suivantes
        changent. L'application recharge systématiquement les données après
        une suppression pour recalculer les __id__.
    """
    ws = get_worksheet(spreadsheet_id, sheet_name)
    gsheet_row = row_id + 2  # +1 header, +1 pour passer en 1-indexed
    ws.delete_rows(gsheet_row)


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE — REMPLACEMENT TOTAL D'UNE FEUILLE
# ══════════════════════════════════════════════════════════════════════════════

@retry_on_quota()
def overwrite_sheet(spreadsheet_id: str, sheet_name: str, df: pd.DataFrame, columns: list):
    """
    Remplace tout le contenu d'un onglet par le DataFrame fourni.

    Utilisé pour les opérations complexes (réorganisation, nettoyage).
    Plus lent qu'un update ciblé mais garantit la cohérence.

    Args:
        spreadsheet_id : ID du Google Spreadsheet
        sheet_name     : Nom de l'onglet
        df             : DataFrame à écrire (sans __id__)
        columns        : Colonnes à écrire (dans l'ordre)
    """
    ws = get_worksheet(spreadsheet_id, sheet_name)

    # Construire la liste de listes [en-tête, ligne1, ligne2, ...]
    df_clean = df.drop(columns=["__id__"], errors="ignore")
    df_clean = df_clean.reindex(columns=columns)  # garantir l'ordre et les colonnes

    # Convertir les dates en chaînes lisibles
    for col in df_clean.columns:
        if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].dt.strftime("%d/%m/%Y").fillna("")

    # Remplacer NaN par des chaînes vides
    df_clean = df_clean.fillna("")

    data = [columns] + df_clean.values.tolist()

    # Effacer le sheet puis réécrire
    ws.clear()
    ws.update("A1", data, value_input_option="USER_ENTERED")


# ══════════════════════════════════════════════════════════════════════════════
# INITIALISATION DU SPREADSHEET
# ══════════════════════════════════════════════════════════════════════════════

def init_spreadsheet(spreadsheet_id: str):
    """
    Vérifie que le spreadsheet contient les onglets requis avec les bons en-têtes.
    Crée les onglets manquants si nécessaire.

    À appeler au démarrage de l'application.

    Args:
        spreadsheet_id : ID du Google Spreadsheet

    Returns:
        True si tout est OK, False si une erreur est survenue
    """
    try:
        spreadsheet = get_spreadsheet(spreadsheet_id)
        existing_sheets = [ws.title for ws in spreadsheet.worksheets()]

        # ── Créer l'onglet BDD s'il n'existe pas ─────────────────────────────
        if SHEET_BDD not in existing_sheets:
            ws = spreadsheet.add_worksheet(title=SHEET_BDD, rows=500, cols=20)
            ws.update("A1", [ALL_COLUMNS])
            logger.info(f"Onglet '{SHEET_BDD}' créé avec les en-têtes.")

        # ── Créer l'onglet Demandes s'il n'existe pas ─────────────────────────
        if SHEET_DEMANDES not in existing_sheets:
            ws = spreadsheet.add_worksheet(title=SHEET_DEMANDES, rows=500, cols=10)
            ws.update("A1", [DEMANDES_COLUMNS])
            logger.info(f"Onglet '{SHEET_DEMANDES}' créé avec les en-têtes.")

        return True

    except Exception as e:
        logger.error(f"Erreur lors de l'initialisation du spreadsheet : {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# IMPORT INITIAL DEPUIS EXCEL  (migration one-shot)
# ══════════════════════════════════════════════════════════════════════════════

def import_excel_to_gsheets(spreadsheet_id: str, excel_path: str) -> dict:
    """
    Importe les données d'un fichier Excel local vers Google Sheets.
    Utile pour la migration initiale one-shot.

    Args:
        spreadsheet_id : ID du Google Spreadsheet cible
        excel_path     : Chemin vers le fichier BDD.xlsx local

    Returns:
        {"bdd": n_lignes_importées, "demandes": n_lignes_importées, "errors": [...]}
    """
    import os
    results = {"bdd": 0, "demandes": 0, "errors": []}

    try:
        # ── Import de la BDD ──────────────────────────────────────────────────
        df_bdd = pd.read_excel(excel_path, dtype=str)
        df_bdd = df_bdd.reindex(columns=ALL_COLUMNS).fillna("")
        overwrite_sheet(spreadsheet_id, SHEET_BDD, df_bdd, ALL_COLUMNS)
        results["bdd"] = len(df_bdd)
        logger.info(f"{len(df_bdd)} lignes importées dans l'onglet BDD.")
    except Exception as e:
        results["errors"].append(f"BDD : {e}")

    # ── Import des demandes (si le fichier existe) ────────────────────────────
    demandes_path = os.path.join(os.path.dirname(excel_path), "demandes.xlsx")
    if os.path.exists(demandes_path):
        try:
            df_dem = pd.read_excel(demandes_path, dtype=str)
            df_dem = df_dem.reindex(columns=DEMANDES_COLUMNS).fillna("")
            overwrite_sheet(spreadsheet_id, SHEET_DEMANDES, df_dem, DEMANDES_COLUMNS)
            results["demandes"] = len(df_dem)
            logger.info(f"{len(df_dem)} demandes importées.")
        except Exception as e:
            results["errors"].append(f"Demandes : {e}")

    return results
