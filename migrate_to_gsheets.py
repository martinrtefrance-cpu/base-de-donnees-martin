"""
migrate_to_gsheets.py
══════════════════════════════════════════════════════════════════════════════
Script de migration one-shot : importe les données Excel locales vers
Google Sheets.

Lancez ce script UNE SEULE FOIS pour initialiser votre Google Spreadsheet
avec les données existantes.

Usage :
    python migrate_to_gsheets.py --excel data/BDD.xlsx --sheet-id VOTRE_ID

Pré-requis :
    pip install gspread google-auth pandas openpyxl
    Fichier credentials/service_account.json présent
══════════════════════════════════════════════════════════════════════════════
"""

import argparse
import sys
import os
import json
import pandas as pd

# Ajouter le répertoire courant au path pour importer les modules locaux
sys.path.insert(0, os.path.dirname(__file__))


def main():
    parser = argparse.ArgumentParser(
        description="Migration Excel → Google Sheets pour l'application RTE BDD Matériel"
    )
    parser.add_argument(
        "--excel",
        required=True,
        help="Chemin vers le fichier BDD.xlsx (ex: data/BDD.xlsx)",
    )
    parser.add_argument(
        "--sheet-id",
        required=True,
        help="ID du Google Spreadsheet cible",
    )
    parser.add_argument(
        "--creds",
        default="credentials/service_account.json",
        help="Chemin vers le fichier credentials JSON (défaut: credentials/service_account.json)",
    )
    args = parser.parse_args()

    # ── Vérifications préalables ──────────────────────────────────────────────
    if not os.path.exists(args.excel):
        print(f"❌ Fichier Excel introuvable : {args.excel}")
        sys.exit(1)

    if not os.path.exists(args.creds):
        print(f"❌ Fichier de credentials introuvable : {args.creds}")
        print("   Placez votre service_account.json dans le dossier credentials/")
        sys.exit(1)

    print(f"📊 Fichier source     : {args.excel}")
    print(f"🔑 Credentials       : {args.creds}")
    print(f"📋 Spreadsheet ID    : {args.sheet_id}")
    print()

    # ── Import des modules (après avoir vérifié les fichiers) ─────────────────
    import gspread
    from google.oauth2.service_account import Credentials
    from google_sheets_connector import (
        SCOPES, SHEET_BDD, SHEET_DEMANDES,
        ALL_COLUMNS, DEMANDES_COLUMNS,
        init_spreadsheet,
        overwrite_sheet,
    )

    # ── Connexion ─────────────────────────────────────────────────────────────
    print("🔌 Connexion à Google Sheets...")
    creds = Credentials.from_service_account_file(args.creds, scopes=SCOPES)
    client = gspread.authorize(creds)
    print("✅ Connexion réussie")

    # Monkey-patch pour que get_gspread_client() utilise notre client local
    import google_sheets_connector as gsc
    gsc.get_gspread_client = lambda: client

    # ── Initialisation des onglets ────────────────────────────────────────────
    print("\n🏗️  Initialisation des onglets...")
    ok = init_spreadsheet(args.sheet_id)
    if ok:
        print("✅ Onglets 'BDD' et 'Demandes' prêts")
    else:
        print("⚠️  Vérifiez les permissions du spreadsheet")

    # ── Import de la BDD ──────────────────────────────────────────────────────
    print(f"\n📥 Lecture de {args.excel}...")
    df_bdd = pd.read_excel(args.excel, dtype=str)
    print(f"   {len(df_bdd)} lignes trouvées")

    # S'assurer que toutes les colonnes attendues sont présentes
    df_bdd = df_bdd.reindex(columns=ALL_COLUMNS).fillna("")

    print("📤 Envoi vers Google Sheets (onglet 'BDD')...")
    spreadsheet = client.open_by_key(args.sheet_id)
    ws_bdd = spreadsheet.worksheet(SHEET_BDD)
    ws_bdd.clear()

    data = [ALL_COLUMNS] + df_bdd.values.tolist()
    ws_bdd.update("A1", data, value_input_option="USER_ENTERED")
    print(f"✅ {len(df_bdd)} lignes importées dans l'onglet 'BDD'")

    # ── Import des demandes (si le fichier existe) ────────────────────────────
    demandes_path = os.path.join(os.path.dirname(args.excel), "demandes.xlsx")
    if os.path.exists(demandes_path):
        print(f"\n📥 Lecture de {demandes_path}...")
        df_dem = pd.read_excel(demandes_path, dtype=str)
        df_dem = df_dem.reindex(columns=DEMANDES_COLUMNS).fillna("")
        print(f"   {len(df_dem)} demandes trouvées")

        ws_dem = spreadsheet.worksheet(SHEET_DEMANDES)
        ws_dem.clear()
        data_dem = [DEMANDES_COLUMNS] + df_dem.values.tolist()
        ws_dem.update("A1", data_dem, value_input_option="USER_ENTERED")
        print(f"✅ {len(df_dem)} demandes importées dans l'onglet 'Demandes'")
    else:
        print(f"\nℹ️  Pas de fichier demandes.xlsx trouvé — onglet 'Demandes' laissé vide")

    print("\n" + "═" * 60)
    print("🎉 Migration terminée avec succès !")
    print(f"   Ouvrez votre spreadsheet :")
    print(f"   https://docs.google.com/spreadsheets/d/{args.sheet_id}/edit")
    print("═" * 60)


if __name__ == "__main__":
    main()
