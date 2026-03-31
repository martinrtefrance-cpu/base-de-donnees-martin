# RTE BDD Matériel — Version Google Sheets

Application Streamlit de gestion des équipements réseau RTE.
**Les données sont stockées dans Google Sheets** (plus de fichiers locaux).

---

## 🏗️ Architecture

```
app.py                         ← Application Streamlit (inchangée)
utils.py                       ← Interface métier (même API qu'avant)
google_sheets_connector.py     ← Couche d'accès Google Sheets API
migrate_to_gsheets.py          ← Script de migration one-shot
.streamlit/secrets.toml        ← Secrets (non commité)
credentials/service_account.json ← Credentials locales (non commité)
```

**Flux de données :**
```
Streamlit UI → utils.py → google_sheets_connector.py → Google Sheets API
```

---

## ⚙️ Configuration Google Cloud — Guide pas à pas

### Étape 1 — Créer un projet Google Cloud

1. Allez sur **https://console.cloud.google.com**
2. Cliquez sur le sélecteur de projet en haut → **Nouveau projet**
3. Nom : `rte-bdd-materiel` (ou ce que vous voulez)
4. Cliquez **Créer**
5. Sélectionnez ce nouveau projet

### Étape 2 — Activer les APIs nécessaires

Dans le menu → **APIs et services** → **Bibliothèque** :

1. Cherchez **Google Sheets API** → cliquez → **Activer**
2. Cherchez **Google Drive API** → cliquez → **Activer**

> ⚠️ Les deux APIs sont nécessaires : Sheets pour lire/écrire, Drive pour lister les fichiers.

### Étape 3 — Créer un compte de service (Service Account)

Le compte de service est un "robot" Google qui agit au nom de votre application.

1. Menu → **APIs et services** → **Identifiants**
2. Cliquez **Créer des identifiants** → **Compte de service**
3. Nom : `rte-bdd-app`
4. Description : `Compte de service pour l'application RTE BDD`
5. Cliquez **Créer et continuer**
6. Rôle : **Éditeur** (ou laisser vide, les permissions sont gérées via le partage du sheet)
7. Cliquez **Continuer** puis **Terminer**

### Étape 4 — Télécharger le fichier de credentials JSON

1. Dans la liste des comptes de service, cliquez sur `rte-bdd-app`
2. Onglet **Clés** → **Ajouter une clé** → **Créer une clé**
3. Format : **JSON** → **Créer**
4. Le fichier JSON se télécharge automatiquement
5. **Renommez-le** `service_account.json`
6. **Placez-le** dans le dossier `credentials/` de votre projet

```
app/
├── credentials/
│   └── service_account.json  ← ici (jamais dans git !)
├── app.py
└── ...
```

### Étape 5 — Créer le Google Spreadsheet

1. Allez sur **https://sheets.google.com**
2. Créez un nouveau spreadsheet
3. Renommez-le : `RTE BDD Matériel`
4. **Créez deux onglets** (en bas) :
   - Renommez `Feuille 1` en **`BDD`**
   - Ajoutez un onglet nommé **`Demandes`**
5. **Copiez l'ID** depuis l'URL :
   ```
   https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
                                          ↑ C'est cet ID ↑
   ```

### Étape 6 — Partager le sheet avec le compte de service

> C'est l'étape la plus souvent oubliée !

1. Dans votre Google Spreadsheet, cliquez **Partager** (bouton vert en haut à droite)
2. Dans le champ email, collez l'adresse du compte de service.
   Elle ressemble à :
   ```
   rte-bdd-app@votre-projet.iam.gserviceaccount.com
   ```
   Trouvez-la dans le fichier `service_account.json` à la clé `client_email`.
3. Rôle : **Éditeur**
4. Décochez "Notifier" (le compte de service ne reçoit pas d'emails)
5. Cliquez **Partager**

### Étape 7 — Configurer les secrets

**En développement local** — éditez `.streamlit/secrets.toml` :

```toml
[google_sheets]
spreadsheet_id = "VOTRE_SPREADSHEET_ID"

[gcp_service_account]
type                        = "service_account"
project_id                  = "rte-bdd-materiel"
private_key_id              = "..."
private_key                 = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email                = "rte-bdd-app@rte-bdd-materiel.iam.gserviceaccount.com"
client_id                   = "..."
auth_uri                    = "https://accounts.google.com/o/oauth2/auth"
token_uri                   = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url        = "..."
```

> 💡 Copiez les valeurs directement depuis votre fichier `service_account.json`.
> La `private_key` doit avoir les `\n` littéraux (pas de vraies nouvelles lignes).

**En production (Streamlit Cloud)** :
1. Allez dans votre app sur share.streamlit.io
2. **Settings** → **Secrets**
3. Collez le contenu de `secrets.toml`

---

## 🚀 Migration des données existantes

Si vous avez déjà des données dans `BDD.xlsx`, exécutez le script de migration :

```bash
# 1. Installer les dépendances
pip install -r requirements.txt

# 2. Lancer la migration
python migrate_to_gsheets.py \
  --excel data/BDD.xlsx \
  --sheet-id VOTRE_SPREADSHEET_ID \
  --creds credentials/service_account.json
```

Le script :
- Lit `BDD.xlsx` et `demandes.xlsx`
- Initialise les onglets si nécessaires
- Importe toutes les données dans Google Sheets

---

## ▶️ Lancer l'application

```bash
# Installer les dépendances
pip install -r requirements.txt

# Lancer
streamlit run app.py
```

---

## 📊 Structure du Google Spreadsheet

### Onglet `BDD` (données principales)

| Colonne | Type | Description |
|---|---|---|
| CRPT | Texte | Site géographique |
| Poste | Texte | Identifiant du poste |
| Nom de projet | Texte | Intitulé complet |
| RUO | Texte | **Clé EOTP2 pour l'import MADU** |
| Programme industrielle | Texte | Programme de référence |
| ... | ... | 16 colonnes au total |

### Onglet `Demandes` (file d'attente admin)

| Colonne | Type | Description |
|---|---|---|
| id_demande | Texte | Identifiant unique |
| type | Texte | Modification / Suppression / Ajout |
| id_ligne | Texte | Numéro de ligne BDD concernée |
| details_ligne | Texte | Résumé de la ligne avant modif |
| description | Texte | Détail de la demande |
| date_demande | Texte | Horodatage (AAAA-MM-JJ HH:MM) |
| statut | Texte | En attente / Acceptée / Refusée |

---

## 🔒 Sécurité

| Fichier | Doit être dans Git ? |
|---|---|
| `credentials/service_account.json` | ❌ Non (dans .gitignore) |
| `.streamlit/secrets.toml` | ❌ Non (dans .gitignore) |
| `app.py`, `utils.py`, etc. | ✅ Oui |

---

## 🐛 Dépannage

| Erreur | Cause | Solution |
|---|---|---|
| `SpreadsheetNotFound` | ID incorrect ou sheet non partagé | Vérifier l'ID + partager avec le compte de service |
| `WorksheetNotFound` | Onglet mal nommé | Créer les onglets "BDD" et "Demandes" exactement |
| `APIError 403` | Permissions insuffisantes | Partager le sheet en mode Éditeur |
| `APIError 429` | Quota dépassé | L'app retry automatiquement (backoff exponentiel) |
| `FileNotFoundError: credentials` | Fichier JSON absent | Placer service_account.json dans credentials/ |
| `KeyError: gcp_service_account` | Secrets non configurés | Compléter .streamlit/secrets.toml |

---

## 📈 Limites de l'API Google Sheets

| Limite | Valeur | Impact |
|---|---|---|
| Requêtes en lecture | 300/min par projet | Cache 30s dans Streamlit |
| Requêtes en écriture | 300/min par projet | Retry automatique |
| Cellules par sheet | 10 millions | Largement suffisant |
| Taille max par cellule | 50 000 caractères | `details_ligne` tronqué à 500 |
