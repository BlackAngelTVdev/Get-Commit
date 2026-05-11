#!/bin/bash

#
# Journal de travail Git → CSV → Excel
# Mode APPEND : complète le CSV avec les nouveaux commits uniquement
# Les termes techniques sont traduits en langage métier dans l'Excel
#

# ─── Config ───────────────────────────────────────────────────────────────────
SINCE="2026-01-13"
OUTPUT_DIR="Doc"
OUTPUT_FILE="$OUTPUT_DIR/journal_de_travail.csv"
OUTPUT_XLSX="$OUTPUT_DIR/journal_de_travail.xlsx"

mkdir -p "$OUTPUT_DIR"

echo "🚀 Extraction en mode append..."

# ─── 1. Récupérer la dernière date du CSV existant ────────────────────────────
LAST_DATE=""
if [ -f "$OUTPUT_FILE" ]; then
    LAST_DATE=$(awk -F',' 'NF && $1 !~ /^Date$/ {
        date = $1
        gsub(/^"|"$/, "", date)
        gsub(/^ +| +$/, "", date)
        if (date != "") { last_date = date }
    } END { print last_date }' "$OUTPUT_FILE" 2>/dev/null)
    echo "ℹ️  Fichier existant trouvé. Dernière date enregistrée : $LAST_DATE"
else
    echo "Date,Nom,Temps,État,Description" > "$OUTPUT_FILE"
    echo "✅ Nouveau fichier CSV créé"
fi

# Convertir la dernière date en format ISO pour git --since
if [ -n "$LAST_DATE" ]; then
    if [[ "$LAST_DATE" =~ ^([0-9]{2})/([0-9]{2})/([0-9]{4})$ ]]; then
        day="${BASH_REMATCH[1]}"
        month="${BASH_REMATCH[2]}"
        year="${BASH_REMATCH[3]}"
        SINCE="$year-$month-$day 00:00:00"
    else
        SINCE="$LAST_DATE"
    fi
else
    SINCE="2026-01-13"
fi

# ─── 2. Extraction git log → CSV ──────────────────────────────────────────────
EXISTING_ENTRIES=$(awk -F',' 'NR>1 {
    date = $1; nom = $2
    gsub(/^"|"$/, "", date); gsub(/^"|"$/, "", nom)
    gsub(/^ +| +$/, "", date); gsub(/^ +| +$/, "", nom)
    if (date != "" && nom != "") { print date "|" nom }
}' "$OUTPUT_FILE" 2>/dev/null)

git log --all --since="$SINCE" --reverse \
    --date=format:'%d/%m/%Y' \
    --pretty=format:%ad%x1f%s%x1f%b%x1e \
| awk -v RS='\036' -F '\037' -v EXISTING="$EXISTING_ENTRIES" '
BEGIN {
    split(EXISTING, existing_arr, "\n")
    for (i in existing_arr) { existing[existing_arr[i]] = 1 }
}
NF {
    current_date = $1
    nom_commit   = $2
    body         = $3

    gsub(/\r|\n/, "",  current_date)
    gsub(/^ +| +$/, "", current_date)
    gsub(/\r|\n/, " ", nom_commit)
    gsub(/\r|\n/, " ", body)

    full_text = nom_commit " " body
    temps = "[?]"
    etat  = "[?]"

    # Extraction Temps [1h45min]
    if (match(full_text, /\[[0-9hmin]+\]/)) {
        temps = substr(full_text, RSTART, RLENGTH)
        sub(/\[[0-9hmin]+\]/, "", full_text)
    }

    # Extraction État [DONE]
    if (match(full_text, /\[[A-Z]+\]/)) {
        etat = substr(full_text, RSTART, RLENGTH)
        sub(/\[[A-Z]+\]/, "", full_text)
    }

    gsub(/"/, "\"\"", nom_commit)
    gsub(/"/, "\"\"", full_text)
    gsub(/^ +| +$/, "", nom_commit)
    gsub(/^ +| +$/, "", full_text)

    entry_key = current_date "|" nom_commit
    if (entry_key in existing) { next }

    if (last_date != "" && last_date != current_date) {
        print "" >> "'$OUTPUT_FILE'"
    }
    last_date = current_date

    printf "\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n",
        current_date, nom_commit, temps, etat, full_text >> "'$OUTPUT_FILE'"
    existing[entry_key] = 1
}
'

echo "📄 CSV généré : $OUTPUT_FILE"

# ─── 3. Conversion CSV → Excel avec traduction des termes techniques ───────────
if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_BIN="python"
else
    PYTHON_BIN=""
fi

if [ -n "$PYTHON_BIN" ]; then
    PY_OUTPUT=$(OUTPUT_FILE="$OUTPUT_FILE" OUTPUT_XLSX="$OUTPUT_XLSX" "$PYTHON_BIN" - << 'PYEOF' 2>&1

import csv
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

src = os.environ["OUTPUT_FILE"]
dst = os.environ["OUTPUT_XLSX"]

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Alignment, Font, PatternFill
except Exception:
    sys.exit(2)


# ──────────────────────────────────────────────────────────────────────────────
# TABLE DE CORRESPONDANCE : termes techniques → termes métier
#
# Formats reconnus pour chaque préfixe (toutes ces variantes sont capturées) :
#   feat:           feat!:          feat(auth):       feat(auth)!:
#   FEAT:           Feat(login):    feat(ui/modal):   feat(scope)!:
#
# Exemples complets :
#   "feat: ajout login"                   → "Fonctionnalité : ajout connexion"
#   "feat(auth): ajout token JWT"         → "Fonctionnalité (auth) : ajout jeton d'authentification JWT"
#   "feat(auth)!: refonte login"          → "Fonctionnalité (auth) : refonte connexion"
#   "fix: crash au démarrage"             → "Correction : crash au démarrage"
#   "fix(paiement): bug montant"          → "Correction (paiement) : anomalie montant"
#   "chore: nettoyage logs"               → "Tâche : nettoyage journal"
#   "chore(ci): update pipeline"          → "Tâche (ci) : mise à jour chaîne de traitement"
#   "doc: guide installation"             → "Documentation : guide installation"
#   "docs: mise à jour README"            → "Documentation : mise à jour lisez-moi"
#   "docs(api): endpoints paiement"       → "Documentation (api) : endpoints paiement"
#   "refactor(auth): réécriture module"   → "Refactorisation (auth) : réécriture module"
#   "style(form): indentation CSS"        → "Mise en forme (form) : indentation CSS"
#   "test(paiement): ajout mock"          → "Test (paiement) : ajout simulation"
#   "perf(db): cache requêtes"            → "Performance (db) : cache requêtes"
#   "ci(github): ajout step lint"         → "Intégration continue (github) : ajout step lint"
#   "build(webpack): upgrade v5"          → "Build (webpack) : mise à niveau v5"
#   "revert: annulation merge"            → "Annulation : fusion"
#   "wip(router): refacto en cours"       → "En cours (router) : refacto en cours"
#   "hotfix(auth): token expiré"          → "Correctif urgent (auth) : jeton d'authentification expiré"
#   "deploy(prod): release v2.3"          → "Déploiement (prod) : version publiée v2.3"
#   "release(v1.2.0): changelog"          → "Nouvelle version (v1.2.0) : journal des modifications"
# ──────────────────────────────────────────────────────────────────────────────

# Regex générique : préfixe optionnellement suivi de (scope) et/ou ! puis :
# Groupe 1 = scope entre parenthèses (sans les parenthèses), ou None
_SCOPE = r"(?:\(([^)]*)\))?"   # capture optionnelle du scope
_BANG  = r"[!]?"               # breaking change optionnel
_SEP   = r":\s*"               # deux-points + espaces


def _prefix_re(keyword: str) -> str:
    """Construit la regex pour un préfixe de commit conventionnel."""
    return rf"^{keyword}{_SCOPE}{_BANG}{_SEP}"


# Table des préfixes → label métier
PREFIX_MAP = [
    # Exemple : "feat(auth): ajout login"   → "Fonctionnalité (auth) : ajout connexion"
    (_prefix_re(r"feat"),     "Fonctionnalité"),
    # Exemple : "fix(paiement): bug montant" → "Correction (paiement) : anomalie montant"
    (_prefix_re(r"fix"),      "Correction"),
    # Exemple : "chore(ci): update deps"    → "Tâche (ci) : mise à jour dépendances"
    (_prefix_re(r"chore?"),   "Tâche"),
    # Exemple : "doc: guide"  ou  "docs(api): endpoints"  → "Documentation ..."
    (_prefix_re(r"docs?"),    "Documentation"),
    # Exemple : "refactor(auth): réécriture" → "Refactorisation (auth) : réécriture"
    (_prefix_re(r"refactor"), "Refactorisation"),
    # Exemple : "style(form): indentation"  → "Mise en forme (form) : indentation"
    (_prefix_re(r"style"),    "Mise en forme"),
    # Exemple : "test(paie): ajout mock"    → "Test (paie) : ajout simulation"
    (_prefix_re(r"test"),     "Test"),
    # Exemple : "perf(db): cache"           → "Performance (db) : cache"
    (_prefix_re(r"perf"),     "Performance"),
    # Exemple : "ci(github): lint"          → "Intégration continue (github) : lint"
    (_prefix_re(r"ci"),       "Intégration continue"),
    # Exemple : "build(webpack): upgrade"   → "Build (webpack) : mise à niveau"
    (_prefix_re(r"build"),    "Build"),
    # Exemple : "revert: annulation merge"  → "Annulation : fusion"
    (_prefix_re(r"revert"),   "Annulation"),
    # Exemple : "wip(router): refacto"      → "En cours (router) : refacto"
    (_prefix_re(r"wip"),      "En cours"),
    # Exemple : "hotfix(auth): token"       → "Correctif urgent (auth) : jeton d'authentification"
    (_prefix_re(r"hotfix"),   "Correctif urgent"),
    # Exemple : "deploy(prod): v2.3"        → "Déploiement (prod) : v2.3"
    (_prefix_re(r"deploy"),   "Déploiement"),
    # Exemple : "release(v1.2): changelog"  → "Nouvelle version (v1.2) : journal des modifications"
    (_prefix_re(r"release"),  "Nouvelle version"),
]


def _replace_prefix(text: str) -> str:
    """Remplace le préfixe de commit conventionnel, en conservant le scope si présent."""
    for pattern, label in PREFIX_MAP:
        m = re.match(pattern, text, flags=re.IGNORECASE)
        if m:
            scope   = m.group(1)  # None si pas de (scope)
            reste   = text[m.end():]
            if scope:
                return f"{label} ({scope}) : {reste}"
            else:
                return f"{label} : {reste}"
    return text


TRANSLATIONS = [

    # ── Mots isolés dans le corps du texte ────────────────────────────────────
    # Exemple : "correction du bug de paiement"  →  "correction de l'anomalie de paiement"
    (r"\bbug\b",              "anomalie"),
    # Exemple : "hotfix sur le token"  →  "correctif urgent sur le jeton d'authentification"
    (r"\bhotfix\b",           "correctif urgent"),
    # Exemple : "merge de la branche feature"  →  "fusion de la branche fonctionnalité"
    (r"\bmerge\b",            "fusion"),
    # Exemple : "création d'une nouvelle branch"  →  "création d'une nouvelle branche"
    (r"\bbranch\b",           "branche"),
    # Exemple : "rollback suite à l'erreur"  →  "retour arrière suite à l'erreur"
    (r"\brollback\b",         "retour arrière"),
    # Exemple : "nouveau deploy en production"  →  "nouveau déploiement en production"
    (r"\bdeploy(?:ment)?\b",  "déploiement"),
    # Exemple : "la pipeline CI a échoué"  →  "la chaîne de traitement CI a échoué"
    (r"\bpipeline\b",         "chaîne de traitement"),
    # Exemple : "revue de la release v2"  →  "revue de la version publiée v2"
    (r"\brelease\b",          "version publiée"),
    # Exemple : "mise à jour du readme"  →  "mise à jour du lisez-moi"
    (r"\breadme\b",           "lisez-moi"),
    # Exemple : "mise à jour du changelog"  →  "mise à jour du journal des modifications"
    (r"\bchangelog\b",        "journal des modifications"),
    # Exemple : "modification de la config"  →  "modification de la configuration"
    (r"\bconfig\b",           "configuration"),
    # Exemple : "ajout variable d'env"  →  "ajout variable d'environnement"
    (r"\benv\b",              "environnement"),
    # Exemple : "script de setup"  →  "script de mise en place"
    (r"\bsetup\b",            "mise en place"),
    # Exemple : "install des dépendances"  →  "installation des dépendances"
    (r"\binstall\b",          "installation"),
    # Exemple : "update librairie axios"  →  "mise à jour librairie axios"
    (r"\bupdate\b",           "mise à jour"),
    # Exemple : "upgrade node 18→20"  →  "mise à niveau node 18→20"
    (r"\bupgrade\b",          "mise à niveau"),
    # Exemple : "amélioration backend API"  →  "amélioration serveur interface"
    (r"\bbackend\b",          "serveur"),
    # Exemple : "correction frontend"  →  "correction interface utilisateur"
    (r"\bfrontend\b",         "interface utilisateur"),
    # Exemple : "nouveau composant UI"  →  "nouveau composant interface"
    (r"\bui\b",               "interface"),
    # Exemple : "amélioration UX formulaire"  →  "amélioration expérience utilisateur formulaire"
    (r"\bux\b",               "expérience utilisateur"),
    # Exemple : "optimisation requête DB"  →  "optimisation requête base de données"
    (r"\b(?:db|database)\b",  "base de données"),
    # Exemple : "lenteur sur la query"  →  "lenteur sur la requête"
    (r"\bquery\b",            "requête"),
    # Exemple : "invalidation du cache"  →  "invalidation du cache"),
    (r"\bcache\b",            "cache"),
    # Exemple : "refresh du token expiré"  →  "refresh du jeton d'authentification expiré"
    (r"\btoken\b",            "jeton d'authentification"),
    # Exemple : "module auth incomplet"  →  "module authentification incomplet"
    (r"\bauth\b",             "authentification"),
    # Exemple : "page de login"  →  "page de connexion"
    (r"\blogin\b",            "connexion"),
    # Exemple : "bouton logout manquant"  →  "bouton déconnexion manquant"
    (r"\blogout\b",           "déconnexion"),
    # Exemple : "reset password oublié"  →  "reset mot de passe oublié"
    (r"\bpassword\b",         "mot de passe"),
    # Exemple : "ouverture d'une modal"  →  "ouverture d'une fenêtre modale"
    (r"\bmodal\b",            "fenêtre modale"),
    # Exemple : "nouveau button valider"  →  "nouveau bouton valider"
    (r"\bbutton\b",           "bouton"),
    # Exemple : "validation du form"  →  "validation du formulaire"
    (r"\bform\b",             "formulaire"),
    # Exemple : "champ input vide"  →  "champ de saisie vide"
    (r"\binput\b",            "champ de saisie"),
    # Exemple : "fix router broken"  →  "correction routeur broken"
    (r"\brouter\b",           "routeur"),
    # Exemple : "nouveau component DatePicker"  →  "nouveau composant DatePicker"
    (r"\bcomponent\b",        "composant"),
    # Exemple : "extraction dans un module"  →  "extraction dans un module"
    (r"\bmodule\b",           "module"),
    # Exemple : "appel au service externe"  →  "appel au service externe"
    (r"\bservice\b",          "service"),
    # Exemple : "erreur dans le controller"  →  "erreur dans le contrôleur"
    (r"\bcontroller\b",       "contrôleur"),
    # Exemple : "update du schema Prisma"  →  "update du schéma Prisma"
    (r"\bschema\b",           "schéma"),
    # Exemple : "nouvelle migration SQL"  →  "nouvelle migration SQL"
    (r"\bmigration\b",        "migration"),
    # Exemple : "ajout seed pour les tests"  →  "ajout données initiales pour les tests"
    (r"\bseed\b",             "données initiales"),
    # Exemple : "mock du service mail"  →  "simulation du service mail"
    (r"\bmock\b",             "simulation"),
    # Exemple : "stub API externe"  →  "bouchon API externe"
    (r"\bstub\b",             "bouchon"),
    # Exemple : "vérification des logs"  →  "vérification du journal"
    (r"\blogs?\b",            "journal"),
    # Exemple : "erreur 500 non capturée"  →  "erreur 500 non capturée"
    (r"\berror\b",            "erreur"),
    # Exemple : "warning deprecation"  →  "avertissement deprecation"
    (r"\bwarning\b",          "avertissement"),
    # Exemple : "timeout sur l'appel API"  →  "délai dépassé sur l'appel API"
    (r"\btimeout\b",          "délai dépassé"),
    # Exemple : "logique de retry"  →  "logique de nouvelle tentative"
    (r"\bretry\b",            "nouvelle tentative"),
    # Exemple : "ajout d'un tag v2.0"  →  "ajout d'une étiquette v2.0"
    (r"\btag\b",              "étiquette"),
    # Exemple : "refacto du script bash"  →  "refacto du script bash"
    (r"\bscript\b",           "script"),
    # Exemple : "appel API paiement"  →  "appel interface paiement"
    (r"\bapi\b",              "interface"),
    # Exemple : "test de régression complet"  →  "test de régression complet"
    (r"\bregression\b",       "régression"),
    # Exemple : "modèle User mis à jour"  →  "modèle utilisateur mis à jour"
    (r"\bmodel\b",            "modèle"),
]


def humanize(text: str) -> str:
    """Traduit les termes techniques en langage métier.
    1) Remplace le préfixe de commit (feat/fix/chore... avec scope optionnel)
    2) Applique la table de mots courants sur le reste
    """
    if not text:
        return text
    result = _replace_prefix(text)
    for pattern, replacement in TRANSLATIONS:
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    return result


# ─── Parsing du temps ─────────────────────────────────────────────────────────
def parse_minutes(raw_value):
    if not raw_value:
        return 0
    value = str(raw_value).strip().strip("[]").lower()
    if value == "?":
        return 0
    compact = value.replace(" ", "")
    match_hm = re.fullmatch(r"(\d+)h(?:(\d+)(?:min)?)?", compact)
    if match_hm:
        return int(match_hm.group(1)) * 60 + int(match_hm.group(2) or 0)
    match_min = re.fullmatch(r"(\d+)min", compact)
    if match_min:
        return int(match_min.group(1))
    if compact.isdigit():
        return int(compact)
    return 0


def to_hhmm(total_minutes):
    total = int(total_minutes or 0)
    hours, minutes = divmod(total, 60)
    return f"{hours}h{minutes:02d}"


# ─── Chargement ou création du workbook ───────────────────────────────────────
excel_exists = os.path.exists(dst)
header_fill  = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
day_fill     = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
row_fill     = PatternFill(start_color="EEF5FC", end_color="EEF5FC", fill_type="solid")
total_fill   = PatternFill(start_color="D6E8FA", end_color="D6E8FA", fill_type="solid")

if excel_exists:
    try:
        wb = load_workbook(dst)
        ws = wb.active
        ws2 = wb["Tableau de bord"] if "Tableau de bord" in wb.sheetnames else None
        print("APPEND_MODE=true")
    except Exception as e:
        excel_exists = False
        print(f"EXCEL_CORRUPTED={e}")

if not excel_exists:
    wb = Workbook()
    ws = wb.active
    ws.title = "Journal"
    ws2 = None
    print("CREATE_MODE=true")

# ─── Lecture CSV ──────────────────────────────────────────────────────────────
rows = []
with open(src, newline="", encoding="utf-8") as f:
    for row in csv.reader(f):
        if not row or all(str(c).strip() == "" for c in row):
            continue
        rows.append(row)

# ─── En-tête (mode création) ──────────────────────────────────────────────────
if not excel_exists:
    header_row = rows[0] if rows else ["Date", "Nom", "Temps", "État", "Description"]
    ws.append(header_row)
    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

# ─── Groupement par jour ──────────────────────────────────────────────────────
grouped_rows = defaultdict(list)
day_order    = []

for row in rows[1:]:
    if len(row) < 5:
        continue
    date_label = str(row[0]).strip()
    if not date_label:
        continue
    if date_label not in grouped_rows:
        day_order.append(date_label)
    grouped_rows[date_label].append(row)

# ─── Remplissage de l'onglet Journal ──────────────────────────────────────────
if excel_exists and ws.max_row > 1:
    ws.delete_rows(2, ws.max_row - 1)

for day in day_order:
    # Ligne de séparateur de jour
    ws.append([day, "", "", "", ""])
    day_row_idx = ws.max_row
    ws.merge_cells(
        start_row=day_row_idx, start_column=1,
        end_row=day_row_idx,   end_column=5
    )
    day_cell            = ws.cell(row=day_row_idx, column=1)
    day_cell.font       = Font(bold=True, color="FFFFFF", name="Arial")
    day_cell.fill       = day_fill
    day_cell.alignment  = Alignment(horizontal="left")

    day_total_minutes = 0

    for entry in grouped_rows[day]:
        # ← Traduction appliquée ici sur le Nom et la Description
        nom_traduit   = humanize(str(entry[1]))
        desc_traduite = humanize(str(entry[4]))

        ws.append(["", nom_traduit, entry[2], entry[3], desc_traduite])
        data_row_idx = ws.max_row
        for col_idx in range(1, 6):
            c = ws.cell(row=data_row_idx, column=col_idx)
            c.fill = row_fill
            c.font = Font(name="Arial")
        day_total_minutes += parse_minutes(entry[2])

    # Ligne total du jour
    ws.append(["", "Total du jour",
               f"{to_hhmm(day_total_minutes)} ({day_total_minutes} min)", "", ""])
    total_row_idx = ws.max_row
    for col_idx in range(1, 6):
        ws.cell(row=total_row_idx, column=col_idx).fill = total_fill
    ws.cell(row=total_row_idx, column=2).font = Font(bold=True, name="Arial")
    ws.cell(row=total_row_idx, column=3).font = Font(bold=True, name="Arial")

    ws.append(["", "", "", "", ""])

ws.freeze_panes                = "A2"
ws.auto_filter.ref             = f"A1:E{ws.max_row}"
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 50
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 12
ws.column_dimensions["E"].width = 90

# ─── Onglet Tableau de bord ───────────────────────────────────────────────────
daily = defaultdict(lambda: {"minutes": 0})
for row in rows[1:]:
    if len(row) < 3:
        continue
    date_label = str(row[0]).strip()
    if not date_label:
        continue
    daily[date_label]["minutes"] += parse_minutes(row[2])

if ws2 is None:
    ws2 = wb.create_sheet("Tableau de bord")

ws2.delete_rows(1, ws2.max_row)
ws2.append(["Date", "Total (min)", "Total (h)", "Total (h:min)"])

sorted_days = sorted(
    daily.items(),
    key=lambda item: datetime.strptime(item[0], "%d/%m/%Y")
)

for day, stats in sorted_days:
    ws2.append([day, stats["minutes"],
                round(stats["minutes"] / 60, 2), to_hhmm(stats["minutes"])])

for col in ws2.columns:
    max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
    ws2.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

if sorted_days:
    total_minutes = sum(s["minutes"] for _, s in sorted_days)
    nb_days       = len(sorted_days)
    top_day, top_stats = max(sorted_days, key=lambda item: item[1]["minutes"])

    for cell in ws2[1]:
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    ws2.freeze_panes   = "A2"
    ws2.auto_filter.ref = f"A1:D{ws2.max_row}"

    ws2["F1"] = "Indicateur"
    ws2["G1"] = "Valeur"
    ws2["F2"] = "Jours suivis"
    ws2["G2"] = nb_days
    ws2["F3"] = "Temps total"
    ws2["G3"] = to_hhmm(total_minutes)
    ws2["F4"] = "Moyenne / jour"
    ws2["G4"] = to_hhmm(round(total_minutes / nb_days))
    ws2["F5"] = "Jour le plus chargé"
    ws2["G5"] = f"{top_day} ({to_hhmm(top_stats['minutes'])})"

    for row in ws2["F1:G1"]:
        for cell in row:
            cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center")
    for row_idx in range(2, 6):
        ws2[f"F{row_idx}"].font = Font(bold=True, name="Arial")

    max_row  = ws2.max_row
    date_ref = Reference(ws2, min_col=1, min_row=2, max_row=max_row)

    bar      = BarChart()
    bar.title           = "Charge par jour (heures)"
    bar.y_axis.title    = "Heures"
    bar.x_axis.title    = "Date"
    bar_data            = Reference(ws2, min_col=3, min_row=1, max_row=max_row)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(date_ref)
    bar.style  = 10
    bar.height = 8
    bar.width  = 16
    ws2.add_chart(bar, "M8")

    ws2["I1"] = "Date"
    ws2["J1"] = "Cumul (h)"
    ws2["I1"].font = Font(bold=True, name="Arial")
    ws2["J1"].font = Font(bold=True, name="Arial")

    running_minutes = 0
    for idx, (day, stats) in enumerate(sorted_days, start=2):
        running_minutes    += stats["minutes"]
        ws2[f"I{idx}"]      = day
        ws2[f"J{idx}"]      = round(running_minutes / 60, 2)

    line              = LineChart()
    line.title        = "Cumul des heures"
    line.y_axis.title = "Heures"
    line.x_axis.title = "Date"
    line_data         = Reference(ws2, min_col=10, min_row=1, max_row=max_row)
    line.add_data(line_data, titles_from_data=True)
    line.set_categories(date_ref)
    line.smooth       = True
    if line.series:
        line.series[0].marker = None
    line.height = 8
    line.width  = 16
    ws2.add_chart(line, "M26")

# ─── Sauvegarde ───────────────────────────────────────────────────────────────
try:
    wb.save(dst)
except PermissionError:
    base, ext  = os.path.splitext(dst)
    fallback   = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    wb.save(fallback)
    print(f"FALLBACK_XLSX={fallback}")
except Exception as exc:
    print(f"SAVE_ERROR={exc}")
    sys.exit(4)

PYEOF
)

    PY_STATUS=$?
    if [ $PY_STATUS -eq 0 ]; then
        if grep -q '^FALLBACK_XLSX=' <<< "$PY_OUTPUT"; then
            FALLBACK_FILE=$(grep '^FALLBACK_XLSX=' <<< "$PY_OUTPUT" | sed 's/^FALLBACK_XLSX=//')
            echo "⚠️  Le fichier principal est verrouillé (probablement ouvert)."
            echo "✅ Excel généré avec un nom alternatif : $FALLBACK_FILE"
        elif grep -q '^APPEND_MODE=true' <<< "$PY_OUTPUT"; then
            echo "♻️  Excel mis à jour (mode append) : $OUTPUT_XLSX"
        elif grep -q '^CREATE_MODE=true' <<< "$PY_OUTPUT"; then
            echo "✅ Nouvel Excel généré : $OUTPUT_XLSX"
        else
            echo "✅ Excel généré : $OUTPUT_XLSX"
        fi
    elif [ $PY_STATUS -eq 2 ]; then
        echo "⚠️  Module Python 'openpyxl' manquant. Installe-le avec :"
        echo "   pip install openpyxl"
    else
        [ -n "$PY_OUTPUT" ] && echo "$PY_OUTPUT"
        echo "⚠️  Erreur lors de la génération Excel. Vérifie que le fichier .xlsx n'est pas ouvert."
    fi
else
    echo "⚠️  Python introuvable, conversion Excel ignorée."
fi

echo "✨ C'est fini !"
