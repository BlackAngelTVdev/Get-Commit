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
    echo "Date,Nom,Temps,État,Description,Type" > "$OUTPUT_FILE"
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

    gsub(/\r|\n/, "",   current_date)
    gsub(/^ +| +$/, "", current_date)
    gsub(/\r|\n/, " ", nom_commit)
    gsub(/\r|\n/, " ", body)

    # ── Nettoyage du body ──────────────────────────────────────────────────
    body_clean = body
    gsub(/^ +| +$/, "", body_clean)

    # Git répète souvent le sujet en tête du body → le supprimer
    nom_trim = nom_commit
    gsub(/^ +| +$/, "", nom_trim)
    if (length(nom_trim) > 0 && index(body_clean, nom_trim) == 1) {
        body_clean = substr(body_clean, length(nom_trim) + 1)
    }

    # Si le body est entièrement contenu dans le nom (sous-chaîne), vider
    if (length(body_clean) > 0 && index(nom_trim, body_clean) > 0) {
        body_clean = ""
    }

    # Supprimer séparateurs parasites en début : tirets, pipes, espaces
    gsub(/^[ \t]*[-–—|]+[ \t]*/, "", body_clean)
    gsub(/^[ \t]+/, "", body_clean)

    # Remplacer les tirets séparateurs internes par " | " (plus lisible)
    gsub(/[ \t]+-[ \t]+/, " | ", body_clean)

    # Nettoyer les espaces multiples
    gsub(/  +/, " ", body_clean)
    gsub(/^ +| +$/, "", body_clean)

    # ── Extraction Temps + État depuis nom + body combinés ─────────────────
    full_text = nom_commit " " body
    temps = "[?]"
    etat  = "[?]"

    if (match(full_text, /\[[0-9hmin]+\]/)) {
        temps = substr(full_text, RSTART, RLENGTH)
    }
    if (match(full_text, /\[[A-Z]+\]/)) {
        etat = substr(full_text, RSTART, RLENGTH)
    }

    # Retirer les tags du nom et de la description
    gsub(/\[[0-9hmin]+\]/, "", nom_commit)
    gsub(/\[[A-Z]+\]/,     "", nom_commit)
    gsub(/\[[0-9hmin]+\]/, "", body_clean)
    gsub(/\[[A-Z]+\]/,     "", body_clean)
    gsub(/^ +| +$/, "", nom_commit)
    gsub(/^ +| +$/, "", body_clean)
    gsub(/  +/, " ", nom_commit)
    gsub(/  +/, " ", body_clean)

    # Échappement CSV
    gsub(/"/, "\"\"", nom_commit)
    gsub(/"/, "\"\"", body_clean)

    entry_key = current_date "|" nom_commit
    if (entry_key in existing) { next }

    if (last_date != "" && last_date != current_date) {
        print "" >> "'$OUTPUT_FILE'"
    }
    last_date = current_date

    printf "\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n",
        current_date, nom_commit, temps, etat, body_clean >> "'$OUTPUT_FILE'"
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
# ──────────────────────────────────────────────────────────────────────────────

_SCOPE = r"(?:\(([^)]*)\))?"   # capture optionnelle du scope
_BANG  = r"[!]?"               # breaking change optionnel
_SEP   = r":\s*"               # deux-points + espaces


def _prefix_re(keyword: str) -> str:
    """Construit la regex pour un préfixe de commit conventionnel."""
    return rf"^{keyword}{_SCOPE}{_BANG}{_SEP}"


# Table des préfixes → label métier
PREFIX_MAP = [
    (_prefix_re(r"feat"),     "Fonctionnalité"),
    (_prefix_re(r"fix"),      "Correction"),
    (_prefix_re(r"chore?"),   "Tâche"),
    (_prefix_re(r"docs?"),    "Documentation"),
    (_prefix_re(r"refactor"), "Refactorisation"),
    (_prefix_re(r"style"),    "Mise en forme"),
    (_prefix_re(r"test"),     "Test"),
    (_prefix_re(r"perf"),     "Performance"),
    (_prefix_re(r"ci"),       "Intégration continue"),
    (_prefix_re(r"build"),    "Build"),
    (_prefix_re(r"revert"),   "Annulation"),
    (_prefix_re(r"wip"),      "En cours"),
    (_prefix_re(r"hotfix"),   "Correctif urgent"),
    (_prefix_re(r"deploy"),   "Déploiement"),
    (_prefix_re(r"release"),  "Nouvelle version"),
]


def split_commit(raw_nom: str) -> tuple:
    for pattern, label in PREFIX_MAP:
        m = re.match(pattern, raw_nom, flags=re.IGNORECASE)
        if m:
            titre = raw_nom[m.end():].strip()
            return label, titre if titre else (m.group(1) or "")

    for pattern, label in PREFIX_MAP:
        kw_match = re.search(r'\^(\w+)', pattern)
        if kw_match:
            kw = kw_match.group(1)
            scope_m = re.match(rf"^{kw}\(([^)]+)\)\s*$", raw_nom, flags=re.IGNORECASE)
            if scope_m:
                return label, scope_m.group(1)

    return "", raw_nom.strip()


def humanize(text: str) -> str:
    if not text:
        return text
    result = text
    for pattern, replacement in TRANSLATIONS:
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    return result


TRANSLATIONS = [
    (r"\bfeat\b",             "Fonctionnalité"),
    (r"\bfix\b",              "Correction"),
    (r"\bchore?\b",           "Tâche"),
    (r"\bdocs?\b",            "Documentation"),
    (r"\brefactor\b",         "Refactorisation"),
    (r"\bstyle\b",            "Mise en forme"),
    (r"\btest\b",             "Test"),
    (r"\bperf\b",             "Performance"),
    (r"\bci\b",               "Intégration continue"),
    (r"\bbuild\b",            "Build"),
    (r"\brevert\b",           "Annulation"),
    (r"\bwip\b",              "En cours"),
    (r"\bhotfix\b",           "Correctif urgent"),
    (r"\bdeploy\b",           "Déploiement"),
    (r"\brelease\b",          "Nouvelle version"),
    (r"\bbug\b",              "anomalie"),
    (r"\bhotfix\b",           "correctif urgent"),
    (r"\bmerge\b",            "fusion"),
    (r"\bbranch\b",           "branche"),
    (r"\brollback\b",         "retour arrière"),
    (r"\bdeploy(?:ment)?\b",  "déploiement"),
    (r"\bpipeline\b",         "chaîne de traitement"),
    (r"\brelease\b",          "version publiée"),
    (r"\breadme\b",           "lisez-moi"),
    (r"\bchangelog\b",        "journal des modifications"),
    (r"\bconfig\b",           "configuration"),
    (r"\benv\b",              "environnement"),
    (r"\bsetup\b",            "mise en place"),
    (r"\binstall\b",          "installation"),
    (r"\bupdate\b",           "mise à jour"),
    (r"\bupgrade\b",          "mise à niveau"),
    (r"\bbackend\b",          "serveur"),
    (r"\bfrontend\b",         "interface utilisateur"),
    (r"\bui\b",               "interface"),
    (r"\bux\b",               "expérience utilisateur"),
    (r"\b(?:db|database)\b",  "base de données"),
    (r"\bquery\b",            "requête"),
    (r"\bcache\b",            "cache"),
    (r"\btoken\b",            "jeton d'authentification"),
    (r"\bauth\b",             "authentification"),
    (r"\blogin\b",            "connexion"),
    (r"\blogout\b",           "déconnexion"),
    (r"\bpassword\b",         "mot de passe"),
    (r"\bmodal\b",            "fenêtre modale"),
    (r"\bbutton\b",           "bouton"),
    (r"\bform\b",             "formulaire"),
    (r"\binput\b",            "champ de saisie"),
    (r"\brouter\b",           "routeur"),
    (r"\bcomponent\b",        "composant"),
    (r"\bmodule\b",           "module"),
    (r"\bservice\b",          "service"),
    (r"\bcontroller\b",       "contrôleur"),
    (r"\bschema\b",           "schéma"),
    (r"\bmigration\b",        "migration"),
    (r"\bseed\b",             "données initiales"),
    (r"\bmock\b",             "simulation"),
    (r"\bstub\b",             "bouchon"),
    (r"\blogs?\b",            "journal"),
    (r"\berror\b",            "erreur"),
    (r"\bwarning\b",          "avertissement"),
    (r"\btimeout\b",          "délai dépassé"),
    (r"\bretry\b",            "nouvelle tentative"),
    (r"\btag\b",              "étiquette"),
    (r"\bscript\b",           "script"),
    (r"\bapi\b",              "interface"),
    (r"\bregression\b",       "régression"),
    (r"\bmodel\b",            "modèle"),
]


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

# ─── En-tête (Structure stricte à 5 colonnes avec Description en C) ───────────
if not excel_exists:
    ws.append(["Date", "Type", "Description", "Temps", "État"])
    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")
else:
    # Si le fichier existe déjà, on s'assure que la colonne C s'appelle "Description"
    ws.cell(row=1, column=3).value = "Description"

# Couleurs par type pour les badges
TYPE_COLORS = {
    "Fonctionnalité":      "2E7D32",
    "Correction":          "C62828",
    "Tâche":               "6A1B9A",
    "Documentation":       "1565C0",
    "Refactorisation":     "E65100",
    "Mise en forme":       "00838F",
    "Test":                "4527A0",
    "Performance":         "558B2F",
    "Intégration continue":"37474F",
    "Build":               "4E342E",
    "Annulation":          "757575",
    "En cours":            "F57F17",
    "Correctif urgent":    "B71C1C",
    "Déploiement":         "1A237E",
    "Nouvelle version":    "004D40",
    "":                    "455A64",
}

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

# Supprime l'ancienne colonne F résiduelle (6ème colonne) des en-têtes du modèle
if ws.max_column >= 6:
    ws.delete_cols(6)

for day in day_order:
    # Ligne de séparateur de jour (5 colonnes au total)
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
        raw_nom = str(entry[1])
        raw_desc = str(entry[4]) if len(entry) > 4 else ""

        type_label, titre = split_commit(raw_nom)
        desc_finale = raw_desc if raw_desc.strip() else titre
        desc_traduite = humanize(desc_finale)

        temps_valeur = entry[2]
        etat_valeur  = entry[3]

        type_color = TYPE_COLORS.get(type_label, TYPE_COLORS[""])
        type_fill  = PatternFill(start_color=type_color, end_color=type_color, fill_type="solid")

        # Insertion des données sur 5 colonnes (C = Description)
        ws.append(["", type_label, desc_traduite, temps_valeur, etat_valeur])
        
        data_row_idx = ws.max_row
        for col_idx in range(1, 6):
            c = ws.cell(row=data_row_idx, column=col_idx)
            c.fill = row_fill
            c.font = Font(name="Arial")
            
        type_cell       = ws.cell(row=data_row_idx, column=2)
        type_cell.fill  = type_fill
        type_cell.font  = Font(bold=True, color="FFFFFF", name="Arial")
        type_cell.alignment = Alignment(horizontal="center")

        day_total_minutes += parse_minutes(temps_valeur)

    # Ligne total du jour (5 colonnes)
    ws.append(["", "", "Total du jour", f"{to_hhmm(day_total_minutes)} ({day_total_minutes} min)", ""])
    total_row_idx = ws.max_row
    for col_idx in range(1, 6):
        ws.cell(row=total_row_idx, column=col_idx).fill = total_fill
    ws.cell(row=total_row_idx, column=3).font = Font(bold=True, name="Arial")
    ws.cell(row=total_row_idx, column=4).font = Font(bold=True, name="Arial")

    ws.append(["", "", "", "", ""])

ws.freeze_panes                 = "A2"
ws.auto_filter.ref              = f"A1:E{ws.max_row}"

# Redimensionnement propre des 5 colonnes
ws.column_dimensions["A"].width = 14   # Date
ws.column_dimensions["B"].width = 18   # Type
ws.column_dimensions["C"].width = 75   # Description
ws.column_dimensions["D"].width = 16   # Temps
ws.column_dimensions["E"].width = 12   # État

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