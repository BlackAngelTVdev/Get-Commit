#!/bin/bash

#
# Journal de travail Git → CSV → Excel
#
# Conforme au livrable « journal de travail individuel » :
#   • chaque commit significatif est référencé (date + hash court + fichiers/répertoires)
#   • l AUTEUR du commit est indiqué (qui a fait quoi)
#   • durée et état de la tâche sont extraits des tags ([1h30] [DONE])
#   • colonnes à remplir à la main par jour :
#       activités planifiées, aide extérieure reçue,
#       imprévus / heures supplémentaires, succès, échecs,
#       appréciation critique et réflexion personnelle
#
# Mode APPEND : complète le CSV avec les nouveaux commits uniquement.
# Les colonnes de notes remplies à la main sont préservées.
#
# Utilisation :
#   bash ./get_commit.sh                   # met à jour le journal
#   SINCE=2026-01-13 bash ./get_commit.sh  # force une date de départ
#

# ─── Config ───────────────────────────────────────────────────────────────────
OUTPUT_DIR="Doc"
OUTPUT_FILE="$OUTPUT_DIR/journal_de_travail.csv"
OUTPUT_XLSX="$OUTPUT_DIR/journal_de_travail.xlsx"
MAX_FILES=20          # nb max de fichiers listés par commit avant résumé

CSV_HEADER='Date,Auteur,Commit,Type,Description,Temps,État,Fichiers,Activités planifiées,Aide extérieure,Imprévus / heures supplémentaires,Succès,Échecs,Appréciation critique / réflexion'

SINCE_FROM="${SINCE:-}"

mkdir -p "$OUTPUT_DIR"

echo "🚀 Extraction en mode append..."

# ─── 1. Migration / création du CSV ───────────────────────────────────────────
REBUILD=0
if [ -f "$OUTPUT_FILE" ]; then
    # En-tête normalisé : Excel / LibreOffice peuvent ajouter des guillemets et des CR.
    CURRENT_HEADER=$(head -n 1 "$OUTPUT_FILE" | tr -d '\r' | sed 's/"//g')
    if [ "$CURRENT_HEADER" != "$CSV_HEADER" ]; then
        BACKUP="${OUTPUT_FILE}.bak_$(date +%Y%m%d_%H%M%S)"
        mv "$OUTPUT_FILE" "$BACKUP"
        REBUILD=1
        echo "♻️  Ancien format de CSV détecté."
        echo "    Sauvegarde de l'ancien fichier : $BACKUP"
        echo "    Reconstruction du CSV depuis l'historique Git (auteur, commit, fichiers)."
    fi
fi

if [ ! -f "$OUTPUT_FILE" ]; then
    printf '%s\n' "$CSV_HEADER" > "$OUTPUT_FILE"
    echo "✅ Nouveau fichier CSV créé : $OUTPUT_FILE"
fi

# ─── 2. Point de départ de l'extraction ───────────────────────────────────────
SINCE=""
if [ "$REBUILD" -eq 0 ]; then
    LAST_DATE=$(awk -F',' 'NR > 1 {
        d = $1
        gsub(/"/, "", d)
        if (d != "") last = d
    } END { print last }' "$OUTPUT_FILE" 2>/dev/null)

    if [ -n "$LAST_DATE" ]; then
        if [[ "$LAST_DATE" =~ ^([0-9]{2})/([0-9]{2})/([0-9]{4})$ ]]; then
            SINCE="${BASH_REMATCH[3]}-${BASH_REMATCH[2]}-${BASH_REMATCH[1]} 00:00:00"
            echo "ℹ️  Fichier existant. Dernière date enregistrée : $LAST_DATE"
        fi
    fi
fi

if [ -n "$SINCE_FROM" ]; then
    SINCE="$SINCE_FROM"
fi

if ! git rev-parse --git-dir >/dev/null 2>&1; then
    echo "❌ Ce dossier n'est pas un dépôt Git. Lance le script dans un dépôt (dossier .git)."
    exit 1
fi

GIT_LOG_ARGS=(log --all --reverse --use-mailmap
              --date=format:%d/%m/%Y
              --pretty=format:%x1e%H%x1f%ad%x1f%an%x1f%s%x1f%b%x1f
              --name-only)

if [ -n "$SINCE" ]; then
    GIT_LOG_ARGS+=(--since="$SINCE")
fi

# ─── 3. Extraction git log → CSV ──────────────────────────────────────────────
git "${GIT_LOG_ARGS[@]}" \
| awk -v RS="$(printf '\036')" -v FS="$(printf '\037')" \
      -v csvfile="$OUTPUT_FILE" -v out="$OUTPUT_FILE" -v maxf="$MAX_FILES" '

# ── Échappement CSV ───────────────────────────────────────────────────────────
function csvq(v) {
    gsub(/"/, "\"\"", v)
    return "\"" v "\""
}

# ── Lecture dune ligne CSV en respectant les guillemets ──────────────────────
function csv_split(line, parts,   i, ch, nx, field, inq, n) {
    n = 0; field = ""; inq = 0
    for (i = 1; i <= length(line); i++) {
        ch = substr(line, i, 1)
        if (inq) {
            if (ch == "\"") {
                nx = substr(line, i + 1, 1)
                if (nx == "\"") { field = field "\""; i++ } else { inq = 0 }
            } else {
                field = field ch
            }
        } else if (ch == "\"") {
            inq = 1
        } else if (ch == ",") {
            parts[++n] = field; field = ""
        } else {
            field = field ch
        }
    }
    parts[++n] = field
    return n
}

# ── Convention de commit : préfixe → libellé métier ──────────────────────────
function prefix_label(k) {
    k = tolower(k)
    if (k == "feat")                        return "Fonctionnalité"
    if (k == "fix")                         return "Correction"
    if (k == "chore" || k == "chor")        return "Tâche"
    if (k == "doc"   || k == "docs")        return "Documentation"
    if (k == "refactor")                    return "Refactorisation"
    if (k == "style")                       return "Mise en forme"
    if (k == "test"  || k == "tests")       return "Test"
    if (k == "perf")                        return "Performance"
    if (k == "ci")                          return "Intégration continue"
    if (k == "build")                       return "Build"
    if (k == "revert")                      return "Annulation"
    if (k == "wip")                         return "En cours"
    if (k == "hotfix")                      return "Correctif urgent"
    if (k == "deploy")                      return "Déploiement"
    if (k == "release")                     return "Nouvelle version"
    return ""
}

# Extrait type(scope)! et renvoie le libellé métier (scope conservé)
function commit_type(subject,   kw, rest, label, cp, scope) {
    rest = subject
    if (!match(rest, /^[A-Za-z]+/)) return "Autre"
    kw = substr(rest, 1, RLENGTH)
    rest = substr(rest, RLENGTH + 1)
    scope = ""
    if (substr(rest, 1, 1) == "(") {
        cp = index(rest, ")")
        if (cp > 1) { scope = substr(rest, 2, cp - 2); rest = substr(rest, cp + 1) }
    }
    if (substr(rest, 1, 1) == "!") rest = substr(rest, 2)
    if (substr(rest, 1, 1) != ":") return "Autre"
    label = prefix_label(kw)
    if (label == "") return "Autre"
    if (scope != "") return label " (" scope ")"
    return label
}

# Retire le préfixe conventionnel du sujet
function strip_prefix(subject,   kw, rest, cp) {
    rest = subject
    if (!match(rest, /^[A-Za-z]+/)) return subject
    kw = substr(rest, 1, RLENGTH)
    rest = substr(rest, RLENGTH + 1)
    if (substr(rest, 1, 1) == "(") {
        cp = index(rest, ")")
        if (cp > 1) rest = substr(rest, cp + 1)
    }
    if (substr(rest, 1, 1) == "!") rest = substr(rest, 2)
    if (substr(rest, 1, 1) != ":") return subject
    if (prefix_label(kw) == "") return subject
    rest = substr(rest, 2)
    sub(/^[ \t]+/, "", rest)
    return rest
}

# Index des commits déjà présents dans le CSV (anti-doublons)
BEGIN {
    # Le CSV est un fichier ligne par ligne : RS (séparateur de commits) ne doit
    # pas s appliquer ici, sinon tout le fichier est lu comme un seul enregistrement.
    commit_rs = RS
    RS = "\n"
    while ((getline line < csvfile) > 0) {
        if (line ~ /^[ \t]*$/) continue
        cnt = csv_split(line, F)
        if (cnt < 3) continue
        if (F[1] == "Date") continue
        if (F[3] != "") {
            seen_sha[substr(F[3], 1, 8)] = 1
        } else {
            legacy_key[F[1] "|" F[2] "|" substr(F[4], 1, 60)] = 1
        }
    }
    close(csvfile)
    RS = commit_rs
}

NF {
    sha     = $1
    date    = $2
    author  = $3
    subject = $4
    body    = $5
    files   = $6

    gsub(/\r|\n/, "", date);       gsub(/^[ \t]+|[ \t]+$/, "", date)
    gsub(/\r|\n/, " ", author);    gsub(/^[ \t]+|[ \t]+$/, "", author)
    gsub(/[\r\n]+/, " ", subject); gsub(/^[ \t]+|[ \t]+$/, "", subject)
    gsub(/[\r\n]+/, " ", body);    gsub(/^[ \t]+|[ \t]+$/, "", body)

    if (date == "" || sha == "") next
    sha8 = substr(sha, 1, 8)
    if (author == "") author = "(inconnu)"

    # ── Nettoyage du corps du message ────────────────────────────────────────
    if (body != "" && index(body, subject) == 1) {
        body = substr(body, length(subject) + 1)
        gsub(/^[ \t]+/, "", body)
    }
    if (body != "" && index(subject, body) > 0) body = ""
    gsub(/^[ \t]*[-–—|]+[ \t]*/, "", body)
    gsub(/[ \t]+-[ \t]+/, " | ", body)
    gsub(/  +/, " ", body)
    gsub(/^[ \t]+|[ \t]+$/, "", body)

    # ── Durée + état (tags de la convention de commit) ───────────────────────
    full_text = subject " " body
    temps = "[?]"
    etat  = "[?]"
    if (match(full_text, /\[[0-9]+ *(h[0-9]*[a-z]*|min)?\]/)) {
        temps = substr(full_text, RSTART, RLENGTH)
    }
    if (match(full_text, /\[[A-Z][A-Z_-]*\]/)) {
        etat = substr(full_text, RSTART, RLENGTH)
    }

    # ── Type + description ───────────────────────────────────────────────────
    type_label = commit_type(subject)
    desc = strip_prefix(subject)
    if (body != "") {
        desc = (desc == "") ? body : desc " — " body
    }
    gsub(/\[[0-9]+ *(h[0-9]*[a-z]*|min)?\]/, "", desc)
    gsub(/\[[A-Z][A-Z_-]*\]/, "", desc)
    gsub(/  +/, " ", desc)
    gsub(/^[ \t]+|[ \t]+$/, "", desc)
    sub(/[ \t]*[—–|\-]+[ \t]*$/, "", desc)
    gsub(/^[ \t]+|[ \t]+$/, "", desc)
    if (desc == "") desc = type_label

    # ── Fichiers / répertoires concernés ─────────────────────────────────────
    gsub(/\r/, "", files)
    sub(/^\n+/, "", files)
    sub(/\n+$/, "", files)
    nbfiles = (files == "") ? 0 : split(files, flist, "\n")
    joined = ""
    for (i = 1; i <= nbfiles && i <= maxf; i++) {
        gsub(/^[ \t]+|[ \t]+$/, "", flist[i])
        if (flist[i] == "") continue
        joined = (joined == "") ? flist[i] : joined " ; " flist[i]
    }
    if (nbfiles > maxf) joined = joined " ; (+" (nbfiles - maxf) " autres)"

    # ── Anti-doublons ────────────────────────────────────────────────────────
    # Un commit = un hash : deux commits distincts (rebase, cherry-pick) restent
    # référencés même si le message et le jour sont identiques.
    if (sha8 in seen_sha) next
    if ((date "|" author "|" substr(desc, 1, 60)) in legacy_key) next

    if (last_date != "" && last_date != date) print "" >> out
    last_date = date

    printf "%s,%s,%s,%s,%s,%s,%s,%s,\"\",\"\",\"\",\"\",\"\",\"\"\n",
        csvq(date), csvq(author), csvq(sha8), csvq(type_label),
        csvq(desc), csvq(temps), csvq(etat), csvq(joined) >> out

    seen_sha[sha8] = 1
}
'

echo "📄 CSV à jour : $OUTPUT_FILE"

if head -n 1 "$OUTPUT_FILE" | grep -q '^Date,'; then
    echo "👥 Auteurs référencés dans le journal :"
    awk -F',' 'NR > 1 {
        a = $2
        gsub(/"/, "", a)
        if (a != "") print a
    }' "$OUTPUT_FILE" | sort | uniq -c | sort -rn | sed 's/^/   /'
fi

# ─── 4. Conversion CSV → Excel ────────────────────────────────────────────────
if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_BIN="python"
else
    PYTHON_BIN=""
fi

if [ -z "$PYTHON_BIN" ]; then
    echo "⚠️  Python introuvable, conversion Excel ignorée."
    echo "✨ C'est fini !"
    exit 0
fi

PY_OUTPUT=$(OUTPUT_FILE="$OUTPUT_FILE" OUTPUT_XLSX="$OUTPUT_XLSX" "$PYTHON_BIN" - << 'PYEOF' 2>&1

import csv
import os
import re
import sys
from collections import OrderedDict, defaultdict
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
# Structure du journal (identique au CSV)
# ──────────────────────────────────────────────────────────────────────────────
HEADERS = ["Date", "Auteur", "Type", "Description", "Temps", "État", "Commit", "Fichiers"]

# Colonnes libres à remplir à la main (index dans la ligne CSV)
NOTE_FIELDS = [
    (8,  "Activités planifiées"),
    (9,  "Aide extérieure reçue"),
    (10, "Imprévus / heures supplémentaires"),
    (11, "Succès"),
    (12, "Échecs"),
    (13, "Appréciation critique et réflexion personnelle"),
]

# ──────────────────────────────────────────────────────────────────────────────
# Traduction des termes techniques → langage métier
# ──────────────────────────────────────────────────────────────────────────────
TRANSLATIONS = [
    (r"\bbug\b",              "anomalie"),
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
    (r"\bapi\b",              "interface"),
    (r"\bregression\b",       "régression"),
    (r"\bmodel\b",            "modèle"),
]


def humanize(text):
    if not text:
        return text
    result = text
    for pattern, replacement in TRANSLATIONS:
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    return result


# ─── Parsing de la durée ─────────────────────────────────────────────────────
def parse_minutes(raw_value):
    if raw_value is None:
        return 0
    value = str(raw_value).strip().strip("[]").lower()
    if value in ("", "?"):
        return 0
    compact = value.replace(" ", "")
    match_hm = re.fullmatch(r"(\d+)h(?:(?:(\d+)(?:min)?))?", compact)
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


# ─── Lecture du CSV ──────────────────────────────────────────────────────────
rows = []
with open(src, newline="", encoding="utf-8") as f:
    for row in csv.reader(f):
        if not row or all(str(c).strip() == "" for c in row):
            continue
        if len(row) < 14:
            row = row + [""] * (14 - len(row))
        rows.append(row)

data = rows[1:] if rows and str(rows[0][0]).strip() == "Date" else rows

# ─── Styles ──────────────────────────────────────────────────────────────────
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
day_fill    = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
row_fill    = PatternFill(start_color="EEF5FC", end_color="EEF5FC", fill_type="solid")
note_fill   = PatternFill(start_color="FFF6E5", end_color="FFF6E5", fill_type="solid")
total_fill  = PatternFill(start_color="D6E8FA", end_color="D6E8FA", fill_type="solid")

TYPE_COLORS = {
    "Fonctionnalité":       "2E7D32",
    "Correction":           "C62828",
    "Tâche":                "6A1B9A",
    "Documentation":        "1565C0",
    "Refactorisation":      "E65100",
    "Mise en forme":        "00838F",
    "Test":                 "4527A0",
    "Performance":          "558B2F",
    "Intégration continue": "37474F",
    "Build":                "4E342E",
    "Annulation":           "757575",
    "En cours":             "F57F17",
    "Correctif urgent":     "B71C1C",
    "Déploiement":          "1A237E",
    "Nouvelle version":     "004D40",
    "Autre":                "455A64",
}

# ─── Chargement ou création du classeur ──────────────────────────────────────
excel_exists = os.path.exists(dst)

if excel_exists:
    try:
        wb = load_workbook(dst)
        ws = wb["Journal"] if "Journal" in wb.sheetnames else wb.active
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


def reset_sheet(sheet):
    """Vide une feuille : fusions, graphiques et lignes."""
    if sheet is None:
        return
    for rng in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(rng))
    sheet._charts = []
    sheet._images = []
    if sheet.max_row > 1:
        sheet.delete_rows(1, sheet.max_row)


# ─── Onglet Journal (reconstruit depuis le CSV) ──────────────────────────────
reset_sheet(ws)
if ws.max_column > len(HEADERS):
    ws.delete_cols(len(HEADERS) + 1, ws.max_column - len(HEADERS))

for idx, head in enumerate(HEADERS, start=1):
    cell = ws.cell(row=1, column=idx, value=head)
    cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
    cell.fill      = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

# ─── Groupement par jour ─────────────────────────────────────────────────────
grouped        = OrderedDict()
author_time    = defaultdict(int)
author_commits = defaultdict(int)

for row in data:
    date_label = str(row[0]).strip()
    if not date_label:
        continue
    grouped.setdefault(date_label, []).append(row)
    auteur = str(row[1]).strip() or "(inconnu)"
    author_time[auteur]    += parse_minutes(row[5])
    author_commits[auteur] += 1

# ─── Remplissage du Journal ─────────────────────────────────────────────────
for day, entries in grouped.items():
    ws.append([f"{day}  ·  {len(entries)} tâche(s)", "", "", "", "", "", "", ""])
    day_row = ws.max_row
    ws.merge_cells(start_row=day_row, start_column=1, end_row=day_row, end_column=len(HEADERS))
    day_cell           = ws.cell(row=day_row, column=1)
    day_cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
    day_cell.fill      = day_fill
    day_cell.alignment = Alignment(horizontal="left", vertical="center")

    day_total = 0

    for entry in entries:
        auteur   = str(entry[1]).strip() or "(inconnu)"
        type_lbl = str(entry[3]).strip() or "Autre"
        desc     = humanize(str(entry[4]).strip())
        temps    = str(entry[5]).strip() or "[?]"
        etat     = str(entry[6]).strip() or "[?]"
        commit   = str(entry[2]).strip()
        fichiers = str(entry[7]).strip()

        type_color = TYPE_COLORS.get(type_lbl, TYPE_COLORS["Autre"])
        type_fill  = PatternFill(start_color=type_color, end_color=type_color, fill_type="solid")

        ws.append(["", auteur, type_lbl, desc, temps, etat, commit, fichiers])
        data_row = ws.max_row

        for col_idx in range(1, len(HEADERS) + 1):
            c = ws.cell(row=data_row, column=col_idx)
            c.fill = row_fill
            c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top", wrap_text=(col_idx in (4, 8)))

        ws.cell(row=data_row, column=2).font = Font(bold=True, name="Arial")

        type_cell           = ws.cell(row=data_row, column=3)
        type_cell.fill      = type_fill
        type_cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
        type_cell.alignment = Alignment(horizontal="center", vertical="top")

        ws.cell(row=data_row, column=5).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=data_row, column=6).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=data_row, column=7).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=data_row, column=7).font = Font(name="Consolas", size=9)

        day_total += parse_minutes(temps)

    # Notes du jour (colonnes 9 à 14 remplies à la main dans le CSV)
    for idx, label in NOTE_FIELDS:
        values = []
        for entry in entries:
            value = str(entry[idx]).strip() if len(entry) > idx else ""
            if value and value not in values:
                values.append(value)
        if not values:
            continue
        ws.append(["", label, " • ".join(values), "", "", "", "", ""])
        note_row = ws.max_row
        ws.merge_cells(start_row=note_row, start_column=3, end_row=note_row, end_column=len(HEADERS))
        for col_idx in range(1, len(HEADERS) + 1):
            c = ws.cell(row=note_row, column=col_idx)
            c.fill = note_fill
            c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=note_row, column=2).font = Font(bold=True, name="Arial")

    # Total du jour
    ws.append(["Total du jour", "", "", f"{to_hhmm(day_total)} ({day_total} min)", "", "", "", ""])
    total_row = ws.max_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    for col_idx in range(1, len(HEADERS) + 1):
        c = ws.cell(row=total_row, column=col_idx)
        c.fill = total_fill
        c.font = Font(bold=True, name="Arial")
        c.alignment = Alignment(vertical="top")

    ws.append([""] * len(HEADERS))

ws.freeze_panes    = "A2"
ws.auto_filter.ref = f"A1:{chr(64 + len(HEADERS))}{ws.max_row}"

widths = {"A": 24, "B": 20, "C": 20, "D": 68, "E": 12, "F": 12, "G": 12, "H": 42}
for col_letter, width in widths.items():
    ws.column_dimensions[col_letter].width = width

# ─── Onglet Tableau de bord ─────────────────────────────────────────────────
if ws2 is None:
    ws2 = wb.create_sheet("Tableau de bord")
reset_sheet(ws2)

ws2.append(["Date", "Total (min)", "Total (h)", "Total (h:min)"])
daily = defaultdict(int)
for row in data:
    date_label = str(row[0]).strip()
    if not date_label:
        continue
    daily[date_label] += parse_minutes(row[5])

try:
    sorted_days = sorted(daily.items(), key=lambda item: datetime.strptime(item[0], "%d/%m/%Y"))
except ValueError:
    sorted_days = sorted(daily.items(), key=lambda item: item[0])

for day, minutes in sorted_days:
    ws2.append([day, minutes, round(minutes / 60, 2), to_hhmm(minutes)])

for cell in ws2[1]:
    cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
    cell.fill      = header_fill
    cell.alignment = Alignment(horizontal="center")

ws2["F1"] = "Indicateur"
ws2["G1"] = "Valeur"

if sorted_days:
    total_minutes = sum(m for _, m in sorted_days)
    nb_days       = len(sorted_days)
    top_day, top_minutes = max(sorted_days, key=lambda item: item[1])

    ws2["F2"] = "Jours suivis"
    ws2["G2"] = nb_days
    ws2["F3"] = "Tâches / commits référencés"
    ws2["G3"] = len(data)
    ws2["F4"] = "Temps total"
    ws2["G4"] = to_hhmm(total_minutes)
    ws2["F5"] = "Moyenne / jour"
    ws2["G5"] = to_hhmm(round(total_minutes / nb_days))
    ws2["F6"] = "Jour le plus chargé"
    ws2["G6"] = f"{top_day} ({to_hhmm(top_minutes)})"

    # Répartition par auteur : qui a fait quoi
    ws2["H1"] = "Auteur"
    ws2["I1"] = "Tâches"
    ws2["J1"] = "Temps total"
    ws2["K1"] = "Temps (h)"

    for row in ws2["F1:K1"]:
        for cell in row:
            cell.font      = Font(bold=True, color="FFFFFF", name="Arial")
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center")

    author_rows = sorted(author_time.items(), key=lambda item: item[1], reverse=True)
    for idx, (auteur, minutes) in enumerate(author_rows, start=2):
        ws2[f"H{idx}"] = auteur
        ws2[f"I{idx}"] = author_commits[auteur]
        ws2[f"J{idx}"] = to_hhmm(minutes)
        ws2[f"K{idx}"] = round(minutes / 60, 2)
        ws2[f"H{idx}"].font = Font(bold=True, name="Arial")

    for row_idx in range(2, 7):
        ws2[f"F{row_idx}"].font = Font(bold=True, name="Arial")

    # Table de cumul (pour le graphique)
    ws2["M1"] = "Date"
    ws2["N1"] = "Cumul (h)"
    ws2["M1"].font = Font(bold=True, name="Arial")
    ws2["N1"].font = Font(bold=True, name="Arial")

    running = 0
    for idx, (day, minutes) in enumerate(sorted_days, start=2):
        running += minutes
        ws2[f"M{idx}"] = day
        ws2[f"N{idx}"] = round(running / 60, 2)

    max_row  = ws2.max_row
    date_ref = Reference(ws2, min_col=1, min_row=2, max_row=max_row)

    bar              = BarChart()
    bar.title        = "Charge par jour (heures)"
    bar.y_axis.title = "Heures"
    bar.x_axis.title = "Date"
    bar.add_data(Reference(ws2, min_col=3, min_row=1, max_row=max_row), titles_from_data=True)
    bar.set_categories(date_ref)
    bar.style, bar.height, bar.width = 10, 8, 16
    ws2.add_chart(bar, "F9")

    line              = LineChart()
    line.title        = "Cumul des heures"
    line.y_axis.title = "Heures"
    line.x_axis.title = "Date"
    line.add_data(Reference(ws2, min_col=14, min_row=1, max_row=max_row), titles_from_data=True)
    line.set_categories(date_ref)
    line.smooth = True
    if line.series:
        line.series[0].marker = None
    line.height, line.width = 8, 16
    ws2.add_chart(line, "F27")

    if author_rows:
        abar              = BarChart()
        abar.type         = "bar"
        abar.title        = "Temps par auteur (heures)"
        abar.y_axis.title = "Auteur"
        abar.x_axis.title = "Heures"
        abar.add_data(
            Reference(ws2, min_col=11, min_row=1, max_row=1 + len(author_rows)),
            titles_from_data=True,
        )
        abar.set_categories(Reference(ws2, min_col=8, min_row=2, max_row=1 + len(author_rows)))
        abar.height, abar.width = 8, 16
        ws2.add_chart(abar, "F45")

for col in ws2.iter_cols(min_col=1, max_col=14):
    letter  = col[0].column_letter
    lengths = [len(str(c.value)) for c in col if c.value is not None]
    if lengths:
        ws2.column_dimensions[letter].width = min(max(max(lengths) + 2, 10), 26)

ws2.freeze_panes = "A2"

# ─── Sauvegarde ─────────────────────────────────────────────────────────────
try:
    wb.save(dst)
except PermissionError:
    base, ext = os.path.splitext(dst)
    fallback  = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
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
        echo "♻️  Excel mis à jour : $OUTPUT_XLSX"
    elif grep -q '^CREATE_MODE=true' <<< "$PY_OUTPUT"; then
        echo "✅ Nouvel Excel généré : $OUTPUT_XLSX"
    else
        echo "✅ Excel généré : $OUTPUT_XLSX"
    fi
elif [ $PY_STATUS -eq 2 ]; then
    echo "⚠️  Module Python 'openpyxl' manquant. Installe-le avec :"
    echo "   python3 -m pip install --user --break-system-packages openpyxl"
else
    [ -n "$PY_OUTPUT" ] && echo "$PY_OUTPUT"
    echo "⚠️  Erreur lors de la génération Excel."
fi

echo "✨ C'est fini !"
