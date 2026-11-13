# 🚀 Get-Commit
![Stars](https://img.shields.io/github/stars/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=yellow)
![Commits](https://img.shields.io/github/commit-activity/m/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=blue)
![Issues](https://img.shields.io/github/issues/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=orange)
![Forks](https://img.shields.io/github/forks/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=808080)
![Last Commit](https://img.shields.io/github/last-commit/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=blue)

> **Génère automatiquement un journal de travail (CSV + Excel) à partir de tes commits Git.**
>
> Le script parse tes messages de commit (temps + état) et met à jour un fichier `Doc/journal_de_travail.csv` puis génère un `Doc/journal_de_travail.xlsx` avec mise en forme et tableau de bord.  
> Les termes techniques sont **automatiquement traduits en langage métier** dans l'Excel.

---

## ⚡ Utilisation rapide

> ⚠️ **Le script doit être exécuté dans un vrai dépôt Git** (un dossier qui contient un `.git`).

Dans **Git Bash** :
```bash
curl -L -o main_commit.sh https://raw.githubusercontent.com/BlackAngelTVdev/Get-Commit/main/main_commit.sh
bash main_commit.sh
```

Les fichiers générés/mis à jour seront créés dans `./Doc/` du dépôt courant :
- `Doc/journal_de_travail.csv`
- `Doc/journal_de_travail.xlsx` *(si Python + openpyxl)*

---

## 🧐 Aperçu
Ce projet contient un script `get_commit.sh` qui :
- lit l'historique `git log` de ton dépôt,
- extrait **Date / Auteur / Commit / Type / Description / Temps / État / Fichiers**,
- **indique qui a fait quoi** (nom de l'auteur du commit, visible dans le CSV et l'Excel),
- **référence chaque commit** (hash court) avec les **fichiers et répertoires** concernés,
- détecte des tags dans tes commits comme **temps** (`[1h45min]`) et **état** (`[DONE]`),
- écrit tout ça dans un **CSV** (mode *append*, sans écraser tes modifications),
- et (si Python est installé) génère un **Excel** mis en forme avec un onglet *Tableau de bord*,
- **traduit automatiquement** les termes techniques (`feat`, `fix`, `chore`…) en langage métier dans l'Excel.

> Les fichiers sont générés dans `Doc/` **du dépôt courant** (le dossier depuis lequel tu exécutes le script).

---

## ✨ Fonctionnalités
- ✅ **Mode APPEND (ajout, pas remplacement)** :
  - le CSV n'est **pas recréé** ; seules les nouvelles entrées sont ajoutées
  - tes modifications manuelles dans le CSV sont préservées
- ✅ **Détection anti-doublons** : évite de réécrire un commit déjà présent (clé : hash du commit).
- ✅ **Extraction automatique depuis tes commits** :
  - `Date` : issue de `git log` (format `dd/mm/yyyy`)
  - `Auteur` : auteur du commit (`%an`, normalisé par `.mailmap` si présent)
  - `Commit` : hash court (8 caractères) → le commit est référencé
  - `Fichiers` : fichiers / répertoires touchés par le commit
  - `Type` : libellé métier du préfixe (`feat` → *Fonctionnalité*…, scope conservé)
  - `Description` : sujet nettoyé + corps du message
- ✅ **Parsing des tags dans les messages de commit** :
  - **Temps** : `[1h]`, `[1h30]`, `[45min]`, `[50]`, etc.
  - **État** : `[DONE]`, `[WIP]`, etc. (majuscules)
- ✅ **Colonnes de journal à remplir à la main** (préservées d'un lancement à l'autre) :
  `Activités planifiées`, `Aide extérieure`, `Imprévus / heures supplémentaires`,
  `Succès`, `Échecs`, `Appréciation critique / réflexion`.
- ✅ **Répartition par auteur** dans l'onglet *Tableau de bord* (tâches + temps par personne).
- ✅ **Traduction automatique dans l'Excel** : les préfixes et mots techniques sont remplacés par leur équivalent métier (voir la section [Convention de commit](#-convention-de-commit) ci-dessous).
- ✅ **Génération Excel (.xlsx)** *(si Python 3 + openpyxl)* :
  - onglet **Journal** : regroupé par jour + total du jour
  - onglet **Tableau de bord** : totaux + graphiques (charge/jour, cumul)
  - si le fichier Excel est ouvert/verrouillé : sauvegarde automatique vers un nom alternatif
- ✅ **Auto-install (optionnel)** : tente d'installer `git`, `awk`, `python` (selon l'OS) si manquants.

---

## 📓 Journal de travail (structure du CSV)

Le CSV `Doc/journal_de_travail.csv` sert de source de vérité et couvre les exigences
d'un journal de travail individuel :

| # | Colonne | Rempli par | Contenu |
| :-: | :--- | :--- | :--- |
| 1 | `Date` | script | jour du commit (`dd/mm/yyyy`) |
| 2 | `Auteur` | script | **qui a fait quoi** |
| 3 | `Commit` | script | hash court → **commit référencé** |
| 4 | `Type` | script | catégorie métier de la tâche |
| 5 | `Description` | script | travail réalisé (sujet + corps) |
| 6 | `Temps` | script | **durée** de la tâche |
| 7 | `État` | script | **état** (`[DONE]`, `[WIP]`, `[REVIEW]`…) |
| 8 | `Fichiers` | script | **fichiers / répertoires** concernés |
| 9 | `Activités planifiées` | **à la main** | ce qui était prévu |
| 10 | `Aide extérieure` | **à la main** | aide reçue |
| 11 | `Imprévus / heures supplémentaires` | **à la main** | imprévus, heures en plus |
| 12 | `Succès` | **à la main** | ce qui a marché |
| 13 | `Échecs` | **à la main** | ce qui n'a pas marché |
| 14 | `Appréciation critique / réflexion` | **à la main** | regard critique, réflexion personnelle |

**Remplir les colonnes 9 à 14 :** édite directement le CSV (une ligne = une tâche).
Le script ne réécrit jamais une ligne existante : tes notes sont conservées.
Dans l'Excel, elles apparaissent sous forme de bloc *notes du jour* sous les tâches du jour.

> ⚠️ L'onglet **Journal** de l'Excel est régénéré depuis le CSV à chaque lancement :
> saisis tes notes dans le **CSV**, pas dans l'Excel.

**Migration :** si le CSV a un ancien format, il est sauvegardé en
`journal_de_travail.csv.bak_<date>` puis reconstruit depuis Git (auteur, commit et
fichiers sont ainsi récupérés pour tout l'historique).

---

## 🧠 Convention de commit

### Format général
```
<type>(<scope>): <description> [<temps>] [<état>]
```

| Champ | Obligatoire | Détail |
| :--- | :---: | :--- |
| `type` | ✅ | Préfixe du commit (voir tableau ci-dessous) |
| `(scope)` | ❌ | Module / partie concernée, entre parenthèses |
| `!` | ❌ | Breaking change (changement majeur), juste avant `:` |
| `description` | ✅ | Résumé court de ce qui a été fait |
| `[temps]` | ❌ | Durée du travail : `[1h]`, `[1h30]`, `[45min]` |
| `[état]` | ❌ | Statut en majuscules : `[DONE]`, `[WIP]`, `[REVIEW]` |

---

### Types reconnus (et leur traduction dans l'Excel)

| Type dans le commit | Variantes acceptées | Affiché dans l'Excel |
| :--- | :--- | :--- |
| `feat` | `feat`, `feat(scope)`, `feat!`, `FEAT(scope)!` | **Fonctionnalité** |
| `fix` | `fix`, `fix(scope)`, `Fix(scope)!` | **Correction** |
| `chore` | `chore`, `chor`, `Chor(scope)` | **Tâche** |
| `doc` / `docs` | `doc`, `docs`, `docs(scope)` | **Documentation** |
| `refactor` | `refactor`, `refactor(scope)!` | **Refactorisation** |
| `style` | `style`, `style(scope)` | **Mise en forme** |
| `test` | `test`, `test(scope)` | **Test** |
| `perf` | `perf`, `perf(scope)` | **Performance** |
| `ci` | `ci`, `ci(scope)` | **Intégration continue** |
| `build` | `build`, `build(scope)` | **Build** |
| `revert` | `revert`, `revert(scope)` | **Annulation** |
| `wip` | `wip`, `wip(scope)` | **En cours** |
| `hotfix` | `hotfix`, `hotfix(scope)` | **Correctif urgent** |
| `deploy` | `deploy`, `deploy(scope)` | **Déploiement** |
| `release` | `release`, `release(scope)` | **Nouvelle version** |

> Le scope entre parenthèses est **conservé** dans l'Excel. Exemple : `feat(auth):` → `Fonctionnalité (auth) :`.

---

### Exemples complets

```bash
# Fonctionnalité simple
feat: ajout page de connexion [1h30] [DONE]
# → Excel : Fonctionnalité : ajout page de connexion

# Fonctionnalité avec scope
feat(auth): ajout token JWT [45min] [DONE]
# → Excel : Fonctionnalité (auth) : ajout jeton d'authentification JWT

# Correction avec breaking change
fix(paiement)!: correction anomalie montant [20min] [DONE]
# → Excel : Correction (paiement) : correction anomalie montant

# Tâche (variante courte "Chor")
Chor(JNR): mise à jour dépendances [10min] [DONE]
# → Excel : Tâche (JNR) : mise à jour dépendances

# Sauvegarde cloud
Chor(Save): Save code on cloud [5min] [DONE]
# → Excel : Tâche (sauvegarde) : sauvegarde code on cloud

# Documentation
docs(api): mise à jour README [15min] [DONE]
# → Excel : Documentation (api) : mise à jour lisez-moi

# Travail en cours
wip(router): refacto frontend [30min] [WIP]
# → Excel : En cours (router) : refacto interface utilisateur

# Déploiement
deploy(prod): release v2.3 backend [10min] [DONE]
# → Excel : Déploiement (prod) : version publiée v2.3 serveur
```

---

### Mots traduits automatiquement dans le texte

En plus du préfixe, certains mots courants dans la description sont aussi traduits :

| Mot technique | Traduction |
| :--- | :--- |
| `bug` | anomalie |
| `token` | jeton d'authentification |
| `auth` | authentification |
| `login` / `logout` | connexion / déconnexion |
| `backend` / `frontend` | serveur / interface utilisateur |
| `db` / `database` | base de données |
| `pipeline` | chaîne de traitement |
| `release` | version publiée |
| `readme` | lisez-moi |
| `changelog` | journal des modifications |
| `deploy` | déploiement |
| `merge` | fusion |
| `rollback` | retour arrière |
| `mock` | simulation |
| `config` | configuration |
| `setup` | mise en place |
| `update` / `upgrade` | mise à jour / mise à niveau |
| `modal` | fenêtre modale |
| `form` / `input` | formulaire / champ de saisie |
| `router` | routeur |
| `controller` | contrôleur |
| `timeout` | délai dépassé |
| `password` | mot de passe |

---

## 🛠 Tech Stack
| Technologie | Usage |
| :--- | :--- |
| ![Bash](https://img.shields.io/badge/Bash-1f425f?style=flat-square) | Script principal, extraction `git log`, génération CSV |
| ![Git](https://img.shields.io/badge/Git-F05032?style=flat-square&logo=git&logoColor=white) | Source des commits |
| ![AWK](https://img.shields.io/badge/AWK-000000?style=flat-square) | Parsing/formatage des données |
| ![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) | Conversion CSV → XLSX + traduction métier |
| ![openpyxl](https://img.shields.io/badge/openpyxl-0B7285?style=flat-square) | Création du classeur + styles + graphiques |

---

## 🚀 Installation & Lancement

### Prérequis
- Un dépôt **Git** local (le script doit être exécuté dans un dossier contenant `.git`)
- Sur Windows : **Git Bash** recommandé (le script est en bash)
- Python 3 + `openpyxl` (uniquement pour générer le `.xlsx`) :
  ```bash
  python3 -m pip install --user --break-system-packages openpyxl
  ```

### Lancer le script (dans ce repo cloné)
```bash
bash ./get_commit.sh
```

Le script génère/met à jour :
- `Doc/journal_de_travail.csv`
- `Doc/journal_de_travail.xlsx` (si Python + openpyxl)

### (Optionnel) Désactiver l'auto-install
Par défaut, le script tente d'installer certaines commandes manquantes.
Pour désactiver :
```bash
AUTO_INSTALL=0 bash ./get_commit.sh
```

---

## 📖 Utilisation

### Où lancer le script ?
Tu peux copier `get_commit.sh` **dans n'importe quel dépôt** et le lancer dedans.
Il utilisera `git log` du dépôt courant et écrira dans `./Doc/`.

### Exemple de workflow
1. Tu commits avec la convention :
   ```
   feat(auth): add login page [1h30] [DONE]
   ```
2. Tu lances le script :
   ```bash
   bash ./get_commit.sh
   ```
3. Tu ouvres `Doc/journal_de_travail.xlsx` → journal mis en forme + dashboard.

### Dépannage rapide
- `fatal: not a git repository` → lance le script **dans un dépôt Git**.
- L'Excel ne se génère pas → installe Python 3 puis : `pip install openpyxl`.
- L'Excel est "verrouillé" (ouvert) → le script enregistre un fichier alternatif (`*_YYYYMMDD_HHMMSS.xlsx`).

---

## 🤝 Contribution

### Cloner le projet
```bash
git clone https://github.com/BlackAngelTVdev/Get-Commit.git
cd Get-Commit
```

### Proposer une amélioration
1. Forkez le projet
2. Créez votre branche : `git checkout -b feat/AmazingFeature`
3. Commit : `git commit -m "feat: add AmazingFeature [30min] [DONE]"`
4. Push : `git push origin feat/AmazingFeature`
5. Ouvrez une Pull Request

### 🧑‍💻 Contributors

Merci à toutes les personnes qui contribuent au projet.

[![Contributors](https://contrib.rocks/image?repo=BlackAngelTVdev/Get-Commit)](https://github.com/BlackAngelTVdev/Get-Commit/graphs/contributors)

---

## 👤 Auteur

**BlackAngelTVdev**  
![Follow](https://img.shields.io/github/followers/BlackAngelTVdev?label=Follow%20Me&style=social)

---

## 📄 Licence
Ce projet est sous une licence **non-commerciale** (custom).

- ✅ Tu peux **utiliser** et **modifier** le code.
- ❌ Tu ne peux **pas vendre / monétiser** le projet (ou un fork) sans autorisation écrite.

Détails complets : voir le fichier [`LICENSE`](LICENSE).
