# 🚀 Get-Commit
![Stars](https://img.shields.io/github/stars/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=yellow)
![Commits](https://img.shields.io/github/commit-activity/m/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=blue)
![Issues](https://img.shields.io/github/issues/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=orange)
![Forks](https://img.shields.io/github/forks/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=808080)
![Last Commit](https://img.shields.io/github/last-commit/BlackAngelTVdev/Get-Commit?style=for-the-badge&color=blue)

> **Génère automatiquement un journal de travail (CSV + Excel) à partir de tes commits Git.**
> 
> Le script parse tes messages de commit (temps + état) et met à jour un fichier `Doc/journal_de_travail.csv` puis génère un `Doc/journal_de_travail.xlsx` avec mise en forme et tableau de bord.

---

## ⚡ Utilisation rapide

> ⚠️ **Le script doit être exécuté dans un vrai dépôt Git** (un dossier qui contient un `.git`).
>

Dans **Git Bash** (recommandé sur Windows) :
```bash
curl -L -o get_commit.sh https://raw.githubusercontent.com/BlackAngelTVdev/Get-Commit/main/get_commit.sh
bash get_commit.sh
```

Les fichiers générés/mis à jour seront créés dans `./Doc/` du dépôt courant :
- `Doc/journal_de_travail.csv`
- `Doc/journal_de_travail.xlsx` *(si Python + openpyxl)*

---

## 🧐 Aperçu
Ce projet contient un script `get_commit.sh` qui :
- lit l’historique `git log` de ton dépôt,
- extrait **Date / Titre / Description**,
- détecte des tags dans tes commits comme **temps** (`[1h45min]`) et **état** (`[DONE]`),
- écrit tout ça dans un **CSV** (mode *append*, sans écraser tes modifications),
- et (si Python est installé) génère un **Excel** mis en forme avec un onglet *Tableau de bord*.

> Les fichiers sont générés dans `Doc/` **du dépôt courant** (le dossier depuis lequel tu exécutes le script).

---

## ✨ Fonctionnalités
- ✅ **Mode APPEND (ajout, pas remplacement)** :
  - le CSV n’est **pas recréé** ; seules les nouvelles entrées sont ajoutées
  - tes modifications manuelles dans le CSV sont préservées
- ✅ **Détection anti-doublons** : évite de réécrire une ligne déjà présente (clé : `Date|Nom`).
- ✅ **Extraction automatique depuis tes commits** :
  - `Date` : issue de `git log` (format `dd/mm/yyyy`)
  - `Nom` : sujet du commit
  - `Description` : sujet + body nettoyé
- ✅ **Parsing des tags dans les messages de commit** :
  - **Temps** : `[1h]`, `[1h30]`, `[45min]`, etc.
  - **État** : `[DONE]`, `[WIP]`, etc. (majuscule)
- ✅ **Génération Excel (.xlsx)** *(si Python 3 + openpyxl)* :
  - onglet **Journal** : regroupé par jour + total du jour
  - onglet **Tableau de bord** : totaux + graphiques (charge/jour, cumul)
  - si le fichier Excel est ouvert/verrouillé : sauvegarde automatique vers un nom alternatif
- ✅ **Auto-install (optionnel)** : tente d’installer `git`, `awk`, `python` (selon l’OS) si manquants.

---

## 🧠 Convention de commit recommandée
Le script est prévu pour fonctionner avec un format simple dans tes messages de commit.

Exemples :
- `feat(auth): login page [1h30] [DONE]`
- `fix(api): handle null response [45min] [WIP]`
- `docs: update README [10min] [DONE]`

Notes :
- Le **temps** est détecté via un tag entre crochets (ex: `[1h45min]`).
- L’**état** est détecté via un tag en majuscules (ex: `[DONE]`).
- Si un champ n’est pas trouvé, la valeur par défaut est `"[?]"`.

---

## 🛠 Tech Stack
| Technologie | Usage |
| :--- | :--- |
| ![Bash](https://img.shields.io/badge/Bash-1f425f?style=flat-square) | Script principal, extraction `git log`, génération CSV |
| ![Git](https://img.shields.io/badge/Git-F05032?style=flat-square&logo=git&logoColor=white) | Source des commits |
| ![AWK](https://img.shields.io/badge/AWK-000000?style=flat-square) | Parsing/formatage des données |
| ![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) | Conversion CSV → XLSX |
| ![openpyxl](https://img.shields.io/badge/openpyxl-0B7285?style=flat-square) | Création du classeur + styles + graphiques |

---

## 🚀 Installation & Lancement

### Prérequis
- Un dépôt **Git** local (le script doit être exécuté dans un dossier contenant `.git`)
- Sur Windows : **Git Bash** recommandé (le script est en bash)
- Python 3 (optionnel, uniquement pour générer le `.xlsx`)

### Lancer le script (dans ce repo cloné)
```bash
bash ./get_commit.sh
```

Le script génère/met à jour :
- `Doc/journal_de_travail.csv`
- `Doc/journal_de_travail.xlsx` (si Python + openpyxl)

### (Optionnel) Désactiver l’auto-install
Par défaut, le script tente d’installer certaines commandes manquantes.
Pour désactiver :
```bash
AUTO_INSTALL=0 bash ./get_commit.sh
```

---

## 📖 Utilisation

### Où lancer le script ?
Tu peux copier `get_commit.sh` **dans n’importe quel dépôt** et le lancer dedans.
Il utilisera `git log` du dépôt courant et écrira dans `./Doc/`.

### Exemple de workflow
1. Tu commits avec une convention :
   - `feat: add search [25min] [DONE]`
2. Tu lances le script :
   - `bash ./get_commit.sh`
3. Tu ouvres `Doc/journal_de_travail.xlsx` pour avoir le journal + le dashboard.

### Dépannage rapide
- Si tu vois `fatal: not a git repository` → lance le script **dans un dépôt Git**.
- Si l’Excel ne se génère pas → installe Python 3 puis : `pip install openpyxl`.
- Si l’Excel est “verrouillé” (ouvert) → le script enregistre un fichier alternatif (`*_YYYYMMDD_HHMMSS.xlsx`).

---

## 🤝 Contribution

### Cloner le projet
```bash
git clone https://github.com/BlackAngelTVdev/Get-Commit.git
cd Get-Commit
```

### Proposer une amélioration
1. Forkez le projet
2. Créez votre branche : `git checkout -b feature/AmazingFeature`
3. Commit : `git commit -m "Add some AmazingFeature"`
4. Push : `git push origin feature/AmazingFeature`
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

