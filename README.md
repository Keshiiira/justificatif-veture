# Justificatifs de Vêture – Allocation Mensuelle

Outil Python automatisant la génération du formulaire **"Justificatifs de Vêture – Allocation Mensuelle"** au format Excel, puis la production d'un **PDF final** fusionnant le formulaire et les photos de factures.

---

## Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| 📊 Classeur Excel | Feuilles `Formulaire` + `Configuration` |
| 🔽 Liste déroulante | Sélection de l'enfant depuis la Configuration |
| 🔗 Référent auto | Rempli automatiquement selon l'enfant choisi (formule INDEX/MATCH) |
| 📅 Date automatique | `=TODAY()` – mise à jour à chaque ouverture |
| 📄 Export PDF | Via LibreOffice (cross-platform) |
| 🖼️ Fusion factures | Images JPG/PNG/HEIC → PDF → fusion avec formulaire |

---

## Prérequis

- **Python 3.9+**
- **LibreOffice** (pour la conversion Excel → PDF) :
  - Linux : `sudo apt install libreoffice`
  - macOS : [libreoffice.org](https://www.libreoffice.org/download/download/)
  - Windows : [libreoffice.org](https://www.libreoffice.org/download/download/) (ou utiliser Excel COM, voir ci-dessous)

---

## Installation

```bash
# 1. Cloner le dépôt
git clone https://github.com/Keshiiira/justificatif-veture.git
cd justificatif-veture

# 2. Créer un environnement virtuel
python -m venv venv
source venv/bin/activate      # Windows : venv\Scripts\activate

# 3. Installer les dépendances
pip install -r requirements.txt

# Support optionnel HEIC (photos iPhone) :
pip install pillow-heif
```

---

## Structure du projet

```
justificatif-veture/
├── src/
│   ├── generate_excel.py   # Génère le classeur Excel
│   └── generate_pdf.py     # Exporte en PDF et fusionne les factures
├── invoices/               # Déposez ici les photos de factures
│   └── .gitkeep
├── output/                 # Créé automatiquement (ignoré par git)
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Configuration initiale

Ouvrez `output/justificatif_veture.xlsx` (après génération) et allez sur l'onglet **Configuration** :

### Informations fixes (lignes 4–8)

| Champ | Valeur par défaut | À modifier |
|---|---|---|
| Nom / Prénom AF | `Dupont Marie` | ✅ |
| Adresse | `12 Rue des Lilas, 09000 Foix` | ✅ |
| Téléphone | `06 12 34 56 78` | ✅ |
| Email | `marie.dupont@email.fr` | ✅ |
| Fait à (ville) | `Foix` | Si besoin |

> Vous pouvez aussi modifier les valeurs par défaut directement dans `src/generate_excel.py` (dictionnaire `CONFIG_AF`).

### Ajouter des enfants et référents (lignes 12–29)

Dans l'onglet **Configuration**, remplissez le tableau :

| Colonne A – Enfant | Colonne B – Référent social |
|---|---|
| Martin Lucas | Mme Bernard Sophie |
| Petit Emma | M. Leclerc Paul |
| *(ajoutez ici)* | *(référent correspondant)* |

La liste déroulante de l'onglet **Formulaire** se met à jour automatiquement.

---

## Utilisation mensuelle

### Étape 1 – Générer le classeur Excel

```bash
python src/generate_excel.py
# Génère : output/justificatif_veture.xlsx
```

Options :
```bash
python src/generate_excel.py --output chemin/vers/fichier.xlsx
```

### Étape 2 – Remplir le formulaire

Ouvrez `output/justificatif_veture.xlsx` dans Excel ou LibreOffice Calc :

1. **Onglet Formulaire** :
   - Sélectionnez l'enfant dans la liste déroulante (cellule `C11`) → le référent social se remplit automatiquement (`C12`).
   - Saisissez les factures dans le tableau (lignes 16–25) : type d'achat, fournisseur, date, coût.
   - Vérifiez la date (`E31`) et la ville (`C31`).
   - Signez (cellule `C32`).

### Étape 3 – Ajouter les photos de factures

Copiez vos photos dans le dossier `invoices/` :
```
invoices/
  facture_vetements_jan.jpg
  facture_chaussures.png
  ticket_caisse.jpeg
```

Formats supportés : `.jpg`, `.jpeg`, `.png` (et `.heic`/`.heif` avec `pillow-heif`).

### Étape 4 – Générer le PDF final

```bash
python src/generate_pdf.py
# Génère :
#   output/formulaire.pdf  (export Excel)
#   output/final.pdf       (formulaire + toutes les factures)
```

Options :
```bash
python src/generate_pdf.py \
  --excel   output/justificatif_veture.xlsx \
  --invoices invoices/ \
  --output  output/final.pdf
```

---

## Export Excel → PDF : détail des méthodes

### LibreOffice (recommandé, cross-platform)

Le script détecte automatiquement `libreoffice` ou `soffice` dans le PATH.

```bash
# Vérification :
libreoffice --version
```

### Windows avec Excel COM (manuel)

Si LibreOffice n'est pas installé sous Windows, vous pouvez exporter manuellement depuis Excel :
`Fichier → Exporter → Créer un document PDF/XPS`

Puis utilisez uniquement la fusion :
```bash
python src/generate_pdf.py --excel ""   # (sauter l'étape Excel)
```
> Dans ce cas, copiez votre `formulaire.pdf` manuellement dans `output/` avant de lancer la fusion.

---

## Support HEIC (photos iPhone)

```bash
pip install pillow-heif
```

Décommentez la ligne correspondante dans `requirements.txt`, puis relancez. Les fichiers `.heic` et `.heif` seront automatiquement détectés.

---

## Commandes de validation complète

```bash
# 1. Générer l'Excel
python src/generate_excel.py

# 2. (Optionnel) Ajouter des images test dans invoices/

# 3. Générer le PDF final
python src/generate_pdf.py

# 4. Vérifier les sorties
ls -lh output/
```
