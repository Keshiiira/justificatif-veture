# Justificatifs de Vêture – Allocation Mensuelle

Outil Python qui génère un **classeur Excel** (`.xlsx`) tout-en-un pour la gestion des justificatifs de vêture.

---

## Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| 📊 Classeur unique `.xlsx` | Feuilles `Formulaire` + `Configuration` |
| 🔽 Liste déroulante | Sélection de l'enfant (depuis `Configuration`) |
| 🔗 Référent automatique | Rempli via formule INDEX/MATCH |
| 📅 Date automatique | `=TODAY()` mise à jour à chaque ouverture |

---

## Prérequis

- **Python 3.9+** avec `openpyxl` : `pip install openpyxl`
- **Microsoft Excel** (Windows/macOS) ou tout tableur compatible `.xlsx`

---

## Installation

```bash
git clone https://github.com/Keshiiira/justificatif-veture.git
cd justificatif-veture
pip install -r requirements.txt
```

---

## Générer le classeur

```bash
python src/generate_excel.py
# Génère : output/justificatif_veture.xlsx
```

Ouvrez le fichier dans **Excel** (Mac, Web, Windows) ou tout tableur compatible.

---

## Configuration initiale

Onglet **Configuration** :

| Champ | Description |
|---|---|
| Nom / Prénom AF | Votre nom complet |
| Adresse | Votre adresse |
| Téléphone | Votre téléphone |
| Email | Votre email |
| Fait à (ville) | Ville de signature |
| Tableau enfants | Colonne A = enfant, Colonne B = référent social |

---

## Utilisation mensuelle (onglet Formulaire)

### 1. Remplir le formulaire

1. Renseigner le **Mois et Année** en haut du formulaire
2. Sélectionner l'enfant dans la liste déroulante (`C13`) → le référent se remplit automatiquement
3. Saisir les factures dans le tableau (lignes 18–27) : nature, fournisseur, date, montant
4. Renseigner le champ M./Mme (`C34`) si nécessaire

---

## Structure du projet

```
justificatif-veture/
├── src/
│   ├── generate_excel.py   # Génère le classeur .xlsx
│   ├── generate_pdf.py     # (optionnel) fusion PDF via Python
│   └── vba_builder.py      # Réservé pour les macros (à venir)
├── invoices/               # Photos de factures
│   └── .gitkeep
├── output/                 # Créé automatiquement
├── requirements.txt
└── README.md
```
