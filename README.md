# Justificatifs de Vêture – Allocation Mensuelle

Outil Python qui génère un **classeur Excel macro-activé** (`.xlsm`) tout-en-un pour la gestion des justificatifs de vêture.

---

## Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| 📊 Classeur unique `.xlsm` | Feuilles `Formulaire` + `Configuration` |
| 🔽 Liste déroulante | Sélection de l'enfant (depuis `Configuration`) |
| 🔗 Référent automatique | Rempli via formule INDEX/MATCH |
| 📅 Date automatique | `=TODAY()` mise à jour à chaque ouverture |
| 📎 Bouton « Ajouter des factures » | Copie les photos sélectionnées dans `invoices/` |
| 📄 Bouton « Exporter en PDF » | Export natif Excel (mode impression), propose la réinitialisation après |
| 🔄 Réinitialisation automatique | Efface enfant, factures saisies, images `invoices/` en un clic |

---

## Prérequis

- **Python 3.9+** avec `openpyxl` : `pip install openpyxl`
- **Microsoft Excel** (Windows/macOS) pour utiliser les macros et l'export PDF natif

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
# Génère : output/justificatif_veture.xlsm
```

Ouvrez le fichier dans **Excel**, activez les macros si demandé.

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

1. Sélectionner l'enfant dans la liste déroulante (`C11`) → le référent se remplit automatiquement
2. Saisir les factures dans le tableau (lignes 16–25) : type, fournisseur, date, coût
3. Renseigner le champ M./Mme (`C32`) si nécessaire

### 2. (Optionnel) Ajouter des photos de factures

Cliquer sur le bouton **📎 Ajouter des factures** → sélectionner les images.  
Elles sont copiées dans le dossier `invoices/` à côté du classeur.

### 3. Exporter en PDF

Cliquer sur le bouton **📄 Exporter en PDF** :
- Une fenêtre de sauvegarde s'ouvre (nom pré-rempli avec le nom de l'enfant)
- Le PDF est créé via l'export natif Excel (mise en page impression)
- Le PDF s'ouvre automatiquement
- Une boîte de dialogue propose de **réinitialiser le formulaire** pour le prochain enfant

### 4. Réinitialisation pour le prochain enfant

En cliquant **Oui** après l'export PDF, le formulaire efface :
- L'identité de l'enfant (`C11`)
- Le détail des factures (lignes 16–25)
- Le champ M./Mme (`C32`)
- Toutes les images dans `invoices/`

---

## Structure du projet

```
justificatif-veture/
├── src/
│   ├── generate_excel.py   # Génère le classeur .xlsm
│   ├── generate_pdf.py     # (optionnel) fusion PDF via Python
│   └── vba_builder.py      # Constructeur OLE2 du vbaProject.bin
├── invoices/               # Photos de factures (géré par le bouton)
│   └── .gitkeep
├── output/                 # Créé automatiquement
├── requirements.txt
└── README.md
```
