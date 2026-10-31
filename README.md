# 📄 PDF OCR to Excel Converter & PAC Interne Automation

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python" />
  <img src="https://img.shields.io/badge/Flask-3.0.0-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask" />
  <img src="https://img.shields.io/badge/OCR-EasyOCR_%26_PyMuPDF-FF6F00?style=for-the-badge&logo=pytorch&logoColor=white" alt="EasyOCR" />
  <img src="https://img.shields.io/badge/Data-Pandas_%26_OpenPyXL-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas" />
  <img src="https://img.shields.io/badge/Automation-VBA_%26_Win32COM-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Excel VBA" />
  <img src="https://img.shields.io/badge/Institution-EMSI_%26_Orange_Maroc-FF6600?style=for-the-badge" alt="Orange & EMSI" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License" />
</p>

<p align="center">
  <img src="docs/screenshots/emsi_logo.png" alt="EMSI" height="60" />
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <img src="docs/screenshots/orange_logo.png" alt="Orange" height="60" />
</p>

---

## 📌 Présentation Générale

**PDF OCR to Excel Converter** est une solution logicielle complète développée en **Python (Flask)** pour automatiser la chaîne de traitement, d'extraction, de conversion et d'analyse budgétaire des **Bons de Commande (BDC)** numérisés sous forme d'images ou de PDF scannés.

Conçu et validé en conditions réelles au sein de la **Direction des Systèmes d'Information & Infrastructure d'ORANGE Maroc (MEDITELECOM)**, ce système permet de remplacer la saisie manuelle fastidieuse et sujette aux erreurs par un pipeline d'ingénierie documentaire intelligent :
1. **Reconnaissance Optique de Caractères (OCR)** basée sur l'apprentissage profond (**EasyOCR**) pour extraire avec précision les métadonnées (fournisseurs, numéro de commande/contrat, dates) et les tableaux d'articles (désignations, quantités, prix unitaire, prix total).
2. **Interface Web interactive (Flask)** permettant la vérification humaine, la modification à la volée et la validation des données extraites.
3. **Génération automatisée de classeurs Excel multi-onglets structurés**, avec une feuille dédiée par bon de commande et calculs arithmétiques automatiques.
4. **Tableau de bord budgétaire consolidé ("Main")** offrant une vue macroscopique et en temps réel des budgets engagés par fournisseur et contrat.
5. **Génération automatisée de Procès-Verbaux de Réception ("PAC Interne")** via une macro VBA intégrée (`PACInterneModule`) et automatisation Word (`python-docx` / `win32com`).

---

## 👥 Cadre Académique & Professionnel

| Information | Détails |
|---|---|
| **Établissement** | **EMSI** — École Marocaine des Sciences de l'Ingénieur (Filière Informatique et Réseaux) |
| **Organisme d'Accueil** | **ORANGE (MEDITELECOM)** — Direction des Systèmes d'Information & Infrastructure |
| **Service d'Accueil** | Solution Réingénierie des Sites & Projets Spéciaux |
| **Encadrant Professionnel** | **M. Rabii Hanif** (Manager Solution Réingénierie des Sites & Projets Spéciaux — Orange) |
| **Cadre du Mémoire** | Mémoire de Projet de Fin d'Année (3ème année — Session Septembre 2025) |
| **Année Universitaire** | 2024 – 2025 |

---

## 📸 Galerie & Démonstration Visuelle

### 1. Interface Web d'Importation & Drag & Drop
Zone de glisser-déposer intuitive permettant le téléversement simultané de plusieurs fichiers PDF scannés.
<p align="center">
  <img src="docs/screenshots/fig07_interface_upload.png" alt="Interface Upload Drag & Drop" width="85%" />
</p>

### 2. Traitement & Extraction OCR en Temps Réel
Conversion des pages PDF en images haute résolution avec **PyMuPDF**, puis détection et transcription textuelle via **EasyOCR**.
<p align="center">
  <img src="docs/screenshots/fig08_extraction_chargement.png" alt="Extraction OCR en cours" width="85%" />
</p>

### 3. Contrôle & Validation des Données Extraites
Tableau web dynamique présentant les champs détectés (fournisseur, numéro BDC, contrat, date) et la liste détaillée des articles avec calcul automatique du total par ligne et possibilité d'édition manuelle.
<p align="center">
  <img src="docs/screenshots/fig09_resultat_extraction_table.png" alt="Vérification et validation des articles" width="85%" />
</p>

### 4. Classeurs Excel Multi-Onglets Générés
Génération automatique d'un onglet par bon de commande (nommé d'après le numéro de commande) avec mise en page soignée, calcul des totaux et bouton interactif **« PAC Interne »**.
<p align="center">
  <img src="docs/screenshots/fig10_excel_commande_1.png" alt="Feuille Excel Bon de Commande 1" width="48%" />
  <img src="docs/screenshots/fig11_excel_commande_2.png" alt="Feuille Excel Bon de Commande 2" width="48%" />
</p>

### 5. Procès-Verbal de Réception Interne ("PAC Interne")
Document officiel Word généré automatiquement au format contractuel de l'entreprise, intégrant les références du bon de commande, les articles réceptionnés et les encarts de signature.
<p align="center">
  <img src="docs/screenshots/fig12_proces_verbal_pac_interne.png" alt="Procès-Verbal de Réception PAC Interne" width="75%" />
</p>

### 6. Tableau de Bord Budgétaire Consolidé (Feuille "Main")
Synthèse financière regroupant par fournisseur et contrat les montants engagés, réceptionnés et les reliquats pour un pilotage budgétaire instantané.
<p align="center">
  <img src="docs/screenshots/fig13_tableau_de_bord_budgets.png" alt="Tableau de Bord Budgétaire" width="85%" />
</p>

---

## 📐 Conception & Modélisation UML

### Diagramme de Séquence Global
<p align="center">
  <img src="docs/screenshots/fig01_diagramme_sequence.png" alt="Diagramme de Séquence" width="90%" />
</p>

### Cycle de Vie du Fichier Excel & Activité du Système
<p align="center">
  <img src="docs/screenshots/fig02_diagramme_etat_excel.png" alt="Diagramme d'état Excel" width="45%" />
  &nbsp;&nbsp;&nbsp;&nbsp;
  <img src="docs/screenshots/fig03_diagramme_activite.png" alt="Diagramme d'activité" width="45%" />
</p>

### Diagramme d'État-Transition
<p align="center">
  <img src="docs/screenshots/fig04_diagramme_etat_transition.png" alt="Diagramme d'état transition" width="70%" />
</p>

### Diagramme de Classes & Modèle Relationnel
<p align="center">
  <img src="docs/screenshots/fig05_diagramme_classes.png" alt="Diagramme de classes" width="48%" />
  <img src="docs/screenshots/fig06_modele_relationnel.png" alt="Modèle relationnel" width="48%" />
</p>

---

## 🛠️ Stack Technologique & Justification des Choix

| Composant | Technologie | Justification |
|---|---|---|
| **Backend & API** | **Python & Flask** | Micro-framework léger, asynchrone et modulaire permettant une intégration native avec les bibliothèques d'analyse de données et d'OCR. |
| **Moteur OCR** | **EasyOCR** (PyTorch) | Basé sur des réseaux de neurones profonds (CNN pour la détection spatiale + RNN/LSTM pour la reconnaissance). Supérieur à Tesseract sur des scans bruités, inclinés ou avec polices non standardisées. |
| **Prétraitement PDF** | **PyMuPDF (`fitz`)** | Rendu ultra-rapide des pages PDF sous forme d'images matricielles haute résolution avant passage dans l'OCR. |
| **Structuration Données** | **Pandas & OpenPyXL** | Manipulation vectorielle de données tabulaires, calcul arithmétique rigoureux et export Excel multi-feuilles avec styles personnalisés. |
| **Automatisation Bureautique** | **VBA & `win32com` / `python-docx`** | Injection automatique de modules macro dans les fichiers `.xlsm` et génération de documents Word officiels conformes aux chartes administratives. |

---

## 🗂️ Structure du Projet

```
EXTRACTOR/
├── app.py                         # Application principale Flask & moteur d'extraction OCR
├── pac_interne_integrated.py      # Module d'injection VBA & génération PAC Interne via Excel
├── pac_interne_processor.py       # Traitement des métadonnées et génération directe Word
├── add_pac_button.py              # Script utilitaire d'ajout du bouton macro dans Excel
├── PAC_Interne_Macro.bas          # Code source VBA de la macro PAC Interne
├── PAC_Interne_VBA_Code.txt       # Instructions textuelles d'injection de la macro
├── Add_PAC_Button_Instructions.txt# Guide pas-à-pas pour l'intégration bureautique
├── test_ocr.py                    # Script de test unitaire pour l'extraction OCR
├── requirements.txt               # Dépendances Python requises
├── run_app.bat                    # Script de lancement en un clic (Windows)
├── .gitignore                     # Exclusion des fichiers temporaires et données privées
├── templates/                     # Gabarits HTML de l'interface web Flask
│   ├── index.html                 # Interface utilisateur principale moderne
│   └── index_old.html
└── docs/
    └── screenshots/               # 17 captures d'écran et schémas d'architecture du mémoire
```

---

## 🚀 Installation & Guide d'Utilisation

### 1. Prérequis Système
- **Système d'exploitation** : Windows 10/11 (requis pour l'automatisation Office Word/Excel via COM) ou Linux/macOS pour le mode conversion web standard.
- **Python** : Version 3.8 ou supérieure (testé sous Python 3.10 et 3.11).
- **Microsoft Office** : Excel et Word installés pour l'exécution des macros PAC.

### 2. Cloner le Projet

```bash
git clone https://github.com/LibrarianPoms/PDF-OCR-to-EXCEL-Converter.git
cd PDF-OCR-to-EXCEL-Converter
```

### 3. Installer les Dépendances

```bash
pip install -r requirements.txt
```

> **Note :** Si vous traitez des formats PDF vectoriels nécessitant un convertisseur d'images rasterisé externe, installez [poppler-utils for Windows](https://github.com/oschwartz10612/poppler-windows/releases/) et ajoutez son dossier `bin` au `PATH`.

### 4. Démarrer l'Application

Sous Windows, double-cliquez simplement sur **`run_app.bat`** ou exécutez dans votre terminal :

```bash
python app.py
```

L'application démarre immédiatement sur le port **5001** :
👉 **http://localhost:5001**

### 5. Déroulement d'un Traitement Type

1. **Chargement** : Glissez-déposez un ou plusieurs fichiers PDF scannés dans la zone prévue à cet effet.
2. **Extraction** : Patientez quelques instants pendant l'analyse EasyOCR.
3. **Revue** : Vérifiez les champs du fournisseur, numéro de commande et articles dans le tableau interactif. Corrigez ou ajoutez des lignes si besoin.
4. **Export Excel** : Cliquez sur **« Générer Excel (Feuilles séparées) »**. Le fichier est téléchargé instantanément.
5. **Génération PAC** : Dans le classeur Excel généré, cliquez sur le bouton macro **« PAC Interne »** pour créer le Procès-Verbal de Réception Word pré-rempli.

---

## 🔒 Confidentialité & Données Privées

Conformément aux directives de sécurité et de confidentialité :
- Les dossiers d'exploitation temporaires (`temp/` et `uploads/`) contenant les fichiers de travail et les factures/scans réels sont **strictement exclus du contrôle de version via `.gitignore`**.
- Les captures d'écran d'illustration présentées dans ce dépôt et dans le mémoire ont été anonymisées (*« Anonyme 1 »*, *« Anonyme 2 »*) afin de protéger les données commerciales des fournisseurs et de l'opérateur.

---

## 📄 Licence

Ce projet est distribué sous licence **MIT**. Libre d'utilisation, de modification et de distribution dans le cadre de projets d'automatisation documentaire et académiques.
