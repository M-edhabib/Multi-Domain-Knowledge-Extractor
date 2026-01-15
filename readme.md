# 📊 Analyseur Automatisé de Bilans Financiers (DPS)

Ce projet est une **Preuve de Concept (PoC)** permettant d'automatiser l'extraction, l'analyse et l'interprétation de rapports financiers au format PDF (type DPS). 

L'outil transforme des documents bruts en données structurées (Excel), génère un script d'analyse textuelle et crée même une version audio (MP3) pour une écoute rapide des conclusions.

## ✨ Fonctionnalités

- **Extraction Intelligente** : Utilise `PyMuPDF` et `RegEx` pour cibler des indicateurs précis (CA, EBE, Marge brute, etc.) malgré la complexité des PDF.
- **Analyse Métier** : Calcule automatiquement si les performances sont en amélioration (modeste ou significative) par rapport à l'année précédente et à la médiane du marché.
- **Reporting Multi-format** :
  - **Excel** : Exportation des données nettoyées et segmentées par onglets.
  - **Texte** : Génération d'un script d'interprétation synthétique.
  - **Audio (TTS)** : Conversion du rapport en fichier MP3 via `pyttsx3`.
- **Interface Graphique** : GUI simple avec `Tkinter` pour sélectionner les dossiers à traiter.

## 🛠️ Technologies utilisées

- **Langage** : Python 3.x
- **Bibliothèques clés** :
  - `pandas` & `xlsxwriter` : Manipulation et export de données.
  - `PyMuPDF (fitz)` : Lecture de fichiers PDF.
  - `pyttsx3` : Synthèse vocale.
  - `Tkinter` : Interface utilisateur.

## 🚀 Installation et Utilisation

### 1. Prérequis
Assurez-vous d'avoir Python installé. Puis, installez les dépendances :

```bash
pip install pandas PyMuPDF xlsxwriter pyttsx3
