# Site Analyzer

## 📝 Description
Cette application Streamlit automatise l’analyse de la performance d'une installation solaire hybride de façon globale. Elle permet de charger un fichier de données, de les analyser, et de générer un rapport détaillé au format Word avec graphiques et tableaux.

## 🚀 Fonctionnalités principales

1- **Importation des données**  
(détails et consignes disponibles dans l'onglet "Indications" de l'application | Fichiers test disponibles dans le fichier data)

2- **Calculs et visualisations interactives** 
 
    - Identification de la production énergétique globale
    - Mis en évidence de l'état de fonctionnement
    - Suivi temporel de la performance

3- **Génération automatisée d’un rapport Word personnalisable incluant tableaux et graphiques**

4- **Analyse personnalisée avec sélection flexible**   
    - Sélection précise de la période temporelle  
    - Filtrage des courbes à inclure dans certaines visualisations  
    - Choix des sections à inclure dans le rapport


## ▶️ Mode d'emploi

### 📌 Prérequis

  - Python 3.8 ou plus récent installé
  - Librairies Python (voir fichier `requirements.txt`)
  - Accès à un terminal ou à VS Code


### 📌 Etapes à suivre

1. #### Télécharger le projet
  - **Option A : Cloner le dépôt GitHub si vous avez git**
    --> git clone https://lien_du_depot.git
    --> cd nom_du_dossier
  - **Option B : Télécharger le dossier compressé (.zip) et l’extraire**


2. #### Se placer dans le dossier du projet
  - **Option 1 : Utilisation du terminal classique du système**
    --> cd chemin/vers/le/dossier_du_projet
  - **Option 2 : Utilisation de VSCode**
    --> Fichier > Ouvrir un dossier et sélectionner le dossier du projet

3. #### Créer un environnement virtuel 
  - **Option 1 : Utilisation du terminal classique du système**
    --> python -m venv env
  - **Option 2 : Utilisation de VSCode**
    (*ouvrir le terminal intégré de VS Code*)
    --> python -m venv env

4. #### Activer l'environnement virtuel
  - **Sur Windows**
    --> .\env\Scripts\activate
  - **Sur macOS/Linux**
    --> source env/bin/activate

5. #### Installer les dépendances
   --> pip install -r requirements.txt

7. #### Lancer l’application
   --> streamlit run app.py

## 👩‍💻 Auteur & Contact
Développé par Amboara RASOLOFOARIMANANA  
amboara.rasolofo@gmail.com
