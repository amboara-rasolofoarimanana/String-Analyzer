# String Analyzer
 

## 📝 Description
Cette application Streamlit automatise l’analyse de la performance des panneaux solaires sur un site photovoltaïque. Elle permet de charger différents types de données, de les analyser, et de générer un rapport détaillé au format Word avec graphiques et tableaux.

---

## 🚀 Fonctionnalités principales

1- **Importation des données relatives à la production et à la caractérisation des panneaux photovoltaïques**  
(détails et consignes disponibles dans l'onglet "Indications" de l'application)

2- **Calculs et visualisations interactives** 
 
    - Identification de la puissance et de l'énergie générées par string et par onduleur 
    - Comparaison entre la puissance moyenne réelle et théorique attendue
    - Comparaison entre l'énergie totale réelle et théorique attendue
    - Suivi temporel de la puissance générée
    - Identification des ratios de performance  
    - Identification des strings les plus performants et les moins performants 
    - Mise en évidence des anomalies et détection des strings suspects 
    - Suivi de l’évolution mensuelle des ratios de performance
    - Comparaison par string des ratios de performance des différents onduleurs du site

3- **Génération automatisée d’un rapport Word personnalisable incluant tableaux et graphiques**

4- **Analyse personnalisée avec sélection flexible**   
    - Choix de l'onduleur à analyser  
    - Sélection précise de la période temporelle  
    - Filtrage des strings photovoltaïques à inclure dans certaines visualisations  
    - Choix des sections à inclure dans le rapport

Cette application offre ainsi un outil puissant et intuitif pour optimiser le suivi et la maintenance des installations photovoltaïques.


---

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
    *ouvrir le terminal intégré de VS Code  (Ctrl + ù ou Terminal > Nouveau terminal)*
    --> python -mvenv env

4. #### Activer l'environnement virtuel
*Valable pour Option 1 et 2*
  - **Sur Windows**
    --> .\env\Scripts\activate
  - **Sur macOS/Linux**
    --> source env/bin/activate

5. #### Installer les dépendances 
  *Valable pour Option 1 et 2*  
  --> pip install -r requirements.txt

6. #### Lancer l’application
  *Valable pour Option 1 et 2*  
  --> streamlit run app.py

## 👩‍💻 Auteur & Contact
Développée par Amboara RASOLOFOARIMANANA
amboara.rasolofo@gmail.com
