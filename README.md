# LPVT - Outil d'Analyse de Rapports de Test

## Description

Ce projet fournit un outil basé sur Python avec une interface graphique (GUI) pour analyser les rapports de test HTML générés par les systèmes LPVT, en se concentrant spécifiquement sur les séquences `SEQ-01` et `SEQ-02`. Il offre deux fonctionnalités principales :

1.  **Résumé Statistique :** Agrège des mesures spécifiques (tensions, résistances, etc.) à partir de plusieurs rapports de test dans une seule feuille de calcul Excel pour une analyse statistique sur différentes unités (identifiées par numéro de série et date/heure du test).
2.  **Rapports Détaillés des Échecs :** Génère des rapports texte individuels pour chaque numéro de série, résumant les résultats de tous les tests effectués sur cette unité, mettant en évidence les échecs et extrayant des informations clés comme les valeurs de résistance et les marquages CMS (pour SEQ-01).

L'outil est conçu pour rationaliser le processus d'examen des résultats des tests, d'identification des tendances et de diagnostic des échecs.

## Fonctionnalités

*   **Interface graphique moderne :** Construite avec Tkinter et ttk pour une expérience utilisateur claire.
*   **Sélection de Répertoire :** Sélectionnez facilement le répertoire parent contenant les rapports de test.
*   **Détection Automatique des Rapports :** Trouve les fichiers `SEQ-01*.html` et `SEQ-02*.html` dans la structure du répertoire sélectionné.
*   **Analyse Statistique Sélective (`LPVT_Gestion_Rapports.py`) :**
    *   Choisissez des paramètres de test prédéfinis spécifiques (tensions, résistances) à inclure dans le rapport statistique.
    *   Recherchez et filtrez la liste des paramètres disponibles.
    *   Option pour trier les résultats chronologiquement par numéro de série et date/heure du test.
    *   Génère un fichier `.xlsx` consolidé (utilisant pandas) dans le répertoire source.
    *   Ouvre automatiquement le fichier Excel généré.
*   **Rapports Détaillés par Numéro de Série (`Affiche_resultats.py` - intégré à l'interface graphique) :**
    *   Traite les rapports organisés en sous-répertoires nommés par numéro de série.
    *   Extrait le statut global du test (`Passed`/`Failed`/`Terminated`).
    *   Pour `SEQ-01`, extrait les résistances calculées, les résistances recommandées à monter et les marquages CMS correspondants.
    *   Détaille les étapes de test échouées/terminées (`Failed`/`Terminated`), incluant la configuration (ex: `ROUE CODEUSE`), les mesures et les messages d'erreur.
    *   Génère un rapport résumé `.txt` pour chaque numéro de série, ordonnant chronologiquement les exécutions de test.
    *   Ouvre automatiquement les rapports `.txt` générés (utilisant Notepad sous Windows).
    *   Inclut une fenêtre de progression lors de la génération de rapports détaillés pour plusieurs numéros de série.
*   **Génération Autonome de Rapports Détaillés :** `Affiche_resultats.py` peut aussi être exécuté indépendamment pour traiter un répertoire parent sélectionné.
*   **Aide & Statut :** Message d'aide de base disponible via un bouton et une barre d'état pour le retour utilisateur.
*   **Gestion de l'Encodage :** Tente de gérer les problèmes potentiels d'encodage de caractères.

## Prérequis

*   **Python :** Version 3.6 ou supérieure recommandée.
*   **Bibliothèques Requises :**
    *   `pandas`
    *   `beautifulsoup4`
    *   `openpyxl` (pour écrire les fichiers `.xlsx` avec pandas)

*(Tkinter est généralement inclus avec les distributions Python standard)*

## Installation

1.  **Clonez le dépôt :**
    ```bash
    git clone <url-de-votre-depot>
    cd <repertoire-du-depot>
    ```

2.  **Installez les dépendances :**
    Il est recommandé d'utiliser un environnement virtuel.
    ```bash
    # Créez un environnement virtuel (optionnel mais recommandé)
    python -m venv venv
    # Activez-le (Windows)
    .\venv\Scripts\activate
    # Activez-le (Linux/macOS)
    source venv/bin/activate

    # Installez les paquets requis
    pip install -r requirements.txt
    ```

    Créez un fichier `requirements.txt` à la racine du projet avec le contenu suivant :
    ```txt
    # requirements.txt
    pandas
    beautifulsoup4
    openpyxl
    ```

## Utilisation

### Application Principale (Rapports Statistiques & Détaillés)

1.  **Exécutez le script principal :**
    ```bash
    python LPVT_Gestion_Rapports.py
    ```
2.  **Sélectionnez le Répertoire :** Cliquez sur "Sélectionner un répertoire" et choisissez le **répertoire parent** qui contient les sous-répertoires nommés d'après les numéros de série (ex: `C:\RapportsTests`). L'application s'attend à une structure comme :
    ```
    <Répertoire Parent>/
    ├── SN12345/
    │   ├── SEQ-01_LPVT_Report[...][...].html
    │   ├── SEQ-02_LPVT_Report[...][...].html
    │   └── ...
    ├── SN67890/
    │   ├── SEQ-01_LPVT_Report[...][...].html
    │   └── ...
    └── ...
    ```
3.  **Sélectionnez les Tests (pour le Rapport Statistique) :**
    *   La liste sous "Sélection des tests à analyser" se remplira avec les valeurs extractibles prédéfinies.
    *   Utilisez la barre de recherche pour filtrer la liste.
    *   Sélectionnez les tests/valeurs désirés en utilisant Maj+clic ou Ctrl+clic.
    *   Utilisez "Tout sélectionner" / "Tout désélectionner" si besoin.
    *   Cochez/décochez "Tri chronologique par n° de série" pour contrôler le tri dans la sortie Excel.
4.  **Générez le Rapport Statistique :** Cliquez sur "Générer le rapport statistique". Un fichier Excel nommé `statistiques_SEQ01_SEQ02.xlsx` sera créé dans le répertoire parent sélectionné et ouvert automatiquement.
5.  **Générez les Rapports Détaillés :** Cliquez sur "Générer les rapports détaillés".
    *   L'application va parcourir chaque sous-répertoire (numéro de série) dans le répertoire parent sélectionné.
    *   Pour chaque numéro de série, elle analyse tous les rapports HTML trouvés, génère un fichier `rapport_<NumeroDeSerie>.txt` dans le sous-répertoire de ce numéro de série, et l'ouvre (dans Notepad).
    *   Une fenêtre de progression montrera l'état.

### Génération Autonome de Rapports Détaillés

Si vous avez seulement besoin des rapports `.txt` détaillés par numéro de série :

1.  **Exécutez le script `Affiche_resultats.py` :**
    ```bash
    python Affiche_resultats.py
    ```
2.  **Sélectionnez le Répertoire :** Une boîte de dialogue vous invitera à sélectionner le **répertoire parent** contenant les sous-répertoires des numéros de série (même structure que ci-dessus).
3.  Le script traitera chaque sous-répertoire, générera les rapports `.txt`, et les ouvrira, en affichant une fenêtre de progression.

## Fonctionnement

*   **`LPVT_Gestion_Rapports.py` :** Utilise `tkinter` pour l'interface graphique. Il parcourt le répertoire sélectionné pour trouver les fichiers pertinents. Pour les rapports statistiques, il utilise des **Expressions Régulières (`re`)** pour extraire des points de données spécifiques du contenu HTML en fonction de la liste de tests prédéfinie. `pandas` est utilisé pour structurer ces données et les exporter vers Excel. Il appelle `Affiche_resultats.traiter_repertoire_serie` pour générer les rapports détaillés.
*   **`Affiche_resultats.py` :** Utilise **BeautifulSoup (`bs4`)** pour analyser la structure HTML des rapports de test. Il extrait le statut global, les données de résistance (pour SEQ-01) et les détails des tests échoués en naviguant dans les balises et classes HTML. Il met en forme ces informations dans un rapport texte lisible. La `ProgressWindow` utilise `tkinter.Toplevel`. `subprocess` est utilisé pour ouvrir les fichiers texte générés.

## Améliorations Futures Possibles

*   Gestion des erreurs plus robuste lors de l'analyse des fichiers.
*   Permettre la personnalisation de la liste des tests prédéfinis (ex: via un fichier de configuration).
*   Support d'autres formats de rapports si nécessaire.
*   Ajouter des calculs statistiques plus avancés ou des graphiques de base.
*   Empaqueter l'application en un exécutable autonome (ex: en utilisant PyInstaller).
*   Utiliser des méthodes indépendantes de la plateforme pour ouvrir les fichiers générés (au lieu de coder en dur `notepad.exe`).
*   Refactoriser la logique d'extraction pour une meilleure maintenabilité.

## Licence

[Spécifiez votre licence ici, ex: Licence MIT, GPL, Propriétaire, etc.]
