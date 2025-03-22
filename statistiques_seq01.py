import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from bs4 import BeautifulSoup
import pandas as pd
from typing import Dict, List, Any, Set, Optional, Tuple
import io

class StatsTestsWindow:
    """
    Fenêtre principale pour l'analyse statistique des tests SEQ-01 et SEQ-02
    """
    def __init__(self):
        """Initialise l'application"""
        # Configuration de l'encodage
        self.configurer_encodage()
        
        # Variables de l'application
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        self.tests_disponibles = []  # Liste de tuples (nom_complet_test, test_parent, identifiant)
        self.tests_selectionnes = []
        self.data_tests = {}
        self.repertoire_parent = ""
        
        # Interface graphique
        self.root = tk.Tk()
        self.root.title("Statistiques SEQ-01 et SEQ-02")
        self.root.geometry("900x700")
        self.creer_interface()
        
    def configurer_encodage(self):
        """Configure l'encodage pour les caractères accentués"""
        try:
            if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except (AttributeError, TypeError):
            # En cas d'erreur, on continue sans changer l'encodage
            pass

    def creer_interface(self):
        """Crée l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Bouton pour sélectionner le répertoire
        btn_repertoire = ttk.Button(
            main_frame, 
            text="Sélectionner répertoire", 
            command=self.selectionner_repertoire
        )
        btn_repertoire.pack(pady=10)
        
        # Label pour afficher le répertoire sélectionné
        self.lbl_repertoire = ttk.Label(main_frame, text="Aucun répertoire sélectionné")
        self.lbl_repertoire.pack(pady=5)
        
        # Frame pour la sélection des tests
        frame_selection = ttk.LabelFrame(main_frame, text="Sélection des tests")
        frame_selection.pack(fill=tk.BOTH, expand=True, pady=10, padx=5)
        
        # Liste des tests disponibles avec scrollbar
        frame_tests = ttk.Frame(frame_selection)
        frame_tests.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scroll_y = ttk.Scrollbar(frame_tests, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox_tests = tk.Listbox(
            frame_tests, 
            selectmode=tk.MULTIPLE, 
            yscrollcommand=scroll_y.set,
            font=("Courier", 10)  # Police à largeur fixe pour faciliter la lecture
        )
        self.listbox_tests.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.config(command=self.listbox_tests.yview)
        
        # Boutons de sélection
        frame_btns = ttk.Frame(frame_selection)
        frame_btns.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(
            frame_btns, 
            text="Sélectionner tout",
            command=self.selectionner_tous_tests
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            frame_btns, 
            text="Désélectionner tout",
            command=self.deselectionner_tous_tests
        ).pack(side=tk.LEFT, padx=5)
        
        # Bouton pour générer le rapport statistique
        btn_generer = ttk.Button(
            main_frame, 
            text="Générer statistiques", 
            command=self.generer_statistiques
        )
        btn_generer.pack(pady=10)
    
    def selectionner_tous_tests(self):
        """Sélectionne tous les tests dans la liste"""
        self.listbox_tests.selection_set(0, tk.END)
        
    def deselectionner_tous_tests(self):
        """Désélectionne tous les tests dans la liste"""
        self.listbox_tests.selection_clear(0, tk.END)
        
    def selectionner_repertoire(self):
        """Permet à l'utilisateur de sélectionner un répertoire et charge les tests disponibles"""
        repertoire = filedialog.askdirectory(
            title="Sélectionnez le répertoire contenant les fichiers SEQ-01/SEQ-02"
        )
        if not repertoire:
            return
            
        self.repertoire_parent = repertoire
        self.lbl_repertoire.config(text=f"Répertoire: {os.path.basename(repertoire)}")
        
        # Rechercher les fichiers SEQ-01/SEQ-02 et charger les tests disponibles
        self.charger_fichiers_seq(repertoire)
        
    def charger_fichiers_seq(self, repertoire):
        """Cherche les fichiers SEQ-01/SEQ-02 et charge les tests disponibles"""
        # Chercher les fichiers SEQ-01 et SEQ-02
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        
        # Chercher récursivement dans les sous-répertoires
        for dossier_racine, sous_dossiers, fichiers in os.walk(repertoire):
            for fichier in fichiers:
                if fichier.startswith("SEQ-01") and fichier.lower().endswith(".html"):
                    chemin_complet = os.path.join(dossier_racine, fichier)
                    self.fichiers_seq01.append(chemin_complet)
                elif fichier.startswith("SEQ-02") and fichier.lower().endswith(".html"):
                    chemin_complet = os.path.join(dossier_racine, fichier)
                    self.fichiers_seq02.append(chemin_complet)
        
        total_fichiers = len(self.fichiers_seq01) + len(self.fichiers_seq02)
        if total_fichiers == 0:
            messagebox.showinfo("Information", "Aucun fichier SEQ-01 ou SEQ-02 trouvé.")
            return
        
        # Analyser un fichier pour récupérer la liste des tests disponibles
        messagebox.showinfo("Information", f"{len(self.fichiers_seq01)} fichiers SEQ-01 et {len(self.fichiers_seq02)} fichiers SEQ-02 trouvés.\nChargement des tests disponibles...")
        
        # Créer manuellement la liste des tests disponibles selon les spécifications
        self.creer_liste_tests_predefinies()
        self.afficher_tests_disponibles()
    
    def creer_liste_tests_predefinies(self):
        """
        Crée une liste prédéfinie des tests à extraire, basée sur les spécifications
        """
        self.tests_disponibles = [
            # Format: (nom_complet, test_parent, identifiant)
            # --------- Tests SEQ-01 ---------
            # Test des alimentations à 24VDC
            ("SEQ-01 > Test des alimentations à 24VDC > Lecture mesure +16V AG34461A", 
                "seq01_24vdc", "24VDC_+16V"),
            ("SEQ-01 > Test des alimentations à 24VDC > Lecture mesure -16V AG34461A", 
                "seq01_24vdc", "24VDC_-16V"),
            ("SEQ-01 > Test des alimentations à 24VDC > Lecture mesure +5V AG34461A", 
                "seq01_24vdc", "24VDC_+5V"),
            ("SEQ-01 > Test des alimentations à 24VDC > Lecture mesure -5V AG34461A", 
                "seq01_24vdc", "24VDC_-5V"),
            
            # Test des alimentations à 115VAC
            ("SEQ-01 > Test des alimentations à 115VAC > Lecture mesure +16V AG34461A", 
                "seq01_115vac", "115VAC_+16V"),
            ("SEQ-01 > Test des alimentations à 115VAC > Lecture mesure -16V AG34461A", 
                "seq01_115vac", "115VAC_-16V"),
            
            # Calcul des résistances
            ("SEQ-01 > Calcul des résistances > Résistance R46 calculée", 
                "seq01_resistances", "R46_calculee"),
            ("SEQ-01 > Calcul des résistances > Résistance R46 à monter", 
                "seq01_resistances", "R46_monter"),
            ("SEQ-01 > Calcul des résistances > Résistance R47 calculée", 
                "seq01_resistances", "R47_calculee"),
            ("SEQ-01 > Calcul des résistances > Résistance R47 à monter", 
                "seq01_resistances", "R47_monter"),
            ("SEQ-01 > Calcul des résistances > Résistance R48 calculée", 
                "seq01_resistances", "R48_calculee"),
            ("SEQ-01 > Calcul des résistances > Résistance R48 à monter", 
                "seq01_resistances", "R48_monter"),
                
            # --------- Tests SEQ-02 ---------
            # Tests de rapport de transfert
            ("SEQ-02 > Test 1.9Un sur 2 voies en 19VDC", 
                "seq02_transfert", "Test_19VDC"),
            ("SEQ-02 > Test 1.9Un sur 2 voies en 115VAC", 
                "seq02_transfert", "Test_115VAC"),
                
            # Précision du rapport de transfert pour chaque gamme et voie
            ("SEQ-02 > Roue codeuse F (Gamme 1) > Voie U", 
                "seq02_precision", "F1_U"),
            ("SEQ-02 > Roue codeuse F (Gamme 1) > Voie V", 
                "seq02_precision", "F1_V"),
            ("SEQ-02 > Roue codeuse F (Gamme 1) > Voie W", 
                "seq02_precision", "F1_W"),
            ("SEQ-02 > Roue codeuse E (Gamme 2) > Voie U", 
                "seq02_precision", "E2_U"),
            ("SEQ-02 > Roue codeuse E (Gamme 2) > Voie V", 
                "seq02_precision", "E2_V"),
            ("SEQ-02 > Roue codeuse E (Gamme 2) > Voie W", 
                "seq02_precision", "E2_W"),
            # Ajout des autres gammes...
            ("SEQ-02 > Roue codeuse D (Gamme 3) > Voie U", 
                "seq02_precision", "D3_U"),
            ("SEQ-02 > Roue codeuse D (Gamme 3) > Voie V", 
                "seq02_precision", "D3_V"),
            ("SEQ-02 > Roue codeuse D (Gamme 3) > Voie W", 
                "seq02_precision", "D3_W"),
            ("SEQ-02 > Roue codeuse 1 (Gamme 15) > Voie U", 
                "seq02_precision", "115_U"),
            ("SEQ-02 > Roue codeuse 1 (Gamme 15) > Voie V", 
                "seq02_precision", "115_V"),
            ("SEQ-02 > Roue codeuse 1 (Gamme 15) > Voie W", 
                "seq02_precision", "115_W"),
        ]
    
    def afficher_tests_disponibles(self):
        """Affiche les tests disponibles dans la listbox avec formatage hiérarchique"""
        # Vider la listbox
        self.listbox_tests.delete(0, tk.END)
        
        # Ajouter les tests avec indentation pour indiquer la hiérarchie
        for nom_complet, test_parent, identifiant in self.tests_disponibles:
            self.listbox_tests.insert(tk.END, nom_complet)
    
    def generer_statistiques(self):
        """Génère les statistiques pour les tests sélectionnés"""
        # Récupérer les indices des tests sélectionnés
        indices = self.listbox_tests.curselection()
        if not indices:
            messagebox.showinfo("Information", "Veuillez sélectionner au moins un test.")
            return
        
        # Récupérer les noms des tests sélectionnés
        self.tests_selectionnes = [self.tests_disponibles[i] for i in indices]
        
        # Analyser les fichiers pour récupérer les données
        messagebox.showinfo("Information", f"Génération des statistiques pour {len(self.tests_selectionnes)} tests...")
        self.analyser_fichiers()
    
    def extraire_valeur_seq01(self, html_content, test_parent, identifiant):
        """
        Extrait la valeur d'une mesure spécifique du SEQ-01 à partir du contenu HTML en utilisant des expressions régulières
        """
        # Traitement pour les alimentations 24VDC
        if test_parent == "seq01_24vdc":
            # Extraire le bloc correspondant à ce test
            block_match = re.search(r"Test des alimentations à 24VDC(.*?)Test des alimentations à 115VAC", html_content, re.DOTALL)
            if not block_match:
                return None
            
            block = block_match.group(1)
            
            # Extraire les valeurs selon l'identifiant
            if identifiant == "24VDC_+16V":
                m_plus16 = re.search(r"Lecture mesure \+16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_plus16.group(1).strip() if m_plus16 else None
            elif identifiant == "24VDC_-16V":
                m_minus16 = re.search(r"Lecture mesure -16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_minus16.group(1).strip() if m_minus16 else None
            elif identifiant == "24VDC_+5V":
                m_plus5 = re.search(r"Lecture mesure \+5V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_plus5.group(1).strip() if m_plus5 else None
            elif identifiant == "24VDC_-5V":
                m_minus5 = re.search(r"Lecture mesure -5V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_minus5.group(1).strip() if m_minus5 else None
        
        # Traitement pour les alimentations 115VAC
        elif test_parent == "seq01_115vac":
            # Extraire le bloc correspondant à ce test
            block_match = re.search(r"Test des alimentations à 115VAC(.*?)Calcul des résistances", html_content, re.DOTALL)
            if not block_match:
                return None
            
            block = block_match.group(1)
            
            # Extraire les valeurs selon l'identifiant
            if identifiant == "115VAC_+16V":
                m_plus16 = re.search(r"Lecture mesure \+16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_plus16.group(1).strip() if m_plus16 else None
            elif identifiant == "115VAC_-16V":
                m_minus16 = re.search(r"Lecture mesure -16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
                return m_minus16.group(1).strip() if m_minus16 else None
        
        # Traitement pour le calcul des résistances
        elif test_parent == "seq01_resistances":
            if identifiant == "R46_calculee":
                r46_calc = re.search(r"Résistance R46 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html_content, re.DOTALL)
                return r46_calc.group(1).strip() if r46_calc else None
            elif identifiant == "R46_monter":
                r46_monter = re.search(r"Résistance R46 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r46_monter.group(1).strip() if r46_monter else None
            elif identifiant == "R47_calculee":
                r47_calc = re.search(r"Résistance R47 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html_content, re.DOTALL)
                return r47_calc.group(1).strip() if r47_calc else None
            elif identifiant == "R47_monter":
                r47_monter = re.search(r"Résistance R47 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r47_monter.group(1).strip() if r47_monter else None
            elif identifiant == "R48_calculee":
                r48_calc = re.search(r"Résistance R48 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html_content, re.DOTALL)
                return r48_calc.group(1).strip() if r48_calc else None
            elif identifiant == "R48_monter":
                r48_monter = re.search(r"Résistance R48 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r48_monter.group(1).strip() if r48_monter else None
                
        return None
        
    def extraire_valeur_seq02(self, html_content, test_parent, identifiant):
        """
        Extrait la valeur d'une mesure spécifique du SEQ-02 à partir du contenu HTML en utilisant des expressions régulières
        """
        # Tests 1.9Un
        if test_parent == "seq02_transfert":
            if identifiant == "Test_19VDC":
                pattern = r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+19VDC.*?Lecture\s+mesure\s+-16V\s+AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>"
                match = re.search(pattern, html_content, re.DOTALL | re.IGNORECASE)
                return match.group(1).strip() if match else None
            elif identifiant == "Test_115VAC":
                pattern = r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+115VAC.*?Lecture\s+mesure\s+-16V\s+AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>"
                match = re.search(pattern, html_content, re.DOTALL | re.IGNORECASE)
                return match.group(1).strip() if match else None
        
        # Précision du rapport de transfert
        elif test_parent == "seq02_precision":
            # Extraire la gamme et la voie à partir de l'identifiant (ex: F1_U)
            if len(identifiant) < 4:
                return None
                
            roue_codeuse = identifiant[0]
            gamme_num = identifiant[1:-2]
            voie = identifiant[-1]
            
            # Pattern pour trouver les sections de précision du rapport de transfert
            pattern = fr"ROUE CODEUSE = {roue_codeuse} \.\. GAMME = {gamme_num} \.\. Voie {voie}.*?Valeur r[^<]*inject[^<]*/ Valeur sortie attendue[^/]*/ Valeur sortie mesur[^/]*/ Pr[^<]*:</td>\s*<td[^>]*>\s*([^<]+)</td>"
            match = re.search(pattern, html_content, re.DOTALL | re.IGNORECASE)
            
            if match:
                # Renvoyer seulement la précision (dernier élément après le /)
                valeurs = match.group(1).strip().split("/")
                if len(valeurs) >= 4:
                    precision = valeurs[3].strip()
                    return precision
            
            return None
            
        return None
    
    def extraire_numero_serie(self, html_content):
        """Extrait le numéro de série depuis le contenu HTML en utilisant des expressions régulières"""
        # Recherche du numéro de série dans l'en-tête
        sn_match = re.search(r'Serial Number:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL)
        if sn_match and sn_match.group(1).strip() != "NONE":
            return sn_match.group(1).strip()
        
        # Recherche dans le contenu pour "Numéro de série de la carte en test"
        serie_match = re.search(r'série[^<]*</td>\s*<td[^>]*class="value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL)
        if serie_match:
            return serie_match.group(1).strip()
        
        # Si rien n'est trouvé, on retourne None
        return None
    
    def extraire_date_heure_du_nom(self, nom_fichier):
        """
        Extrait la date et l'heure du nom du fichier au format SEQ-XX_LPVT_Report[HH MM SS][JJ MM AAAA].html
        Retourne un tuple (date_formatee, heure_formatee)
        """
        # Format attendu: SEQ-XX_LPVT_Report[HH MM SS][JJ MM AAAA].html
        match = re.search(r'\[(\d+ \d+ \d+)\]\[(\d+ \d+ \d+)\]', nom_fichier)
        if match:
            heure_brute = match.group(1)  # HH MM SS
            date_brute = match.group(2)   # JJ MM AAAA
            
            # Formater pour avoir un format plus lisible et utilisable comme identifiant
            heure_formatee = heure_brute.replace(" ", ":")
            date_formatee = date_brute.replace(" ", "/")
            
            return date_formatee, heure_formatee
            
        return "date_inconnue", "heure_inconnue"
    
    def extraire_date_heure(self, html_content, nom_fichier):
        """Extrait la date et l'heure du test depuis le contenu HTML et/ou du nom de fichier"""
        # Tenter d'abord de récupérer du contenu HTML
        date_match = re.search(r'Date:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*<b>([^<]+)</b>', html_content, re.DOTALL)
        date = date_match.group(1).strip() if date_match else None
        
        time_match = re.search(r'Time:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*<b>([^<]+)</b>', html_content, re.DOTALL)
        heure = time_match.group(1).strip() if time_match else None
        
        # Si on n'a pas trouvé dans le HTML, essayer avec le nom de fichier
        if not date or not heure:
            date, heure = self.extraire_date_heure_du_nom(nom_fichier)
        
        return date, heure
    
    def analyser_fichiers(self):
        """Analyse tous les fichiers SEQ-01/SEQ-02 et collecte les données pour les tests sélectionnés"""
        donnees = {}  # Dictionnaire pour stocker les données: {identifiant_unique: {test1: valeur1, test2: valeur2, ...}}
        
        # Analyse des fichiers SEQ-01
        for fichier in self.fichiers_seq01:
            try:
                with open(fichier, "r", encoding="iso-8859-1", errors="replace") as f:
                    html_content = f.read()
                
                # Extraire le numéro de série
                numero_serie = self.extraire_numero_serie(html_content) or os.path.basename(os.path.dirname(fichier))
                
                # Extraire la date et l'heure du nom de fichier
                nom_fichier = os.path.basename(fichier)
                date, heure = self.extraire_date_heure(html_content, nom_fichier)
                
                # Créer un identifiant unique avec le numéro de série, la date et l'heure
                identifiant_unique = f"{numero_serie} [{date}][{heure}]"
                
                # Initialiser le dictionnaire pour ce numéro de série si nécessaire
                if identifiant_unique not in donnees:
                    donnees[identifiant_unique] = {
                        "Numéro de série": numero_serie,
                        "Date": date,
                        "Heure": heure,
                        "Type": "SEQ-01"
                    }
                
                # Extraire les valeurs SEQ-01 pour chaque test sélectionné
                for nom_complet, test_parent, identifiant in self.tests_selectionnes:
                    if test_parent.startswith("seq01_"):
                        valeur = self.extraire_valeur_seq01(html_content, test_parent, identifiant)
                        if valeur:
                            donnees[identifiant_unique][nom_complet] = valeur
            
            except Exception as e:
                print(f"Erreur lors de l'analyse du fichier SEQ-01 {fichier}: {e}")
                import traceback
                traceback.print_exc()
        
        # Analyse des fichiers SEQ-02
        for fichier in self.fichiers_seq02:
            try:
                with open(fichier, "r", encoding="iso-8859-1", errors="replace") as f:
                    html_content = f.read()
                
                # Extraire le numéro de série
                numero_serie = self.extraire_numero_serie(html_content) or os.path.basename(os.path.dirname(fichier))
                
                # Extraire la date et l'heure du nom de fichier
                nom_fichier = os.path.basename(fichier)
                date, heure = self.extraire_date_heure(html_content, nom_fichier)
                
                # Créer un identifiant unique avec le numéro de série, la date et l'heure
                identifiant_unique = f"{numero_serie} [{date}][{heure}]"
                
                # Initialiser le dictionnaire pour ce numéro de série si nécessaire
                if identifiant_unique not in donnees:
                    donnees[identifiant_unique] = {
                        "Numéro de série": numero_serie,
                        "Date": date,
                        "Heure": heure,
                        "Type": "SEQ-02"
                    }
                
                # Extraire les valeurs SEQ-02 pour chaque test sélectionné
                for nom_complet, test_parent, identifiant in self.tests_selectionnes:
                    if test_parent.startswith("seq02_"):
                        valeur = self.extraire_valeur_seq02(html_content, test_parent, identifiant)
                        if valeur:
                            donnees[identifiant_unique][nom_complet] = valeur
            
            except Exception as e:
                print(f"Erreur lors de l'analyse du fichier SEQ-02 {fichier}: {e}")
                import traceback
                traceback.print_exc()
        
        self.data_tests = donnees
        self.creer_tableau_statistiques()
    
    def creer_tableau_statistiques(self):
        """Crée un tableau statistique à partir des données collectées"""
        if not self.data_tests:
            messagebox.showinfo("Information", "Aucune donnée collectée.")
            return
        
        # Créer un DataFrame pandas
        df = pd.DataFrame.from_dict(self.data_tests, orient='index')
        
        # Sauvegarder en Excel
        chemin_excel = os.path.join(self.repertoire_parent, "statistiques_SEQ01_SEQ02.xlsx")
        try:
            # Utilisation d'un nom de feuille sans caractères spéciaux
            df.to_excel(chemin_excel, sheet_name="Statistiques_SEQ01_02")
            messagebox.showinfo(
                "Succès", 
                f"Tableau statistique créé avec succès!\n\nFichier: {chemin_excel}"
            )
            
            # Ouvrir le fichier Excel
            os.startfile(chemin_excel)
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la création du tableau: {e}")
    
    def lancer(self):
        """Lance l'application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = StatsTestsWindow()
    app.lancer()