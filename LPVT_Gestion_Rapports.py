import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from bs4 import BeautifulSoup
import pandas as pd
from typing import Dict, List, Any, Set, Optional, Tuple
import io
import subprocess
from Affiche_resultats import traiter_repertoire_serie, ProgressWindow
import webbrowser
import openpyxl

def extraire_defauts_precision_transfert(html_content):
    """
    Extrait les défauts de la partie 'Vérification de la précision du rapport de transfert'
    Retourne une liste de tuples (nom, precision)
    """
    resultats = []
    soup = BeautifulSoup(html_content, "html.parser")
    # Parcourt tous les tableaux du rapport
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        i = 0
        while i < len(rows):
            tds = rows[i].find_all("td")
            if len(tds) == 2 and "Status:" in tds[0].get_text(strip=True):
                status = tds[1].get_text(strip=True)
                if status.lower() == "failed":
                    # Cherche la ligne "ROUE CODEUSE = ... Voie ..."
                    voie = None
                    precision = None
                    # Cherche en avant dans les 2 prochaines lignes
                    for j in range(i+1, min(i+4, len(rows))):
                        tds2 = rows[j].find_all("td")
                        if len(tds2) == 2 and "ROUE CODEUSE" in tds2[0].get_text():
                            voie = tds2[0].get_text(strip=True)
                        if len(tds2) == 2 and "Précision" in tds2[0].get_text():
                            # Prend le dernier nombre après le dernier '/'
                            valeurs = tds2[1].get_text(strip=True)
                            precision = valeurs.split('/')[-1].strip()
                    if voie and precision:
                        # Extrait RC et Voie
                        import re
                        m = re.search(r'ROUE CODEUSE = ([^\.]+) .. GAMME = [^\.]+ .. Voie (\w):', voie)
                        if m:
                            rc = m.group(1).strip()
                            v = m.group(2).strip()
                            nom = f"RC:{rc} - Voie {v}"
                            resultats.append((nom, precision))
            i += 1
    return resultats

class ModernStatsTestsWindow:
    """
    Fenêtre modernisée pour l'analyse statistique des tests SEQ-01 et SEQ-02
    avec toutes les fonctionnalités originales
    """
    def __init__(self):
        """Initialise l'application avec une interface moderne"""
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
        self.root.title("LPVT - Analyse des tests SEQ-01/SEQ-02")
        self.root.geometry("1100x750")
        self.root.minsize(900, 650)
        
        # Style moderne
        self.setup_styles()
        
        # Création de l'interface
        self.creer_interface_moderne()
        
        # Initialiser la liste des tests prédéfinis
        self.creer_liste_tests_predefinies()
    
    def setup_styles(self):
        """Configure les styles modernes pour l'interface"""
        style = ttk.Style()
        
        # Thème général
        style.theme_use('clam')
        
        # Couleurs
        bg_color = "#f0f0f0"
        primary_color = "#2c3e50"
        secondary_color = "#3498db"
        accent_color = "#e74c3c"
        
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'), foreground=primary_color)
        style.configure('Secondary.TButton', foreground='white', background=secondary_color)
        style.configure('Accent.TButton', foreground='white', background=accent_color)
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=25)
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        style.map('Treeview', background=[('selected', secondary_color)])
        
        self.bg_color = bg_color
        self.primary_color = primary_color
        self.secondary_color = secondary_color
        self.accent_color = accent_color
    
    def creer_interface_moderne(self):
        """Crée l'interface utilisateur moderne"""
        # Frame principal avec padding
        main_frame = ttk.Frame(self.root, padding=(15, 15, 15, 15))
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header avec titre et boutons d'action
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Titre
        title_label = ttk.Label(
            header_frame, 
            text="Analyse des tests SEQ-01/SEQ-02", 
            style='Title.TLabel'
        )
        title_label.pack(side=tk.LEFT)
        
        # Boutons d'action à droite
        action_buttons_frame = ttk.Frame(header_frame)
        action_buttons_frame.pack(side=tk.RIGHT)
        
        ttk.Button(
            action_buttons_frame,
            text="Aide",
            command=self.show_help,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        # Section de sélection du répertoire
        dir_frame = ttk.LabelFrame(
            main_frame, 
            text=" Répertoire source ", 
            padding=(10, 5, 10, 10)
        )
        dir_frame.pack(fill=tk.X, pady=(0, 15))
        
        dir_selection_frame = ttk.Frame(dir_frame)
        dir_selection_frame.pack(fill=tk.X)
        
        # Bouton de sélection
        select_dir_btn = ttk.Button(
            dir_selection_frame,
            text="Sélectionner un répertoire",
            command=self.selectionner_repertoire,
            style='Secondary.TButton'
        )
        select_dir_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Affichage du chemin
        self.dir_path_label = ttk.Label(
            dir_selection_frame, 
            text="Aucun répertoire sélectionné",
            foreground="#7f8c8d",
            font=('Segoe UI', 9)
        )
        self.dir_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Info sur les fichiers trouvés
        self.files_info_label = ttk.Label(
            dir_frame,
            text="",
            foreground=self.secondary_color,
            font=('Segoe UI', 9)
        )
        self.files_info_label.pack(fill=tk.X, pady=(5, 0))
        
        # Section de sélection des tests
        tests_frame = ttk.LabelFrame(
            main_frame, 
            text=" Sélection des tests à analyser ", 
            padding=(10, 5, 10, 10))
        tests_frame.pack(fill=tk.BOTH, expand=True)
        
        # Contrôles de sélection en haut
        tests_controls_frame = ttk.Frame(tests_frame)
        tests_controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(
            tests_controls_frame,
            text="Tout sélectionner",
            command=self.selectionner_tous_tests,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            tests_controls_frame,
            text="Tout désélectionner",
            command=self.deselectionner_tous_tests,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT)
        
        # Checkbox de tri chronologique à droite
        self.tri_chrono_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            tests_controls_frame,
            text="Tri chronologique par n° de série",
            variable=self.tri_chrono_var,
            onvalue=True,
            offvalue=False
        ).pack(side=tk.RIGHT)
        
        # Liste des tests avec scrollbar et recherche
        tests_list_container = ttk.Frame(tests_frame)
        tests_list_container.pack(fill=tk.BOTH, expand=True)
        
        # Barre de recherche
        search_frame = ttk.Frame(tests_list_container)
        search_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(search_frame, text="Rechercher:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(
            search_frame,
            textvariable=self.search_var
        )
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_entry.bind('<KeyRelease>', self.filter_tests_list)
        
        # Liste des tests dans un Treeview
        self.tests_tree = ttk.Treeview(
            tests_list_container,
            selectmode='extended',
            columns=('test_id'),
            show='tree',
            height=15
        )
        
        scroll_y = ttk.Scrollbar(
            tests_list_container, 
            orient=tk.VERTICAL, 
            command=self.tests_tree.yview
        )
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tests_tree.configure(yscrollcommand=scroll_y.set)
        self.tests_tree.pack(fill=tk.BOTH, expand=True)
        
        # Style pour les éléments sélectionnés
        self.tests_tree.tag_configure('selected', background=self.secondary_color, foreground='white')
        
        # Boutons d'action en bas
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(
            action_frame,
            text="Générer le rapport statistique",
            command=self.generer_statistiques,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            action_frame,
            text="Générer les rapports détaillés",
            command=self.generer_rapports_detailles,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT)
        
        # Barre de status
        self.status_bar = ttk.Frame(main_frame, height=25)
        self.status_bar.pack(fill=tk.X, pady=(15, 0))
        
        self.status_label = ttk.Label(
            self.status_bar,
            text="Prêt",
            foreground="#7f8c8d",
            font=('Segoe UI', 9)
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Configurer le redimensionnement
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
    
    def show_help(self):
        """Affiche une aide contextuelle"""
        help_text = """
        Aide - Analyse des tests SEQ-01/SEQ-02
        
        1. Sélectionnez un répertoire contenant les fichiers HTML des tests
        2. Sélectionnez les tests à analyser dans la liste
        3. Cliquez sur "Générer le rapport statistique" pour créer un fichier Excel
        4. Utilisez "Générer les rapports détaillés" pour créer des rapports PDF
        
        Options :
        - Tri chronologique : organise les résultats par numéro de série et date
        - Barre de recherche : filtre la liste des tests disponibles
        """
        messagebox.showinfo("Aide", help_text.strip())
    
    def update_status(self, message):
        """Met à jour la barre de status"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def filter_tests_list(self, event=None):
        """Filtre la liste des tests en fonction de la recherche"""
        search_term = self.search_var.get().lower()
        
        # Effacer la liste actuelle
        for item in self.tests_tree.get_children():
            self.tests_tree.delete(item)
        
        # Ajouter les tests qui correspondent au terme de recherche
        for idx, (nom_complet, test_parent, identifiant) in enumerate(self.tests_disponibles):
            if search_term in nom_complet.lower():
                self.tests_tree.insert('', 'end', text=nom_complet, values=(identifiant))
    
    def selectionner_tous_tests(self):
        """Sélectionne tous les tests dans la liste"""
        for item in self.tests_tree.get_children():
            self.tests_tree.selection_add(item)
    
    def deselectionner_tous_tests(self):
        """Désélectionne tous les tests dans la liste"""
        self.tests_tree.selection_remove(self.tests_tree.get_children())
    
    def configurer_encodage(self):
        """Configure l'encodage pour les caractères accentués"""
        try:
            if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except (AttributeError, TypeError):
            pass

    def creer_liste_tests_predefinies(self):
        """
        Crée une liste prédéfinie des tests à extraire (identique à l'original)
        """
        self.tests_disponibles = [
            # Format: (nom_complet, test_parent, identifiant)
            # --------- Tests SEQ-01 ---------
            ("alim 24VDC +16V", "seq01_24vdc", "24VDC_+16V"),
            ("alim 24VDC -16V", "seq01_24vdc", "24VDC_-16V"),
            ("alim 24VDC +5V", "seq01_24vdc", "24VDC_+5V"),
            ("alim 24VDC -5V", "seq01_24vdc", "24VDC_-5V"),
            ("alim 115VAC +16V", "seq01_115vac", "115VAC_+16V"),
            ("alim 115VAC -16V", "seq01_115vac", "115VAC_-16V"),
            ("R46 calculée", "seq01_resistances", "R46_calculee"),
            ("R46 à monter", "seq01_resistances", "R46_monter"),
            ("R47 calculée", "seq01_resistances", "R47_calculee"),
            ("R47 à monter", "seq01_resistances", "R47_monter"),
            ("R48 calculée", "seq01_resistances", "R48_calculee"),
            ("R48 à monter", "seq01_resistances", "R48_monter"),
            # --------- Tests SEQ-02 ---------
            ("1.9Un en 19VDC", "seq02_transfert", "Test_19VDC"),
            ("1.9Un en 115VAC", "seq02_transfert", "Test_115VAC"),
            ("Défauts précision transfert", "seq02_precision_transfert", "precision_transfert_defauts"),
        ]
        self.filter_tests_list()
    
    def selectionner_repertoire(self):
        """Permet à l'utilisateur de sélectionner un répertoire et charge les tests disponibles"""
        repertoire = filedialog.askdirectory(
            title="Sélectionnez le répertoire contenant les fichiers SEQ-01/SEQ-02"
        )
        if not repertoire:
            return
            
        self.repertoire_parent = repertoire
        short_path = os.path.basename(repertoire)
        if len(repertoire) > 50:
            short_path = "..." + repertoire[-47:]
        
        self.dir_path_label.config(text=short_path, foreground="#2c3e50")
        self.update_status(f"Analyse du répertoire {short_path}...")
        
        # Rechercher les fichiers SEQ-01/SEQ-02 et charger les tests disponibles
        self.charger_fichiers_seq(repertoire)
    
    def charger_fichiers_seq(self, repertoire):
        """Cherche les fichiers SEQ-01/SEQ-02 et charge les tests disponibles"""
        self.update_status("Recherche des fichiers SEQ-01/SEQ-02...")
        
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        
        for dossier_racine, sous_dossiers, fichiers in os.walk(repertoire):
            for fichier in fichiers:
                if fichier.startswith("SEQ-01") and fichier.lower().endswith(".html"):
                    self.fichiers_seq01.append(os.path.join(dossier_racine, fichier))
                elif fichier.startswith("SEQ-02") and fichier.lower().endswith(".html"):
                    self.fichiers_seq02.append(os.path.join(dossier_racine, fichier))
        
        total_fichiers = len(self.fichiers_seq01) + len(self.fichiers_seq02)
        
        if total_fichiers == 0:
            self.files_info_label.config(text="Aucun fichier SEQ-01 ou SEQ-02 trouvé.", foreground=self.accent_color)
            self.update_status("Aucun fichier trouvé")
            return
        
        self.files_info_label.config(
            text=f"{len(self.fichiers_seq01)} fichiers SEQ-01 et {len(self.fichiers_seq02)} fichiers SEQ-02 trouvés",
            foreground="#27ae60"
        )
        self.update_status(f"Prêt - {total_fichiers} fichiers trouvés")
    
    def generer_statistiques(self):
        """Génère les statistiques pour les tests sélectionnés"""
        # Récupérer les tests sélectionnés
        selected_items = self.tests_tree.selection()
        if not selected_items:
            messagebox.showinfo("Information", "Veuillez sélectionner au moins un test.")
            return
        
        # Récupérer les noms des tests sélectionnés
        self.tests_selectionnes = []
        for item in selected_items:
            item_text = self.tests_tree.item(item, 'text')
            # Trouver le test correspondant dans la liste complète
            for test in self.tests_disponibles:
                if test[0] == item_text:
                    self.tests_selectionnes.append(test)
                    break
        
        # Analyser les fichiers pour récupérer les données
        self.update_status("Analyse des fichiers en cours...")
        self.analyser_fichiers()
    
    def extraire_valeur_seq01(self, html_content, test_parent, identifiant):
        """Extraction des valeurs arrondies (Tension mesurée) pour les alimentations"""
        if test_parent == "seq01_24vdc":
            block_match = re.search(r"Test des alimentations à 24VDC(.*?)Test des alimentations à 115VAC", html_content, re.DOTALL)
            if not block_match:
                return None
            block = block_match.group(1)
            if identifiant == "24VDC_+16V":
                m = re.search(r"Tension \+16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "24VDC_-16V":
                m = re.search(r"Tension -16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "24VDC_+5V":
                m = re.search(r"Tension \+5V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "24VDC_-5V":
                m = re.search(r"Tension -5V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None

        elif test_parent == "seq01_115vac":
            block_match = re.search(r"Test des alimentations à 115VAC(.*?)(?:Calcul des résistances|</div>|</body>|$)", html_content, re.DOTALL)
            if not block_match:
                return None
            block = block_match.group(1)
            if identifiant == "115VAC_+16V":
                m = re.search(r"Tension \+16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "115VAC_-16V":
                m = re.search(r"Tension -16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None

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
        """Extraction des valeurs arrondies pour SEQ-02"""
        if test_parent == "seq02_transfert":
            if identifiant == "Test_19VDC":
                # Cherche la valeur arrondie de "Mesure -16V en V:"
                m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", html_content, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "Test_115VAC":
                # Cherche la valeur arrondie de "Mesure -16V en V:" pour 115VAC (adapter si besoin)
                m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", html_content, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
        elif test_parent == "seq02_precision_transfert" and identifiant == "precision_transfert_defauts":
            # Utilise la fonction déjà définie
            defauts = extraire_defauts_precision_transfert(html_content)
            # Retourne sous forme de texte concaténé, ex: "RC:1 - Voie V: -0.2358; RC:2 - Voie I: 0.1234"
            return "; ".join([f"{nom}: {precision}" for nom, precision in defauts]) if defauts else ""
        return None
    
    def extraire_numero_serie(self, html_content):
        """Identique à l'original"""
        sn_match = re.search(r'Serial Number:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL)
        if sn_match and sn_match.group(1).strip() != "NONE":
            return sn_match.group(1).strip()
        
        serie_match = re.search(r'série[^<]*</td>\s*<td[^>]*class="value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL)
        if serie_match:
            return serie_match.group(1).strip()
        
        return None

    def extraire_statut_test(self, html_content):
        """Identique à l'original"""
        status_match = re.search(r"<tr><td class='hdr_name'><b>UUT Result: </b></td><td class='hdr_value'><b><span style=\"color:[^\"]+;\">([^<]+)</span></b></td></tr>", html_content, re.DOTALL)
        if status_match:
            return status_match.group(1).strip()
        
        alt_status_match = re.search(r"UUT Result:.*?<td[^>]*class=\"hdr_value\"[^>]*>.*?<span[^>]*>(Passed|Failed)</span>", html_content, re.DOTALL | re.IGNORECASE)
        if alt_status_match:
            return alt_status_match.group(1).strip()
        
        return "Inconnu"
    
    def extraire_date_heure(self, html_content, nom_fichier):
        """Identique à l'original"""
        date_match = re.search(r'Date:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*<b>([^<]+)</b>', html_content, re.DOTALL)
        date = date_match.group(1).strip() if date_match else None
        
        time_match = re.search(r'Time:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*<b>([^<]+)</b>', html_content, re.DOTALL)
        heure = time_match.group(1).strip() if time_match else None
        
        if not date or not heure:
            date, heure = self.extraire_date_heure_du_nom(nom_fichier)
        
        return date, heure
    
    def extraire_date_heure_du_nom(self, nom_fichier):
        """Identique à l'original"""
        match = re.search(r'\[(\d+ \d+ \d+)\]\[(\d+ \d+ \d+)\]', nom_fichier)
        if match:
            heure_brute = match.group(1)  # HH MM SS
            date_brute = match.group(2)   # JJ MM AAAA
            
            heure_formatee = heure_brute.replace(" ", ":")
            date_formatee = date_brute.replace(" ", "/")
            
            return date_formatee, heure_formatee
            
        return "date_inconnue", "heure_inconnue"
    
    def obtenir_cle_tri_chronologique(self, identifiant_unique):
        """Identique à l'original"""
        try:
            parties = identifiant_unique.split(' [')
            if len(parties) < 2:
                return (identifiant_unique, 0, 0, 0, 0, 0, 0)
            
            numero_serie = parties[0]
            partie_date_heure = '[' + '['.join(parties[1:])
            
            match_date = re.search(r'\[(\d+)/(\d+)/(\d+)\]', partie_date_heure)
            if not match_date:
                return (numero_serie, 0, 0, 0, 0, 0, 0)
            
            jour, mois, annee = map(int, match_date.groups())
            
            match_heure = re.search(r'\[(\d+):(\d+):(\d+)\]', partie_date_heure)
            if not match_heure:
                return (numero_serie, annee, mois, jour, 0, 0, 0)
            
            heure, minute, seconde = map(int, match_heure.groups())
            
            return (numero_serie, annee, mois, jour, heure, minute, seconde)
        except:
            return (identifiant_unique, 0, 0, 0, 0, 0, 0)
    
    def analyser_fichiers(self):
        """Analyse tous les fichiers SEQ-01/SEQ-02 et collecte les données pour les tests sélectionnés"""
        donnees = {}  # Dictionnaire pour stocker les données
        
        # Analyse des fichiers SEQ-01
        for fichier in self.fichiers_seq01:
            try:
                with open(fichier, "r", encoding="iso-8859-1", errors="replace") as f:
                    html_content = f.read()
                
                # Extraire le numéro de série
                numero_serie = self.extraire_numero_serie(html_content) or os.path.basename(os.path.dirname(fichier))
                
                # Extraire la date et l'heure
                nom_fichier = os.path.basename(fichier)
                date, heure = self.extraire_date_heure(html_content, nom_fichier)
                
                # Extraire le statut du test
                statut = self.extraire_statut_test(html_content)
                
                # Créer un identifiant unique
                identifiant_unique = f"{numero_serie} [{date}][{heure}]"
                
                # Initialiser le dictionnaire pour ce numéro de série
                if identifiant_unique not in donnees:
                    donnees[identifiant_unique] = {
                        "Numéro de série": numero_serie,
                        "Date": date,
                        "Heure": heure,
                        "Type": "SEQ-01",
                        "Statut": statut
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
                
                # Extraire la date et l'heure
                nom_fichier = os.path.basename(fichier)
                date, heure = self.extraire_date_heure(html_content, nom_fichier)
                
                # Extraire le statut du test
                statut = self.extraire_statut_test(html_content)
                
                # Créer un identifiant unique
                identifiant_unique = f"{numero_serie} [{date}][{heure}]"
                
                # Initialiser le dictionnaire pour ce numéro de série
                if identifiant_unique not in donnees:
                    donnees[identifiant_unique] = {
                        "Numéro de série": numero_serie,
                        "Date": date,
                        "Heure": heure,
                        "Type": "SEQ-02",
                        "Statut": statut
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

        # Si le tri chronologique est activé, on regroupe par numéro de série et on trie
        if self.tri_chrono_var.get():
            donnees_par_serie = {}
            identifiants_tries = []
            
            for identifiant in sorted(self.data_tests.keys(), key=self.obtenir_cle_tri_chronologique):
                numero_serie = self.data_tests[identifiant]["Numéro de série"]
                
                if numero_serie not in donnees_par_serie:
                    donnees_par_serie[numero_serie] = []
                
                donnees_par_serie[numero_serie].append(identifiant)
                identifiants_tries.append(identifiant)
            
            df = pd.DataFrame.from_dict({id_unique: self.data_tests[id_unique] for id_unique in identifiants_tries}, orient='index')
        else:
            df = pd.DataFrame.from_dict(self.data_tests, orient='index')
        
        # Supprimer les colonnes "Numéro de série", "Date", "Heure" du DataFrame
        colonnes_a_supprimer = ["Numéro de série", "Date", "Heure"]
        df = df.drop(columns=colonnes_a_supprimer, errors='ignore')

        # Sauvegarder en Excel
        chemin_excel = os.path.join(self.repertoire_parent, "statistiques_SEQ01_SEQ02.xlsx")
        try:
            df.to_excel(chemin_excel, sheet_name="Statistiques_SEQ01_02", index=True)

            # Adapter la largeur des colonnes et activer le filtre
            wb = openpyxl.load_workbook(chemin_excel)
            ws = wb.active

            # Activer le filtre automatique sur toutes les colonnes
            ws.auto_filter.ref = ws.dimensions

            # Adapter la largeur de chaque colonne
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ""
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except Exception:
                        pass
                adjusted_width = max_length + 2
                ws.column_dimensions[col_letter].width = adjusted_width

            wb.save(chemin_excel)

            self.update_status(f"Rapport généré: {chemin_excel}")
            webbrowser.open(chemin_excel)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la création du tableau: {e}")
            self.update_status("Erreur lors de la génération du rapport")
    
    def generer_rapports_detailles(self):
        """
        Génère les rapports détaillés en utilisant les fonctions de Affiche_resultats.py
        """
        if not self.repertoire_parent:
            messagebox.showinfo("Information", "Veuillez d'abord sélectionner un répertoire.")
            return
        
        try:
            # Obtenir tous les sous-répertoires directs (les numéros de série)
            sous_repertoires = [os.path.join(self.repertoire_parent, d) for d in os.listdir(self.repertoire_parent) 
                               if os.path.isdir(os.path.join(self.repertoire_parent, d))]
            
            if not sous_repertoires:
                messagebox.showinfo("Information", 
                                    "Aucun sous-répertoire (numéro de série) trouvé dans le répertoire sélectionné.")
                return
            
            # Créer une fenêtre de progression
            progress_window = ProgressWindow(total_files=len(sous_repertoires))
            
            # Traiter chaque répertoire de numéro de série
            for repertoire in sous_repertoires:
                traiter_repertoire_serie(repertoire, progress_window)
            
            # Afficher un message de fin
            progress_window.show_completion()
            self.update_status("Rapports détaillés générés avec succès")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération des rapports: {e}")
            self.update_status("Erreur lors de la génération des rapports")
            import traceback
            traceback.print_exc()
    
    def lancer(self):
        """Lance l'application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernStatsTestsWindow()
    app.lancer()