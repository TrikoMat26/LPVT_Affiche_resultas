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
from openpyxl.styles import Font
import traceback

# Dictionnaire des bornes avec les noms de colonnes exacts
bornes_alims_seq01 = {
    # SEQ-01
    "alim 24VDC +16V": (15.9, 16.3),
    "alim 24VDC -16V": (-16.3, -15.9),
    "alim 24VDC +5V": (5.2, 5.5),
    "alim 24VDC -5V": (-5.5, -5.2),
    "alim 115VAC +16V": (16, 16.3),
    "alim 115VAC -16V": (-16.36, -15),
    # SEQ-02
    "1.9Un en 19VDC": (-16.6, -15),
    "1.9Un en 115VAC": (-16.6, -15),
}

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
    Fenêtre modernisée et professionnelle pour l'analyse statistique des tests SEQ-01 et SEQ-02
    avec toutes les fonctionnalités originales
    """
    def __init__(self):
        """Initialise l'application avec une interface moderne"""
        # Configuration de l'encodage
        self.configurer_encodage()
        
        # Variables de l'application
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        self.tests_disponibles = []
        self.tests_selectionnes = []
        self.data_tests = {}
        self.repertoire_parent = ""
        
        # Interface graphique
        self.root = tk.Tk()
        self.root.title("LPVT Test Analyzer - Interface Moderne")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Style moderne
        self.setup_styles()
        
        # Création de l'interface
        self.creer_interface_moderne()
        
        # Initialiser la liste des tests prédéfinis
        self.creer_liste_tests_predefinies()
    
    def setup_styles(self):
        """Configure les styles modernes et professionnels pour l'interface"""
        style = ttk.Style()
        
        # Thème de base
        style.theme_use('clam')
        
        # Palette de couleurs professionnelle et sobre
        self.colors = {
            'bg_primary': '#fafafa',           # Fond principal très clair
            'bg_secondary': '#ffffff',         # Fond des cartes
            'bg_accent': '#f5f7fa',           # Fond des sections
            'text_primary': '#2d3748',        # Texte principal
            'text_secondary': '#718096',      # Texte secondaire
            'text_muted': '#a0aec0',          # Texte discret
            'border': '#e2e8f0',             # Bordures subtiles
            'primary': '#4a5568',             # Couleur principale
            'success': '#38a169',             # Vert pour les succès
            'warning': '#ed8936',             # Orange pour les avertissements
            'error': '#e53e3e',              # Rouge pour les erreurs
            'info': '#3182ce',               # Bleu pour l'information
            'accent': '#667eea'               # Accent moderne
        }
        
        # Configuration du style global
        style.configure('TFrame', background=self.colors['bg_primary'])
        style.configure('Card.TFrame', 
                       background=self.colors['bg_secondary'],
                       relief='flat',
                       borderwidth=1)
        
        # Typographie
        style.configure('TLabel', 
                       background=self.colors['bg_primary'], 
                       foreground=self.colors['text_primary'],
                       font=('Segoe UI', 10))
        
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 18, 'bold'), 
                       foreground=self.colors['text_primary'],
                       background=self.colors['bg_primary'])
        
        style.configure('Subtitle.TLabel', 
                       font=('Segoe UI', 12, 'bold'), 
                       foreground=self.colors['text_secondary'],
                       background=self.colors['bg_primary'])
        
        style.configure('Caption.TLabel', 
                       font=('Segoe UI', 9), 
                       foreground=self.colors['text_muted'],
                       background=self.colors['bg_primary'])
        
        style.configure('Status.TLabel', 
                       font=('Segoe UI', 9), 
                       foreground=self.colors['info'],
                       background=self.colors['bg_primary'])
        
        # Boutons modernes
        style.configure('Modern.TButton', 
                       font=('Segoe UI', 10),
                       padding=(20, 10),
                       relief='flat',
                       borderwidth=0,
                       focuscolor='none')
        
        style.configure('Primary.TButton', 
                       font=('Segoe UI', 10, 'bold'),
                       padding=(25, 12),
                       relief='flat',
                       borderwidth=0,
                       background=self.colors['primary'],
                       foreground='white',
                       focuscolor='none')
        
        style.configure('Success.TButton', 
                       font=('Segoe UI', 10, 'bold'),
                       padding=(25, 12),
                       relief='flat',
                       borderwidth=0,
                       background=self.colors['success'],
                       foreground='white',
                       focuscolor='none')
        
        style.configure('Info.TButton', 
                       font=('Segoe UI', 10),
                       padding=(15, 8),
                       relief='flat',
                       borderwidth=0,
                       background=self.colors['info'],
                       foreground='white',
                       focuscolor='none')
        
        style.configure('Secondary.TButton', 
                       font=('Segoe UI', 10),
                       padding=(15, 8),
                       relief='flat',
                       borderwidth=1,
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_primary'],
                       focuscolor='none')
        
        # Effets hover pour les boutons
        style.map('Primary.TButton',
                 background=[('active', '#3a4653')])
        style.map('Success.TButton',
                 background=[('active', '#2f7d32')])
        style.map('Info.TButton',
                 background=[('active', '#2c5aa0')])
        style.map('Secondary.TButton',
                 background=[('active', self.colors['bg_accent']),
                           ('pressed', self.colors['border'])])
        
        # TreeView moderne
        style.configure('Modern.Treeview', 
                       font=('Segoe UI', 10),
                       rowheight=30,
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_primary'],
                       fieldbackground=self.colors['bg_secondary'],
                       borderwidth=0,
                       relief='flat')
        
        style.configure('Modern.Treeview.Heading', 
                       font=('Segoe UI', 10, 'bold'),
                       background=self.colors['bg_accent'],
                       foreground=self.colors['text_primary'],
                       relief='flat',
                       borderwidth=0)
        
        style.map('Modern.Treeview',
                 background=[('selected', self.colors['accent'])],
                 foreground=[('selected', 'white')])
        
        # Entrées de texte
        style.configure('Modern.TEntry',
                       fieldbackground=self.colors['bg_secondary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'],
                       font=('Segoe UI', 10),
                       padding=(10, 8))
        
        style.map('Modern.TEntry',
                 bordercolor=[('focus', self.colors['accent'])])
        
        # LabelFrame moderne
        style.configure('Modern.TLabelframe',
                       background=self.colors['bg_primary'],
                       borderwidth=0,
                       relief='flat')
        
        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['bg_primary'],
                       foreground=self.colors['text_secondary'],
                       font=('Segoe UI', 11, 'bold'))

    def creer_interface_moderne(self):
        """Crée l'interface utilisateur moderne et professionnelle"""
        # Configuration de la fenêtre principale
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Container principal avec effet de carte
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # === HEADER SECTION ===
        header_card = ttk.Frame(main_container, style='Card.TFrame')
        header_card.pack(fill=tk.X, pady=(0, 20))
        
        header_content = ttk.Frame(header_card)
        header_content.pack(fill=tk.X, padx=25, pady=20)
        
        # Titre principal avec sous-titre
        title_section = ttk.Frame(header_content)
        title_section.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(title_section, text="LPVT Test Analyzer", style='Title.TLabel').pack(anchor='w')
        ttk.Label(title_section, text="Analyse statistique des tests SEQ-01 et SEQ-02", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(2, 0))
        
        # Boutons d'action dans le header
        header_actions = ttk.Frame(header_content)
        header_actions.pack(side=tk.RIGHT)
        
        ttk.Button(header_actions, text="💡 Aide", command=self.show_help, 
                  style='Info.TButton').pack(side=tk.RIGHT, padx=(10, 0))
        
        # === DIRECTORY SELECTION SECTION ===
        dir_card = ttk.Frame(main_container, style='Card.TFrame')
        dir_card.pack(fill=tk.X, pady=(0, 20))
        
        dir_content = ttk.Frame(dir_card)
        dir_content.pack(fill=tk.X, padx=25, pady=20)
        
        # Section label
        ttk.Label(dir_content, text="📁 Répertoire source", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 15))
        
        # Directory selection row
        dir_row = ttk.Frame(dir_content)
        dir_row.pack(fill=tk.X)
        
        ttk.Button(dir_row, text="Parcourir...", command=self.selectionner_repertoire,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 15))
        
        # Path display avec style moderne
        path_container = ttk.Frame(dir_row)
        path_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.dir_path_label = ttk.Label(path_container, text="Aucun répertoire sélectionné",
                                       style='Caption.TLabel')
        self.dir_path_label.pack(anchor='w')
        
        # Files info
        self.files_info_label = ttk.Label(dir_content, text="", style='Status.TLabel')
        self.files_info_label.pack(anchor='w', pady=(10, 0))
        
        # === TESTS SELECTION SECTION ===
        tests_card = ttk.Frame(main_container, style='Card.TFrame')
        tests_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        tests_content = ttk.Frame(tests_card)
        tests_content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Section header
        tests_header = ttk.Frame(tests_content)
        tests_header.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(tests_header, text="⚡ Sélection des tests à analyser", 
                 style='Subtitle.TLabel').pack(side=tk.LEFT)
        
        # Options à droite
        options_frame = ttk.Frame(tests_header)
        options_frame.pack(side=tk.RIGHT)
        
        self.tri_chrono_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="📅 Tri chronologique par n° de série",
                       variable=self.tri_chrono_var).pack()
        
        # Controls row
        controls_row = ttk.Frame(tests_content)
        controls_row.pack(fill=tk.X, pady=(0, 15))
        
        # Selection buttons
        selection_buttons = ttk.Frame(controls_row)
        selection_buttons.pack(side=tk.LEFT)
        
        ttk.Button(selection_buttons, text="Tout sélectionner", 
                  command=self.selectionner_tous_tests, style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(selection_buttons, text="Tout désélectionner", 
                  command=self.deselectionner_tous_tests, style='Secondary.TButton').pack(side=tk.LEFT)
        
        # Search bar à droite
        search_container = ttk.Frame(controls_row)
        search_container.pack(side=tk.RIGHT)
        
        ttk.Label(search_container, text="🔍", font=('Segoe UI', 12)).pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_container, textvariable=self.search_var, 
                               style='Modern.TEntry', width=25)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind('<KeyRelease>', self.filter_tests_list)
        
        # Tests list avec container moderne
        list_container = ttk.Frame(tests_content)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # TreeView avec style moderne
        self.tests_tree = ttk.Treeview(list_container, selectmode='extended',
                                     columns=('category',), show='tree headings',
                                     style='Modern.Treeview', height=12)
        
        # Colonnes
        self.tests_tree.heading('#0', text='Test', anchor='w')
        self.tests_tree.heading('category', text='Catégorie', anchor='w')
        self.tests_tree.column('#0', width=300, minwidth=200)
        self.tests_tree.column('category', width=200, minwidth=150)
        
        # Scrollbar moderne
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.tests_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tests_tree.configure(yscrollcommand=scrollbar.set)
        self.tests_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # === ACTIONS SECTION ===
        actions_card = ttk.Frame(main_container, style='Card.TFrame')
        actions_card.pack(fill=tk.X)
        
        actions_content = ttk.Frame(actions_card)
        actions_content.pack(fill=tk.X, padx=25, pady=20)
        
        # Actions row
        actions_row = ttk.Frame(actions_content)
        actions_row.pack(fill=tk.X)
        
        # Boutons d'action principaux
        ttk.Button(actions_row, text="📊 Générer le rapport statistique",
                  command=self.generer_statistiques, style='Success.TButton').pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(actions_row, text="📄 Générer les rapports détaillés",
                  command=self.generer_rapports_detailles, style='Primary.TButton').pack(side=tk.LEFT)
        
        # Status bar moderne
        status_container = ttk.Frame(actions_content)
        status_container.pack(fill=tk.X, pady=(15, 0))
        
        status_separator = ttk.Frame(status_container, height=1)
        status_separator.pack(fill=tk.X, pady=(0, 10))
        status_separator.configure(style='Card.TFrame')
        
        self.status_label = ttk.Label(status_container, text="✅ Prêt à traiter vos fichiers", 
                                     style='Status.TLabel')
        self.status_label.pack(side=tk.LEFT)
        
        # Indicateur de version/copyright à droite
        ttk.Label(status_container, text="LPVT Analyzer v2.0", 
                 style='Caption.TLabel').pack(side=tk.RIGHT)

    def show_help(self):
        """Affiche une aide contextuelle moderne"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Aide - LPVT Test Analyzer")
        help_window.geometry("650x550")
        help_window.configure(bg=self.colors['bg_primary'])
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Centrer la fenêtre
        help_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Container principal
        main_frame = ttk.Frame(help_window, style='Card.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        
        # Titre
        ttk.Label(content_frame, text="💡 Guide d'utilisation", 
                 style='Title.TLabel').pack(anchor='w', pady=(0, 20))
        
        # Contenu de l'aide
        help_content = """🎯 OBJECTIF
Analyser les rapports de test HTML SEQ-01 et SEQ-02 pour générer des statistiques et des rapports détaillés.

📋 ÉTAPES D'UTILISATION

1️⃣ SÉLECTION DU RÉPERTOIRE
   • Cliquez sur "Parcourir..." dans la section "Répertoire source"
   • Sélectionnez le dossier parent contenant les sous-dossiers des numéros de série
   • Structure attendue: Dossier_Parent/SN12345/fichiers_HTML

2️⃣ SÉLECTION DES TESTS
   • Choisissez les paramètres à analyser dans la liste organisée par catégories
   • Utilisez la barre de recherche 🔍 pour filtrer rapidement
   • Boutons "Tout sélectionner" / "Tout désélectionner" disponibles

3️⃣ GÉNÉRATION DES RAPPORTS
   📊 Rapport Statistique: Crée un fichier Excel avec toutes les mesures
   📄 Rapports Détaillés: Génère un rapport texte par numéro de série

⚙️ OPTIONS AVANCÉES
   📅 Tri chronologique: Organise les résultats par n° de série et date/heure
   🎯 Filtre par recherche: Trouve rapidement les tests souhaités

💡 CONSEILS
   • Les valeurs hors spécifications apparaissent en rouge dans Excel
   • Les rapports s'ouvrent automatiquement après génération
   • Une barre de progression indique l'avancement du traitement"""
        
        # Zone de texte avec scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Segoe UI', 10),
                             bg=self.colors['bg_secondary'], fg=self.colors['text_primary'],
                             relief='flat', borderwidth=0, padx=15, pady=15)
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.insert('1.0', help_content.strip())
        text_widget.configure(state='disabled')
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bouton fermer
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(button_frame, text="Fermer", command=help_window.destroy,
                  style='Primary.TButton').pack(side=tk.RIGHT)

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
        
        # Grouper les tests par catégorie pour un affichage organisé
        categories = {
            'seq01_24vdc': '🔋 Alimentations 24VDC',
            'seq01_115vac': '⚡ Alimentations 115VAC', 
            'seq01_resistances': '🔧 Résistances',
            'seq02_transfert': '📊 Tests de transfert',
            'seq02_precision_transfert': '🎯 Précision transfert'
        }
        
        # Regrouper les tests filtrés par catégorie
        filtered_by_category = {}
        for nom_complet, test_parent, identifiant in self.tests_disponibles:
            if search_term in nom_complet.lower():
                if test_parent not in filtered_by_category:
                    filtered_by_category[test_parent] = []
                filtered_by_category[test_parent].append((nom_complet, identifiant))
        
        # Ajouter les tests organisés par catégorie
        for category, tests in filtered_by_category.items():
            category_name = categories.get(category, category)
            # Ajouter le nœud de catégorie
            category_node = self.tests_tree.insert('', 'end', text=category_name, 
                                                  values=(f'category_{category}',), open=True)
            
            # Ajouter les tests sous cette catégorie
            for nom_complet, identifiant in tests:
                self.tests_tree.insert(category_node, 'end', text=f"  {nom_complet}", 
                                     values=(identifiant,))
    
    def selectionner_tous_tests(self):
        """Sélectionne tous les tests dans la liste (pas les catégories)"""
        for item in self.tests_tree.get_children():
            # Sélectionner les enfants (tests) mais pas les parents (catégories)
            for child in self.tests_tree.get_children(item):
                self.tests_tree.selection_add(child)
    
    def deselectionner_tous_tests(self):
        """Désélectionne tous les tests dans la liste"""
        self.tests_tree.selection_remove(self.tests_tree.get_children())
        # Aussi désélectionner tous les enfants
        for item in self.tests_tree.get_children():
            for child in self.tests_tree.get_children(item):
                self.tests_tree.selection_remove(child)

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
        
        self.dir_path_label.config(text=short_path, foreground=self.colors['text_primary'])
        self.update_status(f"🔍 Analyse du répertoire {short_path}...")
        
        # Rechercher les fichiers SEQ-01/SEQ-02 et charger les tests disponibles
        self.charger_fichiers_seq(repertoire)

    def charger_fichiers_seq(self, repertoire):
        """Cherche les fichiers SEQ-01/SEQ-02 et charge les tests disponibles"""
        self.update_status("🔍 Recherche des fichiers SEQ-01/SEQ-02...")
        
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
            self.files_info_label.config(text="❌ Aucun fichier SEQ-01 ou SEQ-02 trouvé.", 
                                        foreground=self.colors['error'])
            self.update_status("❌ Aucun fichier trouvé")
            return
        
        self.files_info_label.config(
            text=f"✅ {len(self.fichiers_seq01)} fichiers SEQ-01 et {len(self.fichiers_seq02)} fichiers SEQ-02 trouvés",
            foreground=self.colors['success']
        )
        self.update_status(f"✅ Prêt - {total_fichiers} fichiers trouvés")

    # === CONTINUATION DES FONCTIONS ===
    # (Les fonctions suivantes restent identiques à l'original mais avec des messages de status améliorés)
    
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
            item_text = self.tests_tree.item(item, 'text').strip()
            item_values = self.tests_tree.item(item, 'values')
            
            # Ignorer les catégories (qui commencent par 'category_')
            if item_values and len(item_values) > 0 and item_values[0].startswith('category_'):
                continue
                
            # Nettoyer le texte (enlever les espaces en début si c'est un enfant)
            if item_text.startswith('  '):
                item_text = item_text.strip()
            
            # Trouver le test correspondant dans la liste complète
            for test in self.tests_disponibles:
                if test[0] == item_text:
                    self.tests_selectionnes.append(test)
                    break
        
        if not self.tests_selectionnes:
            messagebox.showinfo("Information", "Aucun test valide sélectionné.")
            return
        
        # Analyser les fichiers pour récupérer les données
        self.update_status("🔄 Analyse des fichiers en cours...")
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
                r46_calc = re.search(r"Résistance R46 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                if r46_calc:
                    return str(int(round(float(r46_calc.group(1).replace(',', '.')))))
                else:
                    return None
            elif identifiant == "R46_monter":
                r46_monter = re.search(r"Résistance R46 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r46_monter.group(1).strip() if r46_monter else None
            elif identifiant == "R47_calculee":
                r47_calc = re.search(r"Résistance R47 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                if r47_calc:
                    return str(int(round(float(r47_calc.group(1).replace(',', '.')))))
                else:
                    return None
            elif identifiant == "R47_monter":
                r47_monter = re.search(r"Résistance R47 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r47_monter.group(1).strip() if r47_monter else None
            elif identifiant == "R48_calculee":
                r48_calc = re.search(r"Résistance R48 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                if r48_calc:
                    return str(int(round(float(r48_calc.group(1).replace(',', '.')))))
                else:
                    return None
            elif identifiant == "R48_monter":
                r48_monter = re.search(r"Résistance R48 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r48_monter.group(1).strip() if r48_monter else None

        return None
        
    def extraire_valeur_seq02(self, html_content, test_parent, identifiant):
        """Extraction des valeurs arrondies pour SEQ-02"""
        if test_parent == "seq02_transfert":
            # Pour 1.9Un en 19VDC
            if identifiant == "Test_19VDC":
                # Cherche le bloc "Test 1.9Un sur 2 voies en 19VDC"
                block_match = re.search(
                    r"Test 1\.9Un sur 2 voies en 19VDC.*?(Mesure -16V en V:.*?Fourchette max en V:.*?Fourchette min en V:.*?)</table>",
                    html_content, re.DOTALL)
                if block_match:
                    block = block_match.group(1)
                    m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                    return m.group(1).replace(',', '.').strip() if m else None
                return None
            # Pour 1.9Un en 115VAC
            elif identifiant == "Test_115VAC":
                # Cherche le bloc "Test 1.9Un sur 2 voies en 115VAC"
                block_match = re.search(
                    r"Test 1\.9Un sur 2 voies en 115VAC.*?(Mesure -16V en V:.*?Fourchette max en V:.*?Fourchette min en V:.*?)</table>",
                    html_content, re.DOTALL)
                if block_match:
                    block = block_match.group(1)
                    m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                    return m.group(1).replace(',', '.').strip() if m else None
                return None
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

        # Déplacer la colonne "Défauts précision transfert" juste après "Statut"
        colonnes = list(df.columns)
        if "Défauts précision transfert" in colonnes and "Statut" in colonnes:
            colonnes.remove("Défauts précision transfert")
            statut_idx = colonnes.index("Statut")
            colonnes.insert(statut_idx + 1, "Défauts précision transfert")
            df = df[colonnes]

        # Supprimer les colonnes "Numéro de série", "Date", "Heure" du DataFrame
        colonnes_a_supprimer = ["Numéro de série", "Date", "Heure"]
        df = df.drop(columns=colonnes_a_supprimer, errors='ignore')

        # Sauvegarder en Excel
        chemin_excel = os.path.join(self.repertoire_parent, "statistiques_SEQ01_SEQ02.xlsx")
        try:
            self.update_status("📊 Génération du fichier Excel...")
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

            # Mapping colonne (titre -> index)
            col_map = {cell.value: idx for idx, cell in enumerate(ws[1])}

            for nom_col, (borne_min, borne_max) in bornes_alims_seq01.items():
                col_idx = col_map.get(nom_col)
                if col_idx is not None:
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        cell = row[col_idx]
                        try:
                            val = float(str(cell.value).replace(",", "."))
                            if not (borne_min <= val <= borne_max):
                                cell.font = Font(color="FF0000")
                        except Exception:
                            pass

            # Coloration en rouge si "à monter" > "calculée"
            col_map = {cell.value: idx for idx, cell in enumerate(ws[1])}
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for prefix in ["R46", "R47", "R48"]:
                    col_calc = col_map.get(f"{prefix} calculée")
                    col_mont = col_map.get(f"{prefix} à monter")
                    if col_calc is not None and col_mont is not None:
                        cell_calc = row[col_calc]
                        cell_mont = row[col_mont]
                        try:
                            # Conversion robuste en int (même si texte)
                            val_calc = int(float(str(cell_calc.value).replace(",", ".").strip()))
                            val_mont = int(float(str(cell_mont.value).replace(",", ".").strip()))
                            if val_mont > val_calc:
                                cell_mont.font = Font(color="FF0000")
                        except Exception:
                            pass

            wb.save(chemin_excel)

            self.update_status(f"✅ Rapport généré avec succès: {os.path.basename(chemin_excel)}")
            webbrowser.open(chemin_excel)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la création du tableau: {e}")
            self.update_status("❌ Erreur lors de la génération du rapport")

    def generer_rapports_detailles(self):
        """
        Génère les rapports détaillés en utilisant les fonctions de Affiche_resultats.py
        """
        if not self.repertoire_parent:
            messagebox.showinfo("Information", "Veuillez d'abord sélectionner un répertoire.")
            return
        
        try:
            self.update_status("🔍 Recherche des sous-répertoires...")
            # Obtenir tous les sous-répertoires directs (les numéros de série)
            sous_repertoires = [os.path.join(self.repertoire_parent, d) for d in os.listdir(self.repertoire_parent) 
                               if os.path.isdir(os.path.join(self.repertoire_parent, d))]
            
            if not sous_repertoires:
                messagebox.showinfo("Information", 
                                    "Aucun sous-répertoire (numéro de série) trouvé dans le répertoire sélectionné.")
                return
            
            self.update_status(f"📄 Génération des rapports pour {len(sous_repertoires)} séries...")
            # Créer une fenêtre de progression
            progress_window = ProgressWindow(total_files=len(sous_repertoires))
            
            # Traiter chaque répertoire de numéro de série
            for repertoire in sous_repertoires:
                traiter_repertoire_serie(repertoire, progress_window)
            
            # Afficher un message de fin
            progress_window.show_completion()
            self.update_status("✅ Rapports détaillés générés avec succès")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération des rapports: {e}")
            self.update_status("❌ Erreur lors de la génération des rapports")
            import traceback
            traceback.print_exc()

    def lancer(self):
        """Lance l'application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernStatsTestsWindow()
    app.lancer()
