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
from Affiche_resultats import traiter_repertoire_serie, ProgressWindow # Assurez-vous que Affiche_resultats.py est accessible
import webbrowser
import openpyxl # Importation de openpyxl
from openpyxl.styles import Font # Pour le style des cellules (couleur)
# from openpyxl.utils.dataframe import dataframe_to_rows # dataframe_to_rows n'est plus utilisé dans la version originale que vous avez fournie
import traceback

# Dictionnaire des bornes avec les noms de colonnes exacts (repris de votre version)
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

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def extraire_defauts_precision_transfert(html_content): # Fonction originale
    resultats = []
    soup = BeautifulSoup(html_content, "html.parser")
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        i = 0
        while i < len(rows):
            tds = rows[i].find_all("td")
            if len(tds) == 2 and "Status:" in tds[0].get_text(strip=True):
                status = tds[1].get_text(strip=True)
                if status.lower() == "failed":
                    voie = None
                    precision = None
                    for j in range(i+1, min(i+4, len(rows))):
                        tds2 = rows[j].find_all("td")
                        if len(tds2) == 2 and "ROUE CODEUSE" in tds2[0].get_text():
                            voie = tds2[0].get_text(strip=True)
                        if len(tds2) == 2 and "Précision" in tds2[0].get_text():
                            valeurs = tds2[1].get_text(strip=True)
                            precision = valeurs.split('/')[-1].strip()
                    if voie and precision:
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
        self.root.geometry("1200x1000")
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
            'bg_primary': '#fafafa',
            'bg_secondary': '#ffffff',
            'bg_accent': '#f5f7fa',
            'text_primary': '#2d3748',
            'text_secondary': '#718096',
            'text_muted': '#a0aec0',
            'border': '#e2e8f0',
            'primary': '#4a5568',
            'success': '#38a169',
            'warning': '#ed8936',
            'error': '#e53e3e',
            'info': '#3182ce',
            'accent': '#667eea'
        }

        style.configure('TFrame', background=self.colors['bg_primary'])
        style.configure('Card.TFrame',
                       background=self.colors['bg_secondary'],
                       relief='flat',
                       borderwidth=1)

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

        style.map('Primary.TButton',
                 background=[('active', '#3a4653')])
        style.map('Success.TButton',
                 background=[('active', '#2f7d32')])
        style.map('Info.TButton',
                 background=[('active', '#2c5aa0')])
        style.map('Secondary.TButton',
                 background=[('active', self.colors['bg_accent']),
                           ('pressed', self.colors['border'])])

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

        style.configure('Modern.TEntry',
                       fieldbackground=self.colors['bg_secondary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'],
                       font=('Segoe UI', 10),
                       padding=(10, 8))

        style.map('Modern.TEntry',
                 bordercolor=[('focus', self.colors['accent'])])

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
        self.root.configure(bg=self.colors['bg_primary'])

        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        header_card = ttk.Frame(main_container, style='Card.TFrame')
        header_card.pack(fill=tk.X, pady=(0, 20))
        header_content = ttk.Frame(header_card)
        header_content.pack(fill=tk.X, padx=25, pady=20)
        title_section = ttk.Frame(header_content)
        title_section.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(title_section, text="LPVT Test Analyzer", style='Title.TLabel').pack(anchor='w')
        ttk.Label(title_section, text="Analyse statistique des tests SEQ-01 et SEQ-02",
                 style='Subtitle.TLabel').pack(anchor='w', pady=(2, 0))
        header_actions = ttk.Frame(header_content)
        header_actions.pack(side=tk.RIGHT)
        ttk.Button(header_actions, text="💡 Aide", command=self.show_help,
                  style='Info.TButton').pack(side=tk.RIGHT, padx=(10, 0))

        dir_card = ttk.Frame(main_container, style='Card.TFrame')
        dir_card.pack(fill=tk.X, pady=(0, 20))
        dir_content = ttk.Frame(dir_card)
        dir_content.pack(fill=tk.X, padx=25, pady=20)
        ttk.Label(dir_content, text="📁 Répertoire source",
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 15))
        dir_row = ttk.Frame(dir_content)
        dir_row.pack(fill=tk.X)
        ttk.Button(dir_row, text="Parcourir...", command=self.selectionner_repertoire,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 15))
        path_container = ttk.Frame(dir_row)
        path_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.dir_path_label = ttk.Label(path_container, text="Aucun répertoire sélectionné",
                                       style='Caption.TLabel')
        self.dir_path_label.pack(anchor='w')
        self.files_info_label = ttk.Label(dir_content, text="", style='Status.TLabel')
        self.files_info_label.pack(anchor='w', pady=(10, 0))

        tests_card = ttk.Frame(main_container, style='Card.TFrame')
        tests_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        tests_content = ttk.Frame(tests_card)
        tests_content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        tests_header = ttk.Frame(tests_content)
        tests_header.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(tests_header, text="⚡ Sélection des tests à analyser",
                 style='Subtitle.TLabel').pack(side=tk.LEFT)
        options_frame = ttk.Frame(tests_header)
        options_frame.pack(side=tk.RIGHT)
        self.tri_chrono_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="📅 Tri chronologique par n° de série",
                       variable=self.tri_chrono_var).pack()
        controls_row = ttk.Frame(tests_content)
        controls_row.pack(fill=tk.X, pady=(0, 15))
        selection_buttons = ttk.Frame(controls_row)
        selection_buttons.pack(side=tk.LEFT)
        ttk.Button(selection_buttons, text="Tout sélectionner",
                  command=self.selectionner_tous_tests, style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(selection_buttons, text="Tout désélectionner",
                  command=self.deselectionner_tous_tests, style='Secondary.TButton').pack(side=tk.LEFT)
        search_container = ttk.Frame(controls_row)
        search_container.pack(side=tk.RIGHT)
        ttk.Label(search_container, text="🔍", font=('Segoe UI', 12)).pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_container, textvariable=self.search_var,
                               style='Modern.TEntry', width=25)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind('<KeyRelease>', self.filter_tests_list)
        list_container = ttk.Frame(tests_content)
        list_container.pack(fill=tk.BOTH, expand=True)
        self.tests_tree = ttk.Treeview(list_container, selectmode='extended',
                                     columns=('category',), show='tree headings',
                                     style='Modern.Treeview', height=12)
        self.tests_tree.heading('#0', text='Test', anchor='w')
        self.tests_tree.heading('category', text='Catégorie', anchor='w')
        self.tests_tree.column('#0', width=300, minwidth=200, stretch=tk.YES)
        self.tests_tree.column('category', width=200, minwidth=150, stretch=tk.YES)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.tests_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tests_tree.configure(yscrollcommand=scrollbar.set)
        self.tests_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        actions_card = ttk.Frame(main_container, style='Card.TFrame')
        actions_card.pack(fill=tk.X)
        actions_content = ttk.Frame(actions_card)
        actions_content.pack(fill=tk.X, padx=25, pady=20)
        actions_row = ttk.Frame(actions_content)
        actions_row.pack(fill=tk.X)
        ttk.Button(actions_row, text="📊 Générer le rapport statistique",
                  command=self.generer_statistiques, style='Success.TButton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(actions_row, text="📄 Générer les rapports détaillés",
                  command=self.generer_rapports_detailles, style='Primary.TButton').pack(side=tk.LEFT)
        status_container = ttk.Frame(actions_content)
        status_container.pack(fill=tk.X, pady=(15, 0))
        status_separator = ttk.Frame(status_container, height=1)
        status_separator.pack(fill=tk.X, pady=(0, 10))
        status_separator.configure(style='Card.TFrame')
        self.status_label = ttk.Label(status_container, text="✅ Prêt à traiter vos fichiers",
                                     style='Status.TLabel')
        self.status_label.pack(side=tk.LEFT)
        ttk.Label(status_container, text="LPVT Analyzer v2.1 (XLSM Support)",
                 style='Caption.TLabel').pack(side=tk.RIGHT)

    def show_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Aide - LPVT Test Analyzer")
        help_window.geometry("650x550")
        help_window.configure(bg=self.colors['bg_primary'])
        help_window.transient(self.root)
        help_window.grab_set()
        help_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        main_frame = ttk.Frame(help_window, style='Card.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        ttk.Label(content_frame, text="💡 Guide d'utilisation",
                 style='Title.TLabel').pack(anchor='w', pady=(0, 20))
        help_content = """🎯 OBJECTIF
Analyser les rapports de test HTML SEQ-01 et SEQ-02 pour générer des statistiques et des rapports détaillés.

📋 ÉTAPES D'UTILISATION

1️⃣ SÉLECTION DU RÉPERTOIRE
   • Cliquez sur "Parcourir..." et sélectionnez le dossier parent contenant les sous-dossiers des numéros de série.
   • Structure attendue: Dossier_Parent/SN12345/fichiers_HTML

2️⃣ SÉLECTION DES TESTS
   • Choisissez les paramètres à analyser. Utilisez la recherche 🔍 pour filtrer.

3️⃣ GÉNÉRATION DES RAPPORTS
   📊 Rapport Statistique: Crée un fichier Excel (.xlsm). Double-cliquez sur un N° de Série (colonne A) pour ouvrir son dossier.
   📄 Rapports Détaillés: Génère un rapport texte par N° de série.

⚙️ OPTIONS
   📅 Tri chronologique: Organise les résultats par N° de série et date/heure.

💡 CONSEILS
   • Valeurs hors spécifications en rouge dans Excel.
   • Rapports ouverts automatiquement.
   • Barre de progression pour les tâches longues."""
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
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(button_frame, text="Fermer", command=help_window.destroy,
                  style='Primary.TButton').pack(side=tk.RIGHT)

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def filter_tests_list(self, event=None):
        search_term = self.search_var.get().lower()
        for item in self.tests_tree.get_children():
            self.tests_tree.delete(item)
        categories = {
            'seq01_24vdc': '🔋 Alimentations 24VDC (SEQ-01)',
            'seq01_115vac': '⚡ Alimentations 115VAC (SEQ-01)',
            'seq01_resistances': '🔧 Résistances (SEQ-01)',
            'seq02_transfert': '📊 Tests de transfert (SEQ-02)',
            'seq02_precision_transfert': '🎯 Précision transfert (SEQ-02)'
        }
        filtered_by_category = {}
        for nom_complet, test_parent, identifiant in self.tests_disponibles:
            if search_term in nom_complet.lower() or search_term in test_parent.lower():
                if test_parent not in filtered_by_category:
                    filtered_by_category[test_parent] = []
                filtered_by_category[test_parent].append((nom_complet, identifiant))
        for category_id, tests_in_cat in filtered_by_category.items():
            category_name = categories.get(category_id, category_id.replace("_", " ").title())
            category_node = self.tests_tree.insert('', 'end', text=category_name,
                                                  values=(f'category_{category_id}',), open=True)
            for nom_complet, identifiant in tests_in_cat:
                self.tests_tree.insert(category_node, 'end', text=f"  {nom_complet}",
                                     values=(identifiant,))

    def selectionner_tous_tests(self):
        for category_node_id in self.tests_tree.get_children():
            for test_node_id in self.tests_tree.get_children(category_node_id):
                self.tests_tree.selection_add(test_node_id)

    def deselectionner_tous_tests(self):
        self.tests_tree.selection_remove(self.tests_tree.get_children())
        for category_node_id in self.tests_tree.get_children():
            for test_node_id in self.tests_tree.get_children(category_node_id):
                self.tests_tree.selection_remove(test_node_id)

    def configurer_encodage(self):
        try:
            if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except (AttributeError, TypeError):
            pass

    def creer_liste_tests_predefinies(self):
        self.tests_disponibles = [
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
            ("1.9Un en 19VDC", "seq02_transfert", "Test_19VDC"),
            ("1.9Un en 115VAC", "seq02_transfert", "Test_115VAC"),
            ("Défauts précision transfert", "seq02_precision_transfert", "precision_transfert_defauts"),
        ]
        self.filter_tests_list()

    def selectionner_repertoire(self):
        repertoire = filedialog.askdirectory(title="Sélectionnez le répertoire contenant les fichiers SEQ-01/SEQ-02")
        if not repertoire: return
        self.repertoire_parent = repertoire
        short_path = "..." + repertoire[-57:] if len(repertoire) > 60 else repertoire
        self.dir_path_label.config(text=short_path, foreground=self.colors['text_primary'])
        self.update_status(f"🔍 Analyse du répertoire {os.path.basename(repertoire)}...")
        self.charger_fichiers_seq(repertoire)

    def charger_fichiers_seq(self, repertoire):
        self.update_status("🔍 Recherche des fichiers SEQ-01/SEQ-02...")
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        for dossier_racine, _, fichiers in os.walk(repertoire):
            for fichier in fichiers:
                if fichier.startswith("SEQ-01") and fichier.lower().endswith(".html"):
                    self.fichiers_seq01.append(os.path.join(dossier_racine, fichier))
                elif fichier.startswith("SEQ-02") and fichier.lower().endswith(".html"):
                    self.fichiers_seq02.append(os.path.join(dossier_racine, fichier))
        total_fichiers = len(self.fichiers_seq01) + len(self.fichiers_seq02)
        if total_fichiers == 0:
            self.files_info_label.config(text="❌ Aucun fichier SEQ-01 ou SEQ-02 trouvé.", foreground=self.colors['error'])
            self.update_status("❌ Aucun fichier trouvé")
            return
        self.files_info_label.config(
            text=f"✅ {len(self.fichiers_seq01)} fichiers SEQ-01 et {len(self.fichiers_seq02)} fichiers SEQ-02 trouvés",
            foreground=self.colors['success'])
        self.update_status(f"✅ Prêt - {total_fichiers} fichiers trouvés")

    def generer_statistiques(self):
        """Génère les statistiques pour les tests sélectionnés.
           Si aucun test n'est sélectionné, tous les tests sont sélectionnés automatiquement."""
        
        selected_items_ids = self.tests_tree.selection() # Récupère les IDs des nœuds sélectionnés

        # --- MODIFICATION ICI ---
        if not selected_items_ids:
            # Aucun test n'est sélectionné, donc on les sélectionne tous
            self.update_status("ℹ️ Aucun test sélectionné, sélection de tous les tests...")
            self.selectionner_tous_tests() # Appelle la méthode qui sélectionne tous les tests dans le TreeView
            self.root.update_idletasks() # Mettre à jour l'interface pour refléter la sélection
            
            # Récupérer à nouveau les items sélectionnés après les avoir tous sélectionnés
            selected_items_ids = self.tests_tree.selection()
            
            # Si même après avoir tout sélectionné, la liste est vide (cas très improbable si tests_disponibles est peuplé)
            if not selected_items_ids:
                 messagebox.showerror("Erreur", "Impossible de sélectionner les tests automatiquement. Aucun test disponible.")
                 self.update_status("❌ Erreur: Aucun test disponible pour sélection automatique.")
                 return
        # --- FIN MODIFICATION ---
        
        self.tests_selectionnes = []
        for item_id in selected_items_ids:
            item_values = self.tests_tree.item(item_id, 'values')
            if not item_values or not isinstance(item_values, (list, tuple)) or len(item_values) == 0:
                continue
            identifiant_unique_extraction = item_values[0]
            if identifiant_unique_extraction.startswith('category_'):
                continue
            found_test = next((test for test in self.tests_disponibles if test[2] == identifiant_unique_extraction), None)
            if found_test:
                self.tests_selectionnes.append(found_test)
        if not self.tests_selectionnes:
            messagebox.showinfo("Information", "Aucun test valide sélectionné.")
            return
        self.update_status("🔄 Analyse des fichiers en cours...")
        self.analyser_fichiers()

    def extraire_valeur_seq01(self, html_content, test_parent, identifiant):
        """Extraction des valeurs arrondies (Tension mesurée) pour les alimentations"""
        if test_parent == "seq01_24vdc": # Corresponds to test_parent_category_id from creer_liste_tests_predefinies
            block_match = re.search(r"Test des alimentations à 24VDC(.*?)Test des alimentations à 115VAC", html_content, re.DOTALL)
            if not block_match: return None
            block = block_match.group(1)
            # Using identifiant from creer_liste_tests_predefinies
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
            if not block_match: return None
            block = block_match.group(1)
            if identifiant == "115VAC_+16V":
                m = re.search(r"Tension \+16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "115VAC_-16V":
                m = re.search(r"Tension -16V mesurée:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block, re.DOTALL)
                return m.group(1).replace(',', '.').strip() if m else None

        elif test_parent == "seq01_resistances":
            if identifiant == "R46_calculee":
                r_match = re.search(r"Résistance R46 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                return str(int(round(float(r_match.group(1).replace(',', '.'))))) if r_match else None
            elif identifiant == "R46_monter":
                r_match = re.search(r"Résistance R46 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r_match.group(1).strip() if r_match else None
            elif identifiant == "R47_calculee":
                r_match = re.search(r"Résistance R47 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                return str(int(round(float(r_match.group(1).replace(',', '.'))))) if r_match else None
            elif identifiant == "R47_monter":
                r_match = re.search(r"Résistance R47 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r_match.group(1).strip() if r_match else None
            elif identifiant == "R48_calculee":
                r_match = re.search(r"Résistance R48 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.,]+)\s*</td>", html_content, re.DOTALL)
                return str(int(round(float(r_match.group(1).replace(',', '.'))))) if r_match else None
            elif identifiant == "R48_monter":
                r_match = re.search(r"Résistance R48 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html_content, re.DOTALL)
                return r_match.group(1).strip() if r_match else None
        return None

    def extraire_valeur_seq02(self, html_content, test_parent, identifiant):
        if test_parent == "seq02_transfert":
            if identifiant == "Test_19VDC":
                block_match = re.search(
                    r"Test 1\.9Un sur 2 voies en 19VDC.*?(Mesure -16V en V:.*?Fourchette max en V:.*?Fourchette min en V:.*?)</table>",
                    html_content, re.DOTALL)
                if block_match:
                    m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block_match.group(1), re.DOTALL)
                    return m.group(1).replace(',', '.').strip() if m else None
            elif identifiant == "Test_115VAC":
                block_match = re.search(
                    r"Test 1\.9Un sur 2 voies en 115VAC.*?(Mesure -16V en V:.*?Fourchette max en V:.*?Fourchette min en V:.*?)</table>",
                    html_content, re.DOTALL)
                if block_match:
                    m = re.search(r"Mesure -16V en V:\s*</td>\s*<td[^>]*>\s*([-\d\.,]+)", block_match.group(1), re.DOTALL)
                    return m.group(1).replace(',', '.').strip() if m else None
        elif test_parent == "seq02_precision_transfert" and identifiant == "precision_transfert_defauts":
            defauts = extraire_defauts_precision_transfert(html_content)
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
        status_match = re.search(r"UUT Result:.*?<td class='hdr_value'><b><span style=\"color:[^\"]+;\">([^<]+)</span></b></td>", html_content, re.DOTALL | re.IGNORECASE)
        if status_match: return status_match.group(1).strip()
        alt_status_match = re.search(r"UUT Result:.*?<td[^>]*class=\"hdr_value\"[^>]*>.*?<span[^>]*>(Passed|Failed|Terminated)</span>", html_content, re.DOTALL | re.IGNORECASE)
        if alt_status_match: return alt_status_match.group(1).strip()
        return "Inconnu"

    def extraire_date_heure(self, html_content, nom_fichier):
        date_match = re.search(r"Date:</td>\s*<td[^>]*class=\"hdr_value\"[^>]*><b>([^<]+)</b></td>", html_content, re.DOTALL)
        date_str = date_match.group(1).strip() if date_match else None
        time_match = re.search(r"Time:</td>\s*<td[^>]*class=\"hdr_value\"[^>]*><b>([^<]+)</b></td>", html_content, re.DOTALL)
        heure_str = time_match.group(1).strip() if time_match else None

        if date_str and heure_str:
            try: # Tentative de standardisation de la date
                parts = re.split(r'[\s/]+', date_str) # Gère "jour mois année" ou "jour/mois/année"
                day, month_name, year = parts[-3], parts[-2], parts[-1]
                months_fr_to_num = {"janvier": "01", "février": "02", "mars": "03", "avril": "04", "mai": "05", "juin": "06",
                                     "juillet": "07", "août": "08", "septembre": "09", "octobre": "10", "novembre": "11", "décembre": "12"}
                month_num_str = months_fr_to_num.get(month_name.lower(), month_name) # Garde le numéro si déjà numérique
                date_std = f"{int(day):02d}/{int(month_num_str):02d}/{year}"
                return date_std, heure_str
            except: # Si le parsing échoue, on utilise le nom du fichier
                pass
        return self.extraire_date_heure_du_nom(nom_fichier)


    def extraire_date_heure_du_nom(self, nom_fichier):
        match = re.search(r'\[(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\]\[(\d{1,2})\s+(\d{1,2})\s+(\d{4})\]', nom_fichier)
        if match:
            h, m, s = match.group(1), match.group(2), match.group(3)
            day, month, year = match.group(4), match.group(5), match.group(6)
            return f"{int(day):02d}/{int(month):02d}/{year}", f"{int(h):02d}:{int(m):02d}:{int(s):02d}"
        return "date_inconnue", "heure_inconnue"

    def obtenir_cle_tri_chronologique(self, identifiant_unique_avec_date_heure):
        """Crée une clé de tri à partir de l'identifiant unique pour un tri chronologique.
           Format attendu: "0178 [JJ/MM/AAAA][HH:MM:SS]" ou "NOM_DOSSIER [JJ/MM/AAAA][HH:MM:SS]"
        """
        try:
            match = re.match(r'^(.*?) \[((\d{2})/(\d{2})/(\d{4}))\]\[((\d{2}):(\d{2}):(\d{2}))\]$', identifiant_unique_avec_date_heure)
            
            if not match:
                sn_part = identifiant_unique_avec_date_heure.split(' [')[0] if ' [' in identifiant_unique_avec_date_heure else identifiant_unique_avec_date_heure
                return (sn_part, 0, 0, 0, 0, 0, 0) 

            numero_serie_str = match.group(1).strip() 
            # date_full = match.group(2) # JJ/MM/AAAA
            jour = int(match.group(4)) # Attention à l'ordre des groupes de capture pour JJ et MM
            mois = int(match.group(3)) # Si le format est JJ/MM/AAAA, groupe 3 est JJ, groupe 4 est MM
            annee = int(match.group(5))
            # heure_full = match.group(6) # HH:MM:SS
            heure = int(match.group(7))
            minute = int(match.group(8))
            seconde = int(match.group(9))
            
            return (numero_serie_str, annee, mois, jour, heure, minute, seconde)
        except Exception as e:
            # print(f"Error parsing sort key for '{identifiant_unique_avec_date_heure}': {e}") # Pour débogage
            sn_part = identifiant_unique_avec_date_heure.split(' [')[0] if ' [' in identifiant_unique_avec_date_heure else identifiant_unique_avec_date_heure
            return (sn_part, 0,0,0,0,0,0)

    def analyser_fichiers(self):
        donnees_collectees = {}
        all_files = self.fichiers_seq01 + self.fichiers_seq02
        # progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate", maximum=len(all_files))
        # progress_bar.pack(pady=10) # Ou placer dans la status bar

        for idx, fichier_path in enumerate(all_files):
            seq_type = "SEQ-01" if fichier_path in self.fichiers_seq01 else "SEQ-02"
            # progress_bar['value'] = idx + 1
            # self.update_status(f"Analyse {seq_type}: {os.path.basename(fichier_path)} ({idx+1}/{len(all_files)})")
            # self.root.update_idletasks()

            try:
                with open(fichier_path, "r", encoding="iso-8859-1", errors="replace") as f_html:
                    html_content = f_html.read()
                numero_serie = self.extraire_numero_serie(html_content) or os.path.basename(os.path.dirname(fichier_path))
                nom_fichier_base = os.path.basename(fichier_path)
                date_test, heure_test = self.extraire_date_heure(html_content, nom_fichier_base)
                statut_test = self.extraire_statut_test(html_content)
                identifiant_unique = f"{numero_serie} [{date_test}][{heure_test}]"

                if identifiant_unique not in donnees_collectees:
                    donnees_collectees[identifiant_unique] = {"Numéro de série": numero_serie, "Date": date_test, "Heure": heure_test, "Type": seq_type, "Statut": statut_test}
                else: # Merge if entry exists (e.g. from different file for same test instance, unlikely here but safe)
                    donnees_collectees[identifiant_unique]["Type"] += f"/{seq_type}" if seq_type not in donnees_collectees[identifiant_unique]["Type"] else ""


                for nom_complet_test, cat_id, id_extract in self.tests_selectionnes:
                    valeur = None
                    if seq_type == "SEQ-01" and cat_id.startswith("seq01_"):
                        valeur = self.extraire_valeur_seq01(html_content, cat_id, id_extract)
                    elif seq_type == "SEQ-02" and cat_id.startswith("seq02_"):
                        valeur = self.extraire_valeur_seq02(html_content, cat_id, id_extract)
                    
                    if valeur is not None: # Store even if empty string (for "precision_transfert_defauts")
                        donnees_collectees[identifiant_unique][nom_complet_test] = valeur
            except Exception as e:
                print(f"Erreur lors de l'analyse du fichier {fichier_path}: {e}")
                traceback.print_exc()
        
        # progress_bar.destroy()
        self.data_tests = donnees_collectees
        self.creer_tableau_statistiques()


    def creer_tableau_statistiques(self):
        """Crée un tableau statistique .xlsm à partir des données collectées."""
        if not self.data_tests:
            messagebox.showinfo("Information", "Aucune donnée collectée.")
            self.update_status("ℹ️ Aucune donnée à exporter.")
            return

        # Création du DataFrame à partir du dictionnaire self.data_tests
        # L'index sera les clés de self.data_tests, c'est-à-dire "SN [Date][Heure]"
        df_full = pd.DataFrame.from_dict(self.data_tests, orient='index')

        if df_full.empty:
            messagebox.showinfo("Information", "Aucune donnée à afficher.")
            self.update_status("ℹ️ DataFrame vide après conversion.")
            return
            
        # --- SECTION CRUCIALE POUR LE TRI ---
        if self.tri_chrono_var.get():
            # Trier l'index du DataFrame en utilisant la clé de tri personnalisée.
            # self.obtenir_cle_tri_chronologique doit retourner un tuple comme
            # (numero_serie_str, annee, mois, jour, heure, minute, seconde)
            # pour que pandas puisse trier correctement.
            
            # Créer une liste de clés de tri à partir de l'index actuel de df_full
            cles_de_tri = [self.obtenir_cle_tri_chronologique(idx_val) for idx_val in df_full.index]
            
            # Créer une série pandas à partir des clés de tri, en conservant l'index original de df_full
            # Cela permet de trier l'index original basé sur les tuples de la clé de tri.
            idx_tries_series = pd.Series(cles_de_tri, index=df_full.index)
            
            # Trier la série de clés, et obtenir le nouvel ordre de l'index
            nouvel_ordre_index = idx_tries_series.sort_values().index
            
            # Réorganiser df_full selon ce nouvel ordre d'index
            df_full = df_full.loc[nouvel_ordre_index]
        else:
            # Tri alphabétique standard de l'index si la case n'est pas cochée
            df_full = df_full.sort_index()
        # --- FIN SECTION CRUCIALE POUR LE TRI ---

        # Les colonnes "Numéro de série brut", "Date", "Heure" sont dans les données,
        # mais l'index contient déjà l'info formatée pour la colonne A.
        cols_to_drop_if_in_df = ["Numéro de série", "Date", "Heure"] 
        df_to_write = df_full.drop(columns=[col for col in cols_to_drop_if_in_df if col in df_full.columns], errors='ignore')

        # Réorganisation des colonnes pour l'affichage Excel (identique à avant)
        standard_cols_excel = ["Type", "Statut"]
        special_col_excel = "Défauts précision transfert"
        selected_test_cols_in_df = [test[0] for test in self.tests_selectionnes if test[0] in df_to_write.columns]
        
        final_col_order_excel = []
        for col in standard_cols_excel:
            if col in df_to_write.columns:
                final_col_order_excel.append(col)
        
        if special_col_excel in df_to_write.columns and special_col_excel not in final_col_order_excel:
            final_col_order_excel.append(special_col_excel)
            
        for col in selected_test_cols_in_df:
            if col not in final_col_order_excel: final_col_order_excel.append(col)
        
        other_existing_cols = [col for col in df_to_write.columns if col not in final_col_order_excel]
        final_col_order_excel.extend(other_existing_cols)
        
        final_col_order_excel = [col for col in final_col_order_excel if col in df_to_write.columns]
        df_to_write = df_to_write[final_col_order_excel]

        # --- DIAGNOSTIC (peut être commenté une fois que tout fonctionne) ---
        print("-" * 50)
        print("DIAGNOSTIC TRI: Contenu de df_to_write APRÈS TRI (AVANT écriture Excel)")
        print("Index de df_to_write (devrait être '0178 [Date][Heure]' ET TRIÉ):")
        print(df_to_write.index)
        print("\nPremières 10 lignes de df_to_write (pour vérifier le tri):")
        print(df_to_write.head(10).to_string())
        print("-" * 50)
        # --- FIN DIAGNOSTIC ---

        # ... (le reste de la fonction pour l'écriture dans Excel, comme dans la version précédente qui fonctionnait pour l'ouverture de dossier) ...
        # ... (nom_fichier_rapport_xlsm, chemin_rapport_xlsm, template_file_path) ...
        # ... (try/except block pour openpyxl.load_workbook, écriture des en-têtes et des données, config_sheet, formatage, wb.save) ...
        # La partie écriture Excel (ws.append, etc.) de la réponse précédente était correcte pour mettre l'index en colonne A.

        nom_fichier_rapport_xlsm = "statistiques_SEQ01_SEQ02.xlsm"
        chemin_rapport_xlsm = os.path.join(self.repertoire_parent, nom_fichier_rapport_xlsm)
        template_file_path = resource_path("template_statistiques.xlsm")

        try:
            self.update_status("📊 Chargement du modèle Excel et écriture des données...")
            wb = openpyxl.load_workbook(template_file_path, keep_vba=True)
            
            data_sheet_name = "Statistiques_SEQ01_02" 
            config_sheet_name = "ConfigSheet"

            ws = wb[data_sheet_name] if data_sheet_name in wb.sheetnames else wb.create_sheet(data_sheet_name)
            if ws.max_row > 1: 
                for r_idx_del in range(ws.max_row, 1, -1): ws.delete_rows(r_idx_del)

            index_header_name = df_to_write.index.name if df_to_write.index.name else "Identifiant Test (SN Date Heure)"
            excel_headers = [index_header_name] + list(df_to_write.columns)
            ws.append(excel_headers)
            
            header_font_style = Font(bold=True)
            for cell_header_obj in ws[1]: cell_header_obj.font = header_font_style
            
            for index_val_excel, data_row_excel in df_to_write.iterrows():
                row_to_write_to_excel = [str(index_val_excel)] + list(data_row_excel.values)
                ws.append(row_to_write_to_excel)

            config_ws = wb[config_sheet_name] if config_sheet_name in wb.sheetnames else wb.create_sheet(config_sheet_name)
            config_ws['A1'] = self.repertoire_parent

            ws.auto_filter.ref = ws.dimensions
            for col_cells_tuple_dim in ws.columns:
                max_len_dim = 0
                col_letter_dim_calc = col_cells_tuple_dim[0].column_letter
                for cell_dim_calc in col_cells_tuple_dim:
                    try:
                        if cell_dim_calc.value: max_len_dim = max(max_len_dim, len(str(cell_dim_calc.value)))
                    except: pass
                ws.column_dimensions[col_letter_dim_calc].width = max_len_dim + 5 if max_len_dim > 0 else 15


            col_map_excel_formatting = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
            for row_num_formatting_loop in range(2, ws.max_row + 1):
                for nom_col_borne_format, (borne_min_format, borne_max_format) in bornes_alims_seq01.items():
                    col_idx_format_lookup = col_map_excel_formatting.get(nom_col_borne_format)
                    if col_idx_format_lookup is not None:
                        cell_to_format_obj = ws.cell(row=row_num_formatting_loop, column=col_idx_format_lookup)
                        try:
                            if cell_to_format_obj.value is not None:
                                val_str_format = str(cell_to_format_obj.value).replace(",", ".").strip()
                                if val_str_format: 
                                    val_float_format = float(val_str_format)
                                    if not (borne_min_format <= val_float_format <= borne_max_format): cell_to_format_obj.font = Font(color="FF0000")
                        except (ValueError, TypeError): pass
                
                for res_prefix_format in ["R46", "R47", "R48"]:
                    col_calc_key_format = f"{res_prefix_format} calculée"
                    col_mont_key_format = f"{res_prefix_format} à monter"
                    col_calc_idx_format = col_map_excel_formatting.get(col_calc_key_format)
                    col_mont_idx_format = col_map_excel_formatting.get(col_mont_key_format)
                    if col_calc_idx_format and col_mont_idx_format:
                        cell_calc_res_obj = ws.cell(row=row_num_formatting_loop, column=col_calc_idx_format)
                        cell_mont_res_obj = ws.cell(row=row_num_formatting_loop, column=col_mont_idx_format)
                        try:
                            if cell_calc_res_obj.value is not None and cell_mont_res_obj.value is not None:
                                val_calc_res_str = str(cell_calc_res_obj.value).replace(",",".").strip()
                                val_mont_res_str = str(cell_mont_res_obj.value).replace(",",".").strip()
                                if val_calc_res_str and val_mont_res_str:
                                    val_calc_res_int = int(float(val_calc_res_str))
                                    val_mont_res_int = int(float(val_mont_res_str))
                                    if val_mont_res_int > val_calc_res_int: cell_mont_res_obj.font = Font(color="FF0000")
                        except (ValueError, TypeError): pass

            # --- AJOUT POUR FIGER LES VOLETS ---
            # Figer la première ligne et la première colonne.
            # La cellule 'B2' signifie que tout ce qui est au-dessus de la ligne 2 (donc la ligne 1)
            # et tout ce qui est à gauche de la colonne B (donc la colonne A) sera figé.
            if ws.max_row > 1 and ws.max_column > 1 : # S'assurer qu'il y a au moins une cellule de données en B2
                 ws.freeze_panes = ws['B2']
            elif ws.max_row > 1 : # S'il n'y a que la colonne A et des lignes, figer juste la ligne 1
                 ws.freeze_panes = ws['A2']
            elif ws.max_column > 1 : # S'il n'y a que la ligne 1 et des colonnes, figer juste la colonne A
                 ws.freeze_panes = ws['B1']
            # Si la feuille est très petite (1x1), ne rien figer.
            # --- FIN AJOUT POUR FIGER LES VOLETS ---


            wb.save(chemin_rapport_xlsm)
            self.update_status(f"✅ Rapport XLSM généré : {os.path.basename(chemin_rapport_xlsm)}")
            try:
                if sys.platform == "win32": os.startfile(chemin_rapport_xlsm)
                elif sys.platform == "darwin": subprocess.call(["open", chemin_rapport_xlsm])
                else: subprocess.call(["xdg-open", chemin_rapport_xlsm])
            except Exception as e_open_final: print(f"Impossible d'ouvrir: {e_open_final}")

        except FileNotFoundError:
             messagebox.showerror("Erreur Modèle", f"template_statistiques.xlsm non trouvé.")
             self.update_status("❌ Erreur: Modèle XLSM manquant")
        except PermissionError:
            messagebox.showerror("Permission Refusée", f"Vérifiez que {nom_fichier_rapport_xlsm} n'est pas ouvert.")
            self.update_status("❌ Erreur Permission XLSM")
        except Exception as e_xlsm_final:
            messagebox.showerror("Erreur XLSM", f"Erreur création XLSM: {e_xlsm_final}")
            self.update_status("❌ Erreur génération XLSM")
            traceback.print_exc()


    def generer_rapports_detailles(self):
        if not self.repertoire_parent:
            messagebox.showinfo("Information", "Veuillez d'abord sélectionner un répertoire.")
            return
        try:
            self.update_status("🔍 Recherche des sous-répertoires...")
            sous_repertoires = [os.path.join(self.repertoire_parent, d) for d in os.listdir(self.repertoire_parent)
                               if os.path.isdir(os.path.join(self.repertoire_parent, d))]
            if not sous_repertoires:
                messagebox.showinfo("Information", "Aucun sous-répertoire (numéro de série) trouvé.")
                self.update_status("ℹ️ Aucun sous-répertoire.")
                return
            self.update_status(f"📄 Génération des rapports pour {len(sous_repertoires)} séries...")
            progress_window = ProgressWindow(total_files=len(sous_repertoires))
            for repertoire in sous_repertoires:
                traiter_repertoire_serie(repertoire, progress_window)
            progress_window.show_completion()
            self.root.wait_window(progress_window.window) # Important pour Tkinter
            self.update_status("✅ Rapports détaillés générés avec succès")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération des rapports: {e}")
            self.update_status("❌ Erreur génération rapports détaillés")
            traceback.print_exc()

    def lancer(self):
        """Lance l'application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernStatsTestsWindow()
    app.lancer()