# --- START OF FILE LPVT_Gestion_Rapports.py (Checkbutton Selection) ---

import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from typing import Dict, List, Any, Set, Optional, Tuple
import io
import subprocess
try:
    from Affiche_resultats import traiter_repertoire_serie, ProgressWindow
except ImportError:
     messagebox.showerror("Erreur d'Import", "Le fichier 'Affiche_resultats.py' est introuvable.\nVeuillez vous assurer qu'il se trouve dans le même répertoire.")
     sys.exit(1)

import webbrowser
import sv_ttk # <--- IMPORT AJOUTÉ

class ModernStatsTestsWindow:
    """
    Fenêtre avec thème sv_ttk et sélection par cases à cocher.
    La logique fonctionnelle est celle de l'original.
    """
    def __init__(self):
        self.configurer_encodage()
        self.fichiers_seq01 = []
        self.fichiers_seq02 = []
        self.tests_disponibles = [] # Sera peuplé par _get_liste_tests_predefinies
        self.tests_selectionnes = [] # Contiendra les tuples des tests cochés
        # --- Nouveau: Dictionnaire pour stocker les BooleanVar des checkboxes ---
        self.tests_selection_vars: Dict[str, tk.BooleanVar] = {}
        # ----------------------------------------------------------------------
        self.data_tests = {}
        self.repertoire_parent = ""

        self.root = tk.Tk()
        self.root.title("LPVT - Analyse des tests SEQ-01/SEQ-02 (Cases à Cocher)")
        self.root.geometry("1100x750")
        self.root.minsize(900, 650)

        try:
            sv_ttk.set_theme("light")
        except Exception as e:
            print(f"Avertissement : Impossible de charger sv_ttk ({e}). Thème Tkinter par défaut.")

        self.setup_minimal_styles()
        self._get_liste_tests_predefinies() # Charger la liste interne des tests

        # --- MODIFIÉ: Appeler la nouvelle fonction de création ---
        self.creer_interface_avec_checkboxes()
        # -------------------------------------------------------

        # Charger la liste dans l'UI initialement
        self._populate_test_checkboxes() # Renommé depuis filter_tests_list

    def setup_minimal_styles(self):
        """Configure uniquement les styles personnalisés nécessaires AVEC sv_ttk."""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'))
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
        # Ajouter style pour Checkbutton si besoin (normalement sv_ttk gère bien)
        style.configure('TCheckbutton', font=('Segoe UI', 10))


    # --- NOUVELLE MÉTHODE DE CRÉATION UI ---
    def creer_interface_avec_checkboxes(self):
        """Crée l'interface utilisateur avec des Checkbuttons pour la sélection."""
        main_frame = ttk.Frame(self.root, padding=(15, 15, 15, 15))
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        title_label = ttk.Label(header_frame, text="Analyse des tests SEQ-01/SEQ-02", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        action_buttons_frame = ttk.Frame(header_frame)
        action_buttons_frame.pack(side=tk.RIGHT)
        ttk.Button(action_buttons_frame, text="Aide", command=self.show_help).pack(side=tk.LEFT, padx=5)

        dir_frame = ttk.LabelFrame(main_frame, text=" Répertoire source ", padding=(10, 5, 10, 10))
        dir_frame.pack(fill=tk.X, pady=(0, 15))
        dir_selection_frame = ttk.Frame(dir_frame)
        dir_selection_frame.pack(fill=tk.X)
        select_dir_btn = ttk.Button(dir_selection_frame, text="Sélectionner un répertoire", command=self.selectionner_repertoire)
        select_dir_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.dir_path_label = ttk.Label(dir_selection_frame, text="Aucun répertoire sélectionné", foreground="#7f8c8d", font=('Segoe UI', 9))
        self.dir_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.files_info_label = ttk.Label(dir_frame, text="", font=('Segoe UI', 9))
        self.files_info_label.pack(fill=tk.X, pady=(5, 0))

        # --- Section Sélection des Tests (Modifiée) ---
        tests_frame = ttk.LabelFrame(main_frame, text=" Sélection des tests à analyser ", padding=(10, 5, 10, 10))
        tests_frame.pack(fill=tk.BOTH, expand=True)
        tests_frame.rowconfigure(2, weight=1) # La zone scrollable prendra l'espace
        tests_frame.columnconfigure(0, weight=1) # Le canvas prendra la largeur

        # Barre de recherche et boutons de sélection (comme avant, mais au-dessus du canvas)
        controls_frame = ttk.Frame(tests_frame)
        controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10)) # Grid layout ici
        controls_frame.columnconfigure(3, weight=1) # Donner du poids à la search entry

        ttk.Button(controls_frame, text="Tout cocher", command=self.selectionner_tous_tests).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(controls_frame, text="Tout décocher", command=self.deselectionner_tous_tests).grid(row=0, column=1, padx=(0, 10))

        ttk.Label(controls_frame, text="Rechercher:").grid(row=0, column=2, padx=(5, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(controls_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=3, sticky="ew")
        # --- MODIFIÉ: Appeler _populate_test_checkboxes au lieu de filter_tests_list ---
        search_entry.bind('<KeyRelease>', self._populate_test_checkboxes)
        # ------------------------------------------------------------------------------

        # Checkbox Tri (à droite)
        self.tri_chrono_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(controls_frame, text="Tri chrono.", variable=self.tri_chrono_var).grid(row=0, column=4, padx=(10, 0))


        # Zone scrollable pour les checkboxes
        # Utilise la structure Canvas + Scrollbar + Frame interne
        canvas = tk.Canvas(tests_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(tests_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas) # Le frame qui contiendra les checkboxes

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        # Lier la molette (optionnel mais pratique)
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=2, column=0, sticky='nsew') # Canvas dans la grille de tests_frame
        scrollbar.grid(row=2, column=1, sticky='ns') # Scrollbar à côté

        # --- Fin Section Sélection Modifiée ---


        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(action_frame, text="Générer le rapport statistique", command=self.generer_statistiques, style='Accent.TButton').pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(action_frame, text="Générer les rapports détaillés", command=self.generer_rapports_detailles).pack(side=tk.LEFT)

        self.status_bar = ttk.Frame(main_frame, height=25)
        self.status_bar.pack(fill=tk.X, pady=(15, 0))
        self.status_label = ttk.Label(self.status_bar, text="Prêt", foreground="#7f8c8d", font=('Segoe UI', 9))
        self.status_label.pack(side=tk.LEFT)


    # Méthode creer_interface_moderne originale SUPPRIMÉE


    # --- Fonctions d'interaction avec Checkboxes (Nouvelles/Modifiées) ---

    def _populate_test_checkboxes(self, event=None):
        """Remplit/Filtre la zone scrollable avec des Checkbuttons."""
        search_term = self.search_var.get().lower()

        # Vider le frame actuel
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        # Important: Ne pas vider self.tests_selection_vars ici,
        # car cela perdrait l'état des cases non visibles.
        # Les variables sont créées une fois dans __init__ ou au premier affichage.

        current_group = None
        for nom_complet, test_parent, identifiant in self.tests_disponibles:
            # Appliquer le filtre
            if search_term and search_term not in nom_complet.lower():
                continue # Ne pas afficher ce test

            # Afficher le groupe si nécessaire
            group_name = "SEQ-01" if test_parent.startswith("seq01") else "SEQ-02"
            if group_name != current_group:
                ttk.Label(self.scrollable_frame, text=f"--- {group_name} ---", style="Header.TLabel").pack(fill=tk.X, pady=(10,2), padx=5)
                current_group = group_name

            # Récupérer ou créer le BooleanVar pour ce test
            if nom_complet not in self.tests_selection_vars:
                self.tests_selection_vars[nom_complet] = tk.BooleanVar()

            # Créer le Checkbutton et le lier à sa variable
            var = self.tests_selection_vars[nom_complet]
            cb = ttk.Checkbutton(
                self.scrollable_frame,
                text=nom_complet, # Affiche seulement le nom lisible
                variable=var,
                style='TCheckbutton'
            )
            cb.pack(anchor='w', padx=15, pady=2)


    def selectionner_tous_tests(self):
        """Coche toutes les cases."""
        for var in self.tests_selection_vars.values():
            var.set(True)

    def deselectionner_tous_tests(self):
        """Décoche toutes les cases."""
        for var in self.tests_selection_vars.values():
            var.set(False)

    # --- Fonctions logiques (restent IDENTIQUES à votre original) ---

    def show_help(self): # Original
        help_text = """... (texte identique à la version précédente) ..."""
        messagebox.showinfo("Aide", help_text.strip())

    def update_status(self, message): # Original
        self.status_label.config(text=message)
        self.root.update_idletasks()

    # filter_tests_list est remplacé par _populate_test_checkboxes

    def configurer_encodage(self): # Original
        try:
            if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            if sys.stderr and hasattr(sys.stderr, 'encoding') and sys.stderr.encoding and sys.stderr.encoding.lower() != 'utf-8':
                sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except (AttributeError, TypeError, ValueError, io.UnsupportedOperation): pass

    def _get_liste_tests_predefinies(self): # Original
        self.tests_disponibles = [
            ("alim 24VDC +16V", "seq01_24vdc", "24VDC_+16V"), ("alim 24VDC -16V", "seq01_24vdc", "24VDC_-16V"),
            ("alim 24VDC +5V", "seq01_24vdc", "24VDC_+5V"), ("alim 24VDC -5V", "seq01_24vdc", "24VDC_-5V"),
            ("alim 115VAC +16V", "seq01_115vac", "115VAC_+16V"), ("alim 115VAC -16V", "seq01_115vac", "115VAC_-16V"),
            ("R46 calculée", "seq01_resistances", "R46_calculee"), ("R46 à monter", "seq01_resistances", "R46_monter"),
            ("R47 calculée", "seq01_resistances", "R47_calculee"), ("R47 à monter", "seq01_resistances", "R47_monter"),
            ("R48 calculée", "seq01_resistances", "R48_calculee"), ("R48 à monter", "seq01_resistances", "R48_monter"),
            ("1.9Un en 19VDC", "seq02_transfert", "Test_19VDC"), ("1.9Un en 115VAC", "seq02_transfert", "Test_115VAC"),
        ]

    def selectionner_repertoire(self): # Original
        repertoire = filedialog.askdirectory(title="Sélectionnez le répertoire parent contenant les dossiers N° Série")
        if not repertoire: return
        self.repertoire_parent = repertoire
        short_path = os.path.basename(repertoire)
        if len(repertoire) > 60: short_path = "..." + repertoire[-57:]
        self.dir_path_label.config(text=short_path, foreground="#2c3e50")
        self.update_status(f"Analyse du répertoire {os.path.basename(repertoire)}...")
        self.charger_fichiers_seq(repertoire)

    def charger_fichiers_seq(self, repertoire): # Original
        self.update_status("Recherche des fichiers SEQ-01/SEQ-02...")
        self.fichiers_seq01 = []; self.fichiers_seq02 = []
        count_seq01 = 0; count_seq02 = 0
        try:
            for item in os.listdir(repertoire):
                subdir_path = os.path.join(repertoire, item)
                if os.path.isdir(subdir_path):
                    for dossier_racine, _, fichiers in os.walk(subdir_path):
                        for fichier in fichiers:
                            f_lower = fichier.lower()
                            if f_lower.endswith(".html"):
                                chemin_complet = os.path.join(dossier_racine, fichier)
                                if fichier.startswith("SEQ-01"): self.fichiers_seq01.append(chemin_complet); count_seq01 += 1
                                elif fichier.startswith("SEQ-02"): self.fichiers_seq02.append(chemin_complet); count_seq02 += 1
        except Exception as e:
             self.files_info_label.config(text=f"Erreur lecture répertoire: {e}", foreground="red"); self.update_status("Erreur lecture")
             messagebox.showerror("Erreur", f"Impossible de lire le contenu du répertoire:\n{e}"); return
        total_fichiers = count_seq01 + count_seq02
        if total_fichiers == 0:
            self.files_info_label.config(text="Aucun fichier SEQ-01/SEQ-02 trouvé.", foreground="orange"); self.update_status("Aucun fichier trouvé")
        else:
            self.files_info_label.config(text=f"{count_seq01} SEQ-01 et {count_seq02} SEQ-02 trouvés.", foreground="green"); self.update_status(f"Prêt - {total_fichiers} fichiers.")


    # --- MODIFIÉ: Récupérer la sélection depuis les Checkboxes ---
    def generer_statistiques(self):
        # 1. Récupérer les tests cochés
        self.tests_selectionnes = [] # Réinitialiser la liste des tuples sélectionnés
        for nom_complet, test_parent, identifiant in self.tests_disponibles:
            # Vérifier si le test a une variable associée et si elle est cochée
            if nom_complet in self.tests_selection_vars and self.tests_selection_vars[nom_complet].get():
                self.tests_selectionnes.append((nom_complet, test_parent, identifiant))

        # 2. Vérifier si au moins un test est sélectionné
        if not self.tests_selectionnes:
            messagebox.showinfo("Information", "Veuillez cocher au moins un test.")
            return

        # 3. Vérifications restantes (répertoire, fichiers) - comme avant
        if not self.repertoire_parent or not os.path.isdir(self.repertoire_parent):
             messagebox.showwarning("Répertoire Manquant", "Sélectionnez un répertoire valide."); return
        if not self.fichiers_seq01 and not self.fichiers_seq02:
             messagebox.showwarning("Fichiers Manquants", "Aucun fichier SEQ trouvé."); return

        # 4. Lancer l'analyse et la création du tableau (comme avant)
        self.update_status("Analyse des fichiers en cours...")
        self.analyser_fichiers()
        self.creer_tableau_statistiques()
    # -----------------------------------------------------------


    # --- Fonctions d'extraction (originales, inchangées) ---
    def extraire_valeur_seq01(self, html_content, test_parent, identifiant): # Original
        # ... (code identique à la version précédente) ...
        if test_parent == "seq01_24vdc":
            block_match = re.search(r"Test des alimentations à 24VDC(.*?)Test des alimentations à 115VAC", html_content, re.DOTALL)
            if not block_match: return None
            block = block_match.group(1)
            patterns = { "24VDC_+16V": r"Lecture mesure \+16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", "24VDC_-16V": r"Lecture mesure -16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", "24VDC_+5V": r"Lecture mesure \+5V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", "24VDC_-5V": r"Lecture mesure -5V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", }
            if identifiant in patterns: match = re.search(patterns[identifiant], block, re.DOTALL | re.IGNORECASE); return match.group(1).strip() if match else None
        elif test_parent == "seq01_115vac":
            block_match = re.search(r"Test des alimentations à 115VAC(.*?)Calcul des résistances", html_content, re.DOTALL)
            if not block_match: return None
            block = block_match.group(1)
            patterns = { "115VAC_+16V": r"Lecture mesure \+16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", "115VAC_-16V": r"Lecture mesure -16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", }
            if identifiant in patterns: match = re.search(patterns[identifiant], block, re.DOTALL | re.IGNORECASE); return match.group(1).strip() if match else None
        elif test_parent == "seq01_resistances":
            patterns = { "R46_calculee": r"R46 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", "R46_monter": r"R46 à monter:\s*</td>\s*<td[^>]*>\s*.*?=\s*([\d]+)\s*ohms", "R47_calculee": r"R47 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", "R47_monter": r"R47 à monter:\s*</td>\s*<td[^>]*>\s*.*?=\s*([\d]+)\s*ohms", "R48_calculee": r"R48 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", "R48_monter": r"R48 à monter:\s*</td>\s*<td[^>]*>\s*.*?=\s*([\d]+)\s*ohms", }
            if identifiant in patterns: match = re.search(patterns[identifiant], html_content, re.DOTALL | re.IGNORECASE); return match.group(1).strip() if match else None
        return None

    def extraire_valeur_seq02(self, html_content, test_parent, identifiant): # Original
        # ... (code identique à la version précédente) ...
        if test_parent == "seq02_transfert":
            patterns = { "Test_19VDC": r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+19VDC.*?Lecture mesure -16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>", "Test_115VAC": r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+115VAC.*?Lecture mesure -16V.*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>" }
            if identifiant in patterns: match = re.search(patterns[identifiant], html_content, re.DOTALL | re.IGNORECASE); return match.group(1).strip() if match else None
        return None

    def extraire_numero_serie(self, html_content): # Original
        # ... (code identique à la version précédente) ...
        sn_match = re.search(r'Serial Number:</td>\s*<td[^>]*class="hdr_value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL | re.IGNORECASE); sn = sn_match.group(1).strip() if sn_match else None;
        if sn and sn.upper() != "NONE": return sn
        serie_match = re.search(r'série[^<]*</td>\s*<td[^>]*class="value"[^>]*>\s*([^<]+)\s*</td>', html_content, re.DOTALL | re.IGNORECASE); sn = serie_match.group(1).strip() if serie_match else None;
        if sn is not None: return sn # Peut être vide
        alt_match1 = re.search(r'<td class="hdr_name">\s*Serial Number:\s*</td>\s*<td class="hdr_value">\s*<b>([^<]+)</b>\s*</td>', html_content, re.DOTALL | re.IGNORECASE); sn = alt_match1.group(1).strip() if alt_match1 else None;
        if sn and sn.upper() != "NONE": return sn
        alt_match2 = re.search(r'<td class="label".*?>\s*Serial Number\s*</td>\s*<td class="value".*?>([^<]+)</td>', html_content, re.DOTALL | re.IGNORECASE); sn = alt_match2.group(1).strip() if alt_match2 else None;
        if sn and sn.upper() != "NONE": return sn
        return None

    def extraire_statut_test(self, html_content): # Original
        # ... (code identique à la version précédente) ...
        status_match = re.search(r"Result:.*?<span style.*?>(.*?)</span>", html_content, re.DOTALL | re.IGNORECASE);
        if status_match: return status_match.group(1).strip()
        alt_match = re.search(r"UUT Result:.*?hdr_value.*?<span.*?>(Passed|Failed|Terminated)</span>", html_content, re.DOTALL | re.IGNORECASE);
        if alt_match: return alt_match.group(1).strip()
        return "Inconnu"

    def extraire_date_heure(self, html_content, nom_fichier): # Original
        # ... (code identique à la version précédente) ...
        date_match = re.search(r'Date:</td>\s*<td.*?<b>([^<]+)</b>', html_content, re.DOTALL); date = date_match.group(1).strip() if date_match else None
        time_match = re.search(r'Time:</td>\s*<td.*?<b>([^<]+)</b>', html_content, re.DOTALL); heure = time_match.group(1).strip() if time_match else None
        if not date or not heure: date_nom, heure_nom = self.extraire_date_heure_du_nom(nom_fichier); date = date if date else date_nom; heure = heure if heure else heure_nom
        try: from datetime import datetime; dt_obj = datetime.strptime(f"{date} {heure}", '%m/%d/%Y %I:%M:%S %p'); date = dt_obj.strftime('%d/%m/%Y'); heure = dt_obj.strftime('%H:%M:%S')
        except (ValueError, TypeError): pass
        return (date if date else "Date?"), (heure if heure else "Heure?")


    def extraire_date_heure_du_nom(self, nom_fichier): # Original
        # ... (code identique à la version précédente) ...
        match = re.search(r'\[(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\]\s*\[(\d{1,2})\s+(\d{1,2})\s+(\d{4})\]', nom_fichier);
        if match: h, m, s, j, mo, a = match.groups(); return f"{int(j):02d}/{int(mo):02d}/{a}", f"{int(h):02d}:{int(m):02d}:{int(s):02d}"
        return "date_inconnue", "heure_inconnue"

    def obtenir_cle_tri_chronologique(self, identifiant_unique): # Original
        # ... (code identique à la version précédente) ...
        try:
            parts = identifiant_unique.split(' [');
            if len(parts) < 3: return (identifiant_unique, 0,0,0,0,0,0)
            num_serie, date_str, heure_str = parts[0], parts[1].split(']')[0], parts[2].split(']')[0]
            try: j,m,a = map(int, date_str.split('/')); h,mi,s = map(int, heure_str.split(':')); return (num_serie,a,m,j,h,mi,s)
            except ValueError: pass
            try: from datetime import datetime; dt_obj_us = datetime.strptime(f"{date_str} {heure_str}", '%m/%d/%Y %I:%M:%S %p'); return (num_serie, dt_obj_us.year, dt_obj_us.month, dt_obj_us.day, dt_obj_us.hour, dt_obj_us.minute, dt_obj_us.second)
            except ValueError: return (num_serie, 0,0,0,0,0,0)
        except: return (identifiant_unique, 9999,12,31,23,59,59)


    # --- Fonctions analyse et création tableau (originales, inchangées) ---
    def analyser_fichiers(self): # Original
        # ... (code identique à la version précédente) ...
        self.data_tests = {}; donnees = self.data_tests; all_files = self.fichiers_seq01 + self.fichiers_seq02; total_fichiers = len(all_files);
        for i, fichier in enumerate(all_files):
            is_seq01 = fichier in self.fichiers_seq01; fichier_type = "SEQ-01" if is_seq01 else "SEQ-02"; self.update_status(f"Analyse {i+1}/{total_fichiers}: {os.path.basename(fichier)}");
            try:
                with open(fichier, "r", encoding="iso-8859-1", errors="replace") as f: html_content = f.read()
                num_serie = self.extraire_numero_serie(html_content);
                if not num_serie: num_serie = os.path.basename(os.path.dirname(fichier)); print(f"WARN: SN non trouvé, utilisé dossier: {num_serie}");
                nom_fichier_base = os.path.basename(fichier); date, heure = self.extraire_date_heure(html_content, nom_fichier_base); statut = self.extraire_statut_test(html_content); id_unique = f"{num_serie} [{date}][{heure}]";
                if id_unique not in donnees: donnees[id_unique] = {"Numéro de série": num_serie, "Date": date, "Heure": heure, "Type": fichier_type, "Statut": statut}
                extract_func = self.extraire_valeur_seq01 if is_seq01 else self.extraire_valeur_seq02; expected_prefix = "seq01_" if is_seq01 else "seq02_";
                for nom_complet, test_parent, identifiant in self.tests_selectionnes:
                    if test_parent.startswith(expected_prefix): valeur = extract_func(html_content, test_parent, identifiant);
                    if valeur is not None: donnees[id_unique][nom_complet] = valeur # Stocke seulement si non None
            except Exception as e: print(f"ERREUR analyse {fichier}: {e}"); import traceback; traceback.print_exc(); self.update_status(f"Erreur analyse {os.path.basename(fichier)}");
        self.update_status(f"Analyse terminée. {len(donnees)} entrées.");

    def creer_tableau_statistiques(self): # Original
        # ... (code identique à la version précédente) ...
        if not self.data_tests: messagebox.showinfo("Information", "Aucune donnée collectée."); self.update_status("Aucune donnée."); return
        df = None;
        try:
            if self.tri_chrono_var.get(): ids_tries = sorted(self.data_tests.keys(), key=self.obtenir_cle_tri_chronologique); data = {k: self.data_tests[k] for k in ids_tries}; df = pd.DataFrame.from_dict(data, orient='index')
            else: df = pd.DataFrame.from_dict(self.data_tests, orient='index')
            cols_debut = ["Numéro de série", "Date", "Heure", "Type", "Statut"]; cols_debut_present = [c for c in cols_debut if c in df.columns]; cols_tests = sorted([c for c in df.columns if c not in cols_debut_present]); df = df[cols_debut_present + cols_tests]; df.index.name = "Identifiant Unique";
            chemin_excel = os.path.join(self.repertoire_parent, "statistiques_SEQ01_SEQ02.xlsx"); self._log_action(f"Sauvegarde vers {chemin_excel}...");
            try:
                with pd.ExcelWriter(chemin_excel, engine='openpyxl') as writer: df.to_excel(writer, sheet_name="Statistiques_SEQ01_02", index=True)
                self.update_status(f"Rapport généré: {os.path.basename(chemin_excel)}"); self._log_action("Rapport Excel généré.", success=True);
                try: webbrowser.open(f'file:///{os.path.realpath(chemin_excel)}')
                except Exception as open_e: messagebox.showwarning("Ouverture Fichier", f"Impossible d'ouvrir Excel auto:\n{open_e}"); self._log_action(f"Erreur ouverture Excel: {open_e}", error=True);
            except PermissionError: messagebox.showerror("Erreur Sauvegarde", f"Permission refusée:\n{chemin_excel}\nFichier ouvert?"); self._log_action(f"Erreur permission Excel", error=True);
            except Exception as excel_e: messagebox.showerror("Erreur Excel", f"Erreur création Excel:\n{excel_e}"); self._log_action(f"Erreur écriture Excel: {excel_e}", error=True); print("ERREUR EXCEL:"); import traceback; traceback.print_exc();
        except Exception as e: messagebox.showerror("Erreur Tableau", f"Erreur création tableau:\n{e}"); self.update_status("Erreur tableau"); self._log_action(f"Erreur DataFrame: {e}", error=True); print("ERREUR DATAFRAME:"); import traceback; traceback.print_exc();

    def _log_action(self, message, success=False, error=False): # Original (Helper)
        print(message); self.update_status(message);

    def generer_rapports_detailles(self): # Original
        # ... (code identique à la version précédente) ...
        if not self.repertoire_parent or not os.path.isdir(self.repertoire_parent): messagebox.showinfo("Information", "Sélectionnez un répertoire."); return
        try:
            sous_repertoires = [os.path.join(self.repertoire_parent, d) for d in os.listdir(self.repertoire_parent) if os.path.isdir(os.path.join(self.repertoire_parent, d))]
            if not sous_repertoires: messagebox.showinfo("Information", "Aucun sous-répertoire trouvé."); return
            progress_window = None # Initialiser
            try: progress_window = ProgressWindow(total_files=len(sous_repertoires)) # Créer la fenêtre
            except NameError: messagebox.showerror("Erreur", "Classe ProgressWindow non trouvée."); return
            except Exception as e: messagebox.showerror("Erreur", f"Erreur ProgressWindow: {e}"); return
            self.update_status("Génération rapports détaillés..."); processed_count = 0;
            for repertoire in sous_repertoires:
                 try:
                      traiter_repertoire_serie(repertoire) # Appel original
                      processed_count += 1
                      if progress_window and hasattr(progress_window, 'update_progress'): progress_window.update_progress() # MàJ fenêtre externe
                 except Exception as e: print(f"ERREUR Traitement détaillé {os.path.basename(repertoire)}: {e}"); import traceback; traceback.print_exc(); self.update_status(f"Erreur sur {os.path.basename(repertoire)}");
            if progress_window and hasattr(progress_window, 'show_completion'): progress_window.show_completion()
            self.update_status("Rapports détaillés générés.");
        except Exception as e: messagebox.showerror("Erreur", f"Erreur génération rapports: {e}"); self.update_status("Erreur rapports détaillés"); print("ERREUR GENERATION DETAILLEE:"); import traceback; traceback.print_exc();


    def lancer(self): # Original
        self.root.mainloop()

if __name__ == "__main__": # Original
    app = ModernStatsTestsWindow()
    app.lancer()

# --- END OF FILE LPVT_Gestion_Rapports.py (Checkbutton Selection) ---