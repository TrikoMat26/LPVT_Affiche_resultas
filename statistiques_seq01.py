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
    Fenêtre principale pour l'analyse statistique des tests SEQ-01
    """
    def __init__(self):
        """Initialise l'application"""
        # Configuration de l'encodage
        self.configurer_encodage()
        
        # Variables de l'application
        self.fichiers_seq01 = []
        self.tests_disponibles = []  # Liste de tuples (nom_complet_test, test_parent, identifiant)
        self.tests_selectionnes = []
        self.data_tests = {}
        self.repertoire_parent = ""
        
        # Interface graphique
        self.root = tk.Tk()
        self.root.title("Statistiques SEQ-01")
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
            title="Sélectionnez le répertoire contenant les fichiers SEQ-01"
        )
        if not repertoire:
            return
            
        self.repertoire_parent = repertoire
        self.lbl_repertoire.config(text=f"Répertoire: {os.path.basename(repertoire)}")
        
        # Rechercher les fichiers SEQ-01 et charger les tests disponibles
        self.charger_fichiers_seq01(repertoire)
        
    def charger_fichiers_seq01(self, repertoire):
        """Cherche les fichiers SEQ-01 et charge les tests disponibles"""
        # Chercher les fichiers SEQ-01
        self.fichiers_seq01 = []
        
        # Chercher récursivement dans les sous-répertoires
        for dossier_racine, sous_dossiers, fichiers in os.walk(repertoire):
            for fichier in fichiers:
                if fichier.startswith("SEQ-01") and fichier.lower().endswith(".html"):
                    chemin_complet = os.path.join(dossier_racine, fichier)
                    self.fichiers_seq01.append(chemin_complet)
        
        if not self.fichiers_seq01:
            messagebox.showinfo("Information", "Aucun fichier SEQ-01 trouvé.")
            return
        
        # Analyser un fichier pour récupérer la liste des tests disponibles
        messagebox.showinfo("Information", f"{len(self.fichiers_seq01)} fichiers SEQ-01 trouvés.\nChargement des tests disponibles...")
        
        # Créer manuellement la liste des tests disponibles selon les spécifications
        self.creer_liste_tests_predefinies()
        self.afficher_tests_disponibles()
    
    def creer_liste_tests_predefinies(self):
        """
        Crée une liste prédéfinie des tests à extraire, basée sur les spécifications
        """
        self.tests_disponibles = [
            # Format: (nom_complet, test_parent, identifiant)
            # Test des alimentations à 24VDC
            ("Test des alimentations à 24VDC > Lecture mesure +16V AG34461A", 
                "Test des alimentations à 24VDC", "24VDC_+16V"),
            ("Test des alimentations à 24VDC > Lecture mesure -16V AG34461A", 
                "Test des alimentations à 24VDC", "24VDC_-16V"),
            ("Test des alimentations à 24VDC > Lecture mesure +5V AG34461A", 
                "Test des alimentations à 24VDC", "24VDC_+5V"),
            ("Test des alimentations à 24VDC > Lecture mesure -5V AG34461A", 
                "Test des alimentations à 24VDC", "24VDC_-5V"),
            
            # Test des alimentations à 115VAC
            ("Test des alimentations à 115VAC > Lecture mesure +16V AG34461A", 
                "Test des alimentations à 115VAC", "115VAC_+16V"),
            ("Test des alimentations à 115VAC > Lecture mesure -16V AG34461A", 
                "Test des alimentations à 115VAC", "115VAC_-16V"),
            
            # Calcul des résistances
            ("Calcul des résistances > Résistance R46 calculée", 
                "Calcul des résistances", "R46_calculee"),
            ("Calcul des résistances > Résistance R46 à monter", 
                "Calcul des résistances", "R46_monter"),
            ("Calcul des résistances > Résistance R47 calculée", 
                "Calcul des résistances", "R47_calculee"),
            ("Calcul des résistances > Résistance R47 à monter", 
                "Calcul des résistances", "R47_monter"),
            ("Calcul des résistances > Résistance R48 calculée", 
                "Calcul des résistances", "R48_calculee"),
            ("Calcul des résistances > Résistance R48 à monter", 
                "Calcul des résistances", "R48_monter"),
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
    
    def extraire_valeur_mesure(self, soup, test_parent, identifiant):
        """
        Extrait la valeur d'une mesure spécifique à partir du fichier HTML
        """
        # Trouver la section correspondant au test parent
        table_parent = None
        
        # Pour "Test des alimentations à 24VDC"
        if test_parent == "Test des alimentations à 24VDC":
            # Trouver le titre principal "Test des alimentations à 24VDC"
            for td in soup.find_all("td", colspan="2", string=lambda s: s and "Test des alimentations à 24VDC" in s):
                if "background-color:#00C4C4" in td.get("style", ""):
                    # Trouver le sous-test dans la séquence "TEST DES ALIMENTATIONS"
                    # d'abord, on cherche la div contenant "TEST DES ALIMENTATIONS" après le test parent
                    current = td
                    while current:
                        next_div = current.find_next("div")
                        if next_div and "Begin Sequence: TEST DES ALIMENTATIONS" in next_div.text:
                            # Chercher le sous-test dans cette div
                            if identifiant == "24VDC_+16V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure +16V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            elif identifiant == "24VDC_-16V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure -16V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            elif identifiant == "24VDC_+5V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure +5V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            elif identifiant == "24VDC_-5V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure -5V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            break
                        current = next_div
                    break
        
        elif test_parent == "Test des alimentations à 115VAC":
            # Trouver le titre principal "Test des alimentations à 115VAC"
            for td in soup.find_all("td", colspan="2", string=lambda s: s and "Test des alimentations à 115VAC" in s):
                if "background-color:#00C4C4" in td.get("style", ""):
                    # Trouver le sous-test dans la séquence "TEST DES ALIMENTATIONS"
                    current = td
                    while current:
                        next_div = current.find_next("div")
                        if next_div and "Begin Sequence: TEST DES ALIMENTATIONS" in next_div.text:
                            # Chercher le sous-test dans cette div
                            if identifiant == "115VAC_+16V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure +16V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            elif identifiant == "115VAC_-16V":
                                sous_test = next_div.find("td", colspan="2", string=lambda s: s and "Lecture mesure -16V AG34461A" in s)
                                if sous_test:
                                    table_parent = sous_test.find_parent("table")
                            break
                        current = next_div
                    break
        
        elif test_parent == "Calcul des résistances":
            for td in soup.find_all("td", colspan="2", string=lambda s: s and "Calcul des résistances" in s):
                if "background-color:#00C4C4" in td.get("style", ""):
                    current = td
                    while current:
                        next_div = current.find_next("div")
                        if next_div and "Begin Sequence: CALCUL DES RESISTANCES" in next_div.text:
                            if identifiant == "R46_calculee":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R46 calculée" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        return value_cell.text.strip()
                            elif identifiant == "R46_monter":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R46 à monter" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        match = re.search(r'(\d+)\s*ohms', value_cell.text.strip())
                                        if match:
                                            return match.group(1)
                                        return value_cell.text.strip()
                            elif identifiant == "R47_calculee":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R47 calculée" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        return value_cell.text.strip()
                            elif identifiant == "R47_monter":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R47 à monter" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        match = re.search(r'(\d+)\s*ohms', value_cell.text.strip())
                                        if match:
                                            return match.group(1)
                                        return value_cell.text.strip()
                            elif identifiant == "R48_calculee":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R48 calculée" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        return value_cell.text.strip()
                            elif identifiant == "R48_monter":
                                label = next_div.find("td", class_="label", string=lambda s: s and "Résistance R48 à monter" in s)
                                if label:
                                    value_cell = label.find_next_sibling("td", class_="value")
                                    if value_cell:
                                        match = re.search(r'(\d+)\s*ohms', value_cell.text.strip())
                                        if match:
                                            return match.group(1)
                                        return value_cell.text.strip()
                            break
                        current = next_div
                    break
        
        # Pour les mesures d'alimentation, extraire la valeur après "Measurement[1]"
        if table_parent and identifiant.startswith(("24VDC_", "115VAC_")):
            # Chercher la ligne contenant "Measurement[1]"
            for tr in table_parent.find_all("tr"):
                if "Measurement[1]" in tr.text:
                    # La valeur est dans la ligne suivante, sous "Data:"
                    next_tr = tr.find_next_sibling("tr")
                    if next_tr:
                        data_label = next_tr.find("td", class_="label", string=lambda s: s and "Data:" in s)
                        if data_label:
                            value_cell = data_label.find_next_sibling("td", class_="value")
                            if not value_cell:
                                # Si on ne trouve pas dans le sibling direct, chercher dans la même ligne
                                value_cell = next_tr.find("td", class_="value")
                                
                            if value_cell:
                                # Chercher la valeur directement dans la cellule ou dans un span
                                span = value_cell.find("span")
                                text_value = span.text.strip() if span else value_cell.text.strip()
                                
                                # S'assurer que les valeurs négatives sont correctement traitées
                                # Vérifier s'il s'agit d'une valeur négative (commence par un signe moins)
                                if text_value.startswith('-'):
                                    return text_value  # Retourner la valeur négative telle quelle
                                else:
                                    return text_value  # Retourner la valeur positive telle quelle
        
        return None
    
    def extraire_numero_serie(self, soup):
        """Extrait le numéro de série depuis la soupe HTML"""
        # D'abord chercher dans les entêtes (Serial Number)
        balise_sn = soup.find("td", class_="hdr_name", string=lambda t: t and "Serial Number:" in t)
        if balise_sn:
            balise_sn_value = balise_sn.find_next_sibling("td", class_="hdr_value")
            if balise_sn_value:
                sn = balise_sn_value.text.strip()
                if sn != "NONE":
                    return sn
        
        # Ensuite chercher dans le contenu pour "Numéro de série de la carte en test"
        for label in soup.find_all("td", class_="label"):
            if "série" in label.text:
                value_cell = label.find_next_sibling("td", class_="value")
                if value_cell:
                    return value_cell.text.strip()
        
        # Si rien n'est trouvé, utiliser le nom du répertoire
        return None
    
    def analyser_fichiers(self):
        """Analyse tous les fichiers SEQ-01 et collecte les données pour les tests sélectionnés"""
        donnees = {}  # Dictionnaire pour stocker les données: {numero_serie: {test1: valeur1, test2: valeur2, ...}}
        
        for fichier in self.fichiers_seq01:
            try:
                with open(fichier, "r", encoding="iso-8859-1", errors="replace") as f:
                    html_content = f.read()
                    soup = BeautifulSoup(html_content, "html.parser")
                
                # Extraire le numéro de série
                numero_serie = self.extraire_numero_serie(soup) or os.path.basename(os.path.dirname(fichier))
                
                # Initialiser le dictionnaire pour ce numéro de série si nécessaire
                if numero_serie not in donnees:
                    donnees[numero_serie] = {}
                
                # Extraire les valeurs pour chaque test sélectionné
                for nom_complet, test_parent, identifiant in self.tests_selectionnes:
                    valeur = self.extraire_valeur_mesure(soup, test_parent, identifiant)
                    if valeur:
                        donnees[numero_serie][nom_complet] = valeur
            
            except Exception as e:
                print(f"Erreur lors de l'analyse du fichier {fichier}: {e}")
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
        chemin_excel = os.path.join(self.repertoire_parent, "statistiques_SEQ01.xlsx")
        try:
            df.to_excel(chemin_excel, sheet_name="Statistiques SEQ-01")
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