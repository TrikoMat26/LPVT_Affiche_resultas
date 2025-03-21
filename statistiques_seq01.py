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
        self.tests_disponibles = []
        self.tests_selectionnes = []
        self.data_tests = {}
        self.repertoire_parent = ""
        
        # Interface graphique
        self.root = tk.Tk()
        self.root.title("Statistiques SEQ-01")
        self.root.geometry("800x600")
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
            yscrollcommand=scroll_y.set
        )
        self.listbox_tests.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.config(command=self.listbox_tests.yview)
        
        # Bouton pour générer le rapport statistique
        btn_generer = ttk.Button(
            main_frame, 
            text="Générer statistiques", 
            command=self.generer_statistiques
        )
        btn_generer.pack(pady=10)
        
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
        
        # Utiliser un échantillon pour extraire les tests disponibles
        try:
            sample_file = self.fichiers_seq01[0]
            self.tests_disponibles = self.extraire_tests_disponibles(sample_file)
            self.afficher_tests_disponibles()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'analyse des tests: {e}")
    
    def extraire_tests_disponibles(self, chemin_fichier):
        """Extrait la liste des tests disponibles dans un fichier SEQ-01"""
        tests = []
        try:
            # Les rapports utilisent souvent iso-8859-1
            with open(chemin_fichier, "r", encoding="iso-8859-1", errors="replace") as f:
                html_content = f.read()
                soup = BeautifulSoup(html_content, "html.parser")
            
            # Rechercher tous les titres de tests (cellules avec colspan="2")
            titres_tests = soup.find_all("td", colspan="2")
            for titre in titres_tests:
                nom_test = titre.text.strip()
                if nom_test and nom_test not in ["", "Begin Sequence", "End Sequence"]:
                    tests.append(nom_test)
            
            # Rechercher aussi les labels qui contiennent des valeurs mesurées
            labels = soup.find_all("td", class_="label")
            for label in labels:
                label_text = label.text.strip()
                
                # Filtrer les labels pertinents
                if any(pattern in label_text for pattern in [
                    "Résistance", "Tension", "Gain", "mesurée", "calculée"
                ]):
                    # Ne pas inclure les labels qui contiennent uniquement "Status:" ou similaires
                    if "Status:" not in label_text and label_text not in tests:
                        tests.append(label_text)
        
        except Exception as e:
            print(f"Erreur lors de l'extraction des tests: {e}")
        
        # Trier alphabétiquement les tests pour faciliter la sélection
        return sorted(set(tests))
    
    def afficher_tests_disponibles(self):
        """Affiche les tests disponibles dans la listbox"""
        # Vider la listbox
        self.listbox_tests.delete(0, tk.END)
        
        # Ajouter les tests
        for test in self.tests_disponibles:
            self.listbox_tests.insert(tk.END, test)
    
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
    
    def extraire_valeur_test(self, soup, nom_test):
        """
        Extrait la valeur d'un test spécifique depuis la soupe HTML
        Retourne None si la valeur n'est pas trouvée
        """
        # Pour les titres de tests (cellules avec colspan="2")
        if "Résistance" not in nom_test and "Tension" not in nom_test:
            # Chercher le titre du test
            titre_cell = soup.find("td", colspan="2", string=lambda t: t and nom_test in t)
            if titre_cell:
                # Trouver la table parente
                table = titre_cell.find_parent("table")
                if table:
                    # Chercher le statut du test
                    status_label = table.find("td", class_="label", string=lambda t: t and "Status:" in t)
                    if status_label:
                        status_value = status_label.find_next_sibling("td", class_="value")
                        if status_value:
                            return status_value.text.strip()
        
        # Pour les labels spécifiques (résistances, tensions, etc.)
        label_cell = soup.find("td", class_="label", string=lambda t: t and nom_test in t)
        if label_cell:
            value_cell = label_cell.find_next_sibling("td", class_="value")
            if value_cell:
                return value_cell.text.strip()
        
        return None
    
    def extraire_numero_serie(self, soup):
        """Extrait le numéro de série depuis la soupe HTML"""
        balise_sn = soup.find("td", class_="hdr_name", string=lambda t: t and "Serial Number:" in t)
        if balise_sn:
            balise_sn_value = balise_sn.find_next_sibling("td", class_="hdr_value")
            if balise_sn_value:
                sn = balise_sn_value.text.strip()
                return sn if sn != "NONE" else None
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
                for test in self.tests_selectionnes:
                    valeur = self.extraire_valeur_test(soup, test)
                    if valeur:
                        # Pour les valeurs numériques, enlever le texte autour
                        if test in donnees[numero_serie]:
                            # Si la donnée existe déjà, prendre la plus récente en supposant que
                            # les fichiers sont traités dans l'ordre chronologique
                            continue
                        donnees[numero_serie][test] = valeur
            
            except Exception as e:
                print(f"Erreur lors de l'analyse du fichier {fichier}: {e}")
        
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