import os
import subprocess
import tkinter as tk
from tkinter import filedialog, Label, Toplevel
from bs4 import BeautifulSoup
import sys
import re
from typing import Tuple, Dict, List, Optional, Any, Set

def formater_nom_fichier(nom_fichier: str) -> Tuple[str, str]:
    """
    Reformate le nom du fichier selon le format désiré
    De: SEQ-01_LPVT_Report[15 35 54][27 01 2025].html
    Vers: SEQ-01_[27 01 2025] [15 35 54]
    
    Retourne un tuple: (nom_formaté, sequence)
    """
    # Enlever l'extension .html
    nom_sans_ext = os.path.splitext(nom_fichier)[0]
    
    # Extraire les parties du nom
    parties = nom_sans_ext.split('_')
    sequence = parties[0] if len(parties) > 0 else ""  # SEQ-01
    
    if len(parties) >= 3:
        # Extraire la date et l'heure entre crochets
        horodatage = parties[-1]  # Report[15 35 54][27 01 2025]
        matches = re.findall(r'\[(.*?)\]', horodatage)
        if len(matches) >= 2:
            heure = matches[0]
            date = matches[1]
            return f"{sequence} [{date}] [{heure}]", sequence
    
    # Si le format n'est pas celui attendu, retourner le nom original
    return nom_fichier, sequence

def obtenir_marquage_cms(valeur_resistance: str) -> Optional[str]:
    """
    Retourne le marquage CMS correspondant à la valeur de résistance
    """
    # Tableau de correspondance entre les valeurs de résistance et les marquages CMS
    correspondance = {
        "221": "34A",
        "332": "3320",
        "475": "66A",
        "681": "81A",
        "1000": "01B"
    }
    
    # Extraire la valeur numérique de la chaîne (par exemple: "Résistance à monter = 475 ohms." -> "475")
    match = re.search(r'(\d+)\s*ohms', valeur_resistance, re.IGNORECASE)
    if match:
        valeur = match.group(1)
        return correspondance.get(valeur)
    
    return None

def lire_valeur_html(element: BeautifulSoup, classe: str, texte_recherche: str) -> Optional[str]:
    """
    Lit une valeur à partir d'un élément HTML en recherchant un texte spécifique.
    Retourne la valeur si trouvée, sinon None.
    """
    element_trouve = element.find("td", class_=classe, string=lambda t: t and texte_recherche in t)
    if element_trouve:
        element_valeur = element_trouve.find_next_sibling("td", class_=f"{classe}_value")
        if element_valeur:
            return element_valeur.text.strip()
    return None

def extraire_informations_resistances(soup: BeautifulSoup) -> Dict[str, Any]:
    """
    Extrait les informations sur les résistances du rapport HTML
    """
    resistances_infos = {
        "R46_calculee": None,
        "R47_calculee": None,
        "R48_calculee": None,
        "R46_monter": None,
        "R47_monter": None,
        "R48_monter": None,
        "R46_marquage": None,
        "R47_marquage": None,
        "R48_marquage": None
    }
    
    # Pattern de recherche pour les résistances
    patterns = {
        "calculee": {
            "R46": "Résistance R46 calculée",
            "R47": "Résistance R47 calculée",
            "R48": "Résistance R48 calculée"
        },
        "monter": {
            "R46": "Résistance R46 à monter",
            "R47": "Résistance R47 à monter",
            "R48": "Résistance R48 à monter"
        }
    }
    
    # Chercher les labels contenant des informations sur les résistances
    for label in soup.find_all("td", class_="label"):
        label_text = label.get_text(strip=True)
        
        # Traiter les résistances calculées
        for r_name, pattern in patterns["calculee"].items():
            if pattern in label_text:
                value = label.find_next_sibling("td", class_="value")
                if value:
                    resistances_infos[f"{r_name}_calculee"] = value.get_text(strip=True)
        
        # Traiter les résistances à monter
        for r_name, pattern in patterns["monter"].items():
            if pattern in label_text:
                value = label.find_next_sibling("td", class_="value")
                if value:
                    valeur_texte = value.get_text(strip=True)
                    resistances_infos[f"{r_name}_monter"] = valeur_texte
                    resistances_infos[f"{r_name}_marquage"] = obtenir_marquage_cms(valeur_texte)
    
    return resistances_infos

def extraire_details_test(table: BeautifulSoup) -> Tuple[List[str], Dict[str, str], Dict[str, str]]:
    """
    Extrait les détails d'un test à partir d'une table HTML
    Retourne (autres_details, roue_codeuse_info, valeurs_info)
    """
    roue_codeuse_info = {}
    valeurs_info = {}
    autres_details = []
    
    for label in table.find_all("td", class_="label"):
        label_text = label.get_text(strip=True)
        
        if "Status:" in label_text:
            continue  # Éviter de répéter le statut
        
        value = label.find_next_sibling("td", class_="value")
        if not value:
            continue
            
        value_text = value.get_text(strip=True)
        if not value_text:
            continue  # Ne pas ajouter les champs vides
        
        # Traiter différemment selon le type d'information
        if "ROUE CODEUSE" in label_text:
            # Extraire les informations de la roue codeuse
            rc_match = re.search(r"ROUE CODEUSE = ([^ \.]+) \.\. GAMME = ([^ \.]+) \.\. Voie ([^ \.]+)", label_text)
            if rc_match:
                roue = rc_match.group(1)
                gamme = rc_match.group(2)
                voie = rc_match.group(3)
                roue_codeuse_info = {
                    "roue": roue,
                    "gamme": gamme,
                    "voie": voie
                }
            else:
                autres_details.append(f"{label_text} {value_text}")
        
        elif "Valeur" in label_text and "injectée" in label_text:
            # Extraire les valeurs mesurées
            valeurs_texte = value_text.strip()
            parties = valeurs_texte.split('/')
            if len(parties) >= 4:
                valeurs_info = {
                    "injectée": parties[0].strip(),
                    "attendue": parties[1].strip(),
                    "mesurée": parties[2].strip(),
                    "précision": parties[3].strip()
                }
            else:
                autres_details.append(f"{label_text} {value_text}")
        
        else:
            # Autres types d'informations
            autres_details.append(f"{label_text} {value_text}")
    
    # Chercher les messages d'erreur spécifiques
    error_label = table.find("td", class_="label", string=lambda t: t and "Error Info:" in t)
    if error_label:
        error_value = error_label.find_next_sibling("td", class_="value")
        if error_value and error_value.get_text(strip=True):
            autres_details.append(f"Erreur: {error_value.get_text(strip=True)}")
    
    return autres_details, roue_codeuse_info, valeurs_info

def formater_details_test(autres_details: List[str], roue_codeuse_info: Dict[str, str], 
                          valeurs_info: Dict[str, str]) -> str:
    """
    Formate les détails du test en une chaîne formatée
    """
    formatted_details = []
    
    # 1. Ajouter en premier les infos générales
    formatted_details.extend(autres_details)
    
    # 2. Ajouter les infos de roue codeuse avec un format amélioré
    if roue_codeuse_info:
        formatted_details.append(f"Configuration: ROUE CODEUSE={roue_codeuse_info.get('roue', '?')} | GAMME={roue_codeuse_info.get('gamme', '?')} | Voie={roue_codeuse_info.get('voie', '?')}")
    
    # 3. Ajouter les valeurs mesurées avec un format amélioré
    if valeurs_info:
        formatted_details.append("Mesures:")
        formatted_details.append(f"  - Valeur injectée : {valeurs_info.get('injectée', '?')}")
        formatted_details.append(f"  - Valeur attendue : {valeurs_info.get('attendue', '?')}")
        formatted_details.append(f"  - Valeur mesurée  : {valeurs_info.get('mesurée', '?')}")
        formatted_details.append(f"  - Précision       : {valeurs_info.get('précision', '?')}")
    
    # Construire la chaîne de détails finale
    return "\n".join(formatted_details) if formatted_details else "Pas de détail disponible"

def analyser_fichier_html(chemin_fichier: str) -> Dict[str, Any]:
    """
    Analyse un rapport HTML et renvoie un dictionnaire contenant les informations pertinentes
    """
    nom_fichier_base = os.path.basename(chemin_fichier)
    nom_format, sequence = formater_nom_fichier(nom_fichier_base)
    
    donnees = {
        "nom_fichier": nom_format,
        "resultat_global": None,
        "tests_echec": [],
        "sequence": sequence,
        "resistances": {},
        "numero_serie": None
    }

    try:
        # Les rapports indiquent souvent charset=iso-8859-1
        with open(chemin_fichier, "r", encoding="iso-8859-1", errors="replace") as f:
            html_content = f.read()
            soup = BeautifulSoup(html_content, "html.parser")

        # 1) Récupérer le "UUT Result" (résultat global)
        balise_resultat = soup.find("td", class_="hdr_name", string=lambda t: t and "UUT Result:" in t)
        if balise_resultat:
            balise_valeur = balise_resultat.find_next_sibling("td", class_="hdr_value")
            if balise_valeur:
                donnees["resultat_global"] = balise_valeur.text.strip()
        
        # 2) Récupérer le numéro de série si disponible
        balise_sn = soup.find("td", class_="hdr_name", string=lambda t: t and "Serial Number:" in t)
        if balise_sn:
            balise_sn_value = balise_sn.find_next_sibling("td", class_="hdr_value")
            if balise_sn_value:
                donnees["numero_serie"] = balise_sn_value.text.strip()
        
        # 3) Si c'est SEQ-01, récupérer les informations sur les résistances
        if sequence == "SEQ-01":
            donnees["resistances"] = extraire_informations_resistances(soup)
        
        # 4) Récupérer les tests en échec (statuts = Failed ou Terminated)
        statuts_echec: Set[str] = {"Failed", "Terminated"}
        
        # Recherche directe des cellules avec statut Failed ou Terminated
        for status_cell in soup.find_all("td", class_="value", string=lambda t: t and t.strip() in statuts_echec):
            status_text = status_cell.text.strip()
            
            # Trouver la table parente qui contient ce statut
            table = status_cell.find_parent("table")
            if not table:
                continue
                
            # Récupérer le titre du test (dans un colspan="2")
            titre_cell = table.find("td", colspan="2")
            if not titre_cell:
                continue
                
            nom_test = titre_cell.get_text(strip=True)
            
            # Extraire les détails du test
            autres_details, roue_codeuse_info, valeurs_info = extraire_details_test(table)
            
            # Formater les détails
            detail_text = formater_details_test(autres_details, roue_codeuse_info, valeurs_info)
            
            # Si le test n'est pas déjà dans la liste, l'ajouter
            if not any(test["nom_test"] == nom_test and test["status"] == status_text for test in donnees["tests_echec"]):
                donnees["tests_echec"].append({
                    "nom_test": nom_test,
                    "status": status_text,
                    "detail": detail_text
                })
                
    except Exception as e:
        print(f"Erreur lors de l'analyse du fichier {chemin_fichier}: {e}")
    
    return donnees

def trouver_fichiers_html(repertoire: str) -> List[str]:
    """
    Trouve tous les fichiers HTML dans le répertoire spécifié (non récursif)
    """
    try:
        fichiers_html = [os.path.join(repertoire, f) for f in os.listdir(repertoire) 
                         if f.lower().endswith('.html')]
        print(f"Trouvé {len(fichiers_html)} fichiers HTML dans {repertoire}")
        return fichiers_html
    except Exception as e:
        print(f"Erreur lors de la recherche de fichiers dans {repertoire}: {e}")
        return []

def formater_section_resistances(resistances: Dict[str, Any], rapport_final: List[str]) -> None:
    """
    Formate et ajoute la section des résistances au rapport
    """
    rapport_final.append("  Informations sur les résistances:")
    
    for r_num in ["R46", "R47", "R48"]:
        if resistances.get(f"{r_num}_calculee"):
            rapport_final.append(f"    - Résistance {r_num} calculée: {resistances[f'{r_num}_calculee']}")
            
    rapport_final.append("  Résistances à monter:")
    for r_num in ["R46", "R47", "R48"]:
        if resistances.get(f"{r_num}_monter"):
            marquage = f" (Marquage CMS: {resistances[f'{r_num}_marquage']})" if resistances.get(f"{r_num}_marquage") else ""
            rapport_final.append(f"    - {r_num}: {resistances[f'{r_num}_monter']}{marquage}")
            
    rapport_final.append("")  # Ligne vide pour séparer

def formater_section_tests_echec(tests_echec: List[Dict[str, Any]], rapport_final: List[str]) -> None:
    """
    Formate et ajoute la section des tests en échec au rapport
    """
    rapport_final.append("  Tests en échec:")
    for test in tests_echec:
        rapport_final.append(f"    - {test['nom_test']} ({test['status']})")
        
        # Formater les détails avec une indentation propre
        for ligne in test['detail'].split('\n'):
            if ligne.strip():  # Ne pas ajouter les lignes vides
                rapport_final.append(f"      * {ligne}")
        
        # Ajouter une ligne vide entre les tests
        rapport_final.append("")

def generer_contenu_rapport(resultats: Dict[str, Any], rapport_final: List[str]) -> None:
    """
    Génère le contenu du rapport pour un fichier analysé
    Modifie la liste rapport_final en place
    """
    statut = resultats["resultat_global"] or "Inconnu"
    rapport_final.append(f"{resultats['nom_fichier']} : {statut}")
    
    # Si c'est une séquence SEQ-01 et qu'il y a des informations sur les résistances
    if resultats["sequence"] == "SEQ-01" and resultats["resistances"]:
        formater_section_resistances(resultats["resistances"], rapport_final)
    
    # Ajouter les détails des tests en échec avec formatage amélioré
    if resultats["tests_echec"]:
        formater_section_tests_echec(resultats["tests_echec"], rapport_final)

def extraire_date_heure(nom_formaté: str) -> tuple:
    """
    Extrait la date et l'heure d'un nom de fichier formaté pour le tri chronologique
    Format attendu: "SEQ-XX [JJ MM AAAA] [HH MM SS]"
    Retourne un tuple permettant le tri chronologique: (AAAA, MM, JJ, HH, MM, SS)
    """
    try:
        # Extraire la partie date [JJ MM AAAA]
        match_date = re.search(r'\[(\d+) (\d+) (\d+)\]', nom_formaté)
        if not match_date:
            return (0, 0, 0, 0, 0, 0)  # Valeur par défaut si le format ne correspond pas
            
        jour, mois, annee = map(int, match_date.groups())
        
        # Extraire la partie heure [HH MM SS]
        match_heure = re.search(r'\[(\d+) (\d+) (\d+)\](?!.*\[\d+ \d+ \d+\])', nom_formaté)
        if not match_heure:
            return (annee, mois, jour, 0, 0, 0)
            
        heure, minute, seconde = map(int, match_heure.groups())
        
        return (annee, mois, jour, heure, minute, seconde)
    except:
        # En cas d'erreur, renvoyer une valeur qui sera en fin de tri
        return (0, 0, 0, 0, 0, 0)

class ProgressWindow:
    """
    Fenêtre non modale pour afficher la progression du traitement
    """
    def __init__(self, total_files: int):
        """Initialise la fenêtre de progression avec le nombre total de fichiers"""
        self.total_files = total_files
        self.processed_files = 0
        
        # Créer une fenêtre Toplevel
        self.window = Toplevel()
        self.window.title("Traitement en cours")
        
        # Définir les dimensions et la position
        window_width = 300
        window_height = 100
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # Empêcher le redimensionnement
        self.window.resizable(False, False)
        
        # S'assurer que la fenêtre reste au-dessus des autres
        self.window.attributes("-topmost", True)
        
        # Créer un frame pour organiser les widgets
        self.frame = tk.Frame(self.window)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Label pour afficher l'information
        self.info_label = Label(
            self.frame, 
            text=f"Traitement en cours...\n\nFichiers restants: {self.total_files - self.processed_files} / {self.total_files}",
            font=("Arial", 10),
            justify="center"
        )
        self.info_label.pack(expand=True)
        
        # Mettre à jour l'interface
        self.window.update()
    
    def update_progress(self) -> None:
        """Met à jour la progression après le traitement d'un fichier"""
        self.processed_files += 1
        remaining = self.total_files - self.processed_files
        
        # Mettre à jour le texte
        self.info_label.config(
            text=f"Traitement en cours...\n\nFichiers restants: {remaining} / {self.total_files}"
        )
        
        # Forcer la mise à jour de l'interface
        self.window.update()
    
    def close(self) -> None:
        """Ferme la fenêtre de progression"""
        self.window.destroy()
    
    def show_completion(self) -> None:
        """Affiche un message de fin de traitement avec un bouton pour fermer"""
        # Supprimer le label existant
        self.info_label.pack_forget()
        
        # Créer un nouveau label avec le message de fin
        self.info_label = Label(
            self.frame, 
            text=f"Traitement terminé !\n\n{self.processed_files} fichiers ont été traités.",
            font=("Arial", 10, "bold"),
            fg="green",
            justify="center"
        )
        self.info_label.pack(expand=True, pady=(10, 20))
        
        # Ajouter un bouton "Fermer"
        self.close_button = tk.Button(
            self.frame,
            text="Fermer",
            command=self.close,
            width=10
        )
        self.close_button.pack(pady=(0, 10))
        
        # La fenêtre n'est plus en topmost pour ne pas gêner l'utilisateur
        self.window.attributes("-topmost", False)
        
        # Redimensionner la fenêtre pour s'adapter au contenu
        self.window.update_idletasks()
        window_width = 300
        window_height = 150  # Hauteur augmentée pour accommoder le bouton
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

def traiter_repertoire_serie(chemin_repertoire: str, progress_window: Optional[ProgressWindow] = None) -> None:
    """
    Traite un répertoire de numéro de série et crée un rapport dans ce répertoire
    Les séquences sont triées par ordre chronologique dans le rapport
    """
    numero_serie = os.path.basename(chemin_repertoire)
    print(f"Traitement du numéro de série: {numero_serie}")
    
    # Trouver tous les fichiers HTML dans ce répertoire
    chemins_html = trouver_fichiers_html(chemin_repertoire)
    
    if not chemins_html:
        print(f"Aucun fichier HTML trouvé pour le numéro de série {numero_serie}")
        # Mettre à jour la fenêtre de progression si elle existe
        if progress_window:
            progress_window.update_progress()
        return
    
    # Collecter tous les résultats pour pouvoir les trier
    resultats_analyses = []
    
    for chemin_complet in chemins_html:
        try:
            resultats = analyser_fichier_html(chemin_complet)
            resultats_analyses.append(resultats)
        except Exception as e:
            print(f"Erreur lors de l'analyse de {chemin_complet}: {e}")
    
    # Trier les résultats par ordre chronologique (date et heure)
    resultats_analyses.sort(key=lambda x: extraire_date_heure(x['nom_fichier']))
    
    # Générer le rapport pour ce numéro de série
    rapport_final = [
        f"Numéro de série : {numero_serie}",
        "-" * 70  # Séparateur
    ]
    
    # Générer le contenu du rapport dans l'ordre chronologique
    for resultats in resultats_analyses:
        generer_contenu_rapport(resultats, rapport_final)
    
    # Créer le fichier de rapport dans le répertoire du numéro de série
    nom_fichier_rapport = f"rapport_{numero_serie}.txt"
    chemin_fichier_rapport = os.path.join(chemin_repertoire, nom_fichier_rapport)
    
    try:
        with open(chemin_fichier_rapport, 'w', encoding='utf-8') as f:
            f.write('\n'.join(rapport_final))
        print(f"Rapport créé: {chemin_fichier_rapport}")
    
        # Ouvrir le rapport
        subprocess.Popen(['notepad.exe', chemin_fichier_rapport])
    except Exception as e:
        print(f"Erreur lors de la création/ouverture du rapport: {e}")
    
    # Mettre à jour la fenêtre de progression si elle existe
    if progress_window:
        progress_window.update_progress()

def configurer_encodage() -> None:
    """Configure l'encodage pour les caractères accentués"""
    try:
        if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except (AttributeError, TypeError):
        # En cas d'erreur, on continue sans changer l'encodage
        pass

def main() -> None:
    # Configuration pour l'affichage des caractères accentués
    configurer_encodage()
    
    # Configuration de l'interface Tkinter
    root = tk.Tk()
    root.withdraw()
    
    # Sélectionner le répertoire parent qui contient les répertoires de numéros de série
    print("Veuillez sélectionner le répertoire contenant les dossiers de numéros de série...")
    repertoire_parent = filedialog.askdirectory(title="Choisissez le répertoire contenant les dossiers de numéros de série")
    if not repertoire_parent:
        print("Aucun répertoire sélectionné, fin du script.")
        return
    
    print(f"Répertoire sélectionné: {repertoire_parent}")
    
    # Obtenir tous les sous-répertoires directs (les numéros de série)
    try:
        sous_repertoires = [os.path.join(repertoire_parent, d) for d in os.listdir(repertoire_parent) 
                            if os.path.isdir(os.path.join(repertoire_parent, d))]
        print(f"Sous-répertoires trouvés: {len(sous_repertoires)}")
    except Exception as e:
        print(f"Erreur lors de la recherche des sous-répertoires: {e}")
        return
    
    if not sous_repertoires:
        print("Aucun sous-répertoire trouvé.")
        return
    
    # Créer la fenêtre de progression
    progress_window = ProgressWindow(total_files=len(sous_repertoires))
    
    # Traiter chaque répertoire de numéro de série indépendamment
    for repertoire in sous_repertoires:
        traiter_repertoire_serie(repertoire, progress_window)
    
    # Afficher un message de fin
    progress_window.show_completion()
    print("Traitement terminé.")
    
    # La fenêtre reste ouverte, l'utilisateur peut la fermer en cliquant sur le bouton
    # Attendre que l'utilisateur ferme la fenêtre avant de terminer le programme
    if progress_window.window.winfo_exists():
        root.wait_window(progress_window.window)

if __name__ == "__main__":
    main()
