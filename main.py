import os
import subprocess
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import sys
import re
from typing import Tuple, Dict, List, Optional, Any

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
    
    return None  # Aucune correspondance trouvée

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
    
    # Chercher les labels contenant des informations sur les résistances
    for label in soup.find_all("td", class_="label"):
        label_text = label.get_text(strip=True)
        
        if "Résistance R46 calculée" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                resistances_infos["R46_calculee"] = value.get_text(strip=True)
        
        elif "Résistance R47 calculée" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                resistances_infos["R47_calculee"] = value.get_text(strip=True)
        
        elif "Résistance R48 calculée" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                resistances_infos["R48_calculee"] = value.get_text(strip=True)
        
        elif "Résistance R46 à monter" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                valeur_texte = value.get_text(strip=True)
                resistances_infos["R46_monter"] = valeur_texte
                resistances_infos["R46_marquage"] = obtenir_marquage_cms(valeur_texte)
        
        elif "Résistance R47 à monter" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                valeur_texte = value.get_text(strip=True)
                resistances_infos["R47_monter"] = valeur_texte
                resistances_infos["R47_marquage"] = obtenir_marquage_cms(valeur_texte)
        
        elif "Résistance R48 à monter" in label_text:
            value = label.find_next_sibling("td", class_="value")
            if value:
                valeur_texte = value.get_text(strip=True)
                resistances_infos["R48_monter"] = valeur_texte
                resistances_infos["R48_marquage"] = obtenir_marquage_cms(valeur_texte)
    
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
        "resistances": {}
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
        statuts_echec = {"Failed", "Terminated"}
        
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

def generer_contenu_rapport(resultats: Dict[str, Any], rapport_final: List[str]) -> None:
    """
    Génère le contenu du rapport pour un fichier analysé
    Modifie la liste rapport_final en place
    """
    statut = resultats["resultat_global"] or "Inconnu"
    rapport_final.append(f"{resultats['nom_fichier']} : {statut}")
    
    # Si c'est une séquence SEQ-01 et qu'il y a des informations sur les résistances
    if resultats["sequence"] == "SEQ-01" and resultats["resistances"]:
        resistances = resultats["resistances"]
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
    
    # Ajouter les détails des tests en échec avec formatage amélioré
    if resultats["tests_echec"]:
        rapport_final.append("  Tests en échec:")
        for test in resultats["tests_echec"]:
            rapport_final.append(f"    - {test['nom_test']} ({test['status']})")
            
            # Formater les détails avec une indentation propre
            for ligne in test['detail'].split('\n'):
                if ligne.strip():  # Ne pas ajouter les lignes vides
                    rapport_final.append(f"      * {ligne}")
            
            # Ajouter une ligne vide entre les tests
            rapport_final.append("")

def traiter_repertoire_serie(chemin_repertoire: str) -> None:
    """
    Traite un répertoire de numéro de série et crée un rapport dans ce répertoire
    """
    numero_serie = os.path.basename(chemin_repertoire)
    print(f"Traitement du numéro de série: {numero_serie}")
    
    # Trouver tous les fichiers HTML dans ce répertoire
    chemins_html = trouver_fichiers_html(chemin_repertoire)
    
    if not chemins_html:
        print(f"Aucun fichier HTML trouvé pour le numéro de série {numero_serie}")
        return
    
    # Générer le rapport pour ce numéro de série
    rapport_final = [
        f"Numéro de série : {numero_serie}",
        "-" * 70  # Séparateur
    ]
    
    for chemin_complet in chemins_html:
        try:
            resultats = analyser_fichier_html(chemin_complet)
            generer_contenu_rapport(resultats, rapport_final)
        except Exception as e:
            print(f"Erreur lors du traitement de {chemin_complet}: {e}")
    
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
    
    # Traiter chaque répertoire de numéro de série indépendamment
    for repertoire in sous_repertoires:
        traiter_repertoire_serie(repertoire)
    
    print("Traitement terminé.")

if __name__ == "__main__":
    main()
