import os
import subprocess
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import sys

def formater_nom_fichier(nom_fichier):
    """
    Reformate le nom du fichier selon le format désiré
    De: SEQ-01_LPVT_Report[15 35 54][27 01 2025].html
    Vers: SEQ-01_[27 01 2025] [15 35 54]
    """
    # Enlever l'extension .html
    nom_sans_ext = os.path.splitext(nom_fichier)[0]
    
    # Extraire les parties du nom
    parties = nom_sans_ext.split('_')
    if len(parties) >= 3:
        sequence = parties[0]  # SEQ-01
        # Extraire la date et l'heure entre crochets
        import re
        horodatage = parties[-1]  # Report[15 35 54][27 01 2025]
        matches = re.findall(r'\[(.*?)\]', horodatage)
        if len(matches) >= 2:
            heure = matches[0]
            date = matches[1]
            return f"{sequence} [{date}] [{heure}]"
    
    # Si le format n'est pas celui attendu, retourner le nom original
    return nom_fichier

def analyser_fichier_html(chemin_fichier):
    """
    Analyse un rapport HTML et renvoie un dictionnaire contenant :
      - nom_fichier : le nom du fichier
      - resultat_global : le statut global (Passed, Failed, Terminated, etc.)
      - tests_echec : une liste de tests en échec (chacun étant un dict avec nom_test, detail, status)
    """
    donnees = {
        "nom_fichier": formater_nom_fichier(os.path.basename(chemin_fichier)),
        "resultat_global": None,
        "tests_echec": []
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
        
        # 3) Récupérer les tests en échec (statuts = Failed ou Terminated)
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
            
            # Collecter tous les détails du test
            details = []
            
            # 1. Récupérer tous les labels et leurs valeurs associées dans cette table
            for label in table.find_all("td", class_="label"):
                if "Status:" not in label.get_text(strip=True):  # Éviter de répéter le statut
                    value = label.find_next_sibling("td", class_="value")
                    if value:
                        label_text = label.get_text(strip=True)
                        value_text = value.get_text(strip=True)
                        if value_text:  # Ne pas ajouter les champs vides
                            details.append(f"{label_text} {value_text}")
            
            # 2. Chercher spécifiquement les informations ROUE CODEUSE et Valeur
            roue_codeuse = None
            valeur_mesure = None
            
            for label in table.find_all("td", class_="label"):
                label_text = label.get_text(strip=True)
                if "ROUE CODEUSE" in label_text:
                    value = label.find_next_sibling("td", class_="value")
                    if value:
                        roue_codeuse = f"{label_text} {value.get_text(strip=True)}"
                        if roue_codeuse not in details:
                            details.append(roue_codeuse)
                
                elif "Valeur" in label_text:
                    value = label.find_next_sibling("td", class_="value")
                    if value:
                        valeur_mesure = f"{label_text} {value.get_text(strip=True)}"
                        if valeur_mesure not in details:
                            details.append(valeur_mesure)
            
            # 3. Chercher les messages d'erreur spécifiques
            error_label = table.find("td", class_="label", string=lambda t: t and "Error Info:" in t)
            if error_label:
                error_value = error_label.find_next_sibling("td", class_="value")
                if error_value:
                    details.append(f"Erreur: {error_value.get_text(strip=True)}")
            
            # Construire la chaîne de détails finale
            detail_text = "\n".join(details) if details else "Pas de détail disponible"
            
            # Si le test n'est pas déjà dans la liste, l'ajouter
            test_existant = False
            for test in donnees["tests_echec"]:
                if test["nom_test"] == nom_test and test["status"] == status_text:
                    test_existant = True
                    break
                    
            if not test_existant:
                donnees["tests_echec"].append({
                    "nom_test": nom_test,
                    "status": status_text,
                    "detail": detail_text
                })
                
    except Exception as e:
        print(f"Erreur lors de l'analyse du fichier {chemin_fichier}: {e}")
    
    return donnees

def trouver_fichiers_html(repertoire):
    """
    Trouve tous les fichiers HTML dans le répertoire spécifié (non récursif)
    """
    fichiers_html = []
    
    try:
        for fichier in os.listdir(repertoire):
            if fichier.lower().endswith('.html'):
                chemin_complet = os.path.join(repertoire, fichier)
                fichiers_html.append(chemin_complet)
        print(f"Trouvé {len(fichiers_html)} fichiers HTML dans {repertoire}")
    except Exception as e:
        print(f"Erreur lors de la recherche de fichiers dans {repertoire}: {e}")
    
    return fichiers_html

def traiter_repertoire_serie(chemin_repertoire):
    """
    Traite un répertoire de numéro de série et crée un rapport dans ce répertoire
    """
    numero_serie = os.path.basename(chemin_repertoire)
    print(f"Traitement du numéro de série: {numero_serie}")
    
    # Trouver tous les fichiers HTML dans ce répertoire (non récursif)
    chemins_html = trouver_fichiers_html(chemin_repertoire)
    
    if not chemins_html:
        print(f"Aucun fichier HTML trouvé pour le numéro de série {numero_serie}")
        return
    
    # Générer le rapport pour ce numéro de série
    rapport_final = []
    rapport_final.append(f"Numéro de série : {numero_serie}")
    rapport_final.append("-" * 70)  # Séparateur
    
    for chemin_complet in chemins_html:
        try:
            resultats = analyser_fichier_html(chemin_complet)
            statut = resultats["resultat_global"] or "Inconnu"
            rapport_final.append(f"{resultats['nom_fichier']} : {statut}")
            
            # Ajouter les détails des tests en échec avec formatage amélioré
            if resultats["tests_echec"]:
                rapport_final.append("  Tests en échec:")
                for test in resultats["tests_echec"]:
                    rapport_final.append(f"    - {test['nom_test']} ({test['status']})")
                    
                    # Formater les détails avec une indentation propre
                    for ligne in test['detail'].split('\n'):
                        if ligne.strip():  # Ne pas ajouter les lignes vides
                            # Ajouter une indentation pour une meilleure lisibilité
                            rapport_final.append(f"      * {ligne}")
                    
                    # Ajouter une ligne vide entre les tests
                    rapport_final.append("")
            
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

def main():
    # Configuration pour l'affichage des caractères accentués
    if sys.stdout.encoding != 'utf-8':
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
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
