import os
import subprocess
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup

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

    # Les rapports indiquent souvent charset=iso-8859-1 ; on peut donc tester cet encodage
    # Ajustez au besoin (cp1252, utf-8, etc.) si les accents ne s'affichent pas correctement
    with open(chemin_fichier, "r", encoding="iso-8859-1", errors="replace") as f:
        soup = BeautifulSoup(f, "html.parser")

    # 1) Récupérer le "UUT Result" (résultat global)
    balise_resultat = soup.find("td", class_="hdr_name", string=lambda t: t and "UUT Result:" in t)
    if balise_resultat:
        balise_valeur = balise_resultat.find_next_sibling("td", class_="hdr_value")
        if balise_valeur:
            donnees["resultat_global"] = balise_valeur.get_text(strip=True)

    # 2) Récupérer les tests en échec (statuts = Failed ou Terminated)
    statuts_echec = {"Failed", "Terminated"}
    for balise_status_label in soup.find_all("td", class_="label"):
        if "Status:" in balise_status_label.get_text(strip=True):
            balise_status_value = balise_status_label.find_next_sibling("td", class_="value")
            if balise_status_value:
                status_text = balise_status_value.get_text(strip=True)
                if status_text in statuts_echec:
                    # On considère ce bloc comme un échec
                    nom_test = "Nom de test inconnu"

                    # Essayer de récupérer le nom du test dans la ligne précédente (ou un <td colspan='2'>)
                    previous_tr = balise_status_label.find_parent("tr").find_previous_sibling("tr")
                    if previous_tr:
                        td_colspan = previous_tr.find("td", colspan="2")
                        if td_colspan:
                            nom_test = td_colspan.get_text(strip=True)

                    # Extraire plus de détails (ex: Valeur réelle injectée, Valeur sortie, etc.)
                    details_test = []
                    next_tr = balise_status_label.find_parent("tr").find_next_sibling("tr")
                    while next_tr:
                        # Si on retombe sur un nouveau "Status:", on arrête
                        candidate_label = next_tr.find("td", class_="label")
                        if candidate_label and "Status:" in candidate_label.get_text(strip=True):
                            break
                        details_test.append(next_tr.get_text(" ", strip=True))
                        next_tr = next_tr.find_next_sibling("tr")

                    detail_texte = "\n".join(details_test)

                    donnees["tests_echec"].append({
                        "nom_test": nom_test,
                        "detail": detail_texte,
                        "status": status_text
                    })

    return donnees

def main():
    root = tk.Tk()
    root.withdraw()
    repertoire = filedialog.askdirectory(title="Choisissez un répertoire contenant les rapports HTML")
    if not repertoire:
        print("Aucun répertoire sélectionné, fin du script.")
        return

    # On récupère le nom du répertoire pour l'utiliser comme numéro de série
    numero_serie = os.path.basename(repertoire)
    
    # Liste tous les fichiers HTML dans le répertoire
    html_files = [f for f in os.listdir(repertoire) if f.lower().endswith(".html")]
    if not html_files:
        print("Aucun fichier HTML trouvé dans ce répertoire.")
        return

    # Prépare le rapport dans une liste de chaînes
    rapport_final = []
    
    # Ajouter une ligne pour le numéro de série
    rapport_final.append(f"Numéro de série : {numero_serie}")
    rapport_final.append("-" * 70)  # Séparateur entre fichiers

    for nom_fic in html_files:
        chemin_complet = os.path.join(repertoire, nom_fic)
        resultats = analyser_fichier_html(chemin_complet)

        rapport_final.append(f"{resultats['nom_fichier']} :\n")

        # Résultat global
        if resultats["resultat_global"]:
            rapport_final.append(f"1. Résultat global du test : \"{resultats['resultat_global']}\"")
            if resultats["resultat_global"].lower() == "terminated":
                rapport_final.append("   Le test global a été interrompu avant d'être complété.")
        else:
            rapport_final.append("1. Résultat global du test : inconnu")

        # Tests en échec
        if not resultats["tests_echec"]:
            rapport_final.append("\n2. Aucun test en échec.\n")
        else:
            rapport_final.append("\n2. Tests en échec :")
            for idx, test_fail in enumerate(resultats["tests_echec"], start=1):
                # Indiquer (rouge) si c'est Failed, ou préciser (terminated) si Terminated
                if test_fail["status"] == "Failed":
                    status_str = "(rouge)"
                else:
                    status_str = f"({test_fail['status'].lower()})"

                rapport_final.append(f"   {idx}) {test_fail['nom_test']}")
                rapport_final.append(f"      Statut : {test_fail['status']} {status_str}")
                if test_fail["detail"]:
                    lignes = test_fail["detail"].split("\n")
                    for ligne in lignes:
                        rapport_final.append(f"      {ligne}")
                rapport_final.append("")

        rapport_final.append("-" * 70)  # Séparateur entre fichiers

    # Nom du fichier = numéro de série + "_resultats_tests.txt"
    fichier_resultat = os.path.join(repertoire, f"{numero_serie}_LPVT.txt")

    # On écrit le rapport dans ce fichier
    with open(fichier_resultat, "w", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(rapport_final))

    # On ouvre le fichier texte dans le Bloc-notes sans attendre sa fermeture
    subprocess.Popen(["notepad", fichier_resultat])

if __name__ == "__main__":
    main()
