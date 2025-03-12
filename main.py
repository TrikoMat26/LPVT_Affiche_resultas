import os
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup

def analyser_fichier_html(chemin_fichier):
    """
    Analyse un rapport HTML (fichier TestStand/NI) pour en extraire :
     - Le résultat global du test
     - Les tests en échec (status = Failed ou autre couleur d'erreur)
    
    Retourne un dictionnaire contenant :
      {
        "nom_fichier": str,
        "resultat_global": str,  # ex: "Terminated", "Passed", ...
        "tests_echec": [
            {
                "nom_test": str,
                "detail": str  # Texte multi-ligne décrivant le test en défaut
            },
            ...
        ]
      }
    """
    donnees = {
        "nom_fichier": os.path.basename(chemin_fichier),
        "resultat_global": None,
        "tests_echec": []
    }

    with open(chemin_fichier, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f, "html.parser")

    # 1) Récupérer le résultat global (UUT Result)
    #    Recherchons la ligne <td class='hdr_name'><b>UUT Result: </b></td> 
    #    puis son voisin <td class='hdr_value'><b><span style="color:#000080;">Terminated</span></b></td>
    balise_resultat = soup.find("td", class_="hdr_name", text=lambda t: t and "UUT Result:" in t)
    if balise_resultat:
        balise_valeur = balise_resultat.find_next_sibling("td", class_="hdr_value")
        if balise_valeur:
            donnees["resultat_global"] = balise_valeur.get_text(strip=True)

    # 2) Parcourir tous les "tests" potentiels : 
    #    On repère chaque bloc où le "Status:" est spécifié,
    #    et si on voit "Failed" ou "Terminated" (ou d'autres statuts d'échec),
    #    on stocke les informations pertinentes.
    #    Dans les rapports, chaque test est souvent délimité par un <tr> "colspan='2'" 
    #    décrivant le nom du test, suivi d'un <tr> décrivant "Status:" etc.

    #    Pour rester simple, on va chercher tous les td.label qui contiennent "Status:",
    #    puis on lit la cellule voisine (td.value). 
    #    Si c'est "Failed" (ou si la couleur est rouge, etc.), on considère que c'est un échec.
    #    Ensuite, on essaie de récupérer le "nom du test" dans la ligne précédente (ou la précédente "colspan=2").
    
    # Liste de mots-clés considérés comme échec :
    statuts_echec = {"Failed", "Terminated"}  # Vous pouvez en ajouter si besoin

    for balise_status_label in soup.find_all("td", class_="label"):
        if "Status:" in balise_status_label.get_text(strip=True):
            # On récupère la cellule voisine
            balise_status_value = balise_status_label.find_next_sibling("td", class_="value")
            if balise_status_value:
                status_text = balise_status_value.get_text(strip=True)
                if status_text in statuts_echec:
                    # On considère qu'on est dans un test en défaut
                    # 2.1) Récupérer le nom du test dans la ligne précédente
                    #      souvent c'est un <tr> colspan='2' style=... qui annonce le test
                    nom_test = "Nom de test inconnu"
                    previous_tr = balise_status_label.find_parent("tr").find_previous_sibling("tr")
                    if previous_tr:
                        # Cherchons éventuellement un <td colspan='2'> 
                        td_colspan = previous_tr.find("td", colspan="2")
                        if td_colspan:
                            nom_test = td_colspan.get_text(strip=True)

                    # 2.2) Récupérer les détails de mesure (Valeur réelle injectée, etc.)
                    #      Dans le rapport fourni, ces infos se situent généralement 
                    #      dans les <td class='label'> / <td class='value'> qui suivent 
                    #      après "Status: ...".  Pour un exemple rapide, on va juste
                    #      extraire quelques lignes de contiguïté.
                    details_test = []
                    # On prend quelques <tr> suivants, jusqu’à trouver un nouveau test ou rien
                    # Pour être plus précis, vous pouvez affiner la logique
                    next_tr = balise_status_label.find_parent("tr").find_next_sibling("tr")
                    while next_tr:
                        # Arrêter si on tombe sur "Status:" => c'est un nouveau test
                        candidate_label = next_tr.find("td", class_="label")
                        if candidate_label and "Status:" in candidate_label.get_text(strip=True):
                            break

                        details_test.append(next_tr.get_text(" ", strip=True))
                        next_tr = next_tr.find_next_sibling("tr")

                    # Conversion en texte lisible
                    detail_texte = "\n".join(details_test)

                    # Ajout à la liste des échecs
                    donnees["tests_echec"].append({
                        "nom_test": nom_test,
                        "detail": detail_texte,
                        "status": status_text
                    })

    return donnees

def main():
    # 1) Ouvre une boîte de dialogue pour choisir un répertoire
    root = tk.Tk()
    root.withdraw()  # On masque la fenêtre principale
    repertoire = filedialog.askdirectory(title="Choisissez un répertoire contenant les rapports HTML")
    if not repertoire:
        print("Aucun répertoire sélectionné, fin du script.")
        return

    # 2) Liste tous les fichiers .html du répertoire
    html_files = [f for f in os.listdir(repertoire) if f.lower().endswith(".html")]

    if not html_files:
        print("Aucun fichier HTML trouvé dans ce répertoire.")
        return

    # 3) Analyse et affichage
    for nom_fic in html_files:
        chemin_complet = os.path.join(repertoire, nom_fic)
        resultats = analyser_fichier_html(chemin_complet)

        print(f"\nVoici les tests en défaut dans le fichier {resultats['nom_fichier']} :\n")

        # Affiche le résultat global (s'il existe)
        if resultats["resultat_global"]:
            print(f"1. Résultat global du test : \"{resultats['resultat_global']}\"")
            if resultats["resultat_global"].lower() == "terminated":
                print("   Le test global a été interrompu avant d'être complété.")
        else:
            print("1. Résultat global du test : inconnu")

        # Affiche les tests en échec
        if not resultats["tests_echec"]:
            print("\n2. Aucun test en échec.\n")
        else:
            print("\n2. Tests en échec :")
            for idx, test_fail in enumerate(resultats["tests_echec"], start=1):
                status_str = f"(rouge)" if test_fail["status"] == "Failed" else f"({test_fail['status'].lower()})"
                print(f"   {idx}) {test_fail['nom_test']}")
                print(f"      Statut : {test_fail['status']} {status_str}")
                # On affiche le détail
                if test_fail["detail"]:
                    lignes = test_fail["detail"].split("\n")
                    for ligne in lignes:
                        print(f"      {ligne}")
                print("")

if __name__ == "__main__":
    main()
