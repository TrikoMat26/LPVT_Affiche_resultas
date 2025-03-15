import os
import subprocess
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import sys
import re
from datetime import datetime
import webbrowser
from collections import defaultdict

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

def extraire_sequence_du_nom(nom_fichier):
    """
    Extrait le numéro de séquence du nom du fichier (SEQ-01, SEQ-02, etc.)
    """
    parties = nom_fichier.split('_')
    if parties and len(parties) > 0:
        return parties[0]
    return ""

def obtenir_marquage_cms(valeur_resistance):
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
        if valeur in correspondance:
            return correspondance[valeur]
    
    return None  # Aucune correspondance trouvée

def analyser_fichier_html(chemin_fichier):
    """
    Analyse un rapport HTML et renvoie un dictionnaire contenant :
      - nom_fichier : le nom du fichier
      - resultat_global : le statut global (Passed, Failed, Terminated, etc.)
      - tests_echec : une liste de tests en échec (chacun étant un dict avec nom_test, detail, status)
      - resistances : informations sur les résistances (pour SEQ-01)
    """
    nom_fichier_base = os.path.basename(chemin_fichier)
    sequence = extraire_sequence_du_nom(nom_fichier_base)
    
    donnees = {
        "nom_fichier": formater_nom_fichier(nom_fichier_base),
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
            resistances_infos = {
                "R46_calculee": None,
                "R47_calculee": None,
                "R48_calculee": None,
                "R46_monter": None,
                "R47_monter": None,
                "R48_monter": None,
                # Ajout des marquages CMS
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
            
            # Ajouter les informations récupérées
            donnees["resistances"] = resistances_infos
        
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
            
            # Collecter tous les détails du test
            details = []
            roue_codeuse_info = {}
            valeurs_info = {}
            autres_details = []
            
            # 1. Récupérer tous les labels et leurs valeurs associées dans cette table
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
                    import re
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
            
            # 3. Chercher les messages d'erreur spécifiques
            error_label = table.find("td", class_="label", string=lambda t: t and "Error Info:" in t)
            if error_label:
                error_value = error_label.find_next_sibling("td", class_="value")
                if error_value and error_value.get_text(strip=True):
                    autres_details.append(f"Erreur: {error_value.get_text(strip=True)}")
            
            # Construire la chaîne de détails avec un format amélioré
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
            detail_text = "\n".join(formatted_details) if formatted_details else "Pas de détail disponible"
            
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

def extraire_date_heure_du_nom(nom_fichier):
    """
    Extrait la date et l'heure du nom du fichier pour le tri chronologique
    """
    matches = re.findall(r'\[(.*?)\]', nom_fichier)
    if len(matches) >= 2:
        heure = matches[0].split()
        date = matches[1].split()
        # Format: jour mois année heure minute seconde
        if len(date) >= 3 and len(heure) >= 3:
            try:
                # Date au format "DD MM YYYY"
                jour, mois, annee = date
                # Heure au format "HH MM SS"
                heures, minutes, secondes = heure
                
                date_heure = f"{annee}-{mois}-{jour} {heures}:{minutes}:{secondes}"
                return datetime.strptime(date_heure, "%Y-%m-%d %H:%M:%S")
            except (ValueError, IndexError):
                pass
    return datetime.min  # Date minimale par défaut

def organiser_defauts_par_type(resultats_tests):
    """
    Organise les défauts par type (roue codeuse, voie, gamme) à travers les différentes séquences de test
    """
    defauts_par_type = defaultdict(list)
    
    # Parcourir tous les résultats de test
    for test_info in resultats_tests:
        nom_fichier = test_info["nom_fichier"]
        date_test = extraire_date_heure_du_nom(nom_fichier)
        
        # Parcourir les tests en échec
        for test_echec in test_info["tests_echec"]:
            nom_test = test_echec["nom_test"]
            status = test_echec["status"]
            detail = test_echec["detail"]
            
            # Identifier le type de défaut (roue codeuse + gamme + voie)
            type_defaut = "Non catégorisé"
            
            # Chercher le pattern de la roue codeuse dans le détail
            rc_match = re.search(r"Configuration: ROUE CODEUSE=([^ |]+) \| GAMME=([^ |]+) \| Voie=([^ |]+)", detail)
            if rc_match:
                roue = rc_match.group(1)
                gamme = rc_match.group(2)
                voie = rc_match.group(3)
                type_defaut = f"ROUE CODEUSE={roue}, GAMME={gamme}, Voie={voie}"
            
            # Ajouter ce défaut à la liste correspondant à ce type
            defauts_par_type[type_defaut].append({
                "nom_fichier": nom_fichier,
                "date_test": date_test,
                "nom_test": nom_test,
                "status": status,
                "detail": detail
            })
    
    # Trier chaque liste de défauts par date de test
    for defaut_type, defauts in defauts_par_type.items():
        defauts_par_type[defaut_type] = sorted(defauts, key=lambda x: x["date_test"])
    
    return defauts_par_type

def generer_rapport_html(chemin_repertoire, resultats_tests):
    """
    Génère un rapport HTML interactif avec une meilleure visualisation des défauts
    """
    numero_serie = os.path.basename(chemin_repertoire)
    
    # Extraire les infos sur les résistances des séquences SEQ-01
    infos_resistances = None
    for test_info in resultats_tests:
        if test_info["sequence"] == "SEQ-01" and test_info["resistances"]:
            infos_resistances = test_info["resistances"]
            break
    
    # Organiser les défauts par type
    defauts_par_type = organiser_defauts_par_type(resultats_tests)
    
    # Générer le contenu HTML
    html_content = [
        '<!DOCTYPE html>',
        '<html lang="fr">',
        '<head>',
        '    <meta charset="UTF-8">',
        '    <meta name="viewport" content="width=device-width, initial-scale=1.0">',
        f'    <title>Rapport de Test - {numero_serie}</title>',
        '    <style>',
        '        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; color: #333; }',
        '        .container { max-width: 1200px; margin: 0 auto; }',
        '        h1, h2, h3 { color: #0066cc; }',
        '        h1 { border-bottom: 2px solid #0066cc; padding-bottom: 10px; }',
        '        .summary { background-color: #f2f2f2; padding: 15px; border-radius: 5px; margin-bottom: 20px; }',
        '        .resistances { background-color: #e6f7ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }',
        '        .defaut-category { background-color: #f5f5f5; margin: 20px 0; padding: 15px; border-left: 5px solid #0066cc; border-radius: 0 5px 5px 0; }',
        '        .test-sequence { margin-bottom: 15px; padding: 10px; border: 1px solid #ddd; border-radius: 5px; }',
        '        .passed { background-color: #e6ffe6; }',
        '        .failed { background-color: #ffe6e6; }',
        '        .terminated { background-color: #e6e6ff; }',
        '        .status-badge { display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 0.8em; margin-left: 5px; }',
        '        .status-passed { background-color: #4CAF50; color: white; }',
        '        .status-failed { background-color: #f44336; color: white; }',
        '        .status-terminated { background-color: #2196F3; color: white; }',
        '        .detail-box { background-color: #f9f9f9; padding: 10px; border-left: 3px solid #ccc; margin-top: 10px; white-space: pre-line; }',
        '        .detail-section { margin-top: 10px; border-top: 1px dotted #ccc; padding-top: 10px; }',
        '        .configuration { font-weight: bold; color: #0066cc; }',
        '        .measurements { margin-left: 20px; }',
        '        .measurement-item { display: flex; justify-content: space-between; max-width: 500px; }',
        '        .measurement-label { font-weight: bold; }',
        '        .timeline { position: relative; margin: 20px 0; padding-left: 30px; }',
        '        .timeline::before { content: ""; position: absolute; left: 10px; top: 0; bottom: 0; width: 2px; background-color: #ddd; }',
        '        .timeline-item { position: relative; margin-bottom: 20px; }',
        '        .timeline-item::before { content: ""; position: absolute; left: -30px; top: 5px; width: 12px; height: 12px; border-radius: 50%; background-color: #0066cc; }',
        '        .timeline-date { font-size: 0.9em; color: #666; margin-bottom: 5px; }',
        '        .collapsible { cursor: pointer; padding: 10px; width: 100%; border: none; text-align: left; outline: none; font-size: 15px; background-color: #f1f1f1; margin-bottom: 5px; }',
        '        .collapsible:after { content: "\\002B"; color: #777; font-weight: bold; float: right; margin-left: 5px; }',
        '        .collapsible.active:after { content: "\\2212"; }',
        '        .content { padding: 0 18px; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out; background-color: white; }',
        '    </style>',
        '</head>',
        '<body>',
        '    <div class="container">',
        f'        <h1>Rapport de Test - Numéro de série: {numero_serie}</h1>',
        '        <div class="summary">',
        '            <h2>Résumé des Tests</h2>',
        f'            <p><strong>Nombre de séquences analysées:</strong> {len(resultats_tests)}</p>',
        f'            <p><strong>Types de défauts identifiés:</strong> {len(defauts_par_type)}</p>',
        '        </div>'
    ]
    
    # Ajouter les infos sur les résistances si disponibles
    if infos_resistances:
        html_content.extend([
            '        <div class="resistances">',
            '            <h2>Informations sur les Résistances (SEQ-01)</h2>',
            '            <h3>Résistances calculées:</h3>',
            '            <ul>'
        ])
        
        if infos_resistances.get("R46_calculee"):
            html_content.append(f'                <li><strong>R46:</strong> {infos_resistances["R46_calculee"]}</li>')
        if infos_resistances.get("R47_calculee"):
            html_content.append(f'                <li><strong>R47:</strong> {infos_resistances["R47_calculee"]}</li>')
        if infos_resistances.get("R48_calculee"):
            html_content.append(f'                <li><strong>R48:</strong> {infos_resistances["R48_calculee"]}</li>')
        
        html_content.extend([
            '            </ul>',
            '            <h3>Résistances à monter:</h3>',
            '            <ul>'
        ])
        
        if infos_resistances.get("R46_monter"):
            marquage = f" (Marquage CMS: {infos_resistances['R46_marquage']})" if infos_resistances.get("R46_marquage") else ""
            html_content.append(f'                <li><strong>R46:</strong> {infos_resistances["R46_monter"]}{marquage}</li>')
        if infos_resistances.get("R47_monter"):
            marquage = f" (Marquage CMS: {infos_resistances['R47_marquage']})" if infos_resistances.get("R47_marquage") else ""
            html_content.append(f'                <li><strong>R47:</strong> {infos_resistances["R47_monter"]}{marquage}</li>')
        if infos_resistances.get("R48_monter"):
            marquage = f" (Marquage CMS: {infos_resistances['R48_marquage']})" if infos_resistances.get("R48_marquage") else ""
            html_content.append(f'                <li><strong>R48:</strong> {infos_resistances["R48_monter"]}{marquage}</li>')
        
        html_content.extend([
            '            </ul>',
            '        </div>'
        ])
    
    # Section pour chaque type de défaut
    html_content.append('        <h2>Évolution des Défauts</h2>')
    
    for defaut_type, defauts in defauts_par_type.items():
        html_content.extend([
            f'        <button class="collapsible">{defaut_type} ({len(defauts)} occurrences)</button>',
            '        <div class="content">',
            f'            <div class="defaut-category">',
            f'                <h3>{defaut_type}</h3>',
            '                <div class="timeline">'
        ])
        
        for defaut in defauts:
            date_str = defaut["date_test"].strftime("%d/%m/%Y %H:%M:%S") if defaut["date_test"] != datetime.min else "Date inconnue"
            status_class = "status-failed" if defaut["status"] == "Failed" else "status-terminated" if defaut["status"] == "Terminated" else ""
            
            html_content.extend([
                '                    <div class="timeline-item">',
                f'                        <div class="timeline-date">{date_str} - {defaut["nom_fichier"]}</div>',
                '                        <div class="test-sequence">',
                f'                            <h4>{defaut["nom_test"]} <span class="status-badge {status_class}">{defaut["status"]}</span></h4>',
                '                            <div class="detail-box">',
            ])
            
            # Formater les détails
            details_lines = defaut["detail"].split("\n")
            for line in details_lines:
                if "Configuration:" in line:
                    html_content.append(f'                                <div class="configuration">{line}</div>')
                elif "Mesures:" in line:
                    html_content.append('                                <div class="detail-section">Mesures:</div>')
                elif line.strip().startswith("- Valeur"):
                    parts = line.strip().split(":", 1)
                    if len(parts) == 2:
                        label, value = parts
                        html_content.append(f'                                <div class="measurements"><div class="measurement-item"><span class="measurement-label">{label.strip()}:</span> <span>{value.strip()}</span></div></div>')
                else:
                    html_content.append(f'                                {line}')
            
            html_content.extend([
                '                            </div>',
                '                        </div>',
                '                    </div>'
            ])
        
        html_content.extend([
            '                </div>',
            '            </div>',
            '        </div>'
        ])
    
    # Ajouter le script JavaScript pour les éléments pliables (collapsible)
    html_content.extend([
        '    </div>',
        '    <script>',
        '        var coll = document.getElementsByClassName("collapsible");',
        '        var i;',
        '',
        '        for (i = 0; i < coll.length; i++) {',
        '            coll[i].addEventListener("click", function() {',
        '                this.classList.toggle("active");',
        '                var content = this.nextElementSibling;',
        '                if (content.style.maxHeight){',
        '                    content.style.maxHeight = null;',
        '                } else {',
        '                    content.style.maxHeight = content.scrollHeight + "px";',
        '                } ',
        '            });',
        '        }',
        '    </script>',
        '</body>',
        '</html>'
    ])
    
    # Enregistrer le fichier HTML
    nom_fichier_html = f"rapport_visuel_{numero_serie}.html"
    chemin_fichier_html = os.path.join(chemin_repertoire, nom_fichier_html)
    
    with open(chemin_fichier_html, 'w', encoding='utf-8') as f:
        f.write("\n".join(html_content))
    
    return chemin_fichier_html

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
    
    # Générer le rapport pour ce numéro de série (texte)
    rapport_final = []
    rapport_final.append(f"Numéro de série : {numero_serie}")
    rapport_final.append("-" * 70)  # Séparateur
    
    # Stocker les résultats de tous les tests pour le rapport HTML
    resultats_tous_tests = []
    
    for chemin_complet in chemins_html:
        try:
            resultats = analyser_fichier_html(chemin_complet)
            resultats_tous_tests.append(resultats)  # Stocker pour le rapport HTML
            
            statut = resultats["resultat_global"] or "Inconnu"
            rapport_final.append(f"{resultats['nom_fichier']} : {statut}")
            
            # Si c'est une séquence SEQ-01 et qu'il y a des informations sur les résistances
            if resultats["sequence"] == "SEQ-01" and resultats["resistances"]:
                resistances = resultats["resistances"]
                rapport_final.append("  Informations sur les résistances:")
                
                if resistances.get("R46_calculee"):
                    rapport_final.append(f"    - Résistance R46 calculée: {resistances['R46_calculee']}")
                if resistances.get("R47_calculee"):
                    rapport_final.append(f"    - Résistance R47 calculée: {resistances['R47_calculee']}")
                if resistances.get("R48_calculee"):
                    rapport_final.append(f"    - Résistance R48 calculée: {resistances['R48_calculee']}")
                    
                rapport_final.append("  Résistances à monter:")
                if resistances.get("R46_monter"):
                    marquage = f" (Marquage CMS: {resistances['R46_marquage']})" if resistances.get("R46_marquage") else ""
                    rapport_final.append(f"    - R46: {resistances['R46_monter']}{marquage}")
                if resistances.get("R47_monter"):
                    marquage = f" (Marquage CMS: {resistances['R47_marquage']})" if resistances.get("R47_marquage") else ""
                    rapport_final.append(f"    - R47: {resistances['R47_monter']}{marquage}")
                if resistances.get("R48_monter"):
                    marquage = f" (Marquage CMS: {resistances['R48_marquage']})" if resistances.get("R48_marquage") else ""
                    rapport_final.append(f"    - R48: {resistances['R48_monter']}{marquage}")
                    
                rapport_final.append("")  # Ligne vide pour séparer
            
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
    
    # Générer le rapport texte
    nom_fichier_rapport = f"rapport_{numero_serie}.txt"
    chemin_fichier_rapport = os.path.join(chemin_repertoire, nom_fichier_rapport)
    
    try:
        with open(chemin_fichier_rapport, 'w', encoding='utf-8') as f:
            f.write('\n'.join(rapport_final))
        print(f"Rapport texte créé: {chemin_fichier_rapport}")
    
        # Ouvrir le rapport texte
        subprocess.Popen(['notepad.exe', chemin_fichier_rapport])
    except Exception as e:
        print(f"Erreur lors de la création/ouverture du rapport texte: {e}")
    
    # Générer le rapport HTML amélioré
    try:
        chemin_html = generer_rapport_html(chemin_repertoire, resultats_tous_tests)
        print(f"Rapport HTML créé: {chemin_html}")
        
        # Ouvrir le rapport HTML dans le navigateur par défaut
        webbrowser.open(f'file://{os.path.abspath(chemin_html)}')
    except Exception as e:
        print(f"Erreur lors de la création/ouverture du rapport HTML: {e}")

def main():
    # Configuration pour l'affichage des caractères accentués
    # Version robuste qui fonctionne aussi en exécutable
    try:
        if sys.stdout and hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except (AttributeError, TypeError):
        # En cas d'erreur, on continue sans changer l'encodage
        pass
    
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
