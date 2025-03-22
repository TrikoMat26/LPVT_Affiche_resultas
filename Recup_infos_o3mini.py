import re
import csv

def extract_measurements(html):
    results = {}

    # Extraction pour le test des alimentations à 24VDC
    block_24 = re.search(r"Test des alimentations à 24VDC(.*?)Test des alimentations à 115VAC", html, re.DOTALL)
    if block_24:
        block = block_24.group(1)
        m_plus16 = re.search(r"Lecture mesure \+16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        m_minus16 = re.search(r"Lecture mesure -16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        m_plus5   = re.search(r"Lecture mesure \+5V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        m_minus5  = re.search(r"Lecture mesure -5V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        results['24VDC'] = {
            '+16V': m_plus16.group(1).strip() if m_plus16 else None,
            '-16V': m_minus16.group(1).strip() if m_minus16 else None,
            '+5V':   m_plus5.group(1).strip()   if m_plus5   else None,
            '-5V':   m_minus5.group(1).strip()  if m_minus5  else None,
        }

    # Extraction pour le test des alimentations à 115VAC
    block_115 = re.search(r"Test des alimentations à 115VAC(.*?)Calcul des résistances", html, re.DOTALL)
    if block_115:
        block = block_115.group(1)
        m_plus16 = re.search(r"Lecture mesure \+16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        m_minus16 = re.search(r"Lecture mesure -16V AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([^<]+)</span>", block, re.DOTALL)
        results['115VAC'] = {
            '+16V': m_plus16.group(1).strip() if m_plus16 else None,
            '-16V': m_minus16.group(1).strip() if m_minus16 else None,
        }

    # Extraction pour le calcul des résistances (recherche directe dans tout le document)
    r46_calc   = re.search(r"Résistance R46 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html, re.DOTALL)
    r46_monter = re.search(r"Résistance R46 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html, re.DOTALL)
    r47_calc   = re.search(r"Résistance R47 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html, re.DOTALL)
    r47_monter = re.search(r"Résistance R47 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html, re.DOTALL)
    r48_calc   = re.search(r"Résistance R48 calculée:\s*</td>\s*<td[^>]*>\s*([\d\.]+)\s*</td>", html, re.DOTALL)
    r48_monter = re.search(r"Résistance R48 à monter:\s*</td>\s*<td[^>]*>\s*Résistance à monter =\s*([\d]+)\s*ohms", html, re.DOTALL)

    results['Resistances'] = {
        'R46_calculee': r46_calc.group(1).strip() if r46_calc else None,
        'R46_a_monter': r46_monter.group(1).strip() if r46_monter else None,
        'R47_calculee': r47_calc.group(1).strip() if r47_calc else None,
        'R47_a_monter': r47_monter.group(1).strip() if r47_monter else None,
        'R48_calculee': r48_calc.group(1).strip() if r48_calc else None,
        'R48_a_monter': r48_monter.group(1).strip() if r48_monter else None,
    }

    return results

def main():
    # Nom du fichier HTML à traiter
    input_file = "SEQ-01_LPVT_Report[11 50 55][31 01 2025].html"
    
    # Lecture du fichier HTML
    with open(input_file, encoding="iso-8859-1") as f:
        html = f.read()
    
    # Extraction des mesures
    results = extract_measurements(html)
    
    # Écriture des résultats dans des fichiers CSV distincts

    # 1. Alimentations 24VDC
    if '24VDC' in results:
        with open("alimentations_24VDC.csv", "w", newline="") as csvfile:
            fieldnames = ["+16V", "-16V", "+5V", "-5V"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(results['24VDC'])
        print("Données 24VDC écrites dans alimentations_24VDC.csv")
    
    # 2. Alimentations 115VAC
    if '115VAC' in results:
        with open("alimentations_115VAC.csv", "w", newline="") as csvfile:
            fieldnames = ["+16V", "-16V"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(results['115VAC'])
        print("Données 115VAC écrites dans alimentations_115VAC.csv")
    
    # 3. Calcul des résistances
    if 'Resistances' in results:
        with open("calcul_resistances.csv", "w", newline="") as csvfile:
            fieldnames = ["R46_calculee", "R46_a_monter", "R47_calculee", "R47_a_monter", "R48_calculee", "R48_a_monter"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(results['Resistances'])
        print("Données résistances écrites dans calcul_resistances.csv")

if __name__ == "__main__":
    main()
