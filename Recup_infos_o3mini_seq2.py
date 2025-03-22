import re
import csv

def extract_seq2_measurements(html):
    results = {}
    
    # Pattern pour le test en 19VDC :
    # On cherche "Test 1.9Un sur 2 voies en 19VDC" suivi (quelques caractères après)
    # de "Lecture mesure -16V AG34461A" et ensuite la valeur numérique dans le span de Measurement[1]
    pattern_19 = r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+19VDC.*?Lecture\s+mesure\s+-16V\s+AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>"
    m_19 = re.search(pattern_19, html, re.DOTALL | re.IGNORECASE)
    if m_19:
        results['Test_19VDC'] = m_19.group(1).strip()
    else:
        results['Test_19VDC'] = None

    # Pattern pour le test en 115VAC :
    # Même logique que pour le bloc précédent, en remplaçant "19VDC" par "115VAC"
    pattern_115 = r"Test\s*1\.9Un\s+sur\s+2\s+voies\s+en\s+115VAC.*?Lecture\s+mesure\s+-16V\s+AG34461A.*?Measurement\[1\].*?Data:\s*</td>\s*<td[^>]*>.*?>([-\d\.]+)</span>"
    m_115 = re.search(pattern_115, html, re.DOTALL | re.IGNORECASE)
    if m_115:
        results['Test_115VAC'] = m_115.group(1).strip()
    else:
        results['Test_115VAC'] = None

    return results

def main():
    input_file = "SEQ-02_LPVT_Report[10 43 07][10 03 2025].html"
    with open(input_file, encoding="iso-8859-1") as f:
        html = f.read()
    
    seq2_results = extract_seq2_measurements(html)
    
    # Écriture des résultats dans un fichier CSV
    with open("seq2_measurements.csv", "w", newline="") as csvfile:
        fieldnames = ["Test_19VDC", "Test_115VAC"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerow(seq2_results)
    
    print("Les mesures de séquence 2 ont été extraites dans seq2_measurements.csv")

if __name__ == "__main__":
    main()
