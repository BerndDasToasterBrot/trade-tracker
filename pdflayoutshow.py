import tkinter as tk
from tkinter import filedialog
import os
from pdfminer.high_level import extract_text, extract_pages
from pdfminer.layout import LAParams, LTTextContainer, LTChar

def analyze_pdf_miner():
    # 1. Dateiauswahl
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Wähle deine PDF für pdfminer aus",
        filetypes=[("PDF Dateien", "*.pdf")]
    )

    if not file_path:
        print("Abbruch: Keine Datei gewählt.")
        return

    print(f"\nANALYSIERE DATEI (pdfminer): {os.path.basename(file_path)}")
    print("="*60)

    # 2. Konfiguration der Layout-Analyse
    # Das ist der entscheidende Teil für die KI-Fütterung.
    # Wenn wir laparams übergeben, versucht pdfminer Spalten zu erkennen.
    laparams = LAParams(
        line_margin=0.5,    # Toleranz für Zeilenabstand
        char_margin=2.0,    # Toleranz für Buchstabenabstand
        all_texts=True      # Auch versteckten Text lesen
    )

    try:
        # Methode A: Der schnelle "Alles auf einmal" String (so wie es oft an LLMs geht)
        print(">>> VARIANTE A: Roher Gesamttext (extract_text) <<<")
        raw_text = extract_text(file_path, laparams=laparams)
        print(raw_text[:2000] + "\n... [Rest abgeschnitten für Übersicht] ...\n")
        print("-" * 60)

        # Methode B: Seitenweise Analyse (Detaillierter)
        print(">>> VARIANTE B: Seitenweise Block-Analyse <<<")
        # Hier iterieren wir durch die Layout-Boxen, um zu sehen, WIE pdfminer gruppiert
        for page_layout in extract_pages(file_path, laparams=laparams):
            print(f"\n--- Seite {page_layout.pageid} ---")
            
            for element in page_layout:
                # Wir schauen nur auf Text-Container (Boxen)
                if isinstance(element, LTTextContainer):
                    # Ausgabe der Box-Koordinaten und des Textes
                    # Das hilft zu verstehen, ob er Header/Footer als eigenen Block erkennt
                    x0, y0, x1, y1 = element.bbox
                    text_content = element.get_text().strip()
                    
                    if text_content:
                        print(f"[Box @ {x0:.1f},{y0:.1f}]: {text_content[:100]}...")

    except Exception as e:
        print(f"Fehler: {e}")

if __name__ == "__main__":
    analyze_pdf_miner()