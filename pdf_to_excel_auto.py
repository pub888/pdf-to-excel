
import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os

def pdf_to_excel(pdf_path):
    try:
        # Nom du fichier Excel de sortie (même dossier, même nom)
        base = os.path.splitext(pdf_path)[0]
        excel_path = base + ".xlsx"

        all_tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for table_num, table in enumerate(tables, start=1):
                    # Vérifie qu’il y a bien des données (pas juste vide)
                    if table and len(table) > 1:
                        # Première ligne = entêtes si cohérentes, sinon numérotées
                        headers = table[0]
                        if any(headers):  # s'il y a des valeurs dans les entêtes
                            df = pd.DataFrame(table[1:], columns=headers)
                        else:
                            df = pd.DataFrame(table[1:])
                        df["Page"] = page_num
                        df["Table"] = table_num
                        all_tables.append(df)

        if not all_tables:
            messagebox.showwarning("Aucun tableau", "Aucun tableau détecté dans le PDF.")
            return

        # Concatène toutes les tables ensemble
        result = pd.concat(all_tables, ignore_index=True)

        # Sauvegarde dans le même dossier
        result.to_excel(excel_path, index=False, engine="openpyxl")
        messagebox.showinfo("Succès", f"✅ Fichier Excel créé :\n{excel_path}")

    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    # Boîte de dialogue pour choisir le fichier PDF
    root = Tk()
    root.withdraw()  # cache la fenêtre principale Tkinter
    pdf_path = filedialog.askopenfilename(
        title="Choisir un fichier PDF à convertir",
        filetypes=[("Fichiers PDF", "*.pdf")]
    )
    if pdf_path:
        pdf_to_excel(pdf_path)
