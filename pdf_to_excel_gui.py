import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber

def convert_pdf_to_excel(pdf_path, excel_path):
    try:
        all_tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for t in tables:
                    if t and len(t) > 1:
                        df = pd.DataFrame(t[1:], columns=t[0])
                        all_tables.append(df)
        if not all_tables:
            messagebox.showwarning("Aucun tableau", "Aucun tableau détecté dans le PDF.")
            return
        writer = pd.ExcelWriter(excel_path, engine='openpyxl')
        for i, df in enumerate(all_tables):
            df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
        writer.close()
        messagebox.showinfo("Succès", f"✅ Conversion terminée :\n{excel_path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

def select_pdf():
    path = filedialog.askopenfilename(filetypes=[("Fichiers PDF", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, path)

def select_output():
    path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                        filetypes=[("Fichier Excel", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, path)

def start_conversion():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()
    if not pdf_path or not excel_path:
        messagebox.showwarning("Champs manquants", "Veuillez sélectionner un PDF et un fichier de sortie.")
        return
    convert_pdf_to_excel(pdf_path, excel_path)

root = tk.Tk()
root.title("Convertisseur PDF → Excel")
root.geometry("500x250")
root.resizable(False, False)

tk.Label(root, text="Fichier PDF à convertir :").pack(pady=5)
pdf_entry = tk.Entry(root, width=60)
pdf_entry.pack()
tk.Button(root, text="Choisir un PDF", command=select_pdf).pack(pady=5)

tk.Label(root, text="Fichier Excel de sortie :").pack(pady=5)
excel_entry = tk.Entry(root, width=60)
excel_entry.pack()
tk.Button(root, text="Choisir l’emplacement", command=select_output).pack(pady=5)

tk.Button(root, text="Convertir maintenant", command=start_conversion,
          bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=15)

root.mainloop()
