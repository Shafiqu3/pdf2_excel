import tkinter as tk
from tkinter import filedialog, messagebox
import tabula
import pandas as pd

def convert_pdf_to_excel(pdf_path, excel_path):
    try:
        dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)
        messagebox.showinfo("Success", "PDF successfully converted to Excel!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def select_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            convert_pdf_to_excel(pdf_path, excel_path)

root = tk.Tk()
root.title("PDF to Excel Converter")

btn_select_pdf = tk.Button(root, text="Select PDF", command=select_pdf)
btn_select_pdf.pack(pady=20)

root.mainloop()