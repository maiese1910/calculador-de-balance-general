import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from .processor import compute_balance_from_diary_and_ledger, export_balance_to_excel


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Calculador de Balance General - Facultad')
        self.geometry('600x200')

        tk.Label(self, text='Libro Diario (.xlsx)').grid(row=0, column=0, sticky='w', padx=8, pady=8)
        self.diary_entry = tk.Entry(self, width=60)
        self.diary_entry.grid(row=0, column=1, padx=8)
        tk.Button(self, text='Seleccionar', command=self.select_diary).grid(row=0, column=2, padx=8)

        tk.Label(self, text='Libro Mayor (.xlsx)').grid(row=1, column=0, sticky='w', padx=8, pady=8)
        self.ledger_entry = tk.Entry(self, width=60)
        self.ledger_entry.grid(row=1, column=1, padx=8)
        tk.Button(self, text='Seleccionar', command=self.select_ledger).grid(row=1, column=2, padx=8)

        tk.Button(self, text='Generar Balance', command=self.generate_balance).grid(row=2, column=1, pady=16)

    def select_diary(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
        if path:
            self.diary_entry.delete(0, tk.END)
            self.diary_entry.insert(0, path)

    def select_ledger(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
        if path:
            self.ledger_entry.delete(0, tk.END)
            self.ledger_entry.insert(0, path)

    def generate_balance(self):
        diary_path = self.diary_entry.get().strip()
        ledger_path = self.ledger_entry.get().strip()
        if not diary_path or not ledger_path:
            messagebox.showwarning('Faltan archivos', 'Seleccione ambos archivos: libro diario y libro mayor')
            return
        try:
            diary_df = pd.read_excel(diary_path)
            ledger_df = pd.read_excel(ledger_path)
            balance = compute_balance_from_diary_and_ledger(diary_df, ledger_df)

            # Ask where to save
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
            if save_path:
                export_balance_to_excel(balance, save_path)
                messagebox.showinfo('Éxito', f'Balance exportado a {save_path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))
