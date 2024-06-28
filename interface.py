import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd

class InterfaceExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface Excel")
        self.df = None
        self.filepath = None

        # Cadre pour les boutons
        frame_buttons = tk.Frame(root)
        frame_buttons.pack(pady=20)

        # Bouton pour charger un fichier Excel
        btn_charger_excel = tk.Button(frame_buttons, text="Charger Excel", command=self.charger_excel)
        btn_charger_excel.pack(side="left", padx=10)

        # Bouton pour enregistrer les modifications
        btn_enregistrer_excel = tk.Button(frame_buttons, text="Enregistrer Excel", command=self.enregistrer_excel)
        btn_enregistrer_excel.pack(side="left", padx=10)

        # Cadre pour la table
        frame_table = tk.Frame(root)
        frame_table.pack()

        # Création de la table
        self.table = ttk.Treeview(frame_table, columns=("Produit", "Prix", "Quantite"), show='headings')
        self.table.pack()

        # Ajouter des colonnes à la table
        self.table.heading("Produit", text="Produit")
        self.table.heading("Prix", text="Prix")
        self.table.heading("Quantite", text="Quantite")

        # Associer la double-clique à la fonction pour éditer la cellule
        self.table.bind('<Double-1>', self.edit_cell)

    def charger_excel(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
        
        if self.filepath:
            try:
                self.df = pd.read_excel(self.filepath)
                self.afficher_donnees(self.df)
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel : {e}")

    def afficher_donnees(self, dataframe):
        # Vider la table existante
        for i in self.table.get_children():
            self.table.delete(i)
        
        # Insérer les nouvelles données
        for index, row in dataframe.iterrows():
            self.table.insert("", "end", values=list(row), iid=index)

    def edit_cell(self, event):
        item = self.table.selection()[0]
        column = self.table.identify_column(event.x)
        column_index = int(column.replace("#", "")) - 1
        value = self.table.item(item, "values")[column_index]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Cell")

        entry = tk.Entry(edit_window)
        entry.pack()
        entry.insert(0, value)

        def save_edit():
            new_value = entry.get()
            current_values = list(self.table.item(item, "values"))
            current_values[column_index] = new_value
            self.table.item(item, values=current_values)

            # Mettre à jour le DataFrame
            row_index = int(item)
            self.df.iat[row_index, column_index] = new_value
            edit_window.destroy()

        btn_save = tk.Button(edit_window, text="Save", command=save_edit)
        btn_save.pack()

    def enregistrer_excel(self):
        if self.df is not None and self.filepath is not None:
            try:
                self.df.to_excel(self.filepath, index=False)
                messagebox.showinfo("Succès", "Modifications enregistrées avec succès !")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement du fichier Excel : {e}")
        else:
            messagebox.showwarning("Attention", "Aucun fichier Excel n'a été chargé.")

