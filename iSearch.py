import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from collections import Counter
import os
from pathlib import Path
import shutil

# Création de la fenêtre principale
window = tk.Tk()

# Fonction appelée lorsque le bouton "Parcourir" est cliqué
def select_file():
    # Récupération du mot saisi par l'utilisateur
    search_word = search_entry.get()

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        # Lecture des données à partir du fichier sélectionné
        data = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)

        # Comptage des occurrences pour toutes les colonnes
        counter = Counter()
        for column in data.columns:
            counter.update(data[column])
    
         # Filtrage des lignes contenant le mot recherché
        mask = data.apply(lambda x: x.astype(str).str.contains(search_word, case=False).any(), axis=1)
        
        filtered_data = data.loc[mask]
        

        # Enregistrement des résultats dans un fichier Excel
        result_filename = os.path.splitext(os.path.basename(file_path))[0] + f"_occurrences_{search_word}.xlsx"
        filtered_data.to_excel(result_filename, index=False)


        # Recherche du mot dans les compteurs
        count = len(filtered_data)

        # Affichage des résultats dans une boîte de dialogue
        result_str = f"Occurrences pour la recherche :{count}. \nVérifiez dans le fichier Excel pour plus de détails"
        messagebox.showinfo("Résultats", result_str)

        # Ajout du nombre d'occurrences à la fin du fichier Excel
        result_df = pd.read_excel(result_filename)
        result_df = pd.concat([pd.DataFrame({"Occurrences": [len(filtered_data)]}), result_df], ignore_index=True)
        result_df.to_excel(result_filename, index=False)
        

# Champ pour entrer le mot pour voir le nombre d'occurrences 
search_entry = tk.Label(text="Taper un mot à rechercher:")
search_entry.pack()
# Création d'un champ de saisie pour la recherche
search_entry = tk.Entry(window)
search_entry.pack()

# Création d'un bouton pour lancer la recherche
search_button = tk.Button(window, text="Rechercher", command=select_file)
search_button.pack()

window.mainloop()
