import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Chemins des fichiers
file_path = "C:/Documents/ENSAT/2A/cours S8/Projets/CNUM - Conception Numérique/test.xlsx"
new_file_path = "C:/Documents/ENSAT/2A/cours S8/Projets/CNUM - Conception Numérique/test_modifie.xlsx" #création nouveau fichier Excel contenant que le tableau d'intérêt
sheet_name = "M1 2324" #sélection de l'emploi du temps du semestre concerné (feuille 1 ici)

# Charger le fichier Excel avec openpyxl pour récupérer les couleurs
wb = load_workbook(file_path)
ws = wb[sheet_name]

#Sélection du tableau précis contenant uniquement les informations nécessaires
# Lire les données avec pandas, en ignorant les 4 premières lignes et en sélectionnant les bonnes colonnes
df = pd.read_excel(file_path, sheet_name=sheet_name,
                   skiprows=4,  # Ignorer les 4 premières lignes
                   usecols="E:O",  # Sélectionner les bonnes colonnes (E:O)
                   header=0,  # Utiliser la ligne 4 comme en-tête
                   engine="openpyxl")

# Limiter le DataFrame aux 30 premières lignes 
df = df.iloc[:29]

# Récupérer les couleurs de remplissage des cellules pour les lignes 5 à 33 et les colonnes E à O
cell_colors = {}
for row_idx, row in enumerate(ws.iter_rows(min_row=3, max_row=35, min_col=4, max_col=15), start=0):  # E=colonne 5, O=colonne 15
    for col_idx, cell in enumerate(row, start=0):  # Index 0 pour Pandas
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb != "00000000":  # Vérifier si une couleur est définie
            # Ajuster le décalage des indices : remonter de 3 lignes et déplacer d'une colonne à gauche
            cell_colors[(row_idx - 3, col_idx - 1)] = cell.fill.fgColor.rgb  # Stocker la couleur ajustée

# Écrire les données modifiées dans un nouveau fichier Excel
df.to_excel(new_file_path, sheet_name="M1 2324_modifie", index=False)

# Charger le nouveau fichier avec openpyxl pour réappliquer les couleurs
new_wb = load_workbook(new_file_path)
new_ws = new_wb["M1 2324_modifie"]

# Appliquer les couleurs récupérées aux nouvelles cellules
for row_idx in range(len(df)):  # Index de 0 à 29
    for col_idx in range(len(df.columns)):  # Index de 0 à 10 (colonnes E à O)
        cell = new_ws.cell(row=row_idx + 2, column=col_idx + 1)  # Ajuster les indices pour correspondre aux nouvelles cellules
        if (row_idx, col_idx) in cell_colors:  # Si une couleur est stockée pour cette cellule
            color = cell_colors[(row_idx, col_idx)]
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# Sauvegarder le fichier modifié
new_wb.save(new_file_path)

print(f"Le fichier Excel modifié a été sauvegardé avec les couleurs sous : {new_file_path}")
