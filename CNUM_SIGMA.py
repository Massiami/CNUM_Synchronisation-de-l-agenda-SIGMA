#%% chemin d'accès LOU
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

#%% chmein d'accès Massiami 
#Gestion des cellules fusionnées

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Chemins des fichiers
file_path = "C:/Users/hp/OneDrive/Documents/COURS IA/SEMESTRE 8/UE-PROJET/CNUM/test.xlsx"
new_file_path = "C:/Users/hp/OneDrive/Documents/COURS IA/SEMESTRE 8/UE-PROJET/CNUM/test_modifie.xlsx"  # Nouveau fichier Excel
sheet_name = "M1 2324"  # Nom de la feuille d'intérêt

# Charger le fichier Excel avec openpyxl pour récupérer couleurs et fusions
wb = load_workbook(file_path)
ws = wb[sheet_name]

# Récupérer les cellules fusionnées dans la zone d'intérêt (colonnes E à O et lignes 5 à 34)
merged_cells = []
for merge_range in ws.merged_cells.ranges:
    if (merge_range.min_col >= 5 and merge_range.max_col <= 15 and 
        merge_range.min_row >= 5 and merge_range.max_row <= 34):
        merged_cells.append(merge_range)

# Charger les données avec pandas :
# - skiprows=4 : on ignore les 4 premières lignes
# - usecols="E:O" : on sélectionne les colonnes E à O
# - header=0 : la première ligne lue (originalement la ligne 5) sera l'en-tête
df = pd.read_excel(file_path, sheet_name=sheet_name,
                   skiprows=4,
                   usecols="E:O",
                   header=0,
                   engine="openpyxl")

# Limiter le DataFrame aux 30 premières lignes de données (originalement lignes 6 à 34)
df = df.iloc[:29]

# Récupérer les couleurs de remplissage pour les cellules de la zone d'intérêt
# On commence à la ligne 6 (plutôt que 5) pour bien coller à la zone des données dans le nouveau fichier.
cell_colors = {}
for row_idx, row in enumerate(ws.iter_rows(min_row=6, max_row=34, min_col=5, max_col=15), start=0):
    for col_idx, cell in enumerate(row, start=0):
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb and cell.fill.fgColor.rgb != "00000000":
            cell_colors[(row_idx, col_idx)] = cell.fill.fgColor.rgb

# Écrire les données dans un nouveau fichier Excel
# Le nouveau fichier aura l'en-tête en ligne 1 et les données à partir de la ligne 2.
df.to_excel(new_file_path, sheet_name="M1 2324_modifie", index=False)

# Charger le nouveau fichier avec openpyxl pour appliquer les couleurs et gérer les cellules fusionnées
new_wb = load_workbook(new_file_path)
new_ws = new_wb["M1 2324_modifie"]

# Appliquer les couleurs récupérées aux cellules du nouveau fichier (données à partir de la ligne 2)
for row_idx in range(len(df)):          # pour chaque ligne de données (0 à 28)
    for col_idx in range(len(df.columns)):  # pour chaque colonne
        cell = new_ws.cell(row=row_idx + 2, column=col_idx + 1)
        if (row_idx, col_idx) in cell_colors:
            color = cell_colors[(row_idx, col_idx)]
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# Pour les cellules fusionnées de l'original :
# Nous récupérons la valeur et la couleur de la cellule en haut à gauche et les appliquons à toutes les cellules de la plage.
for merge_range in merged_cells:
    # Conversion des indices :
    # - nouvelle_ligne = ligne_originale - 4 (puisque l'en-tête remplace la ligne 5)
    # - nouvelle_colonne = colonne_originale - 4 (car la colonne E devient colonne 1)
    new_min_row = merge_range.min_row - 4
    new_max_row = merge_range.max_row - 4
    new_min_col = merge_range.min_col - 4
    new_max_col = merge_range.max_col - 4

    # Récupérer la valeur de la cellule en haut à gauche dans l'original et sa couleur de remplissage
    original_top_left = ws.cell(row=merge_range.min_row, column=merge_range.min_col)
    top_left_value = original_top_left.value
    top_left_color = None
    if original_top_left.fill and original_top_left.fill.fgColor and original_top_left.fill.fgColor.rgb and original_top_left.fill.fgColor.rgb != "00000000":
        top_left_color = original_top_left.fill.fgColor.rgb

    # Appliquer la valeur et, si disponible, le remplissage à chaque cellule de la plage dans le nouveau fichier
    for r in range(new_min_row, new_max_row + 1):
        for c in range(new_min_col, new_max_col + 1):
            new_cell = new_ws.cell(row=r, column=c)
            new_cell.value = top_left_value
            if top_left_color:
                new_cell.fill = PatternFill(start_color=top_left_color, end_color=top_left_color, fill_type="solid")

# Sauvegarder le nouveau fichier modifié
new_wb.save(new_file_path)

print(f"Le fichier Excel modifié a été sauvegardé avec les couleurs et le contenu copié pour les cellules fusionnées sous : {new_file_path}")
