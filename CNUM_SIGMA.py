## CODE 1: SELECTION DU TABLEAU D'INTERET ## 

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Définition des chemins et paramètres de la feuille d'intérêt
file_path = "C:/Users/hp/OneDrive/Documents/COURS IA/SEMESTRE 8/UE-PROJET/CNUM/test.xlsx"
new_file_path = "C:/Users/hp/OneDrive/Documents/COURS IA/SEMESTRE 8/UE-PROJET/CNUM/test_modifie.xlsx"  # Chemin du nouveau fichier Excel qui contiendra le tableau modifié
sheet_name = "M1 2324"  # Nom de la feuille contenant l'emploi du temps ou les données d'intérêt

# ------------------------------------------------------------------
# Chargement du fichier Excel original avec openpyxl pour récupérer
# les informations de formatage (couleurs, commentaires) et les cellules fusionnées.
# ------------------------------------------------------------------
wb = load_workbook(file_path)
ws = wb[sheet_name]

# Récupération des plages de cellules fusionnées dans la zone d'intérêt
# La zone d'intérêt correspond aux colonnes E à O (colonnes 5 à 15) et aux lignes 5 à 34.
merged_cells = []
for merge_range in ws.merged_cells.ranges:
    if (merge_range.min_col >= 5 and merge_range.max_col <= 15 and 
        merge_range.min_row >= 5 and merge_range.max_row <= 34):
        merged_cells.append(merge_range)

# ------------------------------------------------------------------
# Chargement des données dans un DataFrame avec pandas.
# - On ignore les 4 premières lignes (skiprows=4) pour se positionner sur la zone utile.
# - On sélectionne les colonnes E à O (usecols="E:O") pour extraire uniquement le tableau souhaité.
# - La ligne lue en première position (après skiprows) sert d'en-tête (header=0).
# - On limite le DataFrame aux 29 premières lignes de données.
# ------------------------------------------------------------------
df = pd.read_excel(file_path, sheet_name=sheet_name,
                   skiprows=4, usecols="E:O", header=0, engine="openpyxl")
df = df.iloc[:29]

# ------------------------------------------------------------------
# Récupération des couleurs de remplissage et des commentaires des cellules
# dans la zone d'intérêt du fichier original.
# La zone considérée ici est de la ligne 6 à 34 et des colonnes 5 à 15.
# Les indices sont adaptés pour correspondre à la numérotation 0-based utilisée par pandas.
# ------------------------------------------------------------------
cell_colors = {}
cell_comments = {}
for row_idx, row in enumerate(ws.iter_rows(min_row=6, max_row=34, min_col=5, max_col=15), start=0):
    for col_idx, cell in enumerate(row, start=0):  # Les indices commencent à 0 pour aligner avec le DataFrame
        # Vérifier si la cellule possède une couleur de remplissage définie (et non par défaut)
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb and cell.fill.fgColor.rgb != "00000000":
            cell_colors[(row_idx, col_idx)] = cell.fill.fgColor.rgb
        # Si la cellule contient un commentaire, l'enregistrer
        if cell.comment:
            cell_comments[(row_idx, col_idx)] = cell.comment.text
           
# ------------------------------------------------------------------
# Écriture des données (sans formatage) dans un nouveau fichier Excel.
# Le nouveau fichier contiendra l'en-tête en ligne 1 et les données à partir de la ligne 2.
# ------------------------------------------------------------------
df.to_excel(new_file_path, sheet_name="M1 2324_modifie", index=False)

# ------------------------------------------------------------------
# Rechargement du nouveau fichier avec openpyxl afin d'appliquer
# le formatage (couleurs et commentaires) ainsi que la gestion des cellules fusionnées.
# ------------------------------------------------------------------
new_wb = load_workbook(new_file_path)
new_ws = new_wb["M1 2324_modifie"]

# Application des couleurs et commentaires récupérés sur chaque cellule correspondante
# dans le nouveau fichier Excel (les données commencent en ligne 2).
for row_idx in range(len(df)):
    for col_idx in range(len(df.columns)):
        cell = new_ws.cell(row=row_idx + 2, column=col_idx + 1)  # Ajustement : ligne d'en-tête décalée d'une unité
        # Appliquer la couleur si elle a été enregistrée pour cette cellule
        if (row_idx, col_idx) in cell_colors:
            color = cell_colors[(row_idx, col_idx)]
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        # Appliquer le commentaire si présent
        if (row_idx, col_idx) in cell_comments:
            cell.comment = openpyxl.comments.Comment(cell_comments[(row_idx, col_idx)], "Author")

# ------------------------------------------------------------------
# Gestion des cellules fusionnées :
# Pour chaque plage de cellules fusionnées détectée dans l'original,
# on copie la valeur, la couleur et le commentaire de la cellule en haut à gauche
# vers toutes les cellules correspondantes dans la nouvelle feuille.
#
# Les indices sont recalculés car l'en-tête du nouveau fichier remplace la ligne 5 de l'original
# et la colonne E devient la première colonne.
# ------------------------------------------------------------------
for merge_range in merged_cells:
    new_min_row = merge_range.min_row - 4  # Décalage dû à l'en-tête
    new_max_row = merge_range.max_row - 4
    new_min_col = merge_range.min_col - 4  # Décalage des colonnes (colonne E devient colonne 1)
    new_max_col = merge_range.max_col - 4

    # Récupérer la cellule en haut à gauche de la plage fusionnée dans le fichier original
    original_top_left = ws.cell(row=merge_range.min_row, column=merge_range.min_col)
    top_left_value = original_top_left.value
    top_left_color = None
    top_left_comment = original_top_left.comment.text if original_top_left.comment else None
    
    # Si une couleur de remplissage est définie pour la cellule en haut à gauche, la récupérer
    if original_top_left.fill and original_top_left.fill.fgColor and original_top_left.fill.fgColor.rgb and original_top_left.fill.fgColor.rgb != "00000000":
        top_left_color = original_top_left.fill.fgColor.rgb

    # Appliquer la valeur, la couleur et le commentaire à toutes les cellules de la plage fusionnée dans le nouveau fichier
    for r in range(new_min_row, new_max_row + 1):
        for c in range(new_min_col, new_max_col + 1):
            new_cell = new_ws.cell(row=r, column=c)
            new_cell.value = top_left_value
            if top_left_color:
                new_cell.fill = PatternFill(start_color=top_left_color, end_color=top_left_color, fill_type="solid")
            if top_left_comment:
                new_cell.comment = openpyxl.comments.Comment(top_left_comment, "Author")

# ------------------------------------------------------------------
# Sauvegarde du nouveau fichier Excel modifié contenant :
# - Les données extraites,
# - Le formatage (couleurs et commentaires),
# - La gestion des cellules fusionnées avec duplication des informations.
# ------------------------------------------------------------------
new_wb.save(new_file_path)
print(f"Le fichier Excel modifié a été sauvegardé avec les couleurs, la gestion des cellules fusionnées et les commentaires sous : {new_file_path}")



## CODE 2 : CRATION DU FICHIER EXCEL ## 

import openpyxl
import csv
import re
from datetime import datetime, timedelta
import os

# ---------------------------------------------------------------------------
# Paramètres : chemins et dictionnaires
# ---------------------------------------------------------------------------

# Emplacement du fichier Excel
excel_file = r"C:/Users/hp/OneDrive/Documents/COURS IA/SEMESTRE 8/UE-PROJET/CNUM/test_modifie.xlsx"

# Sortie CSV dans le même dossier
output_file = os.path.join(os.path.dirname(excel_file), "output.csv")
with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    
    # Entête
    writer.writerow(["Subject", "Date", "Start Time", "End Time", "Location", "Description"])

# Dictionnaire des horaires (clé = intitulé exact de la colonne, ex : "Lu Matin")
horaires = {
    "Lu Matin": ("08:30", "12:30"),
    "Lu Aprem": ("13:30", "17:30"),
    "Ma Matin": ("08:00", "12:00"),
    "Ma Aprem": ("13:30", "16:00"),
    "Me Matin": ("08:30", "12:30"),
    "Me Aprem": ("13:30", "17:30"),
    "Je Matin": ("08:30", "12:30"),
    "Je Aprem": ("13:30", "17:30"),
    "Ve Matin": ("08:30", "12:30"),
    "Ve Aprem": ("13:30", "17:30"),
}

# Dictionnaire de correspondance couleurs -> lieux
color_to_location = {
    "F8CBAD": "Salle UT2J sans ordi",
    "CCFFCC": "Salle ENSAT sans ordi",
    "99CCFF": "1003-Langue",
    "FF9933": "UT2J GS027",
    "FFCC66": "UT2J GS021",
    "E2F0D9": "703 (projet) ou alternance (entreprise)"
}

# Dictionnaire pour convertir un libellé de mois en nombre
months_fr = {
    "jan": 1, "janv": 1, "janv.": 1, "févr": 2, "fev": 2, "fev.": 2, "févr.": 2,
    "mars": 3, "mars.": 3, "avr": 4, "avr.": 4, "avril": 4,
    "mai": 5, "mai.": 5, "juin": 6, "juin.": 6,
    "juil": 7, "juil.": 7, "juillet": 7,
    "août": 8, "aout": 8, "aout.": 8, "août.": 8,
    "sept": 9, "sept.": 9, "septembre": 9,
    "oct": 10, "oct.": 10, "octobre": 10,
    "nov": 11, "nov.": 11, "novembre": 11,
    "dec": 12, "dec.": 12, "déc": 12, "déc.": 12, "décembre": 12
}

# Pour faire correspondre "Lu" -> 0 (lundi), "Ma" -> 1 (mardi), etc.
day_offsets = {
    "Lu": 0,
    "Ma": 1,
    "Me": 2,
    "Je": 3,
    "Ve": 4
}

# ---------------------------------------------------------------------------
# Chargement du fichier Excel
# ---------------------------------------------------------------------------
try:
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    ws = wb.active  # On suppose que la feuille active est celle à traiter
except Exception as e:
    print(f"Erreur : impossible d'ouvrir '{excel_file}'\n{e}")
    exit(1)

# ---------------------------------------------------------------------------
# Lecture de l'entête (ligne 1)
#    - Col. A = "date" ou quelque chose de similaire
#    - Colonnes B.. = "Lu Matin", "Lu Aprem", "Ma Matin", ...
# ---------------------------------------------------------------------------
headers = [cell.value for cell in ws[1] if cell.value is not None]
if len(headers) < 2:
    print("En-têtes insuffisantes dans la première ligne du fichier Excel.")
    exit(1)

# Les colonnes B, C, ... correspondent aux demi-journées
halfday_headers = headers[1:]  # on enlève la 1re colonne (date)

# ---------------------------------------------------------------------------
# Ouverture du fichier CSV en écriture
# ---------------------------------------------------------------------------
with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    # Écriture de la ligne d'en-tête
    writer.writerow(["Subject", "Date", "Start Time", "End Time", "Location", "Description"])

    # -----------------------------------------------------------------------
    # Parcours des lignes à partir de la 2e
    #    Colonne A : ex "11-15 sept 23"
    #    Colonnes B.. : "Lu Matin", "Lu Aprem", ...
    # -----------------------------------------------------------------------
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        week_cell = row[0]  # cellule de la colonne A
        if not week_cell or not week_cell.value:
            continue

        week_info = str(week_cell.value).strip().lower()
        # On cherche un motif du type "11-15 sept 23" ou "18-22 sept. 2023"
        # Regex : (\d+)\s*-\s*(\d+) -> jour début-fin
        #         ([a-zA-Zéû\.]+)   -> mois
        #         (\d+)             -> année
        match = re.search(r"(\d+)\s*-\s*(\d+)\s+([a-zA-Zéû\.]+)\s+(\d+)", week_info)
        if not match:
            # Format non reconnu
            print(f"Ligne {row_idx}: format de date non reconnu dans '{week_info}'")
            continue

        try:
            day_start = int(match.group(1))         # ex: 11
            # day_end = int(match.group(2))         # ex: 15 (inutile dans le code actuel)
            month_str = match.group(3).replace('.', '')  # ex: sept
            year_str = match.group(4)                    # ex: 23
            month = months_fr.get(month_str, None)
            if not month:
                print(f"Ligne {row_idx}: mois inconnu : {month_str}")
                continue
            # Convertit l'année (si c'est "23", on en fait 2023 ; sinon 2025, etc.)
            if len(year_str) == 2:
                year = 2000 + int(year_str)
            else:
                year = int(year_str)

            # On considère day_start comme le lundi de la semaine
            monday_date = datetime(year, month, day_start)
        except Exception as e:
            print(f"Ligne {row_idx}: impossible de parser la date -> {e}")
            continue

        # -------------------------------------------------------------------
        # Parcours des colonnes B.. (les demi-journées)
        # -------------------------------------------------------------------
        # row[1:] = colonnes B.. de la ligne
        for col_index, cell in enumerate(row[1:], start=1):
            # Le label de demi-journée (ex : "Lu Matin", "Ma Aprem", etc.)
            if col_index - 1 < len(halfday_headers):
                halfday_label = halfday_headers[col_index - 1]
            else:
                continue

            # Si la cellule est vide ET sans commentaire, on ignore
            if cell.value is None and cell.comment is None:
                continue

            # Jour "Lu", "Ma", ...
            day_abbr = halfday_label.split()[0]  # "Lu", "Ma", "Me", ...
            offset = day_offsets.get(day_abbr, None)
            if offset is None:
                # Pas un jour géré
                continue

            # Date de l'événement = lundi + offset
            event_date = monday_date + timedelta(days=offset)
            date_str = event_date.strftime("%Y-%m-%d")

            # Subject = texte de la cellule
            subject = str(cell.value).strip() if cell.value else ""

            # Description = commentaire (s'il existe)
            description = cell.comment.text.strip() if cell.comment else ""

            # Start Time / End Time via le dictionnaire horaires
            if halfday_label in horaires:
                start_time, end_time = horaires[halfday_label]
            else:
                start_time, end_time = ("", "")

            # Location = déterminée via la couleur de fond
            location = ""
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                rgb = cell.fill.fgColor.rgb  # ex: "FFE2F0D9"
                if rgb.startswith("FF") and len(rgb) == 8:
                    color_code = rgb[2:]  # ex: "E2F0D9"
                else:
                    color_code = rgb
                location = color_to_location.get(color_code, "")

            # Écriture dans le CSV
            writer.writerow([subject, date_str, start_time, end_time, location, description])

print(f"Fichier CSV généré : {output_file}")

