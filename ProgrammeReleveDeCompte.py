import openpyxl

def find_value_and_date(worksheet, current_cell, target_value, target_date):
    # Coordonnées de la case actuelle
    current_row, current_column = current_cell.row, current_cell.column

    # Parcourir les 10 cases au-dessus et en dessous de la case actuelle
    for row_offset in range(-10, 11):
        for col_offset in range(-10, 11):
            # Ignorer la case actuelle
            if row_offset == col_offset == 0:
                continue

            # Coordonnées de la case à vérifier
            row_to_check = current_row + row_offset
            col_to_check = current_column + col_offset

            # Récupérer la valeur et la date dans la case à vérifier
            cell_value = worksheet.cell(row=row_to_check, column=col_to_check).value

            # Vérifier si la valeur et la date correspondent aux cibles
            if cell_value == target_value:
                cell_date = worksheet.cell(row=row_to_check, column=col_to_check - 1).value
                if cell_date and cell_date.strftime("%d/%m") == target_date:
                    return worksheet.cell(row=row_to_check, column=col_to_check)

    return None

# Exemple d'utilisation
excel_file_path = 'relevécompteEXCel.xlsx'
sheet_name = 'Sheet0'
target_value = 223293501
target_date = "28/04"  # Date cible au format "DD/MM"

# Charger le classeur Excel
workbook = openpyxl.load_workbook(excel_file_path)

# Sélectionner la feuille appropriée
worksheet = workbook[sheet_name]

# Supposons que la cellule actuelle dont vous voulez partir est A1
current_cell = worksheet['B39']

# Rechercher la case contenant la valeur et la date cibles dans les cases adjacentes
result_cell = find_value_and_date(worksheet, current_cell, target_value, target_date)

if result_cell:
    print("Cellule trouvée :", result_cell.coordinate)
else:
    print("Aucune cellule trouvée contenant la valeur et la date cibles.")
