import openpyxl

def find_value_and_date(worksheet, current_cell, target_value, target_date):
    # Coordonnées de la case actuelle
    current_row= current_cell.row
    # print( "current row")
    # print(current_row)
    # print("target value")
    # print(target_value)
    # # Parcourir les 10 cases au-dessus et en dessous de la case actuelle
    for row_offset in range(-15, 16):
            # Ignorer la case actuelle
            if row_offset == 0:
                continue

            # Coordonnées de la case à vérifier
            row_to_check = current_row + row_offset
            #print('row to check ')
            #print(row_to_check)
            # Récupérer la valeur et la date dans la case à vérifier
            cell_value = worksheet.cell(row=row_to_check, column=2).value
            #print('cell value')
            #print(cell_value)
            # Vérifier si la valeur et la date correspondent aux cibles
            if (target_value in cell_value) & (target_date in cell_value):
                cell_date = worksheet.cell(row=row_to_check, column=2).value
                #print("cell date")
                #print(cell_date)
                print(worksheet.cell(row=row_to_check, column=2))

    return None

# Exemple d'utilisation
excel_file_path = 'relevécompteEXCel.xlsx'
sheet_name = 'Sheet0'
target_value = '2232935'
target_date = "30/05"  # Date cible au format "DD/MM"

# Charger le classeur Excel
workbook = openpyxl.load_workbook(excel_file_path)

# Sélectionner la feuille appropriée
worksheet = workbook[sheet_name]

# Supposons que la cellule actuelle dont vous voulez partir est A1
current_cell = worksheet['B39']

print("target date = ")
print(target_date)

# Rechercher la case contenant la valeur et la date cibles dans les cases adjacentes
result_cell = find_value_and_date(worksheet, current_cell, target_value, target_date)

if result_cell:
    print("Cellule trouvée :", result_cell.coordinate)
else:
    print("Aucune cellule trouvée contenant la valeur et la date cibles.")
