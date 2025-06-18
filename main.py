import PySimpleGUI as sg # type: ignore
import pandas as pd
from openpyxl import load_workbook


# LAYOUT
sg.theme('SystemDefault1')
layout = [
    [sg.Text("Zadejte ID pro vyhledání:", size=(22, 1)), sg.InputText(key='-ID-')],
    [sg.Text("Zdrojový Excel soubor:", size=(22, 1)), 
     sg.Input(key='-SOURCE_FILE-'), 
     sg.FileBrowse("Vybrat...", file_types=(("Excel Files", "*.xlsx;*.xls"),))],
    [sg.Text("Cílový Excel soubor:", size=(22, 1)), 
    sg.Input(key='-TARGET_FILE-', enable_events=True),
    sg.FileBrowse("Vybrat...", file_types=(("Excel Files", "*.xlsx;*.xls"),))],
    [sg.Text("Vyberte cílový list:", size=(22, 1)), 
    sg.Combo([], key='-SHEET_NAME-', size=(40, 1), readonly=True)],
    [sg.Button("Zpracovat data", key='-PROCESS-'), sg.Button("Konec")],
    [sg.Text("Průběh a informace:")],
    [sg.Multiline(size=(85, 15), key='-OUTPUT-', disabled=True, autoscroll=True, reroute_stdout=True)]
]
window = sg.Window("Excel to Excel", layout)


# EVENT LOOP 
while True:
    event, values = window.read()
    if event == '-TARGET_FILE-':
        target_path = values['-TARGET_FILE-']
        if target_path:
            try:
                xls = pd.ExcelFile(target_path)
                sheet_names = xls.sheet_names
                window['-SHEET_NAME-'].update(values=sheet_names, set_to_index=0)
            except Exception as e:
                window['-SHEET_NAME-'].update(values=[], value='')
                print(f"CHYBA při čtení listů: {e}")
    if event == sg.WIN_CLOSED or event == 'Konec':
        break
    if event == '-PROCESS-':
        id_to_find = values['-ID-']
        source_path = values['-SOURCE_FILE-']
        target_path = values['-TARGET_FILE-']
        if not id_to_find or not source_path or not target_path:
            print("CHYBA: Všechna pole musí být vyplněna.")
            continue
        print(f" Hledané ID: {id_to_find}")
        print(f" Zdrojový soubor: {source_path}")
        print(f" Cílový soubor: {target_path}")
        print("-" * 30)
    

        # LOAD FILE
        df_source = None
        try:
            df_source = pd.read_excel(source_path, engine="openpyxl")

        except FileNotFoundError:
            print(f"CHYBA: Zdrojový soubor nebyl nalezen na cestě: {source_path}")
            continue
        except Exception as e:
            print(f"CHYBA: Nepodařilo se načíst zdrojový soubor.")
            print(f"Podrobnosti chyby: {e}")


        # FIND ROW
        SOURCE_ID_COLUMN_NAME = 'ID' # CHANGE TO EXCEL ID COLUMN!!
        COLUMN_MAPPING = {
    'ID': 'F42 ID',
    'Name': 'Function',
    'Name (English)': 'Function (English)',
    'Implementation managers': 'FuReV',
    'FuLi contact person': 'FuReV support',
    'Einsatz zu': 'Einsatz zu',
    'Entfall zu': 'Entfall zu',
    'Funktionscluster (VW) / Solution (CARIAD)': 'Cluster'
}
        found_row = None
        try:
            id_to_find_str = str(id_to_find)
            source_id_column_str = df_source[SOURCE_ID_COLUMN_NAME].astype(str)
            found_data = df_source[source_id_column_str == id_to_find_str]
            if found_data.empty:
                print(f"CHYBA: ID '{id_to_find}' nebylo ve zdrojovém souboru nalezeno.")
                print("-" * 30)
                continue
            found_row = found_data.iloc[0]
            print("Záznam s daným ID byl úspěšně nalezen.")

        except KeyError:
            print(f"CHYBA: V souboru chybí sloupec s názvem '{SOURCE_ID_COLUMN_NAME}'.")
            continue
        except Exception as e:
            print(f"CHYBA: Vyskytla se chyba při hledání dat: {e}")
            continue
        

        # PREPARE ROW AND SAVE WITH OPENPYXL, EXTEND TABLE
        selected_sheet_name = values['-SHEET_NAME-']
        if not selected_sheet_name:
            print("CHYBA: Nebyl vybrán žádný cílový list.")
            continue
        try:
            workbook = load_workbook(target_path)
            sheet = workbook[selected_sheet_name]
            new_row_dict = {}

            for source_col, target_col in COLUMN_MAPPING.items():
                if source_col in found_row and target_col in [cell.value for cell in sheet[1]]:
                    new_row_dict[target_col] = found_row[source_col]

            if not new_row_dict:
                print("CHYBA: Nepodařilo se zkopírovat žádná data.")
                continue

            target_columns_order = [cell.value for cell in sheet[1]]
            new_row_values = [new_row_dict.get(col_name, '') for col_name in target_columns_order]
            sheet.append(new_row_values)
            print("Nový řádek byl přidán.")

            if sheet.tables:
                table_name = list(sheet.tables.keys())[0]
                table = sheet.tables[table_name]
                old_range = table.ref
                new_max_row = sheet.max_row
                new_range = f"{old_range.split(':')[0]}:{old_range.split(':')[1][0]}{new_max_row}"
                from openpyxl.utils import get_column_letter
                start_cell, end_cell = old_range.split(':')
                end_col_letter = get_column_letter(sheet.max_column)
                new_range = f"{start_cell}:{end_col_letter}{new_max_row}"
                table.ref = new_range
                print(f"Rozsah tabulky '{table_name}' byl aktualizován na '{new_range}'.")

            workbook.save(target_path)
            
            print("\nHotovo! Data byla úspěšně zkopírována.")

        except FileNotFoundError:
            print(f"CHYBA: Cílový soubor nebyl nalezen na cestě: {target_path}")
            continue
        except PermissionError:
            print("CHYBA: Cílový soubor je pravděpodobně otevřený v Excelu.")
            print("Zavřete ho a zkuste operaci znovu.")
            continue
        except Exception as e:
            print(f"CHYBA: Nepodařilo se uložit data do cílového souboru: {e}")
            continue

        print("-" * 30)

window.close()