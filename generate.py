import pandas as pd

def generate_excel_files():
    """
    Tato funkce vytvoří dva testovací Excel soubory: zdroj.xlsx a cil.xlsx.
    """
    
    # --- Data pro zdrojový soubor (více sloupců) ---
    source_data = [
        {'ID': 101, 'Jmeno': 'Pavel', 'Prijmeni': 'Novák', 'Vek': 45, 'Mesto': 'Praha', 'Pozice': 'Manažer', 'Telefon': '123-456-789', 'Datum_Nastupu': '2020-01-15'},
        {'ID': 102, 'Jmeno': 'Jana', 'Prijmeni': 'Svobodová', 'Vek': 31, 'Mesto': 'Brno', 'Pozice': 'Analytik', 'Telefon': '987-654-321', 'Datum_Nastupu': '2021-03-22'},
        {'ID': 103, 'Jmeno': 'Tomáš', 'Prijmeni': 'Dvořák', 'Vek': 28, 'Mesto': 'Ostrava', 'Pozice': 'Programátor', 'Telefon': '555-666-777', 'Datum_Nastupu': '2022-11-01'},
        {'ID': 104, 'Jmeno': 'Eva', 'Prijmeni': 'Černá', 'Vek': 52, 'Mesto': 'Plzeň', 'Pozice': 'Účetní', 'Telefon': '111-222-333', 'Datum_Nastupu': '2018-07-19'},
    ]
    df_source = pd.DataFrame(source_data)
    
    # --- Data pro cílový soubor (méně sloupců, některé stejné) ---
    # Tento soubor již obsahuje nějaká data, abychom simulovali přidávání nového řádku.
    target_data = [
        {'ID': 901, 'Jmeno': 'František', 'Prijmeni': 'Veselý', 'Mesto': 'Liberec', 'Stav_Zaznamu': 'Archivováno'},
        {'ID': 902, 'Jmeno': 'Lucie', 'Prijmeni': 'Pokorná', 'Mesto': 'Pardubice', 'Stav_Zaznamu': 'Aktivní'},
    ]
    # Názvy sloupců záměrně v jiném pořadí, abychom otestovali robustnost
    df_target = pd.DataFrame(target_data, columns=['ID', 'Jmeno', 'Prijmeni', 'Mesto', 'Stav_Zaznamu'])

    source_filename = 'zdroj.xlsx'
    target_filename = 'cil.xlsx'

    print("Generuji testovací Excel soubory...")

    try:
        # Uložení zdrojového souboru
        # index=False je důležité, aby se do Excelu neukládal číselný index z pandas
        df_source.to_excel(source_filename, index=False, engine='openpyxl')
        print(f"Úspěšně vytvořen soubor: {source_filename}")

        # Uložení cílového souboru
        df_target.to_excel(target_filename, index=False, engine='openpyxl')
        print(f"Úspěšně vytvořen soubor: {target_filename}")

        print("\nHotovo! Nyní můžete spustit svou hlavní aplikaci a použít tyto soubory.")
        print(f"Zkuste například vyhledat ID '102' ze souboru '{source_filename}' a zkopírovat ho do '{target_filename}'.")

    except PermissionError:
        print("\nCHYBA: Nemohu zapsat soubory. Ujistěte se, že soubory 'zdroj.xlsx' nebo 'cil.xlsx' nejsou otevřené v Excelu.")
    except Exception as e:
        print(f"\nCHYBA: Vyskytla se neočekávaná chyba: {e}")


# Spuštění generátoru, pokud je tento skript spuštěn přímo
if __name__ == "__main__":
    generate_excel_files()