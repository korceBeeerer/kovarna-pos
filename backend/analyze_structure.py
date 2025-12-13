import openpyxl

# Načti Excel soubor
excel_path = r"C:\Users\martin.korcak\Downloads\Tabulka tržeb.xlsx"

try:
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Tržby"]

    print("=" * 80)
    print("ANALÝZA STRUKTURY - HLEDÁNÍ BLOKŮ PRO DNY")
    print("=" * 80)

    # Najdi všechny hlavičky (obsahují "Produkt" ve sloupci B)
    header_rows = []
    for row_idx in range(1, ws.max_row + 1):
        cell_b = ws.cell(row=row_idx, column=2).value
        if cell_b and str(cell_b).strip().lower() == "produkt":
            header_rows.append(row_idx)

    print(f"\nNalezeno {len(header_rows)} bloků (dnů) v měsíci")
    print("-" * 80)

    # Zobraz prvních 5 bloků
    for i, header_row in enumerate(header_rows[:5], 1):
        print(f"\nDEN {i} - Hlavička na řádku {header_row}")

        # Zobraz produkty v tomto bloku
        products_in_block = []
        row_idx = header_row + 1

        while row_idx <= ws.max_row:
            product_name = ws.cell(row=row_idx, column=2).value
            price = ws.cell(row=row_idx, column=3).value

            # Konec bloku - další hlavička nebo prázdný řádek
            if product_name and str(product_name).strip().lower() == "produkt":
                break

            # Pokud je produkt nebo cena, přidej
            if product_name or price:
                products_in_block.append({
                    'row': row_idx,
                    'name': product_name,
                    'price': price
                })
                row_idx += 1
            else:
                # Několik prázdných řádků = konec bloku
                row_idx += 1
                if row_idx > header_row + 60:  # Max 60 řádků na blok
                    break

        print(f"  Počet položek v bloku: {len(products_in_block)}")
        print(f"  První 5 produktů:")
        for prod in products_in_block[:5]:
            name = str(prod['name'])[:30] if prod['name'] else "(prázdný)"
            price = prod['price'] if prod['price'] else 0
            print(f"    Řádek {prod['row']:4}: {name:30} - {price} Kč")

    # Spočítej vzdálenost mezi bloky
    if len(header_rows) > 1:
        print(f"\n{'=' * 80}")
        print("VZDALENOSTI MEZI BLOKY:")
        print("-" * 80)
        for i in range(min(5, len(header_rows) - 1)):
            distance = header_rows[i+1] - header_rows[i]
            print(f"  Den {i+1} -> Den {i+2}: {distance} radku")

    # Načti seznam všech produktů z prvního bloku (den 1)
    print(f"\n{'=' * 80}")
    print("KOMPLETNÍ SEZNAM PRODUKTŮ (z prvního dne):")
    print("-" * 80)

    if header_rows:
        first_header = header_rows[0]
        row_idx = first_header + 1
        all_products = []

        while row_idx <= ws.max_row:
            product_name = ws.cell(row=row_idx, column=2).value
            price = ws.cell(row=row_idx, column=3).value

            # Konec bloku
            if product_name and str(product_name).strip().lower() == "produkt":
                break

            if row_idx > first_header + 60:
                break

            # Přidej všechny řádky (i prázdné)
            all_products.append({
                'row': row_idx,
                'name': product_name,
                'price': price
            })
            row_idx += 1

        print(f"\nCelkem řádků v bloku: {len(all_products)}")
        print("\nVšechny produkty:")
        for i, prod in enumerate(all_products, 1):
            name = str(prod['name'])[:30] if prod['name'] else "(prázdný)"
            price = prod['price'] if prod['price'] else 0
            print(f"  {i:2}. Řádek {prod['row']:4}: {name:30} - {price} Kč")

    wb.close()

except Exception as e:
    print(f"CHYBA: {str(e)}")
    import traceback
    traceback.print_exc()
