import os
import re
import pandas as pd
from pdfminer.high_level import extract_text
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
from datetime import datetime

# Define paths
PDF_FOLDER = os.path.join(os.path.dirname(__file__), 'pdfs')
EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'Trading.xlsx')

def extract_trade_info_transaction_statement(pdf_text: str):
    """Parser für Baader/Scalable 'Transaction Statement: Sale/Purchase'-PDFs."""
    trade_data = {"Source": "transaction_statement"}

    # 1) Trade-Typ UND Datum direkt aus dem Kopf:
    #    "Transaction Statement: Sale\n2025-11-13\n18:04:24:00\n..."
    m_head = re.search(
        r'Transaction Statement:\s*(Sale|Purchase|Buy)\s+(\d{4}-\d{2}-\d{2})',
        pdf_text,
        re.IGNORECASE
    )
    if not m_head:
        print("Could not find transaction header in transaction statement PDF.")
        return None

    word, date_iso = m_head.groups()
    word = word.lower()
    trade_data["Trade Type"] = "Sell" if word == "sale" else "Buy"

    try:
        year, month, day = date_iso.split("-")
        trade_data["Date"] = f"{day}.{month}.{year}"
    except Exception:
        print(f"Could not convert header date '{date_iso}' in transaction statement PDF.")
        return None

    # 2) Block um "Quantity" herum heraus schneiden (Quantity + Asset)
    start = pdf_text.find("Quantity")
    if start == -1:
        print("Could not find 'Quantity' block in transaction statement PDF.")
        return None

    end_candidates = []
    for marker in ["Account Owner", "Market Value", "Value:", "Custody Account"]:
        idx = pdf_text.find(marker, start)
        if idx != -1:
            end_candidates.append(idx)

    end = min(end_candidates) if end_candidates else (start + 400)
    block_text = pdf_text[start:end]
    block_lines = block_text.splitlines()

    # 2a) Stückzahl aus einer Zeile mit "Units <Zahl>"
    quantity = None
    for line in block_lines:
        m = re.search(r'Units\s+([\d\.,]+)', line)
        if m:
            try:
                quantity = float(m.group(1).replace(",", "."))
                break
            except ValueError:
                continue

    if quantity is None:
        print("Could not parse quantity from transaction statement PDF.")
        return None

    trade_data["Quantity"] = quantity

    # 2b) Asset-Name:
    #     Wir sammeln alle Zeilen im Quantity-Block, die keine Metadaten sind,
    #     und nehmen die letzte sinnvolle Zeile ohne ":" (z.B. "HVB Put 17.12.25 NVIDIA 200")
    candidates = []
    for line in block_lines:
        s = line.strip()
        if not s:
            continue
        if s.startswith("Quantity") or s.startswith("Units"):
            continue
        candidates.append(s)

    # zuerst alle ohne ":", falls vorhanden
    no_colon = [s for s in candidates if ":" not in s]
    if no_colon:
        asset_name = no_colon[-1]
    else:
        asset_name = candidates[-1] if candidates else None

    trade_data["Asset Name"] = asset_name

    # 3) Preis pro Einheit:
    #    Wir suchen nach einem "Price"-Label und nehmen die erste Zeile danach,
    #    die wie eine Zahl aussieht (z.B. "1.69")
    lines = pdf_text.splitlines()
    price = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Price"):
            for j in range(i + 1, min(i + 10, len(lines))):
                s = lines[j].strip()
                if re.match(r'^-?\d+[.,]\d+$', s):  # einfache Dezimalzahl
                    try:
                        price = float(s.replace(",", "."))
                        break
                    except ValueError:
                        continue
            if price is not None:
                break

    if price is None:
        print("Could not find price in transaction statement PDF (Price block).")
        return None

    trade_data["Price per Unit"] = price

    # 4) Steuern:
    #    Im Statement gibt es einen Block wie:
    #    43.94
    #    1.84 -
    #    0.16 -
    #    0.10 -
    #    41.84
    #    Wir nehmen die erste Gruppe von Zeilen mit "Zahl -"
    taxes = 0.0
    tax_start = None
    for idx, line in enumerate(lines):
        if re.search(r'([0-9]+[.,][0-9]+)\s*-', line):
            tax_start = idx
            break

    if tax_start is not None:
        for j in range(tax_start, min(tax_start + 10, len(lines))):
            m = re.search(r'([0-9]+[.,][0-9]+)\s*-', lines[j])
            if m:
                try:
                    val = float(m.group(1).replace(",", "."))
                    taxes += abs(val)
                except ValueError:
                    continue

    trade_data["Taxes"] = taxes

    return trade_data

def extract_trade_info_cost_information(pdf_text: str):
    """Parser für Ex-Ante cost information PDFs."""
    trade_data = {"Source": "cost_information"}

    # Trade Type
    if re.search(r'\bOrder\s+Buy\b', pdf_text, re.IGNORECASE) or re.search(r'\bBuy\b', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Buy'
    elif re.search(r'\bOrder\s+Sell\b', pdf_text, re.IGNORECASE) or re.search(r'\bSell\b', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Sell'
    else:
        print("Trade type not found in cost information PDF.")
        return None

    # Asset Name: Zeile nach "Ex-Ante cost information"
    asset_match = re.search(r'Ex-Ante cost information\s*\n\s*(.+)', pdf_text, re.IGNORECASE)
    if asset_match:
        trade_data['Asset Name'] = asset_match.group(1).strip()
    else:
        print("Asset Name not found in cost information PDF.")
        return None

    # Datum – bevorzugt "Date 08.12.2025", sonst Fallback auf dein altes Muster
    date_match = re.search(r'Date\s+(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if not date_match:
        date_match = re.search(r'Date\s*(?:\n.*?){3}\n\s*(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if date_match:
        trade_data['Date'] = date_match.group(1).strip()
    else:
        print("Date not found in cost information PDF.")
        return None

    # Stückzahl – "Quantity 100 Shr."
    quantity_match = re.search(r'Quantity\s+([\d\.,]+)\s*Shr\.', pdf_text)
    if not quantity_match:
        quantity_match = re.search(r'([\d\.,]+)\s*Shr\.', pdf_text)
    if quantity_match:
        quantity = float(quantity_match.group(1).replace(',', '.'))
        trade_data['Quantity'] = quantity
    else:
        print("Quantity not found in cost information PDF.")
        return None

    # Preis pro Stück: Zeile nach "Shr." enthält die geschätzte Order-Summe (z.B. 119.00 €):contentReference[oaicite:0]{index=0}
    price_match = re.search(r'Shr\.\s*\n\s*([\d\.,]+)\s*€', pdf_text)
    if price_match:
        total_order_amount = float(price_match.group(1).replace(',', '.'))
        trade_data['Price per Unit'] = total_order_amount / quantity
    else:
        print("Price / Est. order amount not found in cost information PDF.")
        return None

    # In Cost-Info PDFs keine Steuer -> 0
    trade_data['Taxes'] = 0.0

    return trade_data


def extract_trade_info_contract_note(pdf_text: str):
    """Parser für Contract-Note PDFs (mit Steuern)."""
    trade_data = {"Source": "contract_note"}

    # Zeile beginnt mit "Buy ..." oder "Sell ..."
    trade_line = re.search(r'^(Buy|Sell)\s+(.+)$', pdf_text, re.MULTILINE)
    if not trade_line:
        print("Trade line with Buy/Sell not found in contract note PDF.")
        return None

    trade_data['Trade Type'] = trade_line.group(1)
    trade_data['Asset Name'] = trade_line.group(2).strip()

    # Datum: zuerst "Execution 08.12.2025", ansonsten "Date 08.12.2025"
    date_match = re.search(r'Execution\s+(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if not date_match:
        date_match = re.search(r'Date\s+(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if date_match:
        trade_data['Date'] = date_match.group(1).strip()
    else:
        print("Date not found in contract note PDF.")
        return None

    # Menge + Preis: "100.00 pc. 1.26 EUR 126.00 EUR"
    qp_match = re.search(r'([\d\.,]+)\s*pc\.\s+([\d\.,]+)\s*EUR\s+([\d\.,]+)\s*EUR', pdf_text)
    if qp_match:
        quantity = float(qp_match.group(1).replace(',', '.'))
        price_per_unit = float(qp_match.group(2).replace(',', '.'))
        trade_data['Quantity'] = quantity
        trade_data['Price per Unit'] = price_per_unit
    else:
        print("Quantity/Price line not found in contract note PDF.")
        return None

    # Gebühren (Order fees +/-0.99 EUR)
    fee_match = re.search(r'Order fees\s*([+-]?[\d\.,]+)\s*EUR', pdf_text)
    if fee_match:
        fee_value = float(fee_match.group(1).replace(',', '.'))
        trade_data['Order Fee'] = abs(fee_value)
    else:
        trade_data['Order Fee'] = 0.0

    # Steuern – nur beim Sell: "Taxes -1.98 EUR":contentReference[oaicite:5]{index=5}
    tax_match = re.search(r'Taxes\s*([+-]?[\d\.,]+)\s*EUR', pdf_text)
    if tax_match:
        tax_value = float(tax_match.group(1).replace(',', '.'))
        trade_data['Taxes'] = abs(tax_value)
    else:
        trade_data['Taxes'] = 0.0

    return trade_data


def extract_trade_info(pdf_text: str):
    """Router: entscheidet, welcher Parser genutzt wird."""
    try:
        if re.search(r'Contract note', pdf_text, re.IGNORECASE):
            parser = extract_trade_info_contract_note
        elif re.search(r'Transaction Statement:', pdf_text, re.IGNORECASE):
            parser = extract_trade_info_transaction_statement
        elif re.search(r'Ex-Ante cost information', pdf_text, re.IGNORECASE):
            parser = extract_trade_info_cost_information
        else:
            print("Unrecognized PDF type (neither contract note, transaction statement nor cost info).")
            return None

        trade_data = parser(pdf_text)
        return trade_data

    except Exception as e:
        # Etwas hilfreichere Fehlermeldung, falls doch mal was crasht
        print(f"Error while parsing PDF with {parser.__name__}: {e}")
        snippet = pdf_text[:400].replace('\n', ' ')
        print(f"PDF snippet (first 400 chars): {snippet} ...")
        return None

def get_column_letter_for_header(sheet, header_name: str):
    """Gibt den Spaltenbuchstaben zur Überschrift in Zeile 1 zurück."""
    for cell in sheet[1]:
        if cell.value == header_name:
            try:
                return cell.column_letter  # neuere openpyxl
            except AttributeError:
                return get_column_letter(cell.column)
    return None


def find_next_empty_row_in_column(sheet, column_letter):
    """Find the next empty row in the specified column that has actual data."""
    for row in range(1, sheet.max_row + 1):
        if not sheet[f'{column_letter}{row}'].value:  # Check if the cell in column A is empty
            return row
    return sheet.max_row + 1  # If no empty row is found, return the next available row


def apply_date_format_to_column(sheet, column_letter):
    """Ensure the entire column has the same date format applied."""
    date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
    
    # If the style already exists, don't recreate it
    if "date_style" not in sheet.parent.named_styles:
        sheet.parent.add_named_style(date_style)

    for row in range(2, sheet.max_row + 1):  # Assuming row 1 is headers
        if isinstance(sheet[f'{column_letter}{row}'].value, datetime):
            sheet[f'{column_letter}{row}'].style = date_style


def update_excel(trade_data):
    # Read the existing Excel file or create a new DataFrame
    if os.path.exists(EXCEL_FILE):
        workbook = load_workbook(EXCEL_FILE)
        sheet = workbook.active
    else:
        print("Excel not found")
        return

    trade_type = trade_data['Trade Type']
    asset_name = trade_data['Asset Name']

    # Ensure date format is applied correctly
    date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")

    # Check if the date_style already exists to avoid the ValueError
    if "date_style" not in workbook.named_styles:
        workbook.add_named_style(date_style)

    # NEU: Spalte "Taxes" suchen (Header in Zeile 1 muss exakt so heißen)
    taxes_col = get_column_letter_for_header(sheet, 'Taxes')

    if trade_type == 'Buy':
        next_row = find_next_empty_row_in_column(sheet, 'A')

        sheet[f'A{next_row}'] = trade_data['Asset Name']
        buy_date = trade_data['Date']
        if isinstance(buy_date, str):
            buy_date = datetime.strptime(buy_date, "%d.%m.%Y").date()
        sheet[f'B{next_row}'] = buy_date
        sheet[f'B{next_row}'].style = "date_style"
        sheet[f'C{next_row}'] = trade_data['Quantity']
        sheet[f'D{next_row}'] = trade_data['Price per Unit']

        # Optional: Buy-Zeile mit 0 Steuern initialisieren
        if taxes_col is not None:
            sheet[f'{taxes_col}{next_row}'] = trade_data.get('Taxes', 0.0)

    elif trade_type == 'Sell':
        # Find the row that matches the Buy entry and doesn't have a Sell Date yet
        for row in range(2, sheet.max_row + 1):  # Assuming row 1 is headers
            if sheet[f'A{row}'].value == asset_name and not sheet[f'F{row}'].value:
                sell_date = trade_data['Date']
                if isinstance(sell_date, str):
                    sell_date = datetime.strptime(sell_date, "%d.%m.%Y").date()
                sheet[f'F{row}'] = sell_date
                sheet[f'F{row}'].style = "date_style"
                sheet[f'G{row}'] = trade_data['Quantity']
                sheet[f'H{row}'] = trade_data['Price per Unit']

                # NEU: Steuer in die "Taxes"-Spalte eintragen, falls vorhanden
                if taxes_col is not None and 'Taxes' in trade_data:
                    sheet[f'{taxes_col}{row}'] = trade_data['Taxes']
                break
        else:
            print(f"No matching Buy entry found for asset {asset_name}.")

    workbook.save(EXCEL_FILE)
    print(f"Excel file updated successfully with trade data for {trade_data['Asset Name']}.")

def process_pdfs():
    if not os.path.exists(PDF_FOLDER):
        print(f"PDF folder {PDF_FOLDER} not found")
        return

    trades_by_key = {}

    for filename in os.listdir(PDF_FOLDER):
        if not filename.lower().endswith('.pdf'):
            continue

        pdf_path = os.path.join(PDF_FOLDER, filename)
        print(f"Processing {filename}...")

        try:
            pdf_text = extract_text(pdf_path)
        except Exception as e:
            print(f"Failed to read {filename}: {e}")
            continue

        trade_data = extract_trade_info(pdf_text)
        if not trade_data:
            print(f"Failed to extract data from {filename}.")
            continue

        # Für Debugging ggf. hilfreich
        trade_data['_filename'] = filename

        # Schlüssel, um denselben Trade wiederzuerkennen
        key = (
            trade_data['Trade Type'],
            trade_data['Asset Name'],
            trade_data['Date'],
            trade_data['Quantity'],
        )

        existing = trades_by_key.get(key)
        if existing is None:
            trades_by_key[key] = trade_data
        else:
            # Wenn Cost-Info UND Contract-Note existieren:
            # Contract-Note bevorzugen (weil exakte Gebühren/Steuern)
            if trade_data.get('Source') == 'contract_note' and existing.get('Source') != 'contract_note':
                trades_by_key[key] = trade_data

        # PDF nach erfolgreichem Einlesen löschen (wie bisher)
        try:
            os.remove(pdf_path)
            print(f"{filename} processed and deleted.")
        except OSError as e:
            print(f"Could not delete {filename}: {e}")

    # Jetzt alle Trades in sinnvoller Reihenfolge verarbeiten:
    # 1) alle Buys, 2) alle Sells, jeweils chronologisch nach Datum
    def date_key(date_str: str):
        try:
            d, m, y = map(int, date_str.split("."))
            return (y, m, d)
        except Exception:
            return (9999, 12, 31)

    trades = list(trades_by_key.values())
    trades.sort(key=lambda t: (0 if t['Trade Type'] == 'Buy' else 1, date_key(t['Date'])))

    for t in trades:
        try:
            update_excel(t)
        except Exception as e:
            print(f"Error while updating Excel for trade from file {t.get('_filename')}: {e}")
            print("Trade data that failed:", t)


if __name__ == "__main__":
    process_pdfs()
