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

def parse_number(text):
    """Wandelt Strings wie '1.10', '0,06', '72.00' in float um."""
    if not text:
        return 0.0
    text = text.replace('EUR', '').replace('€', '')
    text = text.replace('pc.', '').replace('pc', '')
    text = text.strip()
    # Deutsch: 1.234,56  -> 1234.56
    if ',' in text and '.' in text and text.find('.') < text.find(','):
        text = text.replace('.', '').replace(',', '.')
    else:
        text = text.replace(',', '.')
    return float(text)


def normalize_asset_name(name):
    """Macht aus dem Titel eine Token-Menge, um fuzzy zu matchen."""
    if not isinstance(name, str):
        return set()
    name = name.lower()
    # offensichtlichen Kram raus
    for ch in ',.$€':
        name = name.replace(ch, ' ')
    cleaned = ''.join(
        ch if (ch.isalnum() or ch.isspace()) else ' '
        for ch in name
    )
    tokens = [t for t in cleaned.split() if t]
    return set(tokens)


def find_best_matching_row(sheet, asset_name, require_empty_sell_date=True, min_score=0.5):
    """
    Sucht die Zeile mit dem besten Token-Overlap.
    require_empty_sell_date=True: nur Zeilen ohne Sell-Datum (für Cost-Info-Sells)
    require_empty_sell_date=False: auch Closed Trades (für Contract-Notes mit Steuern).
    """
    target_tokens = normalize_asset_name(asset_name)
    best_row = None
    best_score = 0.0

    for row in range(2, sheet.max_row + 1):
        if require_empty_sell_date and sheet[f'F{row}'].value:
            # schon verkauft -> für Cost-Info-Sell ignorieren
            continue

        row_name = sheet[f'A{row}'].value
        row_tokens = normalize_asset_name(row_name)
        if not row_tokens:
            continue

        inter = len(target_tokens & row_tokens)
        union = len(target_tokens | row_tokens)
        if union == 0:
            continue

        score = inter / union
        if score > best_score:
            best_score = score
            best_row = row

    if best_score < min_score:
        return None
    return best_row

def extract_cost_info_trade(pdf_text):
    trade_data = {}

    # Quelle markieren
    trade_data['Source'] = 'cost_info'

    # Determine if it's a Buy or Sell
    if re.search(r'Buy', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Buy'
    elif re.search(r'Sell', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Sell'
    else:
        print("Trade type not found.")
        return None

    # Asset Name aus Ex-Ante cost information
    asset_match = re.search(r'Ex-Ante cost information\s*\n\s*(.+)', pdf_text)
    if asset_match:
        trade_data['Asset Name'] = asset_match.group(1).strip()
    else:
        print("Asset Name/ISIN not found.")
        return None

    # Date
    date_match = re.search(r'Date\s*(?:\n.*?){3}\n\s*(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if date_match:
        trade_data['Date'] = date_match.group(1).strip()
    else:
        print("Date not found.")
        return None

    # Quantity
    quantity_match = re.search(r'([\d,\.]+)\s*Shr\.', pdf_text)
    if quantity_match:
        trade_data['Quantity'] = float(quantity_match.group(1).replace(',', '.').strip())
    else:
        print("Quantity not found.")
        return None

    # Price per Unit (aus Gesamtbetrag / Stückzahl)
    price_match = re.search(r'Shr\.\s*\n\s*([\d,\.]+)\s*€', pdf_text)
    if price_match:
        total_order_amount = float(price_match.group(1).replace(',', '.').strip())
        trade_data['Price per Unit'] = total_order_amount / trade_data['Quantity']
    else:
        print("Price per unit not found.")
        return None

    trade_data['Fees'] = 0.99
    print(trade_data)

    return trade_data

def extract_contract_note_trade(pdf_text):
    trade_data = {}
    trade_data['Source'] = 'contract_note'

    # --- Fall 1: Scalable Contract note (neues Format) ---
    if 'Contract note' in pdf_text:
        # Zeile nach "Type Security Quantity Price Amount"
        sec_line_match = re.search(
            r'Type Security Quantity Price Amount\s*\n([^\n]+)',
            pdf_text
        )
        if not sec_line_match:
            print("Could not find security line in Scalable contract note.")
            return None

        first_line = sec_line_match.group(1).strip()
        parts = first_line.split(maxsplit=1)
        if len(parts) < 2:
            print("Unexpected security line format:", first_line)
            return None

        trade_word, asset_name = parts[0], parts[1]
        trade_data['Trade Type'] = 'Buy' if trade_word.lower().startswith('buy') else 'Sell'
        trade_data['Asset Name'] = asset_name.strip()

        # Datum: "Execution 08.12.2025 15:31:57"
        date_match = re.search(r'Execution\s+(\d{2}\.\d{2}\.\d{4})', pdf_text)
        if date_match:
            trade_data['Date'] = date_match.group(1).strip()

        # Menge & Preis: Zeile mit "72.00 pc. 1.74 EUR ..."
        qp_match = re.search(r'\n([\d\.,]+)\s*pc\.?\s*([\d\.,]+)\s*EUR', pdf_text)
        if qp_match:
            qty = parse_number(qp_match.group(1))
            price = parse_number(qp_match.group(2))
            trade_data['Quantity'] = qty
            trade_data['Price per Unit'] = price

        # Steuern von Seite 2
        cg_match = re.search(r'Capital gains tax\s+([\d\.,]+)\s*EUR', pdf_text)
        soli_match = re.search(r'Solidarity surcharge\s+([\d\.,]+)\s*EUR', pdf_text)
        church_match = re.search(r'Church tax\s+([\d\.,]+)\s*EUR', pdf_text)

        if cg_match and soli_match and church_match:
            cg = parse_number(cg_match.group(1))
            soli = parse_number(soli_match.group(1))
            church = parse_number(church_match.group(1))
            trade_data['Capital Gains Tax'] = cg
            trade_data['Solidarity Surcharge'] = soli
            trade_data['Church Tax'] = church
            trade_data['Taxes Sum'] = cg + soli + church

        # Order Fees (falls du sie später in J eintragen willst)
        fee_match = re.search(r'Order fees\s+(-?[\d\.,]+)\s*EUR', pdf_text)
        if fee_match:
            fee = abs(parse_number(fee_match.group(1)))
            trade_data['Order Fee'] = fee

        return trade_data

    # --- Fall 2: Baader / WPABRECHNUNG-060.114 ---
    if 'Transaction Statement: Sale' in pdf_text or 'Transaction Statement: Purchase' in pdf_text:
        if 'Transaction Statement: Sale' in pdf_text:
            trade_data['Trade Type'] = 'Sell'
        else:
            trade_data['Trade Type'] = 'Buy'

                # Asset-Name:
        #   Units 26 UniCredit Bank GmbH
        #   HVB Put 17.12.25 NVIDIA 200
        #   Account Owner
        asset_match = re.search(
            r'Units\s+[\d\.,]+\s+.*\n([^\n]+)\nAccount Owner',
            pdf_text
        )
        if asset_match:
            trade_data['Asset Name'] = asset_match.group(1).strip()
            print(f"[DEBUG] Baader Asset Name: {trade_data['Asset Name']}")
        else:
            print("Asset name not found in Baader contract note.")
            return None
        
        # Datum aus "Order Time:\n2025-11-13\n18:04:24:00"
        date_match = re.search(r'Order Time:\s*[\r\n]+(\d{4}-\d{2}-\d{2})', pdf_text)
        if date_match:
            iso_date = date_match.group(1)
            # in dd.mm.yyyy umwandeln
            dt = datetime.strptime(iso_date, "%Y-%m-%d")
            trade_data['Date'] = dt.strftime("%d.%m.%Y")

        # Menge & Preis aus Details:
        # Details:
        # Units 26 EUR 1.69 GETTEX ...
        details_match = re.search(
            r'Details:\s*[\r\n]+Units\s+([\d\.,]+)\s+EUR\s+([\d\.,]+)',
            pdf_text
        )
        if details_match:
            qty = parse_number(details_match.group(1))
            price = parse_number(details_match.group(2))
            trade_data['Quantity'] = qty
            trade_data['Price per Unit'] = price

        # Steuern:
        # German flat rate tax EUR 1.84 - Church tax EUR 0.16 - Solidarity surcharge EUR 0.10
        taxes_match = re.search(
            r'German flat rate tax\s*EUR\s*([\d\.,]+)\s*-\s*Church tax\s*EUR\s*([\d\.,]+)\s*-\s*Solidarity surcharge\s*EUR\s*([\d\.,]+)',
            pdf_text
        )
        if taxes_match:
            cg = parse_number(taxes_match.group(1))
            church = parse_number(taxes_match.group(2))
            soli = parse_number(taxes_match.group(3))
            trade_data['Capital Gains Tax'] = cg
            trade_data['Church Tax'] = church
            trade_data['Solidarity Surcharge'] = soli
            trade_data['Taxes Sum'] = cg + church + soli

        return trade_data

    # nichts erkannt
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

    trade_type = trade_data['Trade Type']
    asset_name = trade_data['Asset Name']
    source = trade_data.get('Source', 'cost_info')  # 'cost_info' oder 'contract_note'

    # Ensure date format is applied correctly
    date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")

    # Check if the date_style already exists to avoid the ValueError
    if "date_style" not in workbook.named_styles:
        workbook.add_named_style(date_style)

    if trade_type == 'Buy':
        next_row = find_next_empty_row_in_column(sheet, 'A')

        # Insert the new trade data into the predefined table
        sheet[f'A{next_row}'] = trade_data['Asset Name']
        buy_date = trade_data['Date']
        if isinstance(buy_date, str):
            buy_date = datetime.strptime(buy_date, "%d.%m.%Y").date()  # Convert string to date object
        sheet[f'B{next_row}'] = buy_date  # Write as string to avoid datetime conversion
        sheet[f'B{next_row}'].style = "date_style"  # Apply date format to the new row
        sheet[f'C{next_row}'] = trade_data['Quantity']
        sheet[f'D{next_row}'] = trade_data['Price per Unit']
        #sheet[f'K{next_row}'] = trade_data['Fees']

    elif trade_type == 'Sell':
        # Cost-Info-Sell: wir wollen eine Zeile ohne Sell-Datum finden
        # Contract-Note-Sell: wir dürfen auch Trades mit bereits gesetztem Sell-Datum nehmen,
        # um NUR die Steuern nachzutragen.
        require_empty_sell_date = (source == 'cost_info')

        match_row = find_best_matching_row(
            sheet,
            asset_name,
            require_empty_sell_date=require_empty_sell_date,
            min_score=0.5  # kannst du bei Bedarf hoch/runter drehen
        )

        if not match_row:
            print(f"No matching Buy entry found for asset {asset_name}.")
            workbook.save(EXCEL_FILE)
            return

        # Sell-Daten eintragen, aber bei Contract-Notes nur, wenn noch nichts da ist
        sell_date = trade_data.get('Date')
        if sell_date and (source == 'cost_info' or not sheet[f'F{match_row}'].value):
            if isinstance(sell_date, str):
                sell_date = datetime.strptime(sell_date, "%d.%m.%Y").date()
            sheet[f'F{match_row}'] = sell_date
            sheet[f'F{match_row}'].style = "date_style"

        if 'Quantity' in trade_data and (source == 'cost_info' or not sheet[f'G{match_row}'].value):
            sheet[f'G{match_row}'] = trade_data['Quantity']

        if 'Price per Unit' in trade_data and (source == 'cost_info' or not sheet[f'H{match_row}'].value):
            sheet[f'H{match_row}'] = trade_data['Price per Unit']

        # --- Steuern aus Contract Notes ---
        if 'Taxes Sum' in trade_data:
            sheet[f'T{match_row}'] = trade_data['Taxes Sum']

        if 'Capital Gains Tax' in trade_data:
            sheet[f'Y{match_row}'] = trade_data['Capital Gains Tax']

        if 'Solidarity Surcharge' in trade_data:
            sheet[f'Z{match_row}'] = trade_data['Solidarity Surcharge']

        if 'Church Tax' in trade_data:
            sheet[f'AA{match_row}'] = trade_data['Church Tax']

        # Falls du Order Fees aus Contract Notes irgendwann auch automatisch in J summieren willst:
        # if 'Order Fee' in trade_data:
        #     old_fee = sheet[f'J{match_row}'].value or 0
        #     sheet[f'J{match_row}'] = old_fee + trade_data['Order Fee']

        else:
            print(f"No matching Buy entry found for asset {asset_name}.")
            # Optionally, handle this case by logging or adding an error row

    workbook.save(EXCEL_FILE)
    print(f"Excel file updated successfully with trade data for {trade_data['Asset Name']}.")

def process_pdfs():
    for filename in os.listdir(PDF_FOLDER):
        if not filename.lower().endswith('.pdf'):
            continue

        pdf_path = os.path.join(PDF_FOLDER, filename)
        print(f"Processing {filename}...")

        pdf_text = extract_text(pdf_path)

        trade_data = None

        # 1. Cost-Information?
        if 'Ex-Ante cost information' in pdf_text:
            trade_data = extract_cost_info_trade(pdf_text)
        else:
            # 2. Contract Note (Scalable oder Baader)?
            trade_data = extract_contract_note_trade(pdf_text)

        if trade_data:
            update_excel(trade_data)
            os.remove(pdf_path)
            print(f"{filename} processed and deleted.")
        else:
            print(f"Failed to extract data from {filename}.")

if __name__ == "__main__":
    process_pdfs()
