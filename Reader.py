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


def extract_trade_info(pdf_text):
    trade_data = {}

    # Determine if it's a Buy or Sell
    if re.search(r'Buy', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Buy'
    elif re.search(r'Sell', pdf_text, re.IGNORECASE):
        trade_data['Trade Type'] = 'Sell'
    else:
        print("Trade type not found.")
        return None

    # Extract Asset Name or ISIN
    asset_match = re.search(r'Ex-Ante cost information\s*\n\s*(.+)', pdf_text)
    if asset_match:
        trade_data['Asset Name'] = asset_match.group(1).strip()
    else:
        print("Asset Name/ISIN not found.")
        return None

    # Extract Date
    date_match = re.search(r'Date\s*(?:\n.*?){3}\n\s*(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if date_match:
        trade_data['Date'] = date_match.group(1).strip()
    else:
        print("Date not found.")
        return None

    # Extract Quantity
    quantity_match = re.search(r'([\d,\.]+)\s*Shr\.', pdf_text)
    if quantity_match:
        trade_data['Quantity'] = float(quantity_match.group(1).replace(',', '.').strip())
    else:
        print("Quantity not found.")
        return None

    # Extract Price per Unit
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
        # Find the row that matches the Buy entry and doesn't have a Sell Date yet
        for row in range(2, sheet.max_row + 1):  # Assuming row 1 is headers
            if sheet[f'A{row}'].value == asset_name and not sheet[f'F{row}'].value:
                # Update the Sell details
                sell_date = trade_data['Date']
                if isinstance(sell_date, str):
                    sell_date = datetime.strptime(sell_date, "%d.%m.%Y").date()  # Convert string to date object
                sheet[f'F{row}'] = sell_date  # Write as string to avoid datetime conversion
                sheet[f'F{row}'].style = "date_style"  # Apply date format to the Sell Date
                sheet[f'G{row}'] = trade_data['Quantity']  # Units Sold
                sheet[f'H{row}'] = trade_data['Price per Unit']  # Sell Price per Unit
                #sheet[f'K{row}'] = sheet[f'K{row}'].value + trade_data['Fees']  # Add Fees to existing

                # Calculate Profit/Loss (assuming columns E = Total Buy Cost, I = Total Sell Proceeds, K = Fees)
                #total_buy_cost = sheet[f'E{row}'].value
                #total_sell_proceeds = sheet[f'I{row}'].value
                #total_fees = sheet[f'K{row}'].value
                #sheet[f'L{row}'] = total_sell_proceeds - total_buy_cost - total_fees  # Profit/Loss
                break
        else:
            print(f"No matching Buy entry found for asset {asset_name}.")
            # Optionally, handle this case by logging or adding an error row

    workbook.save(EXCEL_FILE)
    print(f"Excel file updated successfully with trade data for {trade_data['Asset Name']}.")

def process_pdfs():
    for filename in os.listdir(PDF_FOLDER):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(PDF_FOLDER, filename)
            print(f"Processing {filename}...")

            # Extract text from the PDF
            pdf_text = extract_text(pdf_path)

            # Extract trade information
            trade_data = extract_trade_info(pdf_text)

            if trade_data:
                # Update the Excel file
                update_excel(trade_data)

                #Delete the processed PDF
                os.remove(pdf_path)
                print(f"{filename} processed and deleted.")
            else:
                print(f"Failed to extract data from {filename}.")

if __name__ == "__main__":
    process_pdfs()
