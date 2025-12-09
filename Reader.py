import os
import re
import pandas as pd
from pdfminer.high_level import extract_text
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font
from datetime import datetime

# --- KONFIGURATION ---
PDF_FOLDER = os.path.join(os.path.dirname(__file__), 'pdfs')
EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'Trading.xlsx')


def parse_german_float(text):
    """Hilfsfunktion: Wandelt deutsche Zahlen (1.234,56) oder englische in float um."""
    if not text:
        return 0.0

    # Bereinigen von Währungssymbolen und Leerzeichen
    text = (
        text.replace('EUR', '')
        .replace('USD', '')
        .replace('€', '')
        .replace('$', '')
        .strip()
    )

    # Fall: 1.234,56 (Deutsch)
    if ',' in text and '.' in text:
        # Wenn der Punkt VOR dem Komma steht (1.200,50), ist es deutsch
        if text.find('.') < text.find(','):
            text = text.replace('.', '').replace(',', '.')
        else:
            # Englisch (1,200.50)
            text = text.replace(',', '')

    # Fall: 1,23 (Deutsch ohne Tausender)
    elif ',' in text:
        text = text.replace(',', '.')

    try:
        return float(text)
    except ValueError:
        return 0.0


# ---------------------------------------------------------
# INTELLIGENTE MATCHING LOGIK
# ---------------------------------------------------------
def are_assets_similar(name1, name2):
    """
    Vergleicht zwei Asset-Namen tolerant.
    Beispiel: 'HVB Put 17.12.25 NVIDIA 200' == 'NVIDIA Put 200,00 $ HVB'
    """
    if not name1 or not name2:
        return False

    # 1. Alles klein schreiben
    s1 = str(name1).lower()
    s2 = str(name2).lower()

    # Wenn exakt gleich -> Treffer
    if s1 == s2:
        return True

    def clean_and_tokenize(text):
        # Datumsformate entfernen (dd.mm.yy, yyyy-mm-dd) um Verwirrung zu vermeiden
        text = re.sub(r'\d{2}\.\d{2}\.\d{2,4}', '', text)
        text = re.sub(r'\d{4}-\d{2}-\d{2}', '', text)

        # Währungszeichen und Füllwörter weg
        text = re.sub(r'[\$€]', '', text)

        # "Put" und "Call" sind wichtig, nicht entfernen!
        text = (
            text.replace('eur', '')
            .replace('usd', '')
            .replace('pc.', '')
            .replace('stk.', '')
        )

        # Kommas durch Punkte ersetzen für Zahlen (200,00 -> 200.00)
        text = text.replace(',', '.')

        # In Wörter splitten
        tokens = text.split()
        cleaned_tokens = set()

        for t in tokens:
            # Versuche Zahlen zu normalisieren (200.00 -> 200.0)
            try:
                val = float(t)
                # Trick: Zahl als String formatieren ohne unnötige Nullen
                t = f"{val:.4f}".rstrip('0').rstrip('.')
            except ValueError:
                pass

            # Kurze Wörter filtern, ABER behalte wichtige wie "Put", "Inc"
            if len(t) > 1 or t.isdigit():
                cleaned_tokens.add(t)

        return cleaned_tokens

    tokens1 = clean_and_tokenize(s1)
    tokens2 = clean_and_tokenize(s2)

    # Schnittmenge bilden
    common = tokens1.intersection(tokens2)

    # Kriterium: Wenn fast alle Wörter des kürzeren Namens im längeren enthalten sind
    min_len = min(len(tokens1), len(tokens2))
    if min_len == 0:
        return False

    # Wir erlauben 1 abweichendes Wort (oft steht irgendwo "Order" oder "De" dabei)
    threshold = min_len if min_len < 3 else min_len - 1

    if len(common) >= threshold:
        return True

    return False


# ---------------------------------------------------------
# PARSER FUNKTIONEN
# ---------------------------------------------------------
def extract_trade_info_transaction_statement(pdf_text: str):
    """Parser für Baader/Scalable 'Transaction Statement'."""
    data = {"Source": "transaction_statement", "Taxes": 0.0, "Fee": 0.0}

    # Typ bestimmen
    if re.search(r'Transaction Statement:\s*Sale', pdf_text, re.IGNORECASE):
        data["Trade Type"] = "Sell"
    elif re.search(r'Transaction Statement:\s*(Purchase|Buy)', pdf_text, re.IGNORECASE):
        data["Trade Type"] = "Buy"
    else:
        return None

    # Datum
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', pdf_text)
    if date_match:
        y, m, d = date_match.group(1).split('-')
        data["Date"] = f"{d}.{m}.{y}"
    else:
        return None

    # Menge (Units)
    units_match = re.search(r'Units\s+([\d\.,]+)', pdf_text)
    if units_match:
        data["Quantity"] = parse_german_float(units_match.group(1))
    else:
        return None

    # --- FIX: Asset Name Extraction ---
    # Das alte Skript hat hier "Execution Venue" erwischt.
    # Strategie: Wir suchen "Price EUR" oder "ISIN". Der Name steht MEISTENS danach.
    lines = pdf_text.splitlines()
    asset_name = None

    # Suchanker finden
    start_idx = -1
    for i, line in enumerate(lines):
        # Der Preisblock ist oft direkt über dem Namen
        if "Price" in line and "EUR" in line:
            start_idx = i
            break

    # Fallback Anker: ISIN
    if start_idx == -1:
        for i, line in enumerate(lines):
            if "ISIN:" in line:
                start_idx = i
                break

    if start_idx != -1:
        # Suche in den folgenden 8 Zeilen nach einem validen Namen
        for i in range(start_idx + 1, min(len(lines), start_idx + 10)):
            line = lines[i].strip()

            # BLACKLIST: Diese Zeilen ignorieren wir strikt!
            if not line:
                continue

            bad_words = [
                "Execution Venue",
                "Market Value",
                "Amount",
                "Order",
                "Account",
                "Baader",
                "Bank",
                "UniCredit",
                "München",
                "Munich",
                "Tax",
                "WKN:",
                "ISIN:",
                "Client",
                "Portfolio",
                "Reference",
            ]
            if any(bw.lower() in line.lower() for bw in bad_words):
                continue

            # Zahlen/Datum ignorieren
            if re.match(r'^[\d\.,\s:-]+$', line):
                continue  # Nur Zahlen/Datum

            if re.search(r'\d{4}-\d{2}-\d{2}', line):
                continue  # Zeile enthält Datum

            # Wenn die Zeile "Price EUR" oder ähnlich ist (falls wir oben ISIN als Anker hatten)
            if "Price" in line and "EUR" in line:
                continue

            # Wenn wir hier sind, ist es sehr wahrscheinlich der Name
            asset_name = line
            break

    # Fallback: Wenn oben nichts gefunden wurde, suchen wir im Quantity Block
    if not asset_name:
        q_idx = pdf_text.find("Quantity")
        if q_idx != -1:
            snippet = pdf_text[q_idx:q_idx + 400]
            for line in snippet.splitlines():
                line = line.strip()
                if "Execution Venue" in line:
                    continue  # WICHTIG
                if len(line) > 3 and not any(
                    x in line for x in ["Quantity", "Units", "Price", "2025-", "Date"]
                ):
                    # Annahme: Der Name ist der längste Text ohne Keywords
                    asset_name = line
                    break

    data["Asset Name"] = asset_name if asset_name else "Unknown Asset"

    # Preis
    price_match = re.search(r'Price\s+EUR\s+([\d\.,]+)', pdf_text)
    if price_match:
        data["Price per Unit"] = parse_german_float(price_match.group(1))
    else:
        return None

    # Steuern – wie im alten Code:
    # Wir suchen die erste Zeile mit "Zahl -" und summieren die
    # nächsten Zeilen, die auch diesem Muster entsprechen.
    taxes = 0.0
    tax_start = None

    for idx, line in enumerate(lines):
        if re.search(r'([0-9]+[.,][0-9]+)\s*-\s*$', line.strip()):
            tax_start = idx
            break

    if tax_start is not None:
        for j in range(tax_start, min(tax_start + 10, len(lines))):
            m = re.search(r'([0-9]+[.,][0-9]+)\s*-\s*$', lines[j].strip())
            if m:
                taxes += parse_german_float(m.group(1))

    data["Taxes"] = taxes
    return data

def extract_trade_info_contract_note(pdf_text: str):
    """Parser für Scalable 'Contract Note'."""
    data = {"Source": "contract_note", "Taxes": 0.0, "Fee": 0.0}

    # Typ & Asset
    head_match = re.search(
        r'^(Buy|Sell)\s+([\d\.,]+)\s*(?:pc\.|Stk\.)?\s+(.+)$',
        pdf_text,
        re.MULTILINE,
    )
    if not head_match:
        head_match = re.search(r'^(Buy|Sell)\s+(.+)$', pdf_text, re.MULTILINE)

    if head_match:
        if len(head_match.groups()) == 3:
            data["Trade Type"] = head_match.group(1).capitalize()
            data["Asset Name"] = head_match.group(3).strip()
        else:
            data["Trade Type"] = head_match.group(1).capitalize()
            data["Asset Name"] = head_match.group(2).strip()
    else:
        return None

    # Datum
    date_match = re.search(
        r'(?:Execution|Date)\s+(\d{2}\.\d{2}\.\d{4})',
        pdf_text,
    )
    if date_match:
        data["Date"] = date_match.group(1)
    else:
        return None

    # Menge & Preis
    qp_match = re.search(
        r'([\d\.,]+)\s*(?:pc\.|Stk\.)\s+([\d\.,]+)\s*EUR',
        pdf_text,
    )
    if qp_match:
        data["Quantity"] = parse_german_float(qp_match.group(1))
        data["Price per Unit"] = parse_german_float(qp_match.group(2))
    else:
        return None

    # Gebühren & Steuern
    fee_match = re.search(r'Order fees\s*([-\d\.,]+)\s*EUR', pdf_text)
    if fee_match:
        data["Fee"] = abs(parse_german_float(fee_match.group(1)))

    tax_match = re.search(r'Taxes\s*([-\d\.,]+)\s*EUR', pdf_text)
    if tax_match:
        data["Taxes"] = abs(parse_german_float(tax_match.group(1)))

    return data


def extract_trade_info_cost_information(pdf_text: str):
    """Parser für 'Ex-Ante cost information'."""
    data = {"Source": "cost_information", "Taxes": 0.0, "Fee": 0.0}

    if "Order Buy" in pdf_text or "Order\nBuy" in pdf_text:
        data["Trade Type"] = "Buy"
    elif "Order Sell" in pdf_text or "Order\nSell" in pdf_text:
        data["Trade Type"] = "Sell"
    else:
        return None

    asset_match = re.search(r'Ex-Ante cost information\s*\n\s*(.+)', pdf_text)
    data["Asset Name"] = asset_match.group(1).strip() if asset_match else "Unknown"

    date_match = re.search(r'Date\s+(\d{2}\.\d{2}\.\d{4})', pdf_text)
    if date_match:
        data["Date"] = date_match.group(1)
    else:
        return None

    qty_match = re.search(r'Quantity\s+([\d\.,]+)', pdf_text)
    if qty_match:
        data["Quantity"] = parse_german_float(qty_match.group(1))
    else:
        return None

    amt_match = re.search(r'Est\. order amount\s+([\d\.,]+)\s*€', pdf_text)
    if amt_match and data["Quantity"] > 0:
        total = parse_german_float(amt_match.group(1))
        data["Price per Unit"] = total / data["Quantity"]
    else:
        data["Price per Unit"] = 0.0

    return data


def extract_trade_info(pdf_path):
    try:
        text = extract_text(pdf_path)
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return None

    if "Transaction Statement" in text:
        return extract_trade_info_transaction_statement(text)
    elif "Contract note" in text or "Abrechnung" in text:
        return extract_trade_info_contract_note(text)
    elif "Ex-Ante cost information" in text or "Kostentransparenz" in text:
        return extract_trade_info_cost_information(text)
    else:
        print(f"Unknown PDF format: {os.path.basename(pdf_path)}")
        return None


# ---------------------------------------------------------
# EXCEL HANDLING
# ---------------------------------------------------------
def ensure_columns_exist(sheet):
    headers = {}
    for cell in sheet[1]:
        if cell.value:
            headers[cell.value] = cell.column_letter

    required_cols = [
        "Taxes",
        "Profit/Loss (Abs.) after tax",
        "Profit/Loss (%) after tax",
        "Profit/Loss (€ per day) after tax",
    ]

    last_col_idx = sheet.max_column

    for col_name in required_cols:
        if col_name not in headers:
            last_col_idx += 1
            new_col_letter = get_column_letter(last_col_idx)
            cell = sheet[f"{new_col_letter}1"]
            cell.value = col_name
            cell.font = Font(bold=True)
            headers[col_name] = new_col_letter

    return headers


def find_buy_row(sheet, asset_name, quantity):
    """Sucht die Kauf-Zeile mit intelligentem Namensvergleich."""
    for row in range(2, sheet.max_row + 1):
        name_cell = sheet[f"A{row}"]
        sell_date_cell = sheet[f"F{row}"]

        # Nur Zeilen ohne Verkaufsdatum betrachten
        if sell_date_cell.value:
            continue

        excel_name = str(name_cell.value).strip() if name_cell.value else ""

        if are_assets_similar(excel_name, asset_name):
            return row

    return None


def write_excel(all_trades):
    if not os.path.exists(EXCEL_FILE):
        print("Excel-Datei nicht gefunden!")
        return

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    # Spalten sicherstellen (fügt sie ggf. neu rechts an)
    header_map = ensure_columns_exist(ws)
    col_taxes = header_map.get("Taxes")
    col_profit_abs_net = header_map.get("Profit/Loss (Abs.) after tax")
    col_profit_pct_net = header_map.get("Profit/Loss (%) after tax")
    col_profit_day_net = header_map.get("Profit/Loss (€ per day) after tax")

    # Datumsformat
    date_style = NamedStyle(name="date_style_final_v2", number_format="DD.MM.YYYY")
    if "date_style_final_v2" not in wb.named_styles:
        wb.add_named_style(date_style)

    for trade in all_trades:
        asset = trade["Asset Name"]
        date_obj = datetime.strptime(trade["Date"], "%d.%m.%Y")
        qty = trade["Quantity"]
        price = trade["Price per Unit"]
        taxes = trade.get("Taxes", 0.0)

        # -------------------
        # BUY – nur A–D setzen
        # -------------------
        if trade["Trade Type"] == "Buy":
            # nächste freie Zeile in Spalte A innerhalb der Tabelle
            row = 2
            while ws[f"A{row}"].value:
                row += 1

            ws[f"A{row}"] = asset
            ws[f"B{row}"] = date_obj
            ws[f"B{row}"].style = "date_style_final_v2"
            ws[f"C{row}"] = qty
            ws[f"D{row}"] = price

            # optional: Steuern für Buys als 0 eintragen (saubere Basis)
            if col_taxes:
                ws[f"{col_taxes}{row}"] = taxes

        # -------------------
        # SELL – nur F–H + neue Spalten
        # -------------------
        elif trade["Trade Type"] == "Sell":
            row = find_buy_row(ws, asset, qty)
            if row:
                # Nur das, was Excel selbst NICHT berechnen kann
                ws[f"F{row}"] = date_obj
                ws[f"F{row}"].style = "date_style_final_v2"
                ws[f"G{row}"] = qty
                ws[f"H{row}"] = price

                # Steuern in die neue Taxes-Spalte
                if col_taxes:
                    ws[f"{col_taxes}{row}"] = taxes

                # After-tax Profit (Abs.)
                if col_profit_abs_net and col_taxes:
                    ws[f"{col_profit_abs_net}{row}"] = f"=N{row}-{col_taxes}{row}"

                # After-tax Profit (%) = After-tax Abs. / Total Buy Cost
                if col_profit_pct_net and col_profit_abs_net:
                    ws[f"{col_profit_pct_net}{row}"] = f"={col_profit_abs_net}{row}/E{row}"
                    ws[f"{col_profit_pct_net}{row}"].number_format = '0.00%'

                # After-tax Profit pro Tag
                if col_profit_day_net and col_profit_abs_net:
                    ws[f"{col_profit_day_net}{row}"] = (
                        f"=IF(K{row}>0, {col_profit_abs_net}{row}/K{row}, {col_profit_abs_net}{row})"
                    )

                print(f"MATCH: '{asset}' gematched mit '{ws[f'A{row}'].value}' in Zeile {row}.")
            else:
                print(f"WARNUNG: Kein Match für '{asset}' (Datum: {trade['Date']}). Bitte Namen in Excel prüfen.")

    try:
        wb.save(EXCEL_FILE)
        print("Update abgeschlossen.")
    except PermissionError:
        print("FEHLER: Bitte Excel-Datei schließen!")


def process_pdfs():
    if not os.path.exists(PDF_FOLDER):
        os.makedirs(PDF_FOLDER)
        return

    files = [
        f for f in os.listdir(PDF_FOLDER)
        if f.lower().endswith('.pdf')
    ]

    if not files:
        print("Keine PDFs gefunden.")
        return

    merged_trades = {}

    for filename in files:
        path = os.path.join(PDF_FOLDER, filename)
        print(f"Lese {filename}...")
        trade_data = extract_trade_info(path)

        if trade_data:
            trade_data['_filename'] = filename

            # Wir nutzen Datum, Typ und Menge als Key, um Duplikate zu mergen (CostInfo vs ContractNote)
            key = (trade_data['Date'], trade_data['Trade Type'], trade_data['Quantity'])

            if key in merged_trades:
                existing = merged_trades[key]

                # Contract Note bevorzugen
                if trade_data['Source'] == 'contract_note':
                    merged_trades[key] = trade_data

                elif trade_data.get('Taxes', 0) > existing.get('Taxes', 0):
                    merged_trades[key] = trade_data

                # Wenn wir nur CostInfo hatten und jetzt TransactionStatement kommt (besserer Name?)
                elif (
                    trade_data['Source'] == 'transaction_statement'
                    and existing['Source'] == 'cost_information'
                ):
                    # Vorsicht: Manchmal ist CostInfo Name besser. Wir bleiben beim ersten oder nehmen den längeren?
                    pass
            else:
                merged_trades[key] = trade_data

        try:
            os.remove(path)
        except Exception:
            pass

    final_list = list(merged_trades.values())

    # Sortierung: Erst alle Buys (0), dann Sells (1) chronologisch
    def sort_algo(t):
        d, m, y = t['Date'].split('.')
        type_prio = 0 if t['Trade Type'] == 'Buy' else 1
        return int(y) * 10000 + int(m) * 100 + int(d), type_prio

    final_list.sort(key=sort_algo)

    if final_list:
        write_excel(final_list)


if __name__ == "__main__":
    process_pdfs()
