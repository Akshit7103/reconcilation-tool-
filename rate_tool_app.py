"""
Dynamic Fee Rate Mapper Tool
Uses tkinter for file upload only, displays results in terminal.
Completely dynamic - no hardcoded values.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import os
from tabulate import tabulate
import re

def select_file(title="Select Excel File"):
    """
    Open file dialog to select Excel file

    Args:
        title (str): Dialog title

    Returns:
        str: Selected file path or None if cancelled
    """
    # Create a root window but hide it
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # Open file dialog
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
    )

    # Destroy the root window
    root.destroy()

    return file_path if file_path else None

def select_multiple_files():
    """
    Allow user to select summary file, card issuance file, and transaction files

    Returns:
        dict: Dictionary with file paths for different data types
    """
    files = {
        'summary': None,
        'card_issuance': None,
        'international_file': None,
        'domestic_file': None,
        'dispute_file': None
    }

    print("\nStep 1: Select the Summary file (contains fee types and rates)")
    summary_file = select_file("Select Summary Excel File")

    if not summary_file:
        return files

    files['summary'] = summary_file
    print(f"Selected summary file: {os.path.basename(summary_file)}")

    print("\nStep 2: Select the Card Issuance Report file")
    card_file = select_file("Select Card Issuance Report Excel File")

    if card_file:
        files['card_issuance'] = card_file
        print(f"Selected card issuance file: {os.path.basename(card_file)}")

    # Ask for international transaction file
    print("\nStep 3: Select International Transaction file")
    print("Required for international license fees calculation")

    international_file = select_file("Select International Transaction Excel File (or Cancel to skip)")
    if international_file:
        files['international_file'] = international_file
        print(f"Selected international transaction file: {os.path.basename(international_file)}")
    else:
        print("No international transaction file selected")

    # Ask for domestic transaction file
    print("\nStep 4: Select Domestic Transaction file")
    print("Required for domestic VISA authorization calculation")

    domestic_file = select_file("Select Domestic Transaction Excel File (or Cancel to skip)")
    if domestic_file:
        files['domestic_file'] = domestic_file
        print(f"Selected domestic transaction file: {os.path.basename(domestic_file)}")
    else:
        print("No domestic transaction file selected")

    # Ask for dispute transaction file
    print("\nStep 5: Select Transaction Dispute file (VROL Report)")
    print("Required for transaction dispute fees calculation")

    dispute_file = select_file("Select Transaction Dispute Excel File (or Cancel to skip)")
    if dispute_file:
        files['dispute_file'] = dispute_file
        print(f"Selected dispute transaction file: {os.path.basename(dispute_file)}")
    else:
        print("No dispute transaction file selected")

    return files

def analyze_excel_structure(file_path):
    """
    Dynamically analyze Excel file structure to find fee types and rates

    Args:
        file_path (str): Path to Excel file

    Returns:
        dict: Structure analysis results
    """
    try:
        # Read Excel file
        xl_file = pd.ExcelFile(file_path)
        sheets = xl_file.sheet_names

        analysis_results = {
            'file_path': file_path,
            'sheets': sheets,
            'mappings': {}
        }

        print(f"\nAnalyzing file: {os.path.basename(file_path)}")
        print(f"Found {len(sheets)} sheet(s): {', '.join(sheets)}")

        # Analyze each sheet
        for sheet_name in sheets:
            print(f"\nAnalyzing sheet: '{sheet_name}'")
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            sheet_analysis = analyze_sheet_for_fee_mapping(df, sheet_name)
            if sheet_analysis['mappings']:
                analysis_results['mappings'][sheet_name] = sheet_analysis

        return analysis_results

    except Exception as e:
        print(f"Error analyzing file: {str(e)}")
        return None

def analyze_sheet_for_fee_mapping(df, sheet_name):
    """
    Dynamically analyze a sheet to find fee type and rate columns

    Args:
        df (DataFrame): Sheet data
        sheet_name (str): Name of the sheet

    Returns:
        dict: Analysis results for the sheet
    """
    result = {
        'sheet_name': sheet_name,
        'shape': df.shape,
        'columns': list(df.columns),
        'mappings': {}
    }

    print(f"   Sheet shape: {df.shape}")
    print(f"   Columns: {len(df.columns)}")

    # Look for potential fee type and rate columns
    fee_mappings = find_fee_rate_pairs(df)

    if fee_mappings:
        result['mappings'] = fee_mappings
        print(f"   Found {len(fee_mappings)} fee-rate mappings")
    else:
        print("   No clear fee-rate mappings found")

    return result

def find_fee_rate_pairs(df):
    """
    Dynamically find fee type and rate column pairs in the dataframe

    Args:
        df (DataFrame): Sheet data

    Returns:
        dict: Fee type to rate mappings
    """
    mappings = {}

    # Keywords that might indicate fee types
    fee_keywords = ['fee', 'charge', 'cost', 'type', 'service', 'transaction']
    rate_keywords = ['rate', 'amount', 'price', 'cost', 'chart', 'value']

    # Try different strategies to find fee-rate pairs

    # Strategy 1: Look for adjacent columns with fee/rate indicators
    for col_idx in range(len(df.columns) - 1):
        col1 = df.columns[col_idx]
        col2 = df.columns[col_idx + 1]

        col1_str = str(col1).lower()
        col2_str = str(col2).lower()

        # Check if columns might be fee type and rate
        is_fee_col1 = any(keyword in col1_str for keyword in fee_keywords)
        is_rate_col2 = any(keyword in col2_str for keyword in rate_keywords)

        if is_fee_col1 and is_rate_col2:
            pair_mappings = extract_mappings_from_columns(df, col1, col2)
            mappings.update(pair_mappings)

    # Strategy 2: Look for data patterns (skip header rows and find actual data)
    if not mappings:
        mappings = extract_mappings_by_pattern(df)

    return mappings

def extract_mappings_from_columns(df, fee_col, rate_col):
    """
    Extract fee-rate mappings from specific columns

    Args:
        df (DataFrame): Sheet data
        fee_col: Fee type column name
        rate_col: Rate column name

    Returns:
        dict: Fee type to rate mappings
    """
    mappings = {}

    for idx, row in df.iterrows():
        fee_type = row[fee_col]
        rate_value = row[rate_col]

        # Skip empty, NaN, or header-like values
        if (pd.notna(fee_type) and pd.notna(rate_value) and
            not str(fee_type).lower().strip() in ['fee type', 'type', 's.no.', 'sno', 'sr.no', 'fee type '] and
            not str(rate_value).lower().strip() in ['rate', 'amount', 'chart', 'rates chart', 'rates chart '] and
            str(fee_type).strip() != '' and str(rate_value).strip() != ''):

            mappings[str(fee_type).strip()] = str(rate_value).strip()

    return mappings

def extract_mappings_by_pattern(df):
    """
    Extract mappings by analyzing data patterns across all columns

    Args:
        df (DataFrame): Sheet data

    Returns:
        dict: Fee type to rate mappings
    """
    mappings = {}

    # Look for patterns where one column has descriptive text and another has rates
    for col_idx in range(len(df.columns) - 1):
        for next_col_idx in range(col_idx + 1, len(df.columns)):
            col1 = df.columns[col_idx]
            col2 = df.columns[next_col_idx]

            # Get non-null values from both columns
            col1_data = df[col1].dropna()
            col2_data = df[col2].dropna()

            # Check if one column looks like descriptions and other like rates
            if len(col1_data) > 0 and len(col2_data) > 0:
                # Skip if columns have different lengths
                min_len = min(len(col1_data), len(col2_data))

                for i in range(min_len):
                    if i < len(col1_data) and i < len(col2_data):
                        val1 = col1_data.iloc[i]
                        val2 = col2_data.iloc[i]

                        # Check if first column looks like text description
                        # and second looks like rate/amount
                        if (isinstance(val1, str) and len(str(val1).strip()) > 3 and
                            not str(val1).lower().strip() in ['fee type', 'type', 's.no.', 'sno', 'fee type ', 'rates chart', 'rates chart '] and
                            not str(val2).lower().strip() in ['rate', 'amount', 'chart', 'rates chart', 'rates chart '] and
                            str(val1).strip() != '' and str(val2).strip() != ''):

                            mappings[str(val1).strip()] = str(val2).strip()

    return mappings

def extract_card_issuance_data(file_path):
    """
    Extract card issuance data from Excel file

    Args:
        file_path (str): Path to card issuance Excel file

    Returns:
        dict: Card issuance data
    """
    try:
        print(f"Analyzing card issuance file: {os.path.basename(file_path)}")

        xl_file = pd.ExcelFile(file_path)
        sheets = xl_file.sheet_names

        card_data = {
            'total_cards': 0,
            'monthly_data': [],
            'raw_data': {}
        }

        for sheet_name in sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Look for card issuance patterns
            cards_found = find_card_issuance_values(df, sheet_name)
            if cards_found:
                card_data['raw_data'][sheet_name] = cards_found
                if 'total_cards' in cards_found:
                    card_data['total_cards'] = max(card_data['total_cards'], cards_found['total_cards'])
                if 'monthly_data' in cards_found:
                    card_data['monthly_data'].extend(cards_found['monthly_data'])

        return card_data

    except Exception as e:
        print(f"Error extracting card issuance data: {str(e)}")
        return None

def find_card_issuance_values(df, sheet_name):
    """
    Find card issuance values in the dataframe

    Args:
        df (DataFrame): Sheet data
        sheet_name (str): Sheet name

    Returns:
        dict: Found card issuance data
    """
    result = {}
    monthly_data = []

    # Look for total cards pattern
    total_keywords = ['total', 'total cards', 'cards issued', 'quarter']

    for idx, row in df.iterrows():
        for col in df.columns:
            cell_value = str(row[col]).lower().strip()

            # Check for total cards
            if any(keyword in cell_value for keyword in total_keywords):
                # Look for number in adjacent cells
                for next_col in df.columns:
                    adjacent_value = row[next_col]
                    if pd.notna(adjacent_value) and str(adjacent_value).isdigit():
                        result['total_cards'] = int(adjacent_value)
                        print(f"   Found total cards: {result['total_cards']}")
                        break

    # Look for monthly/period data
    period_keywords = ['apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec', 'period', 'month']

    for idx, row in df.iterrows():
        period_found = False
        period_name = ""
        cards_count = 0

        for col in df.columns:
            cell_value = str(row[col]).lower().strip()

            # Check if this looks like a period
            if any(keyword in cell_value for keyword in period_keywords) and len(cell_value) < 20:
                period_found = True
                period_name = str(row[col]).strip()

                # Look for number in same row
                for next_col in df.columns:
                    adjacent_value = row[next_col]
                    if pd.notna(adjacent_value) and str(adjacent_value).replace(',', '').isdigit():
                        cards_count = int(str(adjacent_value).replace(',', ''))
                        break

                if cards_count > 0:
                    monthly_data.append({
                        'period': period_name,
                        'cards': cards_count
                    })
                    print(f"   Found period data: {period_name} - {cards_count} cards")
                break

    if monthly_data:
        result['monthly_data'] = monthly_data

    return result

def calculate_fee_amount(rate_formula, card_count, transaction_count=0, transaction_amount=0):
    """
    Dynamically calculate fee amount based on rate formula

    Args:
        rate_formula (str): Rate formula from summary file
        card_count (int): Number of cards
        transaction_count (int): Number of transactions (if applicable)
        transaction_amount (float): Transaction amount (if applicable)

    Returns:
        dict: Calculation result
    """
    try:
        rate_formula = str(rate_formula).strip()

        # Handle different rate formula patterns
        if 'first' in rate_formula.lower() and 'thereafter' in rate_formula.lower():
            # Tiered pricing like "$2 for first 30K cards\n$1.5 per card thereafter"
            return calculate_tiered_card_fee(rate_formula, card_count)

        elif 'per transaction' in rate_formula.lower():
            # Per transaction fee like "Rs 0.25 per transaction"
            return calculate_per_transaction_fee(rate_formula, transaction_count)

        elif 'per dispute' in rate_formula.lower():
            # Per dispute fee like "Rs 250 per dispute"
            return calculate_per_dispute_fee(rate_formula, transaction_count)

        elif 'no of tran' in rate_formula.lower() and '$' in rate_formula:
            # Transaction count based like "No of tran * 5$"
            return calculate_transaction_volume_fee(rate_formula, transaction_count)

        elif 'amount of tran' in rate_formula.lower() or 'amout of tran' in rate_formula.lower():
            # Transaction amount based like "Amount of tran *0.5$" or "Amount of tran *Rs0.25"
            return calculate_transaction_amount_fee(rate_formula, transaction_amount)

        elif rate_formula.isdigit():
            # Fixed amount
            return {
                'calculated_amount': float(rate_formula),
                'calculation_method': 'Fixed Amount',
                'formula_used': rate_formula,
                'currency': '$'
            }

        else:
            return {
                'calculated_amount': 0,
                'calculation_method': 'Unknown Formula',
                'formula_used': rate_formula,
                'error': 'Could not parse rate formula'
            }

    except Exception as e:
        return {
            'calculated_amount': 0,
            'calculation_method': 'Error',
            'formula_used': rate_formula,
            'error': str(e)
        }

def calculate_tiered_card_fee(rate_formula, card_count):
    """
    Calculate tiered card fee like "$2 for first 30K cards, $1.5 per card thereafter"

    Args:
        rate_formula (str): Rate formula
        card_count (int): Number of cards

    Returns:
        dict: Calculation result
    """
    try:
        # Extract numbers and rates from formula
        # Pattern: "$2 for first 30K cards\n$1.5 per card thereafter"

        # Find first tier rate and threshold
        first_rate_match = re.search(r'\$?(\d+(?:\.\d+)?)', rate_formula)
        threshold_match = re.search(r'(\d+)k', rate_formula.lower())

        # Find second tier rate
        lines = rate_formula.split('\n')
        second_rate = 0
        if len(lines) > 1:
            second_rate_match = re.search(r'\$?(\d+(?:\.\d+)?)', lines[1])
            if second_rate_match:
                second_rate = float(second_rate_match.group(1))

        if first_rate_match and threshold_match:
            first_rate = float(first_rate_match.group(1))
            threshold = int(threshold_match.group(1)) * 1000  # Convert K to actual number

            if card_count <= threshold:
                # All cards in first tier
                amount = card_count * first_rate
                method = f"All {card_count} cards at ${first_rate} each"
            else:
                # Split between tiers
                first_tier_amount = threshold * first_rate
                remaining_cards = card_count - threshold
                second_tier_amount = remaining_cards * second_rate
                amount = first_tier_amount + second_tier_amount
                method = f"First {threshold} cards at ${first_rate}, remaining {remaining_cards} cards at ${second_rate}"

            return {
                'calculated_amount': amount,
                'calculation_method': method,
                'formula_used': rate_formula,
                'currency': '$'
            }

    except Exception as e:
        pass

    return {
        'calculated_amount': 0,
        'calculation_method': 'Error in tiered calculation',
        'formula_used': rate_formula
    }

def calculate_per_transaction_fee(rate_formula, transaction_count):
    """Calculate per transaction fee"""
    try:
        rate_match = re.search(r'(\d+(?:\.\d+)?)', rate_formula)
        if rate_match:
            rate = float(rate_match.group(1))
            amount = transaction_count * rate

            # Determine currency
            currency = "Rs" if 'rs' in rate_formula.lower() else "$"

            return {
                'calculated_amount': amount,
                'calculation_method': f"{transaction_count} transactions × {currency}{rate}",
                'formula_used': rate_formula,
                'currency': currency
            }
    except:
        pass

    return {'calculated_amount': 0, 'calculation_method': 'Error', 'formula_used': rate_formula}

def calculate_per_dispute_fee(rate_formula, dispute_count):
    """Calculate per dispute fee"""
    try:
        rate_match = re.search(r'(\d+(?:\.\d+)?)', rate_formula)
        if rate_match:
            rate = float(rate_match.group(1))
            amount = dispute_count * rate

            # Determine currency
            currency = "Rs" if 'rs' in rate_formula.lower() else "$"

            return {
                'calculated_amount': amount,
                'calculation_method': f"{dispute_count} disputes × {currency}{rate}",
                'formula_used': rate_formula,
                'currency': currency
            }
    except:
        pass

    return {'calculated_amount': 0, 'calculation_method': 'Error', 'formula_used': rate_formula}

def calculate_transaction_volume_fee(rate_formula, transaction_count):
    """Calculate transaction volume fee like 'No of tran * 5$'"""
    try:
        rate_match = re.search(r'(\d+(?:\.\d+)?)', rate_formula)
        if rate_match:
            rate = float(rate_match.group(1))
            amount = transaction_count * rate

            # Determine currency (usually $ for volume fees)
            currency = "$"

            return {
                'calculated_amount': amount,
                'calculation_method': f"{transaction_count} transactions × {currency}{rate}",
                'formula_used': rate_formula,
                'currency': currency
            }
    except:
        pass

    return {'calculated_amount': 0, 'calculation_method': 'Error', 'formula_used': rate_formula}

def calculate_transaction_amount_fee(rate_formula, transaction_amount):
    """Calculate transaction amount based fee like 'Amount of tran *0.5$' or 'Amount of tran *Rs0.25'"""
    try:
        # Extract the rate number (could be after * or before currency)
        rate_match = re.search(r'\*\s*(?:Rs|rs|\$)?(\d+(?:\.\d+)?)', rate_formula)
        if not rate_match:
            # Try alternative pattern
            rate_match = re.search(r'(\d+(?:\.\d+)?)', rate_formula)

        if rate_match:
            rate = float(rate_match.group(1))
            amount = transaction_amount * rate

            # Determine currency from formula
            currency_symbol = "$"
            if 'rs' in rate_formula.lower():
                currency_symbol = "Rs"

            return {
                'calculated_amount': amount,
                'calculation_method': f"${transaction_amount:,} × {rate}",
                'formula_used': rate_formula,
                'currency': currency_symbol
            }
    except Exception as e:
        pass

    return {'calculated_amount': 0, 'calculation_method': 'Error', 'formula_used': rate_formula}

def extract_dispute_data_from_vrol(df):
    """
    Extract dispute information from VROL dataframe using working logic from standalone tool

    Args:
        df (DataFrame): VROL sheet data

    Returns:
        dict: Dispute information
    """
    dispute_info = {
        'total_disputes': 0,
        'total_amount': 0,
        'individual_disputes': []
    }

    # Look for dispute count patterns
    for idx, row in df.iterrows():
        for col in df.columns:
            cell_value = str(row[col]).lower().strip()

            # Check for dispute count indicators
            if ('no of disputed' in cell_value or
                'number of disputed' in cell_value or
                'disputed transactions' in cell_value):

                # Look for number in adjacent cells
                for check_col in df.columns:
                    adjacent_value = row[check_col]
                    if (pd.notna(adjacent_value) and
                        str(adjacent_value).replace(',', '').isdigit()):
                        count = int(str(adjacent_value).replace(',', ''))
                        if count > 0 and count < 1000:  # Reasonable dispute count
                            dispute_info['total_disputes'] = max(dispute_info['total_disputes'], count)
                            print(f"   Found dispute count: {count}")

    # Also look for individual disputed transactions (X, Y, etc.)
    transactions = []
    for idx, row in df.iterrows():
        row_values = []
        for col in df.columns:
            val = row[col]
            if pd.notna(val):
                row_values.append(str(val).strip())

        # Look for single letter/short ID followed by amount
        if len(row_values) >= 2:
            for i in range(len(row_values) - 1):
                id_val = row_values[i]
                amount_val = row_values[i + 1]

                # Check if looks like dispute transaction (single letter + amount)
                if (len(id_val) <= 2 and id_val.isalnum() and
                    amount_val.replace(',', '').replace('.', '').isdigit() and
                    float(amount_val.replace(',', '')) > 0):

                    amount = float(amount_val.replace(',', ''))
                    transactions.append({
                        'id': id_val,
                        'amount': amount
                    })
                    print(f"   Found disputed transaction: {id_val} - ${amount:,.0f}")

    if transactions:
        dispute_info['individual_disputes'] = transactions
        dispute_info['total_amount'] = sum(t['amount'] for t in transactions)

        # If we didn't find explicit count, use number of individual transactions
        if dispute_info['total_disputes'] == 0:
            dispute_info['total_disputes'] = len(transactions)

    return dispute_info

def process_specific_transaction_file(file_path, expected_type):
    """
    Process a specific transaction file for a specific transaction type

    Args:
        file_path (str): Path to transaction file
        expected_type (str): Expected transaction type (international, domestic, disputes)

    Returns:
        dict: Transaction data for the specific type
    """
    if not file_path or not os.path.exists(file_path):
        return {'total_amount': 0, 'total_volume': 0, 'transactions': []}

    try:
        print(f"Processing {expected_type} transaction file: {os.path.basename(file_path)}")

        xl_file = pd.ExcelFile(file_path)
        sheets = xl_file.sheet_names

        best_data = {'total_amount': 0, 'total_volume': 0, 'transactions': []}

        for sheet_name in sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if expected_type == 'disputes':
                # Special handling for dispute files using the working logic
                dispute_data = extract_dispute_data_from_vrol(df)
                if dispute_data['total_disputes'] > best_data['total_volume']:
                    best_data['total_amount'] = dispute_data['total_amount']
                    best_data['total_volume'] = dispute_data['total_disputes']
                    best_data['transactions'] = dispute_data['individual_disputes']
            else:
                # Standard handling for international/domestic
                transactions = find_transaction_entries(df)
                totals = find_transaction_totals(df)

                if totals.get('amount', 0) > best_data['total_amount']:
                    best_data['total_amount'] = totals.get('amount', 0)
                    best_data['total_volume'] = totals.get('volume', len(transactions))
                    best_data['transactions'] = transactions

        print(f"   Found {expected_type}: Amount=${best_data['total_amount']:,}, Volume={best_data['total_volume']}")
        return best_data

    except Exception as e:
        print(f"Error processing {expected_type} file: {str(e)}")
        return {'total_amount': 0, 'total_volume': 0, 'transactions': []}


def extract_transactions_from_sheet(df, sheet_name):
    """
    Extract transaction data from a single sheet

    Args:
        df (DataFrame): Sheet data
        sheet_name (str): Sheet name

    Returns:
        dict: Extracted transaction data
    """
    sheet_data = {
        'type': determine_transaction_type(sheet_name, df),
        'transactions': [],
        'total_amount': 0,
        'total_volume': 0
    }

    # Look for transaction patterns
    transactions = find_transaction_entries(df)
    sheet_data['transactions'] = transactions

    # Look for totals
    totals = find_transaction_totals(df)
    if totals:
        sheet_data['total_amount'] = totals.get('amount', 0)
        sheet_data['total_volume'] = totals.get('volume', 0)

    # If no totals found, calculate from individual transactions
    if not sheet_data['total_amount'] and transactions:
        sheet_data['total_amount'] = sum(t.get('amount', 0) for t in transactions)
        sheet_data['total_volume'] = len(transactions)

    print(f"   Sheet '{sheet_name}': Type={sheet_data['type']}, Amount={sheet_data['total_amount']:,}, Volume={sheet_data['total_volume']}")

    return sheet_data

def determine_transaction_type(sheet_name, df):
    """
    Determine the type of transactions in the sheet

    Args:
        sheet_name (str): Sheet name
        df (DataFrame): Sheet data

    Returns:
        str: Transaction type (international, domestic, disputes)
    """
    sheet_name_lower = sheet_name.lower()

    # Check sheet name first
    if 'international' in sheet_name_lower or 'intl' in sheet_name_lower:
        return 'international'
    elif 'domestic' in sheet_name_lower:
        return 'domestic'
    elif 'dispute' in sheet_name_lower or 'vrol' in sheet_name_lower:
        return 'disputes'

    # Check content of the sheet - enhanced detection
    sheet_content = ' '.join(df.astype(str).values.flatten()).lower()

    # Look for specific patterns that indicate transaction type
    if 'international transaction' in sheet_content or 'international transations' in sheet_content:
        return 'international'
    elif 'domestic transaction' in sheet_content:
        return 'domestic'
    elif 'dispute' in sheet_content or 'vrol' in sheet_content or 'transaction dispute' in sheet_content:
        return 'disputes'

    # Additional check: Look at the data patterns
    # If we see transaction IDs like A, B, C... it's likely international
    # If we see transaction IDs like K, L, M... it might be domestic
    transaction_ids = []
    for idx, row in df.iterrows():
        for col in df.columns:
            val = str(row[col]).strip()
            if len(val) == 1 and val.isalpha():
                transaction_ids.append(val)

    if transaction_ids:
        # If we see early alphabet letters, likely international
        if any(tid in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'] for tid in transaction_ids):
            return 'international'
        # If we see later alphabet letters, likely domestic
        elif any(tid in ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'] for tid in transaction_ids):
            return 'domestic'

    return 'unknown'

def find_transaction_entries(df):
    """
    Find individual transaction entries in the dataframe

    Args:
        df (DataFrame): Sheet data

    Returns:
        list: List of transaction dictionaries
    """
    transactions = []

    # Look for transaction ID and amount columns
    transaction_cols = {'id': None, 'amount': None}

    # Find columns that might contain transaction IDs and amounts
    for col in df.columns:
        col_str = str(col).lower()
        if 'transaction' in col_str and 'id' in col_str:
            transaction_cols['id'] = col
        elif 'amount' in col_str:
            transaction_cols['amount'] = col

    # If we found both columns, extract transactions
    if transaction_cols['id'] is not None and transaction_cols['amount'] is not None:
        for idx, row in df.iterrows():
            transaction_id = row[transaction_cols['id']]
            amount = row[transaction_cols['amount']]

            if (pd.notna(transaction_id) and pd.notna(amount) and
                str(transaction_id).strip() not in ['transaction id', 'id', ''] and
                str(amount).replace(',', '').replace('.', '').isdigit()):

                transactions.append({
                    'id': str(transaction_id).strip(),
                    'amount': float(str(amount).replace(',', ''))
                })

    # Alternative: Look for any ID-amount patterns in the data
    if not transactions:
        transactions = find_id_amount_patterns(df)

    return transactions

def find_id_amount_patterns(df):
    """
    Look for ID-amount patterns in the dataframe

    Args:
        df (DataFrame): Sheet data

    Returns:
        list: List of transaction dictionaries
    """
    transactions = []

    for idx, row in df.iterrows():
        row_values = []
        for col in df.columns:
            val = row[col]
            if pd.notna(val):
                row_values.append(str(val).strip())

        # Look for patterns like: single character/short string followed by number
        if len(row_values) >= 2:
            for i in range(len(row_values) - 1):
                id_val = row_values[i]
                amount_val = row_values[i + 1]

                # Check if first value looks like ID and second like amount
                if (len(id_val) <= 3 and id_val.isalnum() and
                    amount_val.replace(',', '').replace('.', '').isdigit() and
                    float(amount_val.replace(',', '')) > 0):

                    transactions.append({
                        'id': id_val,
                        'amount': float(amount_val.replace(',', ''))
                    })
                    break

    return transactions

def find_transaction_totals(df):
    """
    Find total amount and volume in the dataframe

    Args:
        df (DataFrame): Sheet data

    Returns:
        dict: Total amount and volume
    """
    totals = {}

    # Keywords for totals - enhanced for better detection
    total_keywords = ['total', 'sum', 'grand total', 'total of', 'total amount']
    volume_keywords = ['volume', 'count', 'number', 'transactions', 'volume of', 'no of', 'disputed transactions']

    for idx, row in df.iterrows():
        for col in df.columns:
            cell_value = str(row[col]).lower().strip()

            # Check for total amount - enhanced detection
            if any(keyword in cell_value for keyword in total_keywords):
                # Look for number in adjacent cells
                for next_col in df.columns:
                    adjacent_value = row[next_col]
                    if (pd.notna(adjacent_value) and
                        str(adjacent_value).replace(',', '').replace('.', '').isdigit()):
                        amount = float(str(adjacent_value).replace(',', ''))
                        if amount > 1000:  # Likely a total amount
                            totals['amount'] = amount
                            print(f"   Found total amount: {amount:,}")
                        break

            # Check for volume - enhanced detection
            elif any(keyword in cell_value for keyword in volume_keywords):
                # Look for number in adjacent cells
                for next_col in df.columns:
                    adjacent_value = row[next_col]
                    if (pd.notna(adjacent_value) and
                        str(adjacent_value).replace(',', '').isdigit()):
                        volume = int(str(adjacent_value).replace(',', ''))
                        if volume < 10000:  # Likely a volume count
                            totals['volume'] = volume
                            print(f"   Found transaction volume: {volume}")
                        break

            # Special check for dispute-specific patterns
            elif ('no of disputed' in cell_value or 'number of disputed' in cell_value or
                  'disputed transactions' in cell_value):
                # Look for number in adjacent cells
                for next_col in df.columns:
                    adjacent_value = row[next_col]
                    if (pd.notna(adjacent_value) and
                        str(adjacent_value).replace(',', '').isdigit()):
                        volume = int(str(adjacent_value).replace(',', ''))
                        totals['volume'] = volume
                        print(f"   Found disputed transaction count: {volume}")
                        break

    # Additional check: Look for patterns like "Total of Domestic Transactions" with value in next row/column
    for idx, row in df.iterrows():
        for col in df.columns:
            cell_value = str(row[col]).lower().strip()

            # Look for domestic/international transaction totals specifically
            if ('total of' in cell_value and ('domestic' in cell_value or 'international' in cell_value)) or \
               ('total' in cell_value and 'transaction' in cell_value):

                # Check adjacent cells and next row for the actual value
                for check_col in df.columns:
                    # Check same row, different column
                    adjacent_value = row[check_col]
                    if (pd.notna(adjacent_value) and
                        str(adjacent_value).replace(',', '').replace('.', '').isdigit() and
                        float(str(adjacent_value).replace(',', '')) > 1000):
                        totals['amount'] = float(str(adjacent_value).replace(',', ''))
                        print(f"   Found total from pattern matching: {totals['amount']:,}")
                        break

                # Check next row, same column
                if idx + 1 < len(df):
                    next_row_value = df.iloc[idx + 1][col]
                    if (pd.notna(next_row_value) and
                        str(next_row_value).replace(',', '').replace('.', '').isdigit() and
                        float(str(next_row_value).replace(',', '')) > 1000):
                        totals['amount'] = float(str(next_row_value).replace(',', ''))
                        print(f"   Found total from next row: {totals['amount']:,}")
                        break

    return totals

def merge_transaction_data(main_data, sheet_data, sheet_name):
    """
    Merge sheet transaction data into main transaction data

    Args:
        main_data (dict): Main transaction data dictionary
        sheet_data (dict): Sheet transaction data
        sheet_name (str): Sheet name for logging
    """
    transaction_type = sheet_data['type']

    if transaction_type in main_data:
        # For each transaction type, only update if we haven't seen this type before
        # or if the new data is more complete
        if main_data[transaction_type]['total_amount'] == 0:
            # This is the first time we're seeing this transaction type
            main_data[transaction_type]['total_amount'] = sheet_data['total_amount']
            main_data[transaction_type]['total_volume'] = sheet_data['total_volume']
            main_data[transaction_type]['transactions'] = sheet_data['transactions'].copy()
        else:
            # We already have data for this type, only add if amounts are different
            # This handles cases where the same file has multiple sheets
            if sheet_data['total_amount'] != main_data[transaction_type]['total_amount']:
                print(f"   Warning: Found different {transaction_type} amounts in multiple sheets")
                # Take the larger amount (likely the total)
                if sheet_data['total_amount'] > main_data[transaction_type]['total_amount']:
                    main_data[transaction_type]['total_amount'] = sheet_data['total_amount']
                    main_data[transaction_type]['total_volume'] = sheet_data['total_volume']
    else:
        print(f"Warning: Unknown transaction type '{transaction_type}' in sheet '{sheet_name}'")

def calculate_transaction_totals(transaction_data):
    """
    Calculate overall totals from all transaction types

    Args:
        transaction_data (dict): Transaction data dictionary
    """
    total_amount = 0
    total_volume = 0

    for trans_type in ['international', 'domestic', 'disputes']:
        if trans_type in transaction_data:
            total_amount += transaction_data[trans_type]['total_amount']
            total_volume += transaction_data[trans_type]['total_volume']

    transaction_data['all_transactions']['total_amount'] = total_amount
    transaction_data['all_transactions']['total_volume'] = total_volume

    print(f"\nTransaction Data Summary:")
    print(f"International: Amount=${transaction_data['international']['total_amount']:,}, Volume={transaction_data['international']['total_volume']}")
    print(f"Domestic: Amount=${transaction_data['domestic']['total_amount']:,}, Volume={transaction_data['domestic']['total_volume']}")
    print(f"Disputes: Amount=${transaction_data['disputes']['total_amount']:,}, Volume={transaction_data['disputes']['total_volume']}")
    print(f"Total: Amount=${total_amount:,}, Volume={total_volume}")

def display_results(analysis_results, card_data=None, transaction_data=None):
    """
    Display the analysis results in terminal using tabular format with calculations

    Args:
        analysis_results (dict): Analysis results
        card_data (dict): Card issuance data for calculations
        transaction_data (dict): Transaction data for calculations
    """
    if not analysis_results:
        print("No analysis results to display")
        return

    print("\n" + "="*80)
    print("DYNAMIC FEE RATE MAPPING & CALCULATION RESULTS")
    print("="*80)

    if card_data:
        print(f"\nCard Issuance Data:")
        print(f"Total Cards Issued: {card_data['total_cards']:,}")
        if card_data['monthly_data']:
            monthly_summary = ", ".join([f"{item['period']}: {item['cards']:,}" for item in card_data['monthly_data']])
            print(f"Monthly Breakdown: {monthly_summary}")

    if transaction_data:
        print(f"\nTransaction Data Summary:")
        if transaction_data['international']['total_amount'] > 0:
            print(f"International Transactions: Amount=${transaction_data['international']['total_amount']:,}, Volume={transaction_data['international']['total_volume']}")
        if transaction_data['domestic']['total_amount'] > 0:
            print(f"Domestic Transactions: Amount=${transaction_data['domestic']['total_amount']:,}, Volume={transaction_data['domestic']['total_volume']}")
        if transaction_data['disputes']['total_amount'] > 0:
            print(f"Disputes: Amount=${transaction_data['disputes']['total_amount']:,}, Volume={transaction_data['disputes']['total_volume']}")

    if card_data or transaction_data:
        print("-" * 80)

    total_mappings = 0
    all_mappings = []
    total_calculated_amount = 0
    total_final_amount_inr = 0

    for sheet_name, sheet_data in analysis_results['mappings'].items():
        if sheet_data['mappings']:
            print(f"\nSheet: '{sheet_name}'")
            print("-" * 80)

            # Prepare data for tabular display with calculations
            table_data = []
            for i, (fee_type, rate_chart) in enumerate(sheet_data['mappings'].items(), 1):
                calculated_amount = 0
                calculation_method = "N/A"

                if card_data or transaction_data:
                    # Determine appropriate values for calculation
                    card_count = card_data['total_cards'] if card_data else 0

                    # Use dynamic transaction data if available
                    transaction_count = 0
                    transaction_amount = 0

                    if transaction_data:
                        # Choose appropriate transaction data based on fee type
                        fee_type_lower = fee_type.lower()
                        if 'international' in fee_type_lower:
                            transaction_count = transaction_data['international']['total_volume']
                            transaction_amount = transaction_data['international']['total_amount']
                        elif 'domestic' in fee_type_lower:
                            transaction_count = transaction_data['domestic']['total_volume']
                            transaction_amount = transaction_data['domestic']['total_amount']
                        elif 'dispute' in fee_type_lower:
                            transaction_count = transaction_data['disputes']['total_volume']
                            transaction_amount = transaction_data['disputes']['total_amount']
                        else:
                            # Use total values for general fees
                            transaction_count = transaction_data['all_transactions']['total_volume']
                            transaction_amount = transaction_data['all_transactions']['total_amount']

                    # Calculate fee based on rate formula and actual data
                    calc_result = calculate_fee_amount(
                        rate_chart,
                        card_count,
                        transaction_count=transaction_count,
                        transaction_amount=transaction_amount
                    )
                    calculated_amount = calc_result['calculated_amount']
                    calculation_method = calc_result['calculation_method']
                    total_calculated_amount += calculated_amount

                    # Determine exchange rate and calculate final amount
                    currency_symbol = calc_result.get('currency', '$')
                    exchange_rate = 78 if currency_symbol == '$' else 1  # USD to INR = 78, INR = 1

                    formatted_amount = f"{currency_symbol}{calculated_amount:,.2f}" if calculated_amount > 0 else "N/A"
                    final_amount = calculated_amount * exchange_rate if calculated_amount > 0 else 0
                    formatted_final_amount = f"₹{final_amount:,.2f}" if final_amount > 0 else "N/A"

                    # Add to totals
                    total_final_amount_inr += final_amount

                table_data.append([
                    i,
                    fee_type,
                    rate_chart,
                    calculation_method,
                    formatted_amount,
                    exchange_rate,
                    formatted_final_amount
                ])
                total_mappings += 1

            # Display table for this sheet
            headers = ["S.No.", "Fee Type", "Rate Chart", "Calculation Method", "Calculated Amount", "Exchange Rate", "Final Amount (INR)"]
            print(tabulate(table_data, headers=headers, tablefmt="grid",
                         maxcolwidths=[5, 25, 30, 40, 15, 12, 15]))
            print()

            # Store for summary table
            for row in table_data:
                all_mappings.append([sheet_name] + row[1:])

    if total_mappings == 0:
        print("No fee-rate mappings found in any sheet.")
        print("The file might have a different structure than expected.")
        print("Try checking if the file contains fee types and rates in separate columns.")
    else:
        # Display summary table if multiple sheets
        if len(analysis_results['mappings']) > 1:
            print("\n" + "="*80)
            print("SUMMARY - ALL SHEETS")
            print("="*80)
            summary_headers = ["Sheet", "Fee Type", "Rate Chart", "Calculation Method", "Calculated Amount", "Exchange Rate", "Final Amount (INR)"]
            print(tabulate(all_mappings, headers=summary_headers, tablefmt="grid",
                         maxcolwidths=[15, 25, 30, 40, 15, 12, 15]))
            print()

        if (card_data or transaction_data) and total_final_amount_inr > 0:
            print("="*80)
            print(f"TOTAL AMOUNT: ₹{total_final_amount_inr:,.2f}")
            print("="*80)

        print(f"Total mappings found: {total_mappings}")
        print("Analysis completed successfully!")

def main():
    """Main function"""
    print("Dynamic Fee Rate Mapper & Calculator Tool")
    print("=" * 45)
    print("This tool will analyze fee rates and calculate amounts based on card issuance and transaction data.")

    # Select all files using tkinter
    files = select_multiple_files()

    if not files['summary']:
        print("No summary file selected. Exiting...")
        sys.exit(0)

    if not os.path.exists(files['summary']):
        print(f"Summary file not found: {files['summary']}")
        sys.exit(1)

    # Analyze the summary file for fee rates
    print("\n" + "="*60)
    print("ANALYZING SUMMARY FILE FOR FEE RATES")
    print("="*60)
    analysis_results = analyze_excel_structure(files['summary'])

    # Extract card issuance data if card file is provided
    card_data = None
    if files['card_issuance'] and os.path.exists(files['card_issuance']):
        print("\n" + "="*60)
        print("ANALYZING CARD ISSUANCE FILE")
        print("="*60)
        card_data = extract_card_issuance_data(files['card_issuance'])

        if not card_data or card_data['total_cards'] == 0:
            print("Warning: No card issuance data found or total cards is 0")
            print("Proceeding with fee rate mapping only...")
            card_data = None
    elif files['card_issuance']:
        print(f"Warning: Card issuance file not found: {files['card_issuance']}")
        print("Proceeding with fee rate mapping only...")

    # Extract transaction data from separate files
    transaction_data = {
        'international': {'total_amount': 0, 'total_volume': 0, 'transactions': []},
        'domestic': {'total_amount': 0, 'total_volume': 0, 'transactions': []},
        'disputes': {'total_amount': 0, 'total_volume': 0, 'transactions': []},
        'all_transactions': {'total_amount': 0, 'total_volume': 0}
    }

    print("\n" + "="*60)
    print("PROCESSING TRANSACTION FILES")
    print("="*60)

    # Process international transactions
    if files['international_file']:
        transaction_data['international'] = process_specific_transaction_file(
            files['international_file'], 'international'
        )

    # Process domestic transactions
    if files['domestic_file']:
        transaction_data['domestic'] = process_specific_transaction_file(
            files['domestic_file'], 'domestic'
        )

    # Process dispute transactions
    if files['dispute_file']:
        transaction_data['disputes'] = process_specific_transaction_file(
            files['dispute_file'], 'disputes'
        )

    # Calculate totals
    transaction_data['all_transactions']['total_amount'] = (
        transaction_data['international']['total_amount'] +
        transaction_data['domestic']['total_amount'] +
        transaction_data['disputes']['total_amount']
    )
    transaction_data['all_transactions']['total_volume'] = (
        transaction_data['international']['total_volume'] +
        transaction_data['domestic']['total_volume'] +
        transaction_data['disputes']['total_volume']
    )

    print(f"\nTransaction Processing Summary:")
    print(f"International: Amount=${transaction_data['international']['total_amount']:,}, Volume={transaction_data['international']['total_volume']}")
    print(f"Domestic: Amount=${transaction_data['domestic']['total_amount']:,}, Volume={transaction_data['domestic']['total_volume']}")
    print(f"Disputes: Amount=${transaction_data['disputes']['total_amount']:,}, Volume={transaction_data['disputes']['total_volume']}")

    # Set to None if no data found
    if transaction_data['all_transactions']['total_volume'] == 0:
        print("Warning: No transaction data found")
        transaction_data = None

    # Display results with calculations
    display_results(analysis_results, card_data, transaction_data)

if __name__ == "__main__":
    main()
