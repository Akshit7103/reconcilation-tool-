import os
import sys
import io
import tempfile
import shutil
import importlib.util
from contextlib import redirect_stdout
from typing import Optional, Dict

from werkzeug.utils import secure_filename

BASE_DIR = os.path.dirname(__file__)

# Use the local rate_tool_app.py file in the same directory
RATE_APP_FILE = os.path.join(BASE_DIR, 'rate_tool_app.py')

if not os.path.exists(RATE_APP_FILE):
    raise FileNotFoundError(f"Could not locate rate calculator app.py at: {RATE_APP_FILE}")

module_name = 'rate_tool_app'
spec = importlib.util.spec_from_file_location(module_name, RATE_APP_FILE)
rate_tool_app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rate_tool_app)

analyze_excel_structure = rate_tool_app.analyze_excel_structure
extract_card_issuance_data = rate_tool_app.extract_card_issuance_data
process_specific_transaction_file = rate_tool_app.process_specific_transaction_file
calculate_fee_amount = rate_tool_app.calculate_fee_amount

ALLOWED_EXTENSIONS = {"xls", "xlsx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_file(file_storage, base_dir: str) -> Optional[str]:
    if not file_storage or file_storage.filename == "":
        return None

    if not allowed_file(file_storage.filename):
        raise ValueError("Only Excel files (.xls, .xlsx) are supported.")

    filename = secure_filename(file_storage.filename)
    file_path = os.path.join(base_dir, filename)
    file_storage.save(file_path)
    return file_path


def format_currency(amount, currency_symbol):
    if amount is None or amount == 0:
        return "N/A"
    if currency_symbol == "Rs":
        return f"INR {amount:,.2f}"
    return f"USD {amount:,.2f}"


def build_result_context(analysis_results, card_data, transaction_data, warnings):
    context = {
        "has_data": bool(analysis_results and analysis_results.get("mappings")),
        "summary": {
            "sheet_count": 0,
            "total_mappings": 0,
            "total_final_amount_inr": 0.0,
            "total_final_amount_display": "N/A"
        },
        "card": None,
        "transactions": None,
        "sheets": [],
        "warnings": warnings
    }

    if not analysis_results or not analysis_results.get("mappings"):
        return context

    sheet_count = len(analysis_results["mappings"])
    total_final_amount_inr = 0.0
    total_mappings = 0

    sheets_presentations = []

    for sheet_name, sheet_data in analysis_results["mappings"].items():
        rows = []
        mappings = sheet_data.get("mappings", {})
        if not mappings:
            continue

        for fee_type, rate_chart in mappings.items():
            card_count = card_data["total_cards"] if card_data else 0
            transaction_count = 0
            transaction_amount = 0

            if transaction_data:
                fee_type_lower = fee_type.lower()
                if "international" in fee_type_lower:
                    trans_bucket = transaction_data.get("international", {})
                elif "domestic" in fee_type_lower:
                    trans_bucket = transaction_data.get("domestic", {})
                elif "dispute" in fee_type_lower:
                    trans_bucket = transaction_data.get("disputes", {})
                else:
                    trans_bucket = transaction_data.get("all_transactions", {})

                transaction_count = trans_bucket.get("total_volume", 0)
                transaction_amount = trans_bucket.get("total_amount", 0)

            calc_result = calculate_fee_amount(
                rate_chart,
                card_count,
                transaction_count=transaction_count,
                transaction_amount=transaction_amount
            )

            calculated_amount = calc_result.get("calculated_amount", 0)
            calculation_method = calc_result.get("calculation_method", "N/A")
            currency_symbol = calc_result.get("currency", "$") or "$"
            exchange_rate = 78 if currency_symbol == "$" else 1

            if calculated_amount and calculated_amount > 0:
                formatted_amount = format_currency(calculated_amount, "Rs" if currency_symbol == "Rs" else "$")
                final_amount = calculated_amount * exchange_rate
                final_amount_display = f"INR {final_amount:,.2f}"
                total_final_amount_inr += final_amount
            else:
                formatted_amount = "N/A"
                final_amount = None
                final_amount_display = "N/A"

            rows.append({
                "fee_type": fee_type,
                "rate_chart": rate_chart,
                "calculation_method": calculation_method,
                "calculated_amount": calculated_amount,
                "calculated_amount_display": formatted_amount,
                "exchange_rate": exchange_rate if calculated_amount else None,
                "final_amount": final_amount,
                "final_amount_display": final_amount_display,
                "currency_symbol": currency_symbol
            })
            total_mappings += 1

        if rows:
            sheets_presentations.append({
                "name": sheet_name,
                "rows": rows
            })

    context["sheets"] = sheets_presentations
    context["summary"] = {
        "sheet_count": sheet_count,
        "total_mappings": total_mappings,
        "total_final_amount_inr": total_final_amount_inr,
        "total_final_amount_display": f"INR {total_final_amount_inr:,.2f}" if total_final_amount_inr else "N/A"
    }

    if card_data:
        context["card"] = {
            "total_cards": card_data.get("total_cards", 0),
            "monthly_data": card_data.get("monthly_data", [])
        }

    if transaction_data:
        transaction_cards = []
        for key, label in (
            ("international", "International"),
            ("domestic", "Domestic"),
            ("disputes", "Disputes"),
            ("all_transactions", "All Transactions")
        ):
            bucket = transaction_data.get(key)
            if not bucket:
                continue
            if bucket.get("total_volume", 0) == 0 and bucket.get("total_amount", 0) == 0 and key != "all_transactions":
                continue
            transaction_cards.append({
                "label": label,
                "amount": bucket.get("total_amount", 0),
                "amount_display": f"USD {bucket.get('total_amount', 0):,.2f}" if bucket.get("total_amount", 0) else "N/A",
                "volume": bucket.get("total_volume", 0)
            })

        if transaction_cards:
            context["transactions"] = {
                "entries": transaction_cards
            }

    return context


def run_rate_analysis(file_paths: Dict[str, Optional[str]]):
    warnings = []

    with redirect_stdout(io.StringIO()):
        analysis_results = analyze_excel_structure(file_paths.get('summary'))

        card_data = None
        if file_paths.get('card'):
            card_data = extract_card_issuance_data(file_paths['card'])
            if not card_data or card_data.get("total_cards", 0) == 0:
                warnings.append("No card issuance data found or total cards is 0. Proceeding with fee rate mapping only.")
                card_data = None

        transaction_data = {
            "international": {"total_amount": 0, "total_volume": 0, "transactions": []},
            "domestic": {"total_amount": 0, "total_volume": 0, "transactions": []},
            "disputes": {"total_amount": 0, "total_volume": 0, "transactions": []},
            "all_transactions": {"total_amount": 0, "total_volume": 0}
        }

        if file_paths.get('international'):
            transaction_data["international"] = process_specific_transaction_file(
                file_paths['international'], "international"
            )

        if file_paths.get('domestic'):
            transaction_data["domestic"] = process_specific_transaction_file(
                file_paths['domestic'], "domestic"
            )

        if file_paths.get('dispute'):
            transaction_data["disputes"] = process_specific_transaction_file(
                file_paths['dispute'], "disputes"
            )

        transaction_data["all_transactions"]["total_amount"] = (
            transaction_data["international"]["total_amount"] +
            transaction_data["domestic"]["total_amount"] +
            transaction_data["disputes"]["total_amount"]
        )
        transaction_data["all_transactions"]["total_volume"] = (
            transaction_data["international"]["total_volume"] +
            transaction_data["domestic"]["total_volume"] +
            transaction_data["disputes"]["total_volume"]
        )

        if transaction_data["all_transactions"]["total_volume"] == 0:
            warnings.append("No transaction data found.")
            transaction_data = None

    report_context = build_result_context(analysis_results, card_data, transaction_data, warnings)
    return report_context

