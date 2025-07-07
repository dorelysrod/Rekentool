#!/usr/bin/env python3
import pandas as pd
import logging
from typing import List
import os

logging.basicConfig(level=logging.INFO)

def load_cashflows(filepath: str) -> List[float]:
    """
    Laadt initiële investering en jaarlijkse cashflows dynamisch vanuit Excel.
    """
    try:
        df_input = pd.read_excel(filepath, sheet_name='Opdracht_2_input_output', header=None)
        initial_investment_row = df_input[df_input.apply(lambda row: row.astype(str).str.contains('Investering', case=False).any(), axis=1)]
        if not initial_investment_row.empty:
            initial_investment_value = initial_investment_row.iloc[0, initial_investment_row.iloc[0].last_valid_index() + 1]
            initial_investment = -abs(float(str(initial_investment_value).replace(',', '').replace('%', '').strip()))
            logging.info(f"Initiële investering geladen: {initial_investment} EUR")
        else:
            raise ValueError("Investering niet gevonden.")
    except Exception as e:
        logging.error(f"Fout bij laden van initiële investering: {e}")
        raise

    try:
        df_output = pd.read_excel(filepath, sheet_name='Opdracht_2_berekening en output', header=None)
        cashflow_column = df_output.iloc[10:20, 15].dropna().astype(float).tolist()
        logging.info(f"Jaarlijkse cashflows geladen: {cashflow_column}")
    except Exception as e:
        logging.error(f"Fout bij laden van cashflows: {e}")
        raise

    return [initial_investment] + cashflow_column

def export_cashflows_to_excel(cashflows: List[float], output_path: str):
    """
    Exporteer cashflows met cumulatieve som naar Excel.
    """
    years = list(range(len(cashflows)))
    cumulative = []
    total = 0
    for cf in cashflows:
        total += cf
        cumulative.append(total)

    df_export = pd.DataFrame({
        'Jaar': years,
        'Cashflow (EUR)': cashflows,
        'Cumulatieve Cashflow (EUR)': cumulative
    })

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_export.to_excel(output_path, index=False)
    logging.info(f"Cashflowtabel succesvol opgeslagen in {output_path}")

def main(filepath: str = '../data/PraeterBV_Case.xlsx', output_path: str = 'outputs/cashflow_table.xlsx'):
    """
    Pipeline functie.
    """
    cashflows = load_cashflows(filepath)
    export_cashflows_to_excel(cashflows, output_path)

if __name__ == "__main__":
    main()
