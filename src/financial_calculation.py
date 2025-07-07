#!/usr/bin/env python3
import pandas as pd
import numpy_financial as npf
import logging
from typing import List, Optional, Union
import os

logging.basicConfig(level=logging.INFO)

def parse_numeric(value: Union[str, float, int]) -> Optional[float]:
    """
    Converteert strings met komma's of % naar float.
    """
    if isinstance(value, (float, int)):
        return float(value)
    if isinstance(value, str):
        value = value.replace(',', '').replace('%', '').strip()
        try:
            return float(value)
        except ValueError:
            logging.warning(f"parse_numeric: kan '{value}' niet omzetten")
            return None
    return None

def find_value_by_keyword(df: pd.DataFrame, keyword: str) -> Optional[Union[float, str]]:
    """
    Vindt exacte match van keyword (case-insensitive) en retourneert cel rechts ervan.
    """
    for i, row in df.iterrows():
        for j, cell in row.items():
            if pd.notna(cell) and keyword.lower() == str(cell).strip().lower():
                right_value = row.get(j + 1)
                parsed = parse_numeric(right_value)
                if parsed is not None:
                    logging.info(f"[EXACT] Keyword '{keyword}' gevonden op rij {i+1}, kolom {j+1}. Waarde: {parsed}")
                    return parsed
                else:
                    logging.warning(f"[EXACT] Keyword '{keyword}' gevonden op rij {i+1}, kolom {j+1}. Niet numeriek: {right_value}")
                    return str(right_value)
    logging.warning(f"Keyword '{keyword}' niet gevonden.")
    return None

def load_financial_inputs(filepath: str) -> dict:
    """
    Laadt financiële inputparameters vanuit Excel.
    """
    df = pd.read_excel(filepath, sheet_name='Opdracht_2_input_output', header=None)
    inputs = {}
    keywords = ['Inflatie', 'Eigen Vermogen', 'Rente VV', 'Belasting', 'Afschrijving']

    for key in keywords:
        clean_key = key.lower().replace(' ', '_').replace('vv', 'vv')
        inputs[clean_key] = find_value_by_keyword(df, key)

    logging.info(f"Inputs geladen: {inputs}")
    return inputs

def load_cashflows(filepath: str) -> List[float]:
    """
    Laadt investering en jaarlijkse cashflows vanuit Excel.
    """
    df = pd.read_excel(filepath, sheet_name='Opdracht_2_input_output', header=None)

    investment = find_value_by_keyword(df, 'Investering')
    if investment is None:
        raise ValueError("Investering niet gevonden.")
    investment = -abs(investment)

    cashflows = [investment]
    for key in ['Besparing', 'Subsidie (jaarlijks)', 'Eenmalige kosten', 'Vaste exploitatiekosten', 'Herinvestering']:
        value = find_value_by_keyword(df, key)
        if value is not None:
            if any(k in key.lower() for k in ['kosten', 'herinvestering']):
                cashflows.append(-abs(value))
            else:
                cashflows.append(value)

    logging.info(f"Cashflows geladen: {cashflows}")
    return cashflows

def calculate_irr(cashflows: List[float]) -> Optional[float]:
    """
    Berekent de IRR.
    """
    try:
        irr = npf.irr(cashflows)
        if pd.isna(irr):
            return None
        return irr
    except Exception as e:
        logging.error(f"Fout bij berekenen van IRR: {e}")
        return None

def calculate_rev(net_income: float, equity: float) -> Optional[float]:
    """
    Berekent het rendement op eigen vermogen.
    """
    try:
        return net_income / equity if equity else None
    except Exception as e:
        logging.error(f"Fout bij berekenen van REV: {e}")
        return None

def calculate_pat(revenue: float, costs: float, taxes: float) -> Optional[float]:
    """
    Berekent de winst na belasting.
    """
    if None in [revenue, costs, taxes]:
        logging.warning("Ontbrekende waarden voor PAT-berekening.")
        return None
    return revenue - costs - taxes

def calculate_tvt(cashflows: List[float]) -> Union[int, str]:
    """
    Berekent de terugverdientijd.
    """
    cumulative = 0
    for i, cf in enumerate(cashflows):
        cumulative += cf
        if cumulative >= 0:
            return i
    return "Niet terugverdiend"

def calculate_rev_sensitivity(filepath: str, output_path: str = 'outputs/sensitivity_rev_vs_rente.xlsx'):
    """
    Voert een gevoeligheidsanalyse uit van rente VV versus REV en exporteert resultaten naar Excel.
    """
    df = pd.read_excel(filepath, sheet_name='Opdracht_2_input_output', header=None)
    eigen_vermogen = find_value_by_keyword(df, 'Eigen Vermogen') or 0.0
    vreemd_vermogen = find_value_by_keyword(df, 'Vreemd Vermogen') or 0.0
    net_income_base = find_value_by_keyword(df, 'Winst na belasting') or 0.0

    if vreemd_vermogen == 0.0:
        logging.warning("Geen vreemd vermogen gevonden, analyse niet uitgevoerd.")
        return

    rente_range = [round(x * 0.005, 3) for x in range(0, 21)]  # 0% tot 10% in stappen van 0.5%
    results = []

    for rente in rente_range:
        interest_cost = rente * vreemd_vermogen
        adjusted_net_income = net_income_base - interest_cost
        rev = calculate_rev(adjusted_net_income, eigen_vermogen)
        results.append({
            'rente_vv (%)': rente * 100,
            'interest_cost (EUR)': interest_cost,
            'adjusted_net_income (EUR)': adjusted_net_income,
            'rev': rev
        })
        logging.info(f"Rente: {rente*100:.2f}%, Interestkosten: {interest_cost:.2f}, Adjusted net income: {adjusted_net_income:.2f}, REV: {rev:.4f}")

    df_results = pd.DataFrame(results)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_results.to_excel(output_path, index=False)
    logging.info(f"Sensitivity analyse REV vs. Rente VV opgeslagen in {output_path}")

def financial_calculation_pipeline(filepath: str = '../data/PraeterBV_Case.xlsx'):
    """
    Pipeline functie voor financiële berekeningen.
    """
    inputs = load_financial_inputs(filepath)
    cashflows = load_cashflows(filepath)

    df = pd.read_excel(filepath, sheet_name='Opdracht_2_input_output', header=None)
    net_income = find_value_by_keyword(df, 'Winst na belasting') or 0.0
    revenue = find_value_by_keyword(df, 'Besparing') or 0.0
    taxes = find_value_by_keyword(df, 'Belasting') or 0.0

    costs = sum([
        find_value_by_keyword(df, 'Eenmalige kosten') or 0.0,
        find_value_by_keyword(df, 'Vaste exploitatiekosten') or 0.0,
        find_value_by_keyword(df, 'Herinvestering') or 0.0
    ])

    irr = calculate_irr(cashflows)
    rev = calculate_rev(net_income, inputs.get('eigen_vermogen', 0.0))
    pat = calculate_pat(revenue, costs, taxes)
    tvt = calculate_tvt(cashflows)

    logging.info("=== DEFINITIEVE RESULTATEN ===")
    logging.info(f"Totale rendement (IRR): {irr:.2%}" if irr is not None else "Totale rendement (IRR): nan%")
    logging.info(f"Rendement op eigen vermogen (REV): {rev:.2f}" if rev is not None else "Rendement op eigen vermogen (REV): nan")
    logging.info(f"Winst na belasting (PAT): {pat:.2f} EUR" if pat is not None else "Winst na belasting (PAT): nan EUR")
    logging.info(f"Terugverdientijd (TVT): {tvt} jaar")

    calculate_rev_sensitivity(filepath)

if __name__ == "__main__":
    financial_calculation_pipeline()
