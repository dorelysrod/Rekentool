#!/usr/bin/env python3
import pandas as pd
import matplotlib.pyplot as plt
import os
import logging
import re

logging.basicConfig(level=logging.INFO)

def load_client_data(filepath):
    """
    Laadt klantgegevens vanuit Excel.
    """
    try:
        df = pd.read_excel(filepath, sheet_name='Opdracht_1_voorbeeld', header=None)
        df = df.astype(str)

        building_type, area_m2, current_consumption = None, None, None

        for _, row in df.iterrows():
            for col_idx, value in row.items():
                if "Categorie" in value:
                    building_type = row.get(col_idx + 1)
                if "Oppervlakte" in value:
                    try:
                        area_m2 = float(row.get(col_idx + 1))
                    except:
                        pass
                if "elektriciteit" in value:
                    try:
                        current_consumption = float(row.get(col_idx + 1))
                    except:
                        pass

        if not building_type or not area_m2 or not current_consumption:
            raise ValueError("Niet alle klantgegevens konden worden geladen.")

        logging.info(f"Klantgegevens geladen: {building_type}, {area_m2} m2, {current_consumption} kWh/m2")
        return {'building_type': building_type, 'area_m2': area_m2, 'current_consumption': current_consumption}

    except Exception as e:
        logging.error(f"Fout bij laden van klantgegevens: {e}")
        raise

def load_cbs_data_api():
    """
    Laadt CBS data via API en merge dimension tables.
    """
    try:
        url_data = "https://opendata.cbs.nl/ODataApi/OData/83374NED/TypedDataSet"
        df_api_raw = pd.read_json(url_data)
        df_api = pd.json_normalize(df_api_raw['value'])

        url_dim_cat = "https://opendata.cbs.nl/ODataApi/OData/83374NED/UtiliteitsbouwDienstensector"
        df_dim_cat_raw = pd.read_json(url_dim_cat)
        df_dim_cat = pd.json_normalize(df_dim_cat_raw['value'])[['Key', 'Title']]

        url_dim_surface = "https://opendata.cbs.nl/ODataApi/OData/83374NED/Oppervlakteklasse"
        df_dim_surface_raw = pd.read_json(url_dim_surface)
        df_dim_surface = pd.json_normalize(df_dim_surface_raw['value'])[['Key', 'Title']]

        df_api = df_api.merge(df_dim_cat, left_on='UtiliteitsbouwDienstensector', right_on='Key', how='left')
        df_api = df_api.merge(df_dim_surface, left_on='Oppervlakteklasse', right_on='Key', how='left', suffixes=('_cat', '_surface'))

        logging.info("CBS data en dimensietabellen succesvol geladen en gemerged.")
        return df_api

    except Exception as e:
        logging.error(f"Fout bij laden van CBS data via API: {e}")
        raise

def parse_surface_class(surface_class):
    """
    Parseert oppervlakteklasse naar min en max m2.
    """
    try:
        cleaned = surface_class.replace('.', '').replace(',', '').replace(' ', '')
        numbers = re.findall(r'\d+', cleaned)
        if len(numbers) >= 2:
            min_m2 = int(numbers[0])
            max_m2 = int(numbers[1])
            return min_m2, max_m2
        else:
            return None, None
    except:
        return None, None

def classify_and_get_sector_average(client, cbs_df):
    """
    Classificeert klant in sector en haalt sectorgemiddelde op.
    """
    try:
        CATEGORY_COL = 'Title_cat'
        SURFACE_CLASS_COL = 'Title_surface'
        AVG_CONSUMPTION_COL = 'GemiddeldElektriciteitsverbruik_2'

        matches = cbs_df[cbs_df[CATEGORY_COL].str.lower().str.contains(client['building_type'].lower(), na=False)]

        if matches.empty:
            logging.warning(f"Geen categorie match gevonden voor {client['building_type']}.")
            return None

        logging.info(f"Beschikbare oppervlakteranges voor {client['building_type']}:")

        for _, row in matches.iterrows():
            surface_class = row[SURFACE_CLASS_COL]
            min_m2, max_m2 = parse_surface_class(surface_class)
            avg_consumption = row[AVG_CONSUMPTION_COL]

            logging.info(f"Range: {surface_class}, min: {min_m2}, max: {max_m2}, avg: {avg_consumption}")

            if min_m2 is not None and max_m2 is not None:
                if min_m2 <= client['area_m2'] <= max_m2:
                    logging.info(f"Sectorgemiddelde gevonden: {avg_consumption} kWh/m2")
                    return avg_consumption

        logging.warning(f"Geen exacte match, hoogste range wordt als fallback gebruikt.")
        matches = matches.copy()
        matches['max_m2'] = matches[SURFACE_CLASS_COL].apply(lambda x: parse_surface_class(x)[1])
        highest_range = matches.sort_values(by='max_m2', ascending=False).iloc[0]
        highest_avg = highest_range[AVG_CONSUMPTION_COL]

        if pd.notnull(highest_avg):
            logging.info(f"Sectorgemiddelde van hoogste range gebruikt: {highest_avg} kWh/m2")
            return highest_avg
        else:
            logging.warning("Geen sectorgemiddelde beschikbaar in hoogste range.")
            return None

    except Exception as e:
        logging.error(f"Fout bij classificatie: {e}")
        raise

def plot_consumption_comparison(client_value, sector_average, difference, output_path):
    """
    Genereert en slaat een plot op van klantverbruik vs. sectorgemiddelde.
    """
    try:
        labels = ['Klant', 'Gemiddelde sector']
        values = [client_value, sector_average]

        plt.figure(figsize=(6,4))
        bars = plt.bar(labels, values, color=['gray', 'green'])
        plt.ylabel('Verbruik (kWh/m2)')

        if sector_average == 0:
            plt.title('Geen sectorgemiddelde beschikbaar')
        else:
            plt.title(f'Verbruik klant vs. sector (verschil: {difference:.2f}%)')

        plt.grid(axis='y', linestyle='--', alpha=0.7)

        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, height + 1, f'{height:.2f}', ha='center')

        plt.tight_layout()
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        plt.savefig(output_path)
        plt.close()
        logging.info(f"Plot opgeslagen in {output_path}")

    except Exception as e:
        logging.error(f"Fout bij genereren van plot: {e}")
        raise

def sector_comparison_pipeline(filepath, output_path):
    """
    Pipeline functie voor vergelijking klantverbruik met sectorgemiddelde.
    """
    try:
        client = load_client_data(filepath)
        cbs_df = load_cbs_data_api()
        sector_average = classify_and_get_sector_average(client, cbs_df)

        if sector_average is None:
            logging.warning("Geen sectorgemiddelde beschikbaar, plot toont alleen klantverbruik.")
            sector_average = 0

        difference = ((client['current_consumption'] - sector_average) / sector_average) * 100 if sector_average else 0

        plot_consumption_comparison(client['current_consumption'], sector_average, difference, output_path)

        logging.info(f"Klantverbruik: {client['current_consumption']} kWh/m2")
        logging.info(f"Sectorgemiddelde: {sector_average} kWh/m2")
        logging.info(f"Verschil: {difference:.2f}%")

    except Exception as e:
        logging.error(f"Pipeline fout: {e}")

if __name__ == "__main__":
    filepath = '../data/PraeterBV_Case.xlsx'
    output_path = 'outputs/plots/sector_comparison.png'
    sector_comparison_pipeline(filepath, output_path)
