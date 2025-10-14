# parser.py
import pandas as pd
import re
from typing import Dict, Tuple

# --- 1. Definición de Plantillas y Aliases ---
COLUMN_MAP_BASIC = {
    'TITLE': ['Title', 'Event Title', 'Name'],
    'START_TIME': ['StartTime', 'Start Time', 'Start', 'DTSTART'],
    'END_TIME': ['EndTime', 'End Time', 'End', 'DTEND'],
    'LOCATION': ['Location', 'Place', 'Address'],
    'OWNER': ['Owner', 'Organizer', 'Creator'],
    'PARTICIPANTS': ['Participants', 'Attendees', 'Invitees'],
    'DESCRIPTION': ['Description', 'Details'],
    'LINK': ['Link', 'Meeting Link', 'Zoom', 'Meet', 'URL']
}

COLUMN_MAP_MEDIUM = {
    'TITLE': ['Title', 'Event Title', 'Name'],
    'START_DATE': ['Start Date', 'Date Start'],
    'END_DATE': ['End Date', 'Date End'],
    'START_TIME': ['StartTime', 'Start Time', 'Time Start'],
    'END_TIME': ['EndTime', 'End Time', 'Time End'],
    'FREQUENCY': ['Frequency', 'Recurrence', 'Repeats', 'Days'],
    'LOCATION': ['Location', 'Place', 'Address'],
    'OWNER': ['Owner', 'Organizer', 'Creator'],
    'PARTICIPANTS': ['Participants', 'Attendees', 'Invitees'],
    'DESCRIPTION': ['Description', 'Details'],
    'LINK': ['Link', 'Meeting Link', 'Zoom', 'Meet', 'URL']
}


def identify_template_and_map_columns(sheet_df: pd.DataFrame) -> Tuple[str, Dict[str, str]]:
    """Identifica la plantilla y devuelve el mapeo de columnas canónicas a reales."""
    
    # Limpiar columnas para comparación: minusculas y sin caracteres especiales
    actual_cols_cleaned = [re.sub(r'[^a-zA-Z0-9]', '', str(c).lower()) for c in sheet_df.columns]
    
    # Intentar Medium y luego Basic (Medium es un superset de campos complejos)
    for template_name, canonical_map in [("Medium", COLUMN_MAP_MEDIUM), ("Basic", COLUMN_MAP_BASIC)]:
        col_map = {}
        missing_required = False
        
        # Definir campos requeridos mínimos
        required_fields = []
        if template_name == 'Basic':
            required_fields = ['TITLE', 'START_TIME', 'END_TIME']
        elif template_name == 'Medium':
            required_fields = ['TITLE', 'START_DATE', 'END_DATE', 'START_TIME', 'END_TIME']
        
        # Intentar mapear todas las columnas canónicas
        for canonical, aliases in canonical_map.items():
            
            canonical_cleaned = re.sub(r'[^a-zA-Z0-9]', '', canonical.lower())
            alias_cleaned = [re.sub(r'[^a-zA-Z0-9]', '', alias.lower()) for alias in aliases]

            found_index = next((
                i for i, actual_cleaned in enumerate(actual_cols_cleaned) 
                if actual_cleaned == canonical_cleaned or actual_cleaned in alias_cleaned
            ), -1)

            if found_index != -1:
                original_col_name = sheet_df.columns[found_index]
                col_map[canonical] = original_col_name
            elif canonical in required_fields:
                missing_required = True
                break

        # Si no faltan campos requeridos y encontramos todos los mínimos
        if not missing_required and all(req in col_map for req in required_fields):
             return template_name, col_map

    raise ValueError("Could not confidently map sheet columns to Basic or Medium template (missing required fields).")

def load_and_parse_excel(filepath: str) -> Dict[str, Dict]:
    """Carga el archivo Excel y procesa cada hoja, imprimiendo logs de éxito/fallo de hoja."""
    try:
        all_sheets = pd.read_excel(filepath, sheet_name=None)
    except FileNotFoundError:
        raise FileNotFoundError(f"Error: The file '{filepath}' was not found.")
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")
        
    parsed_data = {}
    
    for sheet_name, df in all_sheets.items():
        if df.empty or len(df.columns) < 3:
            print(f"⚠️ Sheet '{sheet_name}' skipped: Sheet is empty or has too few columns.")
            continue
            
        try:
            template, col_map = identify_template_and_map_columns(df)
            
            df = df.rename(columns={v: k for k, v in col_map.items()})
            df = df.dropna(how='all', subset=['TITLE'])
            
            parsed_data[sheet_name] = {'template': template, 'data': df}
            print(f"✅ Sheet '{sheet_name}' identified as **{template}** template.")
            
        except ValueError as e:
            print(f"⚠️ Sheet '{sheet_name}' skipped: {e}")
            
    return parsed_data