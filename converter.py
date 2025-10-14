# converter.py
from icalendar import Calendar, Event, vRecur
from datetime import datetime, time, date, timedelta 
import pytz 
import pandas as pd
from typing import Dict, Tuple

def _create_rrule_for_medium(frequency_str: str) -> vRecur:
    """
    Crea una regla de recurrencia (RRULE) para la plantilla Medium.
    Ahora es robusto contra nombres de días completos y abreviaturas.
    """
    
    # Mapeo de abreviaturas y nombres completos a formato iCalendar (MO, TU, WE, etc.)
    day_map = {
        'MON': 'MO', 'TUE': 'TU', 'WED': 'WE', 'THU': 'TH', 'FRI': 'FR', 'SAT': 'SA', 'SUN': 'SU',
        'MONDAY': 'MO', 'TUESDAY': 'TU', 'WEDNESDAY': 'WE', 'THURSDAY': 'TH', 
        'FRIDAY': 'FR', 'SATURDAY': 'SA', 'SUNDAY': 'SU'
    }
    
    # 1. Dividir la cadena por comas
    days_input = frequency_str.upper().split(',')
    
    byday = []
    for day_name in days_input:
        # 2. Limpiar espacios alrededor del nombre del día (CRÍTICO para " Monday")
        cleaned_day = day_name.strip() 
        
        # 3. Intentar mapeo con el nombre completo o las tres primeras letras
        if cleaned_day in day_map:
            byday.append(day_map[cleaned_day])
        elif cleaned_day[:3] in day_map:
            byday.append(day_map[cleaned_day[:3]])
            
    
    if not byday:
        raise ValueError("Frequency string is invalid or contains unrecognizable days. Check spelling or use abbreviations (e.g., MON, WED).")
        
    # Asumiendo recurrencia semanal
    return vRecur({'FREQ': ['WEEKLY'], 'BYDAY': byday})

# La firma de retorno incluye la lista de errores (row_errors) para el log file
def create_icalendar(parsed_data: Dict[str, Dict]) -> Tuple[str, int, list]: 
    """Convierte las filas de los DataFrames en una cadena ICS."""
    cal = Calendar()
    cal.add('prodid', '-//Cali XLSX to ICS Parser//EN')
    cal.add('version', '2.0')
    event_count = 0
    row_errors = [] # Lista para capturar errores por fila
    
    # Usar UTC como la zona horaria predeterminada
    DEFAULT_TZ = pytz.utc
    
    for sheet_name, data in parsed_data.items():
        df = data['data']
        template = data['template']
        
        for index, row in df.iterrows():
            try:
                event = Event()
                event_uid = f'{sheet_name}-{index}-{datetime.now().timestamp()}@cali.latai-community.com'
                event.add('uid', event_uid)
                event.add('summary', row['TITLE'])
                event.add('location', str(row.get('LOCATION', '')))
                
                description = f"{str(row.get('DESCRIPTION', ''))}\nLink: {str(row.get('LINK', ''))}"
                event.add('description', description)
                
                owner_email = str(row.get("OWNER", "")).strip()
                if owner_email and '@' in owner_email:
                    event.add('organizer', f'mailto:{owner_email}')
                
                # --- Manejo de Fechas/Tiempos ---
                if template == 'Basic':
                    start_dt = pd.to_datetime(row['START_TIME']).to_pydatetime()
                    end_dt = pd.to_datetime(row['END_TIME']).to_pydatetime()
                    
                    event.add('dtstart', start_dt.replace(tzinfo=DEFAULT_TZ))
                    event.add('dtend', end_dt.replace(tzinfo=DEFAULT_TZ))
                
                elif template == 'Medium':
                    start_date: date = pd.to_datetime(row['START_DATE']).date()
                    end_date: date = pd.to_datetime(row['END_DATE']).date()
                    start_time: time = pd.to_datetime(row['START_TIME']).time()
                    end_time: time = pd.to_datetime(row['END_TIME']).time()

                    # Combinar fecha y hora
                    start_dt = datetime.combine(start_date, start_time).replace(tzinfo=DEFAULT_TZ)
                    event.add('dtstart', start_dt)
                    
                    end_dt_first_instance = datetime.combine(start_date, end_time).replace(tzinfo=DEFAULT_TZ)
                    
                    # Si la hora de fin es <= hora de inicio, cruza la medianoche
                    if end_time <= start_time:
                         end_dt_first_instance += timedelta(days=1)

                    event.add('dtend', end_dt_first_instance)
                    
                    # Añadir Recurrencia (RRULE)
                    frequency_val = row.get('FREQUENCY')
                    
                    # CRÍTICO: Manejar NaN/None explícitamente y convertir a string limpio
                    if pd.isna(frequency_val):
                         frequency_val = ''
                         
                    frequency_str = str(frequency_val).strip()

                    if frequency_str:
                        rrule = _create_rrule_for_medium(frequency_str)
                        
                        # El final de la recurrencia es la fecha de FIN (End Date) a las 23:59:59
                        end_until_dt = datetime.combine(end_date, time(23, 59, 59)).replace(tzinfo=DEFAULT_TZ)
                        rrule['UNTIL'] = end_until_dt
                        
                        event.add('rrule', rrule)

                # --- Participantes (Attendees) ---
                participants_val = str(row.get('PARTICIPANTS', '')).strip()
                if participants_val:
                    for email in participants_val.split(','):
                        email = email.strip()
                        if '@' in email:
                             event.add('attendee', f'mailto:{email}')

                cal.add_component(event)
                event_count += 1

            except Exception as e:
                # Log the error, incluyendo el valor que falló para depuración
                error_message = f"‼️ Error processing event in sheet '{sheet_name}', row {index}: {e}. Skipping row. FREQUENCY value was: {str(row.get('FREQUENCY', ''))}"
                print(error_message) 
                row_errors.append(error_message) 

    # FINAL FIX: Retornar la lista de errores
    return cal.to_ical().decode('utf-8'), event_count, row_errors