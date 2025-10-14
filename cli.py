# cli.py
import sys
from datetime import datetime
import os 
# CRITICAL FIX: Ensure these imports match your filenames
from parser import load_and_parse_excel 
from converter import create_icalendar 

# --- Función Auxiliar para Escribir el Log ---
def write_log_file(log_summary: list, log_filepath: str):
    """Escribe el resumen del proceso en un archivo."""
    try:
        with open(log_filepath, 'w', encoding='utf-8') as f:
            f.write("CALI PROCESS LOG\n")
            f.write("="*20 + "\n\n")
            f.write("\n".join(log_summary))
        print(f"\n✅ Log file successfully saved to: {log_filepath}")
    except Exception as e:
        print(f"\n‼️ CRITICAL ERROR: Could not write log file to {log_filepath}. Reason: {e}")

# --- Función Principal Modificada ---
def main_cali_app(filepath: str, output_file: str):
    """Función principal para ejecutar la aplicación Cali."""
    
    base_name = os.path.splitext(output_file)[0]
    log_file = f"{base_name}.log"
    
    print(f"\n--- CALI - XLSX to ICS Parser ---")
    start_time = datetime.now()
    log_summary = [] 

    def log_and_print(message: str, force_print=False):
        """Añade un mensaje a la lista de log y lo imprime en la consola si es necesario."""
        log_summary.append(message)
        if force_print:
            print(message)
    
    try:
        # 1. Load, Map, and Parse Excel
        log_and_print(f"1. Loading and processing '{filepath}'...")
        
        # load_and_parse_excel imprime directamente logs de éxito/fallo de hoja
        parsed_data = load_and_parse_excel(filepath)
        
        if not parsed_data:
            raise Exception("No valid sheets found to process. Please check your Excel headers.")
        
        # 2. Convert DataFrames to ICS
        log_and_print("2. Generating iCalendar content...")
        
        # EXPECT 3 RETURN VALUES: ics content, event count, and row errors
        ics_content, event_count, row_errors = create_icalendar(parsed_data) 
        
        # 3. Save the ICS File
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(ics_content)
        
        end_time = datetime.now()
        
        # 4. Generar Resumen Final
        
        log_summary.append("\n" + "="*40)
        log_summary.append("--- PROCESS COMPLETED ---")
        
        if row_errors:
            log_summary.append(f"Status: SUCCESS with {len(row_errors)} DATA WARNINGS")
        else:
            log_summary.append(f"Status: SUCCESS")
            
        log_summary.append("="*40)
        
        # AÑADIR WARNINGS ESPECÍFICOS AL LOG
        if row_errors:
            log_summary.append("\n--- ROW-LEVEL DATA WARNINGS (SKIPPED EVENTS) ---")
            log_summary.extend(row_errors)
            log_summary.append("-------------------------------------------------\n")

        log_summary.append(f"Input File: {filepath}")
        log_summary.append(f"Output ICS File: {output_file}")
        log_summary.append(f"Total Events Generated: {event_count}")
        log_summary.append(f"Time Taken: {end_time - start_time}")
        log_summary.append("\nProcessed Sheets:")
        
        for sheet_name, data in parsed_data.items():
            log_summary.append(f"- {sheet_name} (Template: {data['template']}, Rows Processed: {len(data['data'])})")
        
        # Imprimir el resumen final en consola
        print("\n" + "="*40)
        if row_errors:
            print(f"--- PROCESS COMPLETED WITH {len(row_errors)} WARNINGS ---")
            print(f"Check {log_file} for details on skipped rows.")
        else:
            print("--- PROCESS COMPLETED SUCCESSFULLY ---")
            
        print(f"Total Events: {event_count}")
        print(f"ICS file: {output_file}")
        print("="*40)
        
    except Exception as e:
        end_time = datetime.now()
        log_summary.append("\n" + "="*40)
        log_summary.append("--- PROCESS FAILED ---")
        log_summary.append(f"Error: {type(e).__name__}: {str(e)}")
        log_summary.append(f"Time Taken: {end_time - start_time}")
        log_summary.append("="*40)
        
        # Imprimir el error principal en consola
        print("\n" + "="*40)
        print("--- PROCESS FAILED ---")
        print(f"Error: {type(e).__name__}: {str(e)}")
        print("="*40)
        
    finally:
        write_log_file(log_summary, log_file)


# --- Bloque de Ejecución (Main) ---
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Error: Invalid number of arguments.")
        print("Usage: python cli.py <input_file.xlsx> <output_file.ics>")
        sys.exit(1)
        
    input_file = sys.argv[1]
    output_file = "output/"+ sys.argv[2]
    
    main_cali_app(input_file, output_file)