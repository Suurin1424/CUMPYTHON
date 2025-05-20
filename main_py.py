#!/usr/bin/env python3
# main.py - Główny skrypt uruchamiający PDF Email Extractor

import os
import sys
import argparse
import json
from datetime import datetime
import logging
from typing import List, Dict, Any

# Importuj nasze moduły
try:
    from pdf_email_extractor import DataExtractor
    from config import *
except ImportError:
    print("Błąd: Nie można zaimportować modułów. Uruchom setup.py")
    sys.exit(1)

# Konfiguracja logowania
def setup_logging():
    """Konfiguruje system logowania"""
    log_dir = os.path.dirname(LOGGING_CONFIG['log_file'])
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    logging.basicConfig(
        level=getattr(logging, LOGGING_CONFIG['level']),
        format=LOGGING_CONFIG['log_format'],
        handlers=[
            logging.FileHandler(LOGGING_CONFIG['log_file']),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

class ExtractorManager:
    """Główna klasa zarządzająca ekstraktorem"""
    
    def __init__(self):
        self.logger = setup_logging()
        self.extractor = DataExtractor()
        self.stats = {
            'processed_files': 0,
            'successful_extractions': 0,
            'failed_extractions': 0,
            'start_time': datetime.now()
        }
    
    def process_single_pdf(self, file_path: str) -> Dict[str, Any]:
        """Przetwarza pojedynczy plik PDF"""
        self.logger.info(f"Przetwarzanie pliku: {file_path}")
        
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Plik nie istnieje: {file_path}")
            
            # Sprawdź rozmiar pliku
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > PERFORMANCE_SETTINGS['max_pdf_size_mb']:
                raise ValueError(f"Plik zbyt duży: {file_size_mb:.1f}MB")
            
            extracted_data = self.extractor.extract_from_pdf(file_path)
            extracted_data['source_file'] = os.path.basename(file_path)
            extracted_data['file_size_mb'] = file_size_mb
            
            # Walidacja danych
            validation_result = self.validate_extracted_data(extracted_data)
            extracted_data['validation'] = validation_result
            
            self.stats['successful_extractions'] += 1
            return extracted_data
            
        except Exception as e:
            self.logger.error(f"Błąd przetwarzania {file_path}: {str(e)}")
            self.stats['failed_extractions'] += 1
            return {
                'source_file': os.path.basename(file_path),
                'error': str(e),
                'validation': {'is_valid': False, 'errors': [str(e)]}
            }
        finally:
            self.stats['processed_files'] += 1
    
    def process_directory(self, directory_path: str) -> List[Dict[str, Any]]:
        """Przetwarza wszystkie pliki PDF w katalogu"""
        self.logger.info(f"Przetwarzanie katalogu: {directory_path}")
        
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
            self.logger.warning(f"Utworzono katalog: {directory_path}")
            return []
        
        pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            self.logger.warning("Nie znaleziono plików PDF w katalogu")
            return []
        
        self.logger.info(f"Znaleziono {len(pdf_files)} plików PDF")
        
        all_data = []
        for filename in pdf_files:
            file_path = os.path.join(directory_path, filename)
            extracted_data = self.process_single_pdf(file_path)
            all_data.append(extracted_data)
        
        return all_data
    
    def validate_extracted_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Waliduje wyekstraktowane dane"""
        validation_result = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'missing_required_fields': []
        }
        
        # Sprawdź wymagane pola
        for field in FIELD_PRIORITIES['required_fields']:
            if not data.get(field):
                validation_result['missing_required_fields'].append(field)
                validation_result['errors'].append(f"Brak wymaganego pola: {field}")
        
        # Walidacja formatu
        if data.get('shipping_postcode'):
            if not self._validate_postcode(data['shipping_postcode']):
                validation_result['errors'].append("Nieprawidłowy format kodu pocztowego")
        
        if data.get('material_code'):
            if not self._validate_ppg_code(data['material_code']):
                validation_result['warnings'].append("Kod materiału nie pasuje do wzorca PPG")
        
        # Sprawdź obecność kluczowych słów
        if data.get('material_description'):
            has_keywords = any(keyword in data['material_description'] 
                             for keyword in VALIDATION_RULES['required_material_keywords'])
            if not has_keywords:
                validation_result['warnings'].append("Brak kluczowych słów w opisie materiału")
        
        if validation_result['errors']:
            validation_result['is_valid'] = False
        
        return validation_result
    
    def _validate_postcode(self, postcode: str) -> bool:
        """Waliduje format kodu pocztowego"""
        import re
        return bool(re.match(VALIDATION_RULES['postcode_format'], postcode.strip()))
    
    def _validate_ppg_code(self, code: str) -> bool:
        """Waliduje format kodu PPG"""
        import re
        return bool(re.match(VALIDATION_RULES['ppg_code_format'], code.strip()))
    
    def process_emails(self, email_config: Dict[str, str]) -> List[Dict[str, Any]]:
        """Przetwarza emaile z serwera IMAP"""
        self.logger.info("Rozpoczynam przetwarzanie emaili")
        
        try:
            emails_data = self.extractor.read_emails_from_imap(
                server=email_config['server'],
                username=email_config['username'],
                password=email_config['password'],
                folder=email_config.get('folder', 'INBOX')
            )
            
            # Walidacja danych z emaili
            for email_data in emails_data:
                validation_result = self.validate_extracted_data(email_data)
                email_data['validation'] = validation_result
            
            self.logger.info(f"Przetworzono {len(emails_data)} emaili")
            return emails_data
            
        except Exception as e:
            self.logger.error(f"Błąd przetwarzania emaili: {str(e)}")
            return []
    
    def export_to_excel(self, data: List[Dict[str, Any]], output_file: str) -> bool:
        """Eksportuje dane do Excel z zaawansowanym formatowaniem"""
        try:
            # Upewnij się, że katalog wyjściowy istnieje
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Zapisz dane
            success = self.extractor.save_to_excel(data, output_file)
            
            if success:
                # Dodaj arkusz statystyk
                self.add_statistics_sheet(output_file, data)
                self.logger.info(f"Dane wyeksportowane do: {output_file}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Błąd eksportu do Excel: {str(e)}")
            return False
    
    def add_statistics_sheet(self, excel_file: str, data: List[Dict[str, Any]]):
        """Dodaje arkusz ze statystykami"""
        try:
            import pandas as pd
            from openpyxl import load_workbook
            
            # Przygotuj statystyki
            total_records = len(data)
            valid_records = sum(1 for d in data if d.get('validation', {}).get('is_valid', True))
            invalid_records = total_records - valid_records
            
            processing_time = datetime.now() - self.stats['start_time']
            
            stats_data = {
                'Statystyka': [
                    'Łączna liczba rekordów',
                    'Poprawne rekordy',
                    'Niepoprawne rekordy',
                    'Procent sukcesu',
                    'Czas przetwarzania',
                    'Data generowania raportu'
                ],
                'Wartość': [
                    total_records,
                    valid_records,
                    invalid_records,
                    f"{(valid_records/total_records*100):.1f}%" if total_records > 0 else "0%",
                    str(processing_time).split('.')[0],
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            
            # Dodaj arkusz statystyk
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
                pd.DataFrame(stats_data).to_excel(writer, sheet_name='Statystyki', index=False)
            
        except Exception as e:
            self.logger.error(f"Błąd dodawania arkusza statystyk: {str(e)}")
    
    def export_to_json(self, data: List[Dict[str, Any]], output_file: str) -> bool:
        """Eksportuje dane do JSON"""
        try:
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2, default=str)
            
            self.logger.info(f"Dane wyeksportowane do JSON: {output_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"Błąd eksportu do JSON: {str(e)}")
            return False
    
    def generate_report(self, data: List[Dict[str, Any]]) -> str:
        """Generuje raport z przetwarzania"""
        total_records = len(data)
        valid_records = sum(1 for d in data if d.get('validation', {}).get('is_valid', True))
        
        report = f"""
=== RAPORT PRZETWARZANIA ===
Data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Łączna liczba rekordów: {total_records}
Poprawne rekordy: {valid_records}
Niepoprawne rekordy: {total_records - valid_records}
Procent sukcesu: {(valid_records/total_records*100):.1f}%

=== STATYSTYKI PLIKÓW ===
Przetworzone pliki: {self.stats['processed_files']}
Udane ekstrakcje: {self.stats['successful_extractions']}
Nieudane ekstrakcje: {self.stats['failed_extractions']}

=== BRAKUJĄCE DANE ===
"""
        
        # Analiza brakujących pól
        missing_fields = {}
        for record in data:
            if record.get('validation'):
                for field in record['validation'].get('missing_required_fields', []):
                    missing_fields[field] = missing_fields.get(field, 0) + 1
        
        if missing_fields:
            for field, count in sorted(missing_fields.items(), key=lambda x: x[1], reverse=True):
                report += f"{field}: {count} razy\n"
        else:
            report += "Brak brakujących wymaganych pól.\n"
        
        return report

def parse_arguments():
    """Parsuje argumenty wiersza poleceń"""
    parser = argparse.ArgumentParser(description='PDF Email Data Extractor')
    parser.add_argument('--mode', choices=['pdf', 'email', 'single', 'interactive'], 
                       default='interactive', help='Tryb przetwarzania')
    parser.add_argument('--input', help='Ścieżka do pliku lub katalogu wejściowego')
    parser.add_argument('--output', help='Ścieżka do pliku wyjściowego')
    parser.add_argument('--format', choices=['excel', 'json', 'both'], 
                       default='excel', help='Format wyjściowy')
    parser.add_argument('--config', help='Ścieżka do pliku konfiguracyjnego')
    parser.add_argument('--verbose', '-v', action='store_true', help='Szczegółowe logowanie')
    
    return parser.parse_args()

def load_email_credentials():
    """Ładuje dane dostępowe do emaila z pliku .env lub z inputu"""
    email_config = {}
    
    # Spróbuj załadować z pliku .env
    if os.path.exists('.env'):
        with open('.env', 'r') as f:
            for line in f:
                if '=' in line and not line.startswith('#'):
                    key, value = line.strip().split('=', 1)
                    if key.startswith('EMAIL_'):
                        email_config[key.lower().replace('email_', '')] = value
    
    # Jeśli brak danych, poproś użytkownika
    if not email_config.get('server'):
        print("\n=== Konfiguracja emaila ===")
        email_config['server'] = input("Serwer IMAP (np. imap.gmail.com): ")
        email_config['username'] = input("Email: ")
        email_config['password'] = input("Hasło: ")
        email_config['folder'] = input("Folder (domyślnie INBOX): ") or 'INBOX'
    
    return email_config

def interactive_mode():
    """Tryb interaktywny"""
    manager = ExtractorManager()
    
    while True:
        print("\n=== PDF Email Data Extractor ===")
        print("1. Przetwarzaj pliki PDF z katalogu")
        print("2. Przetwarzaj pojedynczy plik PDF")
        print("3. Przetwarzaj emaile")
        print("4. Trenuj model ML")
        print("5. Pokaż statystyki")
        print("6. Wyjście")
        
        choice = input("\nWybierz opcję (1-6): ").strip()
        
        if choice == '1':
            input_dir = input(f"Katalog wejściowy (domyślnie {INPUT_PDF_DIR}): ") or INPUT_PDF_DIR
            data = manager.process_directory(input_dir)
            
            if data:
                output_file = input(f"Plik wyjściowy (domyślnie {OUTPUT_EXCEL_FILE}): ") or OUTPUT_EXCEL_FILE
                output_path = os.path.join(OUTPUT_DIR, output_file)
                
                if manager.export_to_excel(data, output_path):
                    print(f"\n✓ Przetwarzanie zakończone. Wyniki w: {output_path}")
                    print(manager.generate_report(data))
        
        elif choice == '2':
            file_path = input("Ścieżka do pliku PDF: ").strip()
            if file_path:
                data = [manager.process_single_pdf(file_path)]
                output_path = os.path.join(OUTPUT_DIR, f"single_file_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                
                if manager.export_to_excel(data, output_path):
                    print(f"\n✓ Plik przetworzony. Wyniki w: {output_path}")
        
        elif choice == '3':
            email_config = load_email_credentials()
            data = manager.process_emails(email_config)
            
            if data:
                output_path = os.path.join(OUTPUT_DIR, f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                if manager.export_to_excel(data, output_path):
                    print(f"\n✓ Emaile przetworzone. Wyniki w: {output_path}")
        
        elif choice == '4':
            manager.extractor.train_model()
            print("\n✓ Model ML został wytrenowany")
        
        elif choice == '5':
            print(f"\n=== STATYSTYKI ===")
            print(f"Przetworzone pliki: {manager.stats['processed_files']}")
            print(f"Udane ekstrakcje: {manager.stats['successful_extractions']}")
            print(f"Nieudane ekstrakcje: {manager.stats['failed_extractions']}")
            print(f"Czas działania: {datetime.now() - manager.stats['start_time']}")
        
        elif choice == '6':
            print("Do widzenia!")
            break
        
        else:
            print("Nieprawidłowy wybór!")

def main():
    """Główna funkcja programu"""
    args = parse_arguments()
    
    # Ustawienia logowania
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    if args.mode == 'interactive':
        interactive_mode()
    else:
        manager = ExtractorManager()
        data = []
        
        if args.mode == 'pdf':
            if args.input:
                if os.path.isdir(args.input):
                    data = manager.process_directory(args.input)
                else:
                    data = [manager.process_single_pdf(args.input)]
            else:
                data = manager.process_directory(INPUT_PDF_DIR)
        
        elif args.mode == 'email':
            email_config = load_email_credentials()
            data = manager.process_emails(email_config)
        
        elif args.mode == 'single':
            if not args.input:
                print("Błąd: Wymagana ścieżka do pliku (--input)")
                sys.exit(1)
            data = [manager.process_single_pdf(args.input)]
        
        # Eksport wyników
        if data:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if args.format in ['excel', 'both']:
                excel_file = args.output or os.path.join(OUTPUT_DIR, f"results_{timestamp}.xlsx")
                manager.export_to_excel(data, excel_file)
            
            if args.format in ['json', 'both']:
                json_file = args.output or os.path.join(OUTPUT_DIR, f"results_{timestamp}.json")
                manager.export_to_json(data, json_file)
            
            # Wyświetl raport
            print(manager.generate_report(data))
        else:
            print("Brak danych do przetworzenia")

if __name__ == "__main__":
    main()
