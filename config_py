# config.py - Plik konfiguracyjny dla Data Extractor

# Konfiguracja katalogów
INPUT_PDF_DIR = "input_pdfs"
OUTPUT_DIR = "output"
OUTPUT_EXCEL_FILE = "extracted_data.xlsx"

# Konfiguracja emaili - UWAGA: Nigdy nie commituj prawdziwych danych logowania!
EMAIL_SETTINGS = {
    'imap_server': 'imap.gmail.com',
    'imap_port': 993,
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'use_ssl': True,
    'folders_to_check': ['INBOX', 'Orders', 'Bestellingen']
}

# Wzorce regex specyficzne dla holenderskich/polskich dokumentów
CUSTOM_PATTERNS = {
    # Holenderskie wzorce
    'dutch_postcode': r'(\d{4}\s*[A-Z]{2})',
    'dutch_address': r'(?:Adres|Afleveradres|Straat):\s*([^\n\r]+)',
    'dutch_customer': r'(?:Klant|Bedrijf|Naam):\s*([^\n\r]+)',
    
    # PPG specyficzne wzorce
    'ppg_code': r'(PPG\d{6})',
    'sigma_product': r'(Sigma\s+\w+\s+\w+\s+[\d.]+\s*Ltr)',
    'ral_color': r'(Ral\s*\d{4})',
    'order_number': r'(?:BS|Order|Bestelling)(\d{6})',
    
    # Ilości i jednostki
    'quantity_pieces': r'(\d+)\s*stuks?',
    'quantity_liters': r'([\d.]+)\s*Ltr',
    
    # Daty
    'dutch_date': r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})',
    
    # Kontakt i pickup
    'pickup_indication': r'(?:Kom\s+ik\s+(?:zo\s+)?ophalen|Afhalen|Pickup)',
    'contact_person': r'(?:Contactpersoon|Contact):\s*([^\n\r]+)'
}

# Konfiguracja Machine Learning
ML_SETTINGS = {
    'model_path': 'models/email_classifier.pkl',
    'vectorizer_path': 'models/vectorizer.pkl',
    'confidence_threshold': 0.7,
    'retrain_on_startup': False,
    
    # Kategorie klasyfikacji
    'email_categories': [
        'order',       # zamówienie
        'quote',       # oferta
        'delivery',    # dostawa
        'pickup',      # odbiór
        'complaint',   # reklamacja
        'invoice',     # faktura
        'inquiry',     # zapytanie
        'other'        # inne
    ]
}

# Mapowanie pól do priorytetów (na podstawie Twojej tabeli)
FIELD_PRIORITIES = {
    # Zielone pola - muszą być w emailu/PDF
    'required_fields': [
        'customer_name',
        'po_number',
        'material_code',
        'material_description',
        'shipping_street',
        'shipping_postcode',
        'colour_code',
        'fan_code',
        'shipping_condition'
    ],
    
    # Żółte pola - pobierane z emaila
    'email_filled_fields': [
        'po_number',
        'project_number',
        'sold_to_party_code'
    ],
    
    # Niebieskie pola - pobierane z systemu
    'system_fields': [
        'customer_name',
        'sales_office',
        'sales_group'
    ],
    
    # Szare pola - pomijane
    'skip_fields': [
        'sales_office_auto',
        'sales_group_auto'
    ]
}

# Ustawienia stylu Excel
EXCEL_STYLING = {
    'header_color': '90EE90',  # jasno zielony
    'required_field_color': 'C6EFCE',  # jasno zielony
    'email_field_color': 'FFFF99',     # żółty
    'system_field_color': 'ADD8E6',    # jasno niebieski
    'auto_adjust_columns': True,
    'max_column_width': 50,
    'add_filters': True,
    'freeze_header': True
}

# Logging configuration
LOGGING_CONFIG = {
    'level': 'INFO',  # DEBUG, INFO, WARNING, ERROR
    'log_to_file': True,
    'log_file': 'logs/extractor.log',
    'log_format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'max_log_size': 10485760,  # 10MB
    'backup_count': 5
}

# Walidacja danych
VALIDATION_RULES = {
    'postcode_format': r'^\d{4}\s*[A-Z]{2}$',
    'ppg_code_format': r'^PPG\d{6}$',
    'email_format': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
    'date_formats': ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'],
    
    # Minimalne wymagania
    'min_customer_name_length': 2,
    'min_order_number_length': 4,
    'required_material_keywords': ['Sigma', 'PPG']
}

# Ustawienia wydajności
PERFORMANCE_SETTINGS = {
    'max_pdf_size_mb': 50,
    'max_email_size_mb': 10,
    'batch_processing_size': 100,
    'parallel_processing': True,
    'max_workers': 4,
    'timeout_seconds': 300
}
