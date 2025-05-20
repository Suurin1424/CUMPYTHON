# PDF i Email Data Extractor
# Autor: Python Script dla ekstrakcji danych z PDF i emaili do Excel

import pandas as pd
import PyPDF2
import pdfplumber
import re
import email
import imaplib
import os
from email.mime.text import MimeText
from email.mime.multipart import MimeMultipart
import smtplib
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import logging
from datetime import datetime
import nltk
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
import pickle

# Konfiguracja logowania
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DataExtractor:
    def __init__(self):
        self.patterns = {
            # Pola wymagane (zielone w specyfikacji)
            'customer_name': r'(?:Klant|Customer|Naam):\s*([^\n\r]+)',
            'po_number': r'(?:Boeknummer|PO|Order):\s*([^\n\r]+)',
            'material_code': r'(?:PPG\d+|Kod\s*materiału):\s*([^\n\r]+)',
            'material_description': r'(?:Sigma.*?(?:\d+\.?\d*\s*Ltr)|Opis.*?(?:\d+\.?\d*\s*Ltr))',
            'shipping_street': r'(?:Adres|Address|Straat):\s*([^\n\r,]+)',
            'shipping_postcode': r'(\d{4}\s*[A-Z]{2})',
            'colour_code': r'(?:Ral\s*\d+|No\.\d+\.?\d*)',
            'fan_code': r'(?:Fan|Waaier):\s*([^\n\r]+)',
            'shipping_condition': r'(?:Verzending|Levering|Shipping):\s*([^\n\r]+)',
            
            # Pola pomocnicze
            'project_number': r'(?:Project|Projectnummer):\s*([^\n\r]+)',
            'date': r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})',
            'quantity': r'(\d+)\s*stuks?',
            'reference_number': r'(?:Ref|Referentie):\s*([^\n\r]+)'
        }
        
        # Model ML do klasyfikacji
        self.ml_model = None
        self.vectorizer = None
        self.load_or_train_model()
    
    def load_or_train_model(self):
        """Ładuje istniejący model ML lub trenuje nowy"""
        try:
            with open('email_classifier.pkl', 'rb') as f:
                self.ml_model = pickle.load(f)
            with open('vectorizer.pkl', 'rb') as f:
                self.vectorizer = pickle.load(f)
            logger.info("Załadowano istniejący model ML")
        except FileNotFoundError:
            logger.info("Trenowanie nowego modelu ML")
            self.train_model()
    
    def train_model(self):
        """Trenuje prosty model do klasyfikacji emaili"""
        # Przykładowe dane treningowe
        training_data = [
            ("bestelling order PPG materiaal", "order"),
            ("aanvraag offerte prijsopgave", "quote"),
            ("levering verzending adres", "delivery"),
            ("factuur betaling rekening", "invoice"),
            ("klacht probleem kwaliteit", "complaint"),
            ("Sigma Rapid Gloss bestelling", "order"),
            ("Kom ik ophalen", "pickup"),
            ("Afleveradres wijziging", "delivery")
        ]
        
        texts, labels = zip(*training_data)
        
        self.vectorizer = TfidfVectorizer(stop_words='english', lowercase=True)
        X = self.vectorizer.fit_transform(texts)
        
        self.ml_model = MultinomialNB()
        self.ml_model.fit(X, labels)
        
        # Zapisz model
        with open('email_classifier.pkl', 'wb') as f:
            pickle.dump(self.ml_model, f)
        with open('vectorizer.pkl', 'wb') as f:
            pickle.dump(self.vectorizer, f)
    
    def classify_email(self, text):
        """Klasyfikuje email za pomocą ML"""
        if self.ml_model and self.vectorizer:
            X = self.vectorizer.transform([text])
            prediction = self.ml_model.predict(X)[0]
            confidence = max(self.ml_model.predict_proba(X)[0])
            return prediction, confidence
        return "unknown", 0.5
    
    def extract_from_pdf(self, pdf_path):
        """Ekstraktuje dane z pliku PDF"""
        extracted_data = {}
        
        try:
            # Próba z pdfplumber (lepsze dla tabel)
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
            
            # Fallback do PyPDF2
            if not text.strip():
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page in pdf_reader.pages:
                        text += page.extract_text() + "\n"
            
            # Ekstrakcja danych z wykorzystaniem regex
            for field, pattern in self.patterns.items():
                matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
                if matches:
                    if field == 'material_description':
                        # Specjalne przetwarzanie dla opisów materiałów
                        descriptions = []
                        for match in matches:
                            descriptions.append(match.strip())
                        extracted_data[field] = "; ".join(descriptions)
                    else:
                        extracted_data[field] = matches[0].strip()
            
            # Specjalne przetwarzanie dla danych z bestelling
            if "BESTELLING" in text:
                extracted_data['document_type'] = 'order'
                
                # Ekstrakcja pozycji zamówienia
                order_items = []
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if 'PPG' in line and ('Sigma' in line or 'Ral' in line):
                        item = {
                            'code': re.search(r'PPG\d+', line).group() if re.search(r'PPG\d+', line) else '',
                            'description': line.strip(),
                            'quantity': '',
                            'color': ''
                        }
                        
                        # Szukaj ilości w kolejnych liniach
                        for j in range(i+1, min(i+3, len(lines))):
                            if re.search(r'\d+\s*stuks?', lines[j]):
                                item['quantity'] = re.search(r'(\d+)', lines[j]).group()
                                break
                        
                        # Szukaj kodu koloru
                        if 'Ral' in line:
                            color_match = re.search(r'Ral\s*\d+', line)
                            if color_match:
                                item['color'] = color_match.group()
                        elif 'No.' in line:
                            color_match = re.search(r'No\.\d+\.?\d*', line)
                            if color_match:
                                item['color'] = color_match.group()
                        
                        order_items.append(item)
                
                extracted_data['order_items'] = order_items
            
            logger.info(f"Pomyślnie wyekstraktowano dane z PDF: {pdf_path}")
            return extracted_data
            
        except Exception as e:
            logger.error(f"Błąd podczas ekstrakcji z PDF {pdf_path}: {str(e)}")
            return {}
    
    def extract_from_email(self, email_content):
        """Ekstraktuje dane z treści emaila"""
        extracted_data = {}
        
        # Klasyfikacja emaila
        email_type, confidence = self.classify_email(email_content)
        extracted_data['email_type'] = email_type
        extracted_data['confidence'] = confidence
        
        # Ekstrakcja danych z wykorzystaniem regex
        for field, pattern in self.patterns.items():
            matches = re.findall(pattern, email_content, re.IGNORECASE | re.MULTILINE)
            if matches:
                extracted_data[field] = matches[0].strip()
        
        # Specjalne przetwarzanie dla różnych typów emaili
        if email_type == 'order':
            # Dodatkowe przetwarzanie dla zamówień
            if 'ophalen' in email_content.lower():
                extracted_data['shipping_condition'] = 'Pickup/Afhalen'
        
        logger.info(f"Wyekstraktowano dane z emaila (typ: {email_type}, pewność: {confidence:.2f})")
        return extracted_data
    
    def read_emails_from_imap(self, server, username, password, folder='INBOX'):
        """Odczytuje emaile z serwera IMAP"""
        emails_data = []
        
        try:
            mail = imaplib.IMAP4_SSL(server)
            mail.login(username, password)
            mail.select(folder)
            
            # Szukaj emaili (możesz dostosować kryteria)
            _, message_numbers = mail.search(None, 'ALL')
            
            for num in message_numbers[0].split():
                _, msg_data = mail.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)
                
                # Wyciągnij treść emaila
                if email_message.is_multipart():
                    content = ""
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            content += part.get_payload(decode=True).decode()
                else:
                    content = email_message.get_payload(decode=True).decode()
                
                # Ekstraktuj dane
                extracted = self.extract_from_email(content)
                extracted['email_subject'] = email_message['Subject']
                extracted['email_from'] = email_message['From']
                extracted['email_date'] = email_message['Date']
                
                emails_data.append(extracted)
            
            mail.close()
            mail.logout()
            
        except Exception as e:
            logger.error(f"Błąd podczas odczytu emaili: {str(e)}")
        
        return emails_data
    
    def save_to_excel(self, data_list, output_file):
        """Zapisuje wyekstraktowane dane do pliku Excel"""
        try:
            # Przygotuj dane dla DataFrame
            rows = []
            for data in data_list:
                row = {
                    # Pola zielone (wymagane)
                    'Customer_Name': data.get('customer_name', ''),
                    'PO_Number': data.get('po_number', ''),
                    'Material_Code': data.get('material_code', ''),
                    'Material_Description': data.get('material_description', ''),
                    'Shipping_Street': data.get('shipping_street', ''),
                    'Shipping_Postcode': data.get('shipping_postcode', ''),
                    'Colour_Code': data.get('colour_code', ''),
                    'Fan_Code': data.get('fan_code', ''),
                    'Shipping_Condition': data.get('shipping_condition', ''),
                    
                    # Pola dodatkowe
                    'Project_Number': data.get('project_number', ''),
                    'Date': data.get('date', ''),
                    'Document_Type': data.get('document_type', ''),
                    'Email_Type': data.get('email_type', ''),
                    'Confidence': data.get('confidence', ''),
                    'Reference_Number': data.get('reference_number', ''),
                    'Order_Items': str(data.get('order_items', '')),
                    'Source_File': data.get('source_file', ''),
                    'Processing_Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                rows.append(row)
            
            # Utwórz DataFrame
            df = pd.DataFrame(rows)
            
            # Zapisz do Excel z formatowaniem
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Extracted_Data', index=False)
                
                # Dodaj formatowanie
                workbook = writer.book
                worksheet = writer.sheets['Extracted_Data']
                
                # Kolor dla nagłówków
                header_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                for cell in worksheet[1]:
                    cell.fill = header_fill
                
                # Automatyczne dopasowanie szerokości kolumn
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            logger.info(f"Dane zapisane do pliku Excel: {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"Błąd podczas zapisywania do Excel: {str(e)}")
            return False
    
    def process_directory(self, directory_path, output_file):
        """Przetwarza wszystkie pliki PDF w katalogu"""
        all_data = []
        
        for filename in os.listdir(directory_path):
            if filename.lower().endswith('.pdf'):
                pdf_path = os.path.join(directory_path, filename)
                extracted_data = self.extract_from_pdf(pdf_path)
                extracted_data['source_file'] = filename
                all_data.append(extracted_data)
        
        if all_data:
            return self.save_to_excel(all_data, output_file)
        else:
            logger.warning("Nie znaleziono danych do przetworzenia")
            return False

# Funkcja główna
def main():
    """Główna funkcja programu"""
    extractor = DataExtractor()
    
    # Konfiguracja - dostosuj do swoich potrzeb
    PDF_DIRECTORY = "pdfs"  # Katalog z plikami PDF
    OUTPUT_FILE = "extracted_data.xlsx"
    
    # Konfiguracja emaili (opcjonalnie)
    EMAIL_CONFIG = {
        'server': 'imap.gmail.com',  # Dostosuj do swojego dostawcy
        'username': 'your_email@gmail.com',
        'password': 'your_password',  # Użyj hasła aplikacji dla Gmail
        'folder': 'INBOX'
    }
    
    print("=== PDF i Email Data Extractor ===")
    print("1. Przetwarzaj pliki PDF z katalogu")
    print("2. Przetwarzaj emaile z serwera IMAP")
    print("3. Przetwarzaj pojedynczy plik PDF")
    print("4. Trenuj model ML")
    
    choice = input("Wybierz opcję (1-4): ")
    
    if choice == '1':
        # Przetwarzanie katalogów PDF
        if not os.path.exists(PDF_DIRECTORY):
            os.makedirs(PDF_DIRECTORY)
            print(f"Utworzono katalog {PDF_DIRECTORY}. Umieść w nim pliki PDF i uruchom ponownie.")
            return
        
        success = extractor.process_directory(PDF_DIRECTORY, OUTPUT_FILE)
        if success:
            print(f"Przetwarzanie zakończone. Wyniki zapisane w {OUTPUT_FILE}")
        else:
            print("Błąd podczas przetwarzania")
    
    elif choice == '2':
        # Przetwarzanie emaili
        print("Podaj dane dostępowe do emaila:")
        EMAIL_CONFIG['server'] = input("Serwer IMAP (np. imap.gmail.com): ") or EMAIL_CONFIG['server']
        EMAIL_CONFIG['username'] = input("Email: ")
        EMAIL_CONFIG['password'] = input("Hasło: ")
        
        emails_data = extractor.read_emails_from_imap(**EMAIL_CONFIG)
        
        if emails_data:
            success = extractor.save_to_excel(emails_data, OUTPUT_FILE)
            if success:
                print(f"Przetwarzanie emaili zakończone. Wyniki zapisane w {OUTPUT_FILE}")
        else:
            print("Nie znaleziono emaili do przetworzenia")
    
    elif choice == '3':
        # Przetwarzanie pojedynczego pliku
        pdf_path = input("Podaj ścieżkę do pliku PDF: ")
        if os.path.exists(pdf_path):
            extracted_data = extractor.extract_from_pdf(pdf_path)
            extracted_data['source_file'] = os.path.basename(pdf_path)
            
            success = extractor.save_to_excel([extracted_data], OUTPUT_FILE)
            if success:
                print(f"Przetwarzanie zakończone. Wyniki zapisane w {OUTPUT_FILE}")
            else:
                print("Błąd podczas przetwarzania")
        else:
            print("Plik nie istnieje")
    
    elif choice == '4':
        # Trenowanie modelu
        extractor.train_model()
        print("Model ML został wytrenowany")
    
    else:
        print("Nieprawidłowy wybór")

if __name__ == "__main__":
    main()
