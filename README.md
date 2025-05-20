# CUMPYTHON
# PDF i Email Data Extractor

Narzędzie do automatycznej ekstrakcji danych z plików PDF i wiadomości email do arkuszy Excel, z funkcją klasyfikacji i analizy zawartości przy pomocy ML.

## 📋 Opis

PDF i Email Data Extractor to zaawansowane narzędzie stworzone do automatyzacji procesu wydobywania istotnych danych biznesowych (takich jak dane klientów, numery zamówień, kody materiałów, itp.) z plików PDF oraz wiadomości email. Program wykrywa i wyodrębnia dane zgodnie z określonymi wzorcami, a następnie zapisuje je w ustrukturyzowanym formacie Excel lub JSON.

Narzędzie jest szczególnie przydatne dla firm handlujących materiałami budowlanymi, farbami i produktami PPG, gdzie wiele zamówień i dokumentów przychodzi w formie plików PDF lub wiadomości email od klientów z Holandii i Polski.

## ✨ Funkcje

- **Ekstrakcja danych z PDF** - wydobywanie danych z plików PDF, takich jak zamówienia, oferty, dokumenty dostawy
- **Przetwarzanie wiadomości email** - łączenie z serwerem IMAP i analiza treści emaili
- **Machine Learning** - klasyfikacja rodzaju dokumentu oraz kontekstu wiadomości
- **Walidacja danych** - sprawdzanie kompletności i poprawności wydobytych informacji
- **Eksport** - zapisywanie danych w formatach Excel i JSON z zaawansowanym formatowaniem
- **Tryb wsadowy** - możliwość przetwarzania całych katalogów plików PDF
- **Raportowanie** - generowanie raportów statystycznych z procesu ekstrakcji
- **Interfejs wiersza poleceń** - obsługa przez interfejs tekstowy oraz argumenty wiersza poleceń

## 🔧 Instalacja

### Wymagania systemowe

- Python 3.8 lub nowszy
- Biblioteki wymienione w pliku `requirements.txt`

### Kroki instalacji

1. Sklonuj lub pobierz repozytorium:
   ```
   git clone https://github.com/twojafirma/pdf-email-extractor.git
   cd pdf-email-extractor
   ```

2. Stwórz i aktywuj wirtualne środowisko Python (opcjonalnie, ale zalecane):
   ```
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Linux/macOS
   source venv/bin/activate
   ```

3. Zainstaluj wymagane zależności:
   ```
   pip install -r requirements.txt
   ```

4. Pierwszy uruchomienie i utworzenie katalogów:
   ```
   python main.py
   ```

## 🚀 Użytkowanie

### Tryb interaktywny

Najłatwiejszy sposób na rozpoczęcie pracy z narzędziem:

```
python main.py
```

Program wyświetli menu z dostępnymi opcjami:
1. Przetwarzaj pliki PDF z katalogu
2. Przetwarzaj pojedynczy plik PDF
3. Przetwarzaj emaile
4. Trenuj model ML
5. Pokaż statystyki
6. Wyjście

### Tryb wiersza poleceń

Dla zautomatyzowanych operacji, możesz wykorzystać argumenty wiersza poleceń:

```
python main.py --mode pdf --input ./dokumenty --output wyniki.xlsx
```

#### Dostępne opcje:

- `--mode` - tryb pracy: `pdf`, `email`, `single`, `interactive` (domyślnie: `interactive`)
- `--input` - ścieżka do pliku lub katalogu wejściowego
- `--output` - ścieżka do pliku wyjściowego
- `--format` - format wyjściowy: `excel`, `json`, `both` (domyślnie: `excel`)
- `--config` - ścieżka do pliku konfiguracyjnego (opcjonalnie)
- `--verbose` lub `-v` - szczegółowe logowanie

### Przykłady użycia

#### Przetwarzanie wszystkich plików PDF z katalogu:

```
python main.py --mode pdf --input ./zamowienia --output wynik.xlsx
```

#### Przetwarzanie emaili:

```
python main.py --mode email --output emaile.xlsx
```

#### Przetwarzanie pojedynczego pliku PDF:

```
python main.py --mode single --input ./dokument.pdf --format both
```

## ⚙️ Konfiguracja

Ustawienia programu znajdują się w pliku `config.py`. Możesz dostosować:

- Ścieżki katalogów wejściowych/wyjściowych
- Ustawienia serwera email
- Wzorce wyrażeń regularnych dla ekstrakcji danych
- Konfigurację machine learning
- Priorytetyzację pól danych
- Style formatowania Excel
- Ustawienia logowania
- Reguły walidacji danych
- Parametry wydajnościowe

### Konfiguracja emaili

Dla bezpieczeństwa, dane logowania do serwera email powinny być przechowywane w pliku `.env` w katalogu głównym aplikacji:

```
EMAIL_SERVER=imap.gmail.com
EMAIL_USERNAME=twoj_email@gmail.com
EMAIL_PASSWORD=twoje_haslo_aplikacji
EMAIL_FOLDER=INBOX
```

**Uwaga:** Dla kont Gmail należy użyć hasła aplikacji, a nie głównego hasła konta.

## 📄 Wymagane pola danych

Pola, które program stara się wyodrębnić:

### Pola wymagane (zielone)
- **customer_name** - nazwa klienta
- **po_number** - numer zamówienia
- **material_code** - kod materiału (np. PPG123456)
- **material_description** - opis materiału
- **shipping_street** - ulica dostawy
- **shipping_postcode** - kod pocztowy (format NL: 1234 AB)
- **colour_code** - kod koloru (np. RAL1234)
- **fan_code** - kod wachlarza
- **shipping_condition** - warunki dostawy

### Pola dodatkowe
- **project_number** - numer projektu
- **date** - data
- **reference_number** - numer referencyjny
- **order_items** - pozycje zamówienia

## 🔍 Działanie mechanizmu ekstrakcji

1. Program otwiera plik PDF lub email
2. Wyodrębnia tekst
3. Stosuje wyrażenia regularne (regex) do znalezienia odpowiednich pól danych
4. W przypadku emaili, używa ML do określenia typu wiadomości
5. Waliduje znalezione dane
6. Zapisuje wyniki do Excel/JSON
7. Generuje raport z przetwarzania

## 🧠 Model ML

Program wykorzystuje prosty klasyfikator Naive Bayes do kategoryzacji emaili:
- **order** - zamówienie
- **quote** - oferta
- **delivery** - dostawa
- **pickup** - odbiór
- **complaint** - reklamacja
- **invoice** - faktura
- **inquiry** - zapytanie
- **other** - inne

Model można dotrenować wybierając opcję "Trenuj model ML" w menu głównym.

## 📝 Logi

Logi działania programu są zapisywane w katalogu `logs/extractor.log` i zawierają informacje o:
- Przetworzonych plikach
- Wykrytych danych
- Napotkanych błędach 
- Ostrzeżeniach i brakujących danych

## 🔧 Rozwiązywanie problemów

### Brak wykrywanych danych
- Sprawdź czy format PDF nie jest zeskanowanym obrazem (wymaga OCR)
- Zweryfikuj wzorce regex w `config.py`
- Sprawdź logi błędów

### Problemy z czytaniem emaili
- Upewnij się, że dane logowania są poprawne
- Dla Gmail włącz "Dostęp mniej bezpiecznych aplikacji" lub użyj hasła aplikacji
- Sprawdź ustawienia serwera IMAP

### Niewystarczająca dokładność ML
- Dotrenuj model na większej liczbie przykładów
- Dostosuj `confidence_threshold` w ustawieniach ML

## 🤝 Wsparcie

W przypadku pytań lub problemów, utwórz Issue w repozytorium projektu lub skontaktuj się z autorem:

- Email: support@twojafirma.com
- Tel: +48 123 456 789

## 📜 Licencja

Ten projekt jest licencjonowany na podstawie licencji MIT - szczegóły w pliku LICENSE.
