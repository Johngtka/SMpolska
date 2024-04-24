<!-- @format -->

### Technical Documentation (general)

Ten skrypt Pythona przetwarza i analizuje dane z plików Excela, a następnie
zapisuje wyniki w formacie JSON. Wykorzystuje biblioteki pandas do manipulacji
danymi i json do obsługi formatu JSON.

Kod wczytuje dane z dwóch arkuszy pliku Excela, przetwarza je i porównuje.
Szczególnie skupia się na kolumnie ‘Członkowie Zarządu’, identyfikując numery
KRS, które występują w obu arkuszach.

Dla każdego wspólnego numeru KRS, kod wyszukuje odpowiadające mu wiersze,
wyświetla je na konsoli i dodaje do bufora.

Po przetworzeniu wszystkich numerów KRS, zawartość bufora jest zapisywana do
pliku JSON. Kod jest przejrzysty i dobrze zorganizowany, co ułatwia zrozumienie
jego działania i celu.

W efekcie, ten skrypt umożliwia efektywne przetwarzanie i analizę danych z
plików Excela, a następnie zapisanie wyników w łatwo dostępnym formacie JSON.
Jest to przykład jak można wykorzystać moc Pythona i jego bibliotek do pracy z
danymi.

### Technical Documentation (technical)

Ten skrypt Pythona składa się z kilku kluczowych komponentów:

1. Importowanie bibliotek: Skrypt importuje biblioteki pandas i json, które są
   niezbędne do przetwarzania danych.
2. Inicjalizacja zmiennych: Skrypt inicjalizuje listę buffer do przechowywania
   danych oraz licznik rowCounter do śledzenia liczby przetworzonych wierszy.
3. Wczytywanie danych: Skrypt wczytuje dane z dwóch arkuszy (‘Index’ i
   ‘CzlonkowieZarzadu’) pliku Excela ‘SMpolska.xls’ do dwóch DataFrame’ów pandas
   (indexKRS i boardMembersKRS). Kolumna ‘Członkowie Zarządu’ jest traktowana
   jako ciąg znaków.

4. Znajdowanie wspólnych numerów KRS: Skrypt tworzy przecięcie dwóch zestawów
   numerów KRS z obu DataFrame’ów, tworząc zestaw relationKRS.
5. Iteracja przez numery KRS: Dla każdego numeru KRS w relationKRS, skrypt
   znajduje odpowiadające mu wiersze w DataFrame boardMembersKRS.
6. Przetwarzanie wierszy: Dla każdego znalezionego wiersza, skrypt wyświetla
   numer wiersza i dane wiersza, a następnie dodaje dane wiersza do bufora jako
   słownik.
7. Zapisywanie danych do pliku JSON: Po przetworzeniu wszystkich numerów KRS,
   skrypt zapisuje dane z bufora do pliku ‘bufor.json’ w formacie JSON.

8. Plik jest otwarty w trybie dołączania (‘a’), co oznacza, że nowe dane są
   dodawane na końcu pliku, a nie nadpisują istniejących danych.
9. W efekcie, ten skrypt pozwala na przetworzenie danych z plików Excela,
   porównanie danych z dwóch arkuszy, wydrukowanie wybranych wierszy i zapisanie
   tych wierszy do pliku JSON.
10. Skrypt jest napisany w języku Python i wykorzystuje biblioteki pandas i json
    do manipulacji danymi i obsługi formatu JSON.
