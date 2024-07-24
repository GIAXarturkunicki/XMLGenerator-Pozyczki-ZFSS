# XML Generator Pożyczek - ZFŚS

## Dokumentacja techniczna

### Przegląd
Ta aplikacja generuje pliki XML dla wniosków pożyczkowych w systemie ZFSS, korzystając z danych z plików Excel. Nazwy kolumn musza sie zgadzac. W plikach *_template zostały wpisane dane testowe umożliwiajace migracje.


### Wymagane kolumny w plikach Excel

#### `kodguid_template.xlsx`
- **Kod**: Kod pracownika
- **Guid**: Guid Pracownika
 
#### `source_template.xlsx`

##### Obowiązkowe
- **Pracownik**: Kod pracownika.(Taki jak w egeri program sam poprawi do kodu z enovy)
- **Imię i Nazwisko**:  imię i nazwisko .
- **Okres**: Okres, w którym obowiązuje umowa.
- **SaldoBO**: Saldo początkowe.
- **Numer umowy/pożyczki**: Numer umowy lub pożyczki.
- **Data**: Data zawarcia umowy.
- **Kwota**: Kwota pożyczki.
- **SplatyOd**: Data rozpoczęcia spłat.
- **IloscRat**: Liczba rat.
- **Typ**: Typ pożyczki.
- **Sposob**: Sposób spłaty.
- **Procent**: Oprocentowanie.
- **Żyrant 1 Imię i Nazwisko**: Imię i nazwisko pierwszego żyranta.
- **Żyrant 2 Imię i Nazwisko**: Imię i nazwisko drugiego żyranta.
- **Żyrant 1 KOD**: Kod pierwszego żyranta. Program uzupelnia to pole sam z pliku kod_guid
- **Żyrant 2 KOD**: Kod drugiego żyranta.    Program uzupelnia to pole sam z pliku kod_guid
- **Definicja**: Definicja umowy.
- **KwotaRaty**: Kwota raty.

