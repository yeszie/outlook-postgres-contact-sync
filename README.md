# Synchronizacja kontaktów Outlook z bazą danych PostgreSQL

## Opis programu
Program synchronizuje kontakty z programu Outlook z bazą danych PostgreSQL oraz odwrotnie. Dodatkowo, obsługuje listę zablokowanych adresów e-mail (blacklist), co pozwala na automatyczne usuwanie niepożądanych kontaktów zarówno z Outlooka, jak i z bazy danych. Program może także aktualizować notatki w kontaktach Outlooka, dodając informację o liczbie dni od ostatniego kontaktu.

## Funkcje programu

### 1. `export_outlook_to_postgres()`
Eksportuje kontakty z programu Outlook do bazy danych PostgreSQL. Pobiera podstawowe informacje, takie jak:
- Imię i nazwisko,
- Adresy e-mail (główne, drugie i trzecie),
- Numery telefonów (praca, komórkowy, domowy, faks),
- Ostatni kontakt (aktualizowany przy każdym eksporcie).

Kontakty są zapisywane do tabeli `contacts` w PostgreSQL, a w przypadku duplikatów następuje aktualizacja istniejących rekordów.

### 2. `sync_postgres_to_outlook()`
Synchronizuje kontakty z bazy danych PostgreSQL do Outlooka. Jeśli dany kontakt istnieje już w Outlooku, jest on aktualizowany, a jeśli nie – dodawany jako nowy. Obsługiwane są również pola takie jak:
- Główne i dodatkowe adresy e-mail,
- Numery telefonów (praca, komórkowy, domowy, faks).

Podczas synchronizacji zapisywane są zmiany w tabeli `contact_history`, aby śledzić, które kontakty były aktualizowane lub dodawane.

### 3. `remove_blacklisted_contacts()`
Usuwa kontakty, których adresy e-mail, domeny lub prefiksy znajdują się na czarnej liście (blacklist). Funkcja ta usuwa zablokowane kontakty zarówno z Outlooka, jak i z bazy danych PostgreSQL. Po usunięciu kontaktu zmiany są zapisywane w `contact_history`.

### 4. `check_recent_emails()`
Sprawdza skrzynkę odbiorczą Outlooka (oraz podfoldery) w poszukiwaniu kontaktów, z którymi była prowadzona korespondencja w ciągu ostatnich 6 miesięcy. Kontakty te są dodawane do bazy danych PostgreSQL, chyba że ich adresy są na czarnej liście.

### 5. `save_recent_contacts_to_db(recent_contacts)`
Zapisuje ostatnie kontakty z wiadomości e-mail do bazy danych PostgreSQL. Przed dodaniem kontaktów do bazy sprawdzana jest czarna lista, aby uniknąć zapisywania niepożądanych adresów.

### 6. `update_contact_notes_in_outlook()`
Aktualizuje notatki w kontaktach Outlooka, dodając informację o liczbie dni od ostatniego kontaktu. Nowa informacja o ostatnim kontakcie jest dodawana na początku notatki, zachowując poprzednią zawartość.

## Struktura bazy danych

### Tabela `contacts`
Przechowuje informacje o kontaktach, takie jak:
- `first_name`: Imię,
- `last_name`: Nazwisko,
- `email`, `email2`, `email3`: Adresy e-mail,
- `phone_work`, `phone_mobile`, `phone_home`: Numery telefonów,
- `last_contact`: Data ostatniego kontaktu.

### Tabela `contact_history`
Zawiera historię zmian kontaktów:
- `contact_id`: ID kontaktu,
- `change_type`: Typ zmiany (`add`, `update`, `delete`),
- `changed_by`: Użytkownik, który dokonał zmiany (lub "system").

### Tabela `blacklist`
Przechowuje informacje o zablokowanych kontaktach:
- `email`: Zablokowany adres e-mail,
- `domain`: Zablokowana domena,
- `prefix`: Zablokowany prefiks (np. `noreply@*`).

## Działanie krok po kroku
1. **Eksport kontaktów z Outlooka**: Program pobiera kontakty z Outlooka i zapisuje je do bazy PostgreSQL.
2. **Synchronizacja z PostgreSQL do Outlooka**: Jeśli kontakt istnieje w bazie, jest aktualizowany lub dodawany do Outlooka.
3. **Sprawdzanie czarnej listy**: Program usuwa kontakty z Outlooka i bazy, których adresy są na czarnej liście.
4. **Sprawdzanie korespondencji e-mail**: Program sprawdza skrzynkę odbiorczą i dodaje nowe kontakty do bazy, o ile nie są na czarnej liście.
5. **Aktualizacja notatek w kontaktach**: Program dodaje informację o liczbie dni od ostatniego kontaktu do notatek w Outlooku.
