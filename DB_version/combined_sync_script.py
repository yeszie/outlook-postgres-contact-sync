import win32com.client
import psycopg2
from datetime import datetime, timedelta
import json, os
 
# Połączenie z bazą danych PostgreSQL
def connect_to_db():
    return psycopg2.connect(
        host="localhost",
        port="5432",
        database="contact",
        user="postgres",
        password="pass"
    )

# Synchronizacja kontaktów z Outlooka do PostgreSQL z dodatkowymi polami
def export_outlook_to_postgres():
    conn = connect_to_db()
    cursor = conn.cursor()

    # Pobranie kontaktów z Outlooka
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    contacts = []
    for contact in contacts_folder.Items:
        try:
            # Pobieranie podstawowych pól
            first_name = getattr(contact, "FirstName", "")
            last_name = getattr(contact, "LastName", "")
            email = getattr(contact, "Email1Address", "")
            email2 = getattr(contact, "Email2Address", "")  # Drugi adres e-mail
            email3 = getattr(contact, "Email3Address", "")  # Trzeci adres e-mail

            # Pobieranie numerów telefonów
            phone_work = getattr(contact, "BusinessTelephoneNumber", "")
            phone_work_2 = getattr(contact, "Business2TelephoneNumber", "")  # Drugi numer pracy
            phone_mobile = getattr(contact, "MobileTelephoneNumber", "")
            phone_mobile_2 = getattr(contact, "OtherTelephoneNumber", "")  # Drugi numer komórkowy
            phone_home = getattr(contact, "HomeTelephoneNumber", "")  # Numer domowy
            phone_fax_work = getattr(contact, "BusinessFaxNumber", "")  # Faks służbowy
            phone_fax_home = getattr(contact, "HomeFaxNumber", "")  # Faks domowy

            last_contact = datetime.now()

            # Zapis do bazy danych PostgreSQL
            cursor.execute("""
                INSERT INTO contacts (first_name, last_name, email, email2, email3, phone_work, phone_work_2, phone_mobile, phone_mobile_2, phone_home, phone_fax_work, phone_fax_home, last_contact)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (email) DO UPDATE SET
                first_name = EXCLUDED.first_name,
                last_name = EXCLUDED.last_name,
                email2 = EXCLUDED.email2,
                email3 = EXCLUDED.email3,
                phone_work = EXCLUDED.phone_work,
                phone_work_2 = EXCLUDED.phone_work_2,
                phone_mobile = EXCLUDED.phone_mobile,
                phone_mobile_2 = EXCLUDED.phone_mobile_2,
                phone_home = EXCLUDED.phone_home,
                phone_fax_work = EXCLUDED.phone_fax_work,
                phone_fax_home = EXCLUDED.phone_fax_home,
                last_contact = EXCLUDED.last_contact;
            """, (first_name, last_name, email, email2, email3, phone_work, phone_work_2, phone_mobile, phone_mobile_2, phone_home, phone_fax_work, phone_fax_home, last_contact))

        except AttributeError:
            continue  # Pomijamy elementy, które nie są kontaktami

    conn.commit()
    cursor.close()
    conn.close()

# Dodanie lub aktualizacja kontaktów w Outlooku na podstawie danych z PostgreSQL
def sync_postgres_to_outlook():
    conn = connect_to_db()
    cursor = conn.cursor()  # Otwieramy kursor na początku

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    # Pobranie kontaktów z PostgreSQL
    cursor.execute("SELECT first_name, last_name, email, email2, email3, phone_work, phone_work_2, phone_mobile, phone_mobile_2, phone_home, phone_fax_work, phone_fax_home FROM contacts")
    contacts = cursor.fetchall()

    for contact in contacts:
        first_name, last_name, email, email2, email3, phone_work, phone_work_2, phone_mobile, phone_mobile_2, phone_home, phone_fax_work, phone_fax_home = contact

        # Sprawdzamy, czy kontakt już istnieje w Outlooku
        existing_contact = None
        for item in contacts_folder.Items:
            if item.Email1Address == email:
                existing_contact = item
                break

        # Jeśli kontakt istnieje, to go aktualizujemy
        if existing_contact:
            if first_name: existing_contact.FirstName = first_name
            if last_name: existing_contact.LastName = last_name
            if email2: existing_contact.Email2Address = email2  # Aktualizacja drugiego adresu e-mail
            if email3: existing_contact.Email3Address = email3  # Aktualizacja trzeciego adresu e-mail
            if phone_work: existing_contact.BusinessTelephoneNumber = phone_work
            if phone_work_2: existing_contact.Business2TelephoneNumber = phone_work_2  # Drugi numer pracy
            if phone_mobile: existing_contact.MobileTelephoneNumber = phone_mobile
            if phone_mobile_2: existing_contact.OtherTelephoneNumber = phone_mobile_2  # Drugi numer komórkowy
            if phone_home: existing_contact.HomeTelephoneNumber = phone_home  # Numer domowy
            if phone_fax_work: existing_contact.BusinessFaxNumber = phone_fax_work  # Faks służbowy
            if phone_fax_home: existing_contact.HomeFaxNumber = phone_fax_home  # Faks domowy
            existing_contact.Save()
            print(f"Zaktualizowano kontakt w Outlooku: {email}")
            
            # Logowanie aktualizacji
            cursor.execute("SELECT id FROM contacts WHERE email = %s", (email,))
            contact_id = cursor.fetchone()[0]
            log_change_to_db(contact_id, 'update')
        else:
            # Jeśli kontakt nie istnieje, dodajemy nowy kontakt
            new_contact = contacts_folder.Items.Add()
            if first_name: new_contact.FirstName = first_name
            if last_name: new_contact.LastName = last_name
            if email: new_contact.Email1Address = email
            if email2: new_contact.Email2Address = email2  # Dodanie drugiego adresu e-mail
            if email3: new_contact.Email3Address = email3  # Dodanie trzeciego adresu e-mail
            if phone_work: new_contact.BusinessTelephoneNumber = phone_work
            if phone_work_2: new_contact.Business2TelephoneNumber = phone_work_2  # Drugi numer pracy
            if phone_mobile: new_contact.MobileTelephoneNumber = phone_mobile
            if phone_mobile_2: new_contact.OtherTelephoneNumber = phone_mobile_2  # Drugi numer komórkowy
            if phone_home: new_contact.HomeTelephoneNumber = phone_home  # Numer domowy
            if phone_fax_work: new_contact.BusinessFaxNumber = phone_fax_work  # Faks służbowy
            if phone_fax_home: new_contact.HomeFaxNumber = phone_fax_home  # Faks domowy
            new_contact.Save()
            print(f"Dodano nowy kontakt w Outlooku: {email}")

            # Logowanie dodania nowego kontaktu
            cursor.execute("SELECT id FROM contacts WHERE email = %s", (email,))
            contact_id = cursor.fetchone()[0]
            log_change_to_db(contact_id, 'add')

    cursor.close()  # Zamknięcie kursora po zakończeniu operacji
    conn.close()  # Zamknięcie połączenia z bazą danych

def update_contact_notes_in_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    # Pobierz kontakty z PostgreSQL
    conn = connect_to_db()
    cursor = conn.cursor()
    cursor.execute("SELECT first_name, last_name, email, last_contact FROM contacts")
    contacts = cursor.fetchall()
    cursor.close()
    conn.close()

    for contact in contacts:
        first_name, last_name, email, last_contact = contact

        # Sprawdzamy, czy kontakt już istnieje w Outlooku
        existing_contact = None
        for item in contacts_folder.Items:
            if item.Email1Address == email:
                existing_contact = item
                break

        if existing_contact:
            # Oblicz ilość dni od ostatniego kontaktu
            days_since_last_contact = (datetime.now() - last_contact).days
            note_text = f"Ostatnia korespondencja: {days_since_last_contact} dni temu"

            # Jeśli notatka już istnieje, dodajemy nową informację na początku
            if existing_contact.Body:
                existing_contact.Body = f"{note_text}\n{existing_contact.Body}"
            else:
                existing_contact.Body = note_text

            # Zapisz zmiany
            existing_contact.Save()

            print(f"Zaktualizowano notatki dla kontaktu {email}: {note_text}")

# Pobranie listy zablokowanych adresów, domen oraz przedrostków z PostgreSQL
def get_blacklist():
    conn = connect_to_db()
    cursor = conn.cursor()
    cursor.execute("SELECT email, domain, prefix FROM blacklist")  # Zakładam, że dodamy kolumnę 'prefix' do tabeli blacklist
    blacklist = cursor.fetchall()
    cursor.close()
    conn.close()
    
    # Oddzielne listy dla adresów e-mail, domen i przedrostków
    emails = [row[0] for row in blacklist if row[0] is not None]
    domains = [row[1] for row in blacklist if row[1] is not None]
    prefixes = [row[2] for row in blacklist if row[2] is not None]
    
    return emails, domains, prefixes

# Dodanie kontaktów do Outlooka z uwzględnieniem blacklisty (e-maile, domeny, prefiksy)
def add_contacts_to_outlook(recent_contacts):
    emails, domains, prefixes = get_blacklist()

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    for contact in recent_contacts:
        email = contact['email']
        email_prefix, email_domain = email.split('@')  # Pobieramy prefiks i domenę z e-maila
        
        # Sprawdzamy, czy adres jest na blacklist (e-mail, domena, prefiks z wildcardami)
        if (email in emails or
            any(email_domain == domain.lstrip('*.') or email_domain.endswith(domain.lstrip('*.')) for domain in domains) or
            any(email_prefix == prefix.rstrip('@*') for prefix in prefixes) or
            any(f"*{email_prefix}@{email_domain}".endswith(f"{wild_prefix}@{wild_domain}".replace('*', '')) for wild_prefix in prefixes for wild_domain in domains)):
            print(f"Adres {email} znajduje się na blacklist i nie zostanie dodany.")
            continue  # Pomijamy ten kontakt, bo jest na blacklist

        # Dodajemy kontakt do Outlooka
        new_contact = contacts_folder.Items.Add()
        new_contact.FullName = contact['name']
        new_contact.Email1Address = email
        new_contact.Save()
        print(f"Dodano nowy kontakt: {email}")

def log_change_to_db(contact_id, change_type):
    # Otwieramy połączenie na nowo wewnątrz tej funkcji
    conn = connect_to_db()
    cursor = conn.cursor()

    # Sprawdzenie, czy contact_id istnieje w tabeli contacts
    cursor.execute("SELECT id FROM contacts WHERE id = %s", (contact_id,))
    result = cursor.fetchone()

    if result:
        try:
            # Pobieramy nazwę użytkownika, który dokonał zmiany
            user = os.getlogin()  # Jeśli użytkownik jest zalogowany, przypiszemy jego nazwę
        except Exception:
            user = "system"  # Jeśli nie ma zalogowanego użytkownika, zapisujemy jako "system"

        cursor.execute("""
            INSERT INTO contact_history (contact_id, change_type, changed_by)
            VALUES (%s, %s, %s);
        """, (contact_id, change_type, user))

        conn.commit()
    else:
        print(f"Kontakt z id={contact_id} nie istnieje w tabeli 'contacts', pominięto zapis do 'contact_history'.")

    cursor.close()  # Zamknięcie kursora
    conn.close()  # Zamknięcie połączenia

# Usuwanie kontaktów z Outlooka i PostgreSQL na podstawie blacklisty
def remove_blacklisted_contacts():
    emails, domains, prefixes = get_blacklist()  # Upewnij się, że przypisujemy 3 wartości
    
    if not emails and not domains and not prefixes:
        print("Brak zablokowanych adresów, domen i prefiksów.")
        return
    
    # Usuwanie kontaktów z Outlooka
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    for contact in contacts_folder.Items:
        email = getattr(contact, "Email1Address", None)
        email_prefix, email_domain = email.split('@') if email else (None, None)
        
        # Sprawdzanie, czy e-mail jest na blacklist (e-mail, domena, prefiks)
        if (email in emails or
            any(email_domain == domain.lstrip('*.') or email_domain.endswith(domain.lstrip('*.')) for domain in domains) or
            any(email_prefix == prefix.rstrip('@*') for prefix in prefixes)):
            print(f"Usuwam kontakt z Outlooka: {email}")
            contact.Delete()  # Usunięcie kontaktu z Outlooka

            # Usunięcie kontaktu z bazy danych PostgreSQL
            conn = connect_to_db()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM contacts WHERE email = %s", (email,))
            conn.commit()
            cursor.close()
            conn.close()
            
            # Logowanie usunięcia
            cursor.execute("SELECT id FROM contacts WHERE email = %s", (email,))
            contact_id = cursor.fetchone()[0]
            log_change_to_db(contact_id, 'delete')

    print("Brak kontaktów w tabeli 'contacts' do usunięcia.")

# Funkcja rekurencyjna do przeszukiwania folderów
def search_emails_in_folder(folder):
    recent_contacts = []
    
    # Przeszukiwanie wiadomości w bieżącym folderze
    for message in folder.Items:
        received_time_naive = message.ReceivedTime.replace(tzinfo=None)
        
        try:
            email = message.SenderEmailAddress
            name = message.SenderName
            recent_contacts.append({'email': email, 'name': name, 'last_contact': received_time_naive})
        except AttributeError:
            continue  # Pomijamy wiadomości bez nadawcy
    
    # Przeszukiwanie podfolderów
    for subfolder in folder.Folders:
        recent_contacts.extend(search_emails_in_folder(subfolder))  # Rekurencyjnie przeszukaj podfoldery

    return recent_contacts

# Funkcja sprawdzająca wiadomości w skrzynce odbiorczej w ciągu ostatnich 6 miesięcy
def check_recent_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 oznacza Skrzynkę Odbiorczą

    # Rekurencyjnie przeszukujemy skrzynkę odbiorczą i wszystkie jej podfoldery
    recent_contacts = search_emails_in_folder(inbox)

    return recent_contacts

# Zapisz kontakty z wiadomości w bazie PostgreSQL
def save_recent_contacts_to_db(recent_contacts):
    emails, domains, prefixes = get_blacklist()  # Pobierz blacklistę przed dodaniem kontaktów

    conn = connect_to_db()
    cursor = conn.cursor()

    for contact in recent_contacts:
        email = contact['email']
        last_contact = contact['last_contact']  # Bierzemy datę z wiadomości e-mail

        # Sprawdzamy, czy adres e-mail zawiera '@'
        if '@' in email:
            email_prefix, email_domain = email.split('@')
        else:
            print(f"Błędny adres e-mail: {email}, pomijam ten kontakt.")
            continue  # Pomijamy ten kontakt, ponieważ nie jest prawidłowy

        # Sprawdzamy, czy adres jest na blackliście
        if (email in emails or
            any(email_domain == domain.lstrip('*.') or email_domain.endswith(domain.lstrip('*.')) for domain in domains) or
            any(email_prefix == prefix.rstrip('@*') for prefix in prefixes)):
            print(f"Adres {email} znajduje się na blacklist i nie zostanie dodany do bazy danych.")
            continue  # Pomijamy ten kontakt, bo jest na blacklist

        # Zapis kontaktu do bazy danych
        cursor.execute("""
            INSERT INTO contacts (email, last_contact)
            VALUES (%s, %s)
            ON CONFLICT (email) DO UPDATE SET
            last_contact = EXCLUDED.last_contact;
        """, (email, last_contact))

        # Logowanie zmiany w contact_history
        cursor.execute("SELECT id FROM contacts WHERE email = %s", (email,))
        contact_id = cursor.fetchone()[0]
        log_change_to_db(contact_id, 'update')

    conn.commit()
    cursor.close()
    conn.close()

# Funkcja do wyszukiwania adresów nieaktualnych (brak kontaktu powyżej 12 miesięcy)
def find_inactive_contacts():
    conn = connect_to_db()
    cursor = conn.cursor()

    one_year_ago = datetime.now() - timedelta(days=365)  # 12 miesięcy
    cursor.execute("SELECT email, last_contact FROM contacts WHERE last_contact <= %s", (one_year_ago,))
    inactive_contacts = cursor.fetchall()

    cursor.close()
    conn.close()
    return inactive_contacts

if __name__ == "__main__":
    # 1. Synchronizacja z Outlooka do PostgreSQL
    export_outlook_to_postgres()

    # 2. Synchronizacja z PostgreSQL do Outlooka
    sync_postgres_to_outlook()

    # 3. Usunięcie zablokowanych kontaktów (Outlook i PostgreSQL)
    remove_blacklisted_contacts()

    # 4. Sprawdzenie wiadomości nie starszych niż xxx miesięcy
    recent_contacts = check_recent_emails()
    save_recent_contacts_to_db(recent_contacts)

    # 5. Wyszukaj kontakty z brakiem kontaktu powyżej 12 miesięcy
    #inactive_contacts = find_inactive_contacts()

    #if inactive_contacts:
    #    print(f"Nieaktualne kontakty (brak kontaktu >12 miesięcy): {inactive_contacts}")

    # 6. Aktualizacja notatek z informacją o liczbie dni od ostatniego kontaktu
    update_contact_notes_in_outlook()  # Wywołanie funkcji aktualizującej notatki