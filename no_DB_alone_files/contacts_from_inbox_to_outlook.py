import win32com.client
import csv
from datetime import datetime, timedelta

def extract_contacts_from_folder(folder, contacts, batch_size=100, file_path=None):
    """
    Ekstrakcja kontaktów z folderu. Zapisuj kontakty co `batch_size` rekordów do pliku CSV.
    """
    six_months_ago = datetime.now() - timedelta(days=180)  # Filtrujemy wiadomości młodsze niż 6 miesięcy
    batch_counter = 0

    # Iteracja przez wszystkie wiadomości w folderze
    for message in folder.Items:
        if message.Class == 43:  # 43 oznacza typ wiadomości (MailItem)
            try:
                # Przekształcamy ReceivedTime (który ma strefę czasową) na offset-naive, czyli bez strefy czasowej
                received_time = message.ReceivedTime.replace(tzinfo=None)

                # Filtrujemy wiadomości na podstawie daty
                if received_time < six_months_ago:
                    continue

                # Sprawdzamy, czy wiadomość ma nadawcę i nadawca nie jest nieznany
                if not message.Sender or not message.SenderEmailAddress:
                    print(f"Pomijam wiadomość bez nadawcy.")
                    continue

                sender_email = message.SenderEmailAddress
                sender_name = message.SenderName
                
                # Sprawdzanie typu nadawcy i pobieranie dodatkowych informacji tylko wtedy, gdy są dostępne
                sender_company = ''
                business_phone = ''
                mobile_phone = ''
                
                if hasattr(message, 'SenderType') and message.SenderType == 0 and message.Sender.GetExchangeUser():
                    sender_company = message.Sender.GetExchangeUser().CompanyName
                    business_phone = message.Sender.GetExchangeUser().BusinessTelephoneNumber
                    mobile_phone = message.Sender.GetExchangeUser().MobileTelephoneNumber

                # Jeśli adres e-mail nie istnieje w kontaktach, dodajemy
                if sender_email not in contacts:
                    contacts[sender_email] = {
                        'name': sender_name,
                        'company': sender_company,
                        'business_phone': business_phone,
                        'mobile_phone': mobile_phone
                    }

                    batch_counter += 1

                # Zapisujemy kontakty co `batch_size` rekordów
                if batch_counter >= batch_size and file_path:
                    save_contacts_to_csv(contacts, file_path, append=True)
                    batch_counter = 0

            except Exception as e:
                print(f"Problem z przetwarzaniem wiadomości: {e}")
                continue

    # Zapisz pozostałe kontakty, jeśli istnieją
    if batch_counter > 0 and file_path:
        save_contacts_to_csv(contacts, file_path, append=True)

    return contacts

def save_contacts_to_csv(contacts, file_path, append=False):
    """
    Zapis kontaktów do pliku CSV. Jeśli `append` jest ustawione na True, zapisuje dodatkowe rekordy bez nadpisywania pliku.
    """
    write_header = not append  # Nagłówki tylko przy pierwszym zapisie
    mode = 'a' if append else 'w'

    # Zapisujemy kontakty do pliku CSV
    with open(file_path, mode=mode, newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if write_header:
            writer.writerow(["Name", "Email", "Company", "Business Phone", "Mobile Phone"])  # Nagłówki kolumn

        for email, info in contacts.items():
            writer.writerow([info['name'], email, info['company'], info['business_phone'], info['mobile_phone']])

    print(f"Kontakty zostały zapisane do pliku {file_path}")

def find_existing_contact(email, contacts_folder):
    """
    Szuka kontaktu o podanym adresie e-mail w folderze kontaktów.
    Zwraca znaleziony kontakt lub None, jeśli kontakt nie istnieje.
    """
    for contact in contacts_folder.Items:
        if contact.Email1Address == email:
            return contact
    return None

def save_contacts_to_outlook(contacts, contacts_folder):
    """
    Zapisuje kontakty bezpośrednio do książki adresowej Outlooka.
    """
    for email, info in contacts.items():
        # Sprawdzamy, czy taki kontakt już istnieje
        contact = find_existing_contact(email, contacts_folder)
        if contact is None:
            contact = contacts_folder.Items.Add()
            print(f"Tworzę nowy kontakt: {info['name']} ({email})")
        else:
            print(f"Aktualizuję istniejący kontakt: {info['name']} ({email})")

        # Uzupełnianie danych kontaktowych
        contact.FirstName = info['name'] if info['name'] else contact.FirstName
        contact.Email1Address = email
        contact.CompanyName = info['company'] if info['company'] else contact.CompanyName
        contact.BusinessTelephoneNumber = info['business_phone'] if info['business_phone'] else contact.BusinessTelephoneNumber
        contact.MobileTelephoneNumber = info['mobile_phone'] if info['mobile_phone'] else contact.MobileTelephoneNumber

        contact.Save()

    print("Kontakty zostały zapisane do Outlooka.")

def extract_and_save_contacts(save_to_outlook=True, file_path=None, batch_size=100):
    """
    Główna funkcja, która zbiera kontakty i zapisuje je co `batch_size` rekordów do CSV lub Outlooka.
    """
    # Połączenie z aplikacją Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Foldery do przeszukania: Skrzynka odbiorcza, Wysłane elementy
    inbox = outlook.GetDefaultFolder(6)  # 6 oznacza skrzynkę odbiorczą
    sent_items = outlook.GetDefaultFolder(5)  # 5 oznacza wysłane elementy
    
    # Zbieramy kontakty z wybranych folderów
    contacts = {}
    contacts = extract_contacts_from_folder(inbox, contacts, batch_size=batch_size, file_path=file_path)
    contacts = extract_contacts_from_folder(sent_items, contacts, batch_size=batch_size, file_path=file_path)

    if save_to_outlook:
        # Zapisujemy kontakty w Outlooku
        contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów
        save_contacts_to_outlook(contacts, contacts_folder)
    else:
        if file_path:
            save_contacts_to_csv(contacts, file_path, append=False)
        else:
            print("Brak ścieżki do zapisu pliku CSV.")

# Zapisz do Outlooka lub do pliku CSV
# Dla CSV:
#extract_and_save_contacts(save_to_outlook=False, file_path='contacts_from_emails.csv', batch_size=10)

# Dla Outlooka:
extract_and_save_contacts(save_to_outlook=True, batch_size=10)
