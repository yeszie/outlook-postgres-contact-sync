import win32com.client
import csv

def export_outlook_contacts(file_path):
    # Połączenie z aplikacją Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Zakładamy, że kontakty znajdują się w domyślnym folderze kontaktów
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów

    # Definiujemy listę atrybutów, które chcemy pobrać z kontaktów
    fields = [
        "FirstName", "LastName", "Email1Address", "Email2Address", "Email3Address",
        "BusinessTelephoneNumber", "HomeTelephoneNumber", "MobileTelephoneNumber", 
        "BusinessFaxNumber", "HomeFaxNumber", "PagerNumber",
        "CompanyName", "JobTitle", "Department", "OfficeLocation", "ManagerName",
        "HomeAddressStreet", "HomeAddressCity", "HomeAddressPostalCode", "HomeAddressCountry", 
        "BusinessAddressStreet", "BusinessAddressCity", "BusinessAddressPostalCode", "BusinessAddressCountry",
        "Birthday", "Anniversary", "Body"  # Body to pole notatek w Outlooku
    ]
    
    # Otwieramy plik CSV do zapisu kontaktów
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Zapisujemy nagłówki (nazwy pól)
        writer.writerow(fields)
        
        # Iterujemy przez wszystkie kontakty i zapisujemy dane do pliku
        for contact in contacts_folder.Items:
            try:
                # Zbieramy dane dla każdego kontaktu
                row = []
                for field in fields:
                    value = getattr(contact, field, '')  # Pobieramy wartość lub pusty string, jeśli pole nie istnieje
                    row.append(value if value else '')  # Jeśli pole ma wartość, dodajemy ją, inaczej pusty string
                writer.writerow(row)
            except AttributeError:
                continue  # Pomijamy elementy, które nie są kontaktami

    print(f"Kontakty zostały wyeksportowane do {file_path}")

# Ścieżka do pliku CSV, w którym będą zapisane kontakty
file_path = 'outlook_contacts_extended.csv'

export_outlook_contacts(file_path)
