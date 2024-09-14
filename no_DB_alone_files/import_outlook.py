import win32com.client
import csv
from datetime import datetime

def find_existing_contact(email, contacts_folder):
    """
    Szuka kontaktu o podanym adresie e-mail w folderze kontaktów.
    Zwraca znaleziony kontakt lub None, jeśli kontakt nie istnieje.
    """
    for contact in contacts_folder.Items:
        if contact.Email1Address == email:
            return contact
    return None

def update_or_create_contact(row, contacts_folder):
    """
    Aktualizuje istniejący kontakt, jeśli znaleziono, lub tworzy nowy, jeśli kontakt nie istnieje.
    """
    # Sprawdź, czy kontakt o podanym adresie e-mail już istnieje
    contact = find_existing_contact(row['Email1Address'], contacts_folder)

    if contact is None:
        # Tworzymy nowy kontakt, jeśli nie znaleziono istniejącego
        contact = contacts_folder.Items.Add()
        print(f"Tworzę nowy kontakt: {row['FirstName']} {row['LastName']}")
    else:
        print(f"Aktualizuję istniejący kontakt: {row['FirstName']} {row['LastName']}")

    # Uzupełnianie danych lub aktualizacja
    contact.FirstName = row['FirstName'] if row['FirstName'] else contact.FirstName
    contact.LastName = row['LastName'] if row['LastName'] else contact.LastName
    contact.Email1Address = row['Email1Address'] if row['Email1Address'] else contact.Email1Address
    contact.Email2Address = row['Email2Address'] if row['Email2Address'] else contact.Email2Address
    contact.Email3Address = row['Email3Address'] if row['Email3Address'] else contact.Email3Address
    contact.BusinessTelephoneNumber = row['BusinessTelephoneNumber'] if row['BusinessTelephoneNumber'] else contact.BusinessTelephoneNumber
    contact.HomeTelephoneNumber = row['HomeTelephoneNumber'] if row['HomeTelephoneNumber'] else contact.HomeTelephoneNumber
    contact.MobileTelephoneNumber = row['MobileTelephoneNumber'] if row['MobileTelephoneNumber'] else contact.MobileTelephoneNumber
    contact.BusinessFaxNumber = row['BusinessFaxNumber'] if row['BusinessFaxNumber'] else contact.BusinessFaxNumber
    contact.HomeFaxNumber = row['HomeFaxNumber'] if row['HomeFaxNumber'] else contact.HomeFaxNumber
    contact.PagerNumber = row['PagerNumber'] if row['PagerNumber'] else contact.PagerNumber
    contact.CompanyName = row['CompanyName'] if row['CompanyName'] else contact.CompanyName
    contact.JobTitle = row['JobTitle'] if row['JobTitle'] else contact.JobTitle
    contact.Department = row['Department'] if row['Department'] else contact.Department
    contact.OfficeLocation = row['OfficeLocation'] if row['OfficeLocation'] else contact.OfficeLocation
    contact.ManagerName = row['ManagerName'] if row['ManagerName'] else contact.ManagerName
    contact.HomeAddressStreet = row['HomeAddressStreet'] if row['HomeAddressStreet'] else contact.HomeAddressStreet
    contact.HomeAddressCity = row['HomeAddressCity'] if row['HomeAddressCity'] else contact.HomeAddressCity
    contact.HomeAddressPostalCode = row['HomeAddressPostalCode'] if row['HomeAddressPostalCode'] else contact.HomeAddressPostalCode
    contact.HomeAddressCountry = row['HomeAddressCountry'] if row['HomeAddressCountry'] else contact.HomeAddressCountry
    contact.BusinessAddressStreet = row['BusinessAddressStreet'] if row['BusinessAddressStreet'] else contact.BusinessAddressStreet
    contact.BusinessAddressCity = row['BusinessAddressCity'] if row['BusinessAddressCity'] else contact.BusinessAddressCity
    contact.BusinessAddressPostalCode = row['BusinessAddressPostalCode'] if row['BusinessAddressPostalCode'] else contact.BusinessAddressPostalCode
    contact.BusinessAddressCountry = row['BusinessAddressCountry'] if row['BusinessAddressCountry'] else contact.BusinessAddressCountry

    # Aktualizacja daty urodzin, jeśli podano nową wartość
    if row['Birthday']:
        try:
            # Konwersja formatu: RRRR-MM-DD HH:MM:SS
            contact.Birthday = datetime.strptime(row['Birthday'], '%Y-%m-%d %H:%M:%S')
        except ValueError:
            print(f"Nieprawidłowy format daty dla kontaktu {contact.FirstName} {contact.LastName}, pomijam urodziny.")

    # Aktualizacja rocznicy, jeśli podano nową wartość
    if row['Anniversary']:
        try:
            # Konwersja formatu: RRRR-MM-DD HH:MM:SS
            contact.Anniversary = datetime.strptime(row['Anniversary'], '%Y-%m-%d %H:%M:%S')
        except ValueError:
            print(f"Nieprawidłowy format daty dla kontaktu {contact.FirstName} {contact.LastName}, pomijam rocznicę.")

    contact.Body = row['Body'] if row['Body'] else contact.Body  # Pole notatek

    contact.Save()  # Zapisujemy zmiany


def import_outlook_contacts(file_path):
    # Połączenie z aplikacją Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Zakładamy, że kontakty mają być importowane do domyślnego folderu kontaktów
    contacts_folder = outlook.GetDefaultFolder(10)  # 10 oznacza folder kontaktów
    
    # Otwieramy plik CSV z kontaktami do importu
    with open(file_path, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        
        for row in reader:
            update_or_create_contact(row, contacts_folder)

    print(f"Kontakty zostały zaimportowane z {file_path}")

# Ścieżka do pliku CSV, z którego będą importowane kontakty
file_path = 'outlook_contacts_extended.csv'

import_outlook_contacts(file_path)
