import os
import csv
from certificate import *
from docx import Document
from docx2pdf import convert

# Create output folders if they don't exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass

def get_participants(file_path):
    data = []
    with open(file_path, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row)
    return data

def create_docx_files(template_path, participants):
    event = input("Enter the event name: ")
    ambassador = input("Enter Host Ambassador Name: ")
    cohost = input("Enter Co-host Ambassador Name: ")
    event_date = input("Enter the event date (e.g., October 27, 2024): ")

    for participant in participants:
        doc = Document(template_path)

        name = participant["Name"].strip()
        # email = participant["Email"].strip()

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)
        replace_cohost_name(doc, cohost)
        replace_event_date(doc, event_date)

        doc.save(f'Output/Doc/{name}.docx')

        print(f"Creating PDF for {name}")
        convert(f'Output/Doc/{name}.docx', f'Output/Pdf/{name}.pdf')

# Set certificate template and participant list paths
certificate_file = "Data/Event Certificate Template.docx"
print("NOTE: Selecting Test Mode as 'N' will use actual data from Data/Participant List.csv. Test mode 'Y' uses temp.csv, generating only 1 certificate for testing.")
participate_file = "Data/" + ("Participant List.csv" if input("Test Mode (Y/N): ").lower().startswith("n") else "temp.csv")

# Get participants and process data
participants = get_participants(participate_file)
create_docx_files(certificate_file, participants)
