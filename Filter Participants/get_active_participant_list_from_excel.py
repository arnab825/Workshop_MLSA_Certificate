import pandas as pd
from fuzzywuzzy import fuzz

def get_active_participants_from_excel(excel_path, name_column, email_column, participant_list, output_path, ambassador_host, ambassador_cohost):
    # Read the Excel file
    excel_data = pd.read_excel(excel_path)

    # Convert all data to lowercase for better similarity matching
    excel_data = excel_data.applymap(lambda x: x.lower() if isinstance(x, str) else x)
    excel_data = excel_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Drop duplicate entries based on name and email columns
    unique_data = excel_data.drop_duplicates(subset=[name_column, email_column])
    unique_data.reset_index(drop=True, inplace=True)

    # Extract names and emails from the cleaned Excel data
    excel_names = unique_data[name_column]
    excel_emails = unique_data[email_column]

    # List to store matched participants
    matched_participants = []

    # Match participants from the challenge leaderboard to the Excel data
    for i, name in enumerate(excel_names):
        for participant in participant_list:
            similarity_ratio = fuzz.token_set_ratio(name, participant)
            if similarity_ratio >= 75:  # Threshold for name similarity
                matched_participants.append({"Name": name, "Email": excel_emails[i]})
                break

    # Remove any duplicates from matched participants
    matched_participants_unique = [dict(t) for t in {tuple(d.items()) for d in matched_participants}]

    # Write matched participants with ambassador names to CSV
    with open(output_path, "w") as file:
        # Write header
        file.write("Name,Email,Host Ambassador,Co-host Ambassador\n")
        
        # Write each participant's information with ambassador names
        for participant in matched_participants_unique:
            file.write(f"{participant['Name'].title()},{participant['Email']},{ambassador_host},{ambassador_cohost}\n")
