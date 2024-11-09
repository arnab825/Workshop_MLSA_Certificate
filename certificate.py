import re

def docx_replace_regex(doc_obj, regex, replace, add_space=False, add_tab=False):
    # Define the replacement with added spaces or tabs if specified
    if add_space:
        replace = f"\t{replace} "  # Add spaces on both sides
        print(f"Added space to replace: {replace}")
    
    if add_tab:
        # Debugging: Print the length of the replace string before applying any tab logic
        print(f"Original replace value: '{replace}' (Length: {len(replace)})")
        
        # Check the length of the replace string
        if len(replace) <= 10:
            replace = f"\t{replace}"  # Add two tabs for longer replace strings
            print(f"Added two tabs to replace: {replace}")
        elif len(replace) <= 20:
            replace = f"\t\t{replace}"    # Add one tab for medium-length replace strings
            print(f"Added one tab to replace: {replace}")
        else:
            print(f"No tab added to replace: {replace}")

    # Check paragraphs
    for p in doc_obj.paragraphs:
        full_text = ''.join(run.text for run in p.runs)
        if regex.search(full_text):
            print(f"Found placeholder in paragraph: '{full_text.strip()}'")
            for inline in p.runs:
                if regex.search(inline.text):
                    print(f"Before replacement in paragraph run: '{inline.text}'")
                    inline.text = regex.sub(replace, inline.text)
                    print(f"After replacement in paragraph run: '{inline.text}'")

    # Check tables
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = ''.join(run.text for p in cell.paragraphs for run in p.runs)
                if regex.search(cell_text):
                    print(f"Found placeholder in table cell: '{cell_text.strip()}'")
                    for p in cell.paragraphs:
                        for inline in p.runs:
                            if regex.search(inline.text):
                                print(f"Before replacement in table cell run: '{inline.text}'")
                                inline.text = regex.sub(replace, inline.text)
                                print(f"After replacement in table cell run: '{inline.text}'")

def replace_info(doc, value, placeholder, add_space=False, add_tab=False):
    pattern = r"\{\s*" + re.escape(placeholder.strip("{}")) + r"\s*\}"
    reg = re.compile(pattern, re.IGNORECASE)  # Make regex case-insensitive
    replace = value.strip()
    found = False

    for p in doc.paragraphs:
        combined_text = ''.join(run.text for run in p.runs)
        if reg.search(combined_text.strip()):
            found = True

    if not found:
        print(f"Placeholder '{placeholder}' not found in document.")
        
    docx_replace_regex(doc, reg, replace, add_space=add_space, add_tab=add_tab)

def replace_participant_name(doc, name):
    replace_info(doc, name, "{Name Surname}")

def replace_event_name(doc, event):
    replace_info(doc, event, "{INSERT EVENT NAME}")

def replace_ambassador_name(doc, ambassador):
    replace_info(doc, ambassador, "{AMBASSADOR NAME}", add_space=True)

def replace_cohost_name(doc, cohost):
    # Debugging: Print the original cohost name before any changes
    print(f"Original cohost name: '{cohost}' (Length: {len(cohost)})")
    
    replace_info(doc, cohost, "{AMBASSADOR NAME1}", add_tab=True)  # Pass add_tab to handle logic

def replace_event_date(doc, date):
    replace_info(doc, date, "{date}")
