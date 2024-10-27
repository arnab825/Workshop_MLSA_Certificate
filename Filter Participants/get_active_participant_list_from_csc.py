# Global imports
import time
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# Local Module imports
from get_active_participant_list_from_excel import get_active_participants_from_excel

# Setting variables

# Path to the Excel file having data of participants who registered for the challenge
excel_path = r"C:\Users\Rahul\OneDriveSky\Desktop\AI900ChallengeParticipantData.xlsx"

# Name of columns with name & email information
name_column = "Arnab Roy"
email_column = "Arnab.Roy@studentambassadors.com"

# Path of the file where the final list of participants will be stored
output_path = r"../Data/Participant List.csv"

# Challenge link
challenge_link = "https://teams.microsoft.com/v2/?culture=en-in&country=in"

# Prompt for ambassador names
ambassador_host = input("Enter Host Ambassador Name: ")
ambassador_cohost = input("Enter Co-host Ambassador Name: ")

# Configure Selenium options
options = Options()
options.add_argument('--headless')  # Run Chrome in headless mode, no GUI needed

driver = webdriver.Chrome()

# Open challenge link in Chrome
driver.get(challenge_link)

# Wait for the leaderboard to load
time.sleep(7)

# Get the updated page source after dynamic loading
page_source = driver.page_source

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(page_source, "html.parser")

# Find the div with id "leaderboard-list"
leaderboard_div = soup.find("div", id="leaderboard-list")

# Initialize participant list
participant_list = []

def extract_active_participants(leaderboard_items):
    for item in leaderboard_items:
        user_name = item.find("span", class_="leaderboard-name").text.strip()
        user_score = item.find("span", class_="visually-hidden").text.strip().split(",")[2]
        no_of_modules_completed = int(item.find("span", class_="visually-hidden").text.strip().split(",")[2].split('/')[0])

        # Add participant if they completed all modules
        if no_of_modules_completed == total_no_of_modules:
            participant_list.append(user_name)

    # Close Chrome browser
    driver.quit()

    # Compare and map names from the Excel sheet, saving only matching names to the output file
    get_active_participants_from_excel(
        excel_path, name_column, email_column, participant_list, output_path, ambassador_host, ambassador_cohost
    )

    print("Active Participant list extracted successfully")
    print(f"Please check it at: {output_path}")

    sys.exit()

if leaderboard_div:
    leaderboard_items = leaderboard_div.find_all("li", class_="is-unstyled")
    total_no_of_modules = int(leaderboard_items[0].find("span", "visually-hidden").text.strip().split(",")[2].split('/')[1].split(' ')[0])
    first_page_button = soup.find("button", class_="pagination-link is-current")
    total_pages = int(first_page_button["aria-label"].split()[-1])

    extract_active_participants(leaderboard_items)

    for page in range(2, total_pages + 1):
        next_page_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.pagination-next[aria-label='Next']"))
        )
        next_page_button.click()
        time.sleep(7)
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        leaderboard_div = soup.find("div", id="leaderboard-list")
        leaderboard_items = leaderboard_div.find_all("li", class_="is-unstyled")
        extract_active_participants(leaderboard_items)
else:
    print("Leaderboard div not found or browser session expired.")
    driver.quit()
