import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load credentials
username = 'rajkumarsain.doit'  # Replace with your actual username
password = 'Network@1984'  # Replace with your actual password

# Set up Chrome WebDriver using Service
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Step 1: Open the login page
driver.get("https://sso.rajasthan.gov.in/signin")

# Step 2: Locate the username and password fields and login button
time.sleep(2)  # Adjust sleep time based on the page loading time

# Enter username
username_field = driver.find_element(By.ID, 'cpBody_cpBody_txt_Data1')  # Replace 'username_id' with the actual ID or locator for username
username_field.send_keys(username)

# Enter password
password_field = driver.find_element(By.ID, 'cpBody_cpBody_txt_Data2')  # Replace 'password_id' with the actual ID or locator for password
password_field.send_keys(password)

# Submit the login form (find the login button and click it)
login_button = driver.find_element(By.XPATH, '//*[@id="cpBody_cpBody_btn_LDAPLogin"]')  # Replace 'login_button_id' with the actual ID of the login button
time.sleep(10)
login_button.click()

# Step 3: Wait for the login to complete
time.sleep(10)

# Step 4: Use WebDriverWait to ensure the h3 elements are loaded
try:
    # Wait until h3 elements with class 'filterable' are present
    h3_elements = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'h3.filterable'))
    )

    # Iterate over the elements and check for the text "JAN AADHAAR"
    for element in h3_elements:
        if "JAN AADHAAR" in element.text.upper():  # Compare in uppercase to handle case-insensitivity
            print("Found 'JAN AADHAAR'")
            # You can also click on it if needed
            element.click()
            break
    else:
        print("JAN AADHAAR not found")
         # Step 5: Wait for the new page to load after clicking "Jan Aadhaar"
    time.sleep(5)  # Adjust sleep time based on page loading speed
    
    # Step 6: Find the span element ENROLLMENT with class "dashOptEnrolementEdit" and click it
    span_element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.dashOptEnrolementEdit[title="#"]'))
    )
    span_element.click()
    print("Clicked on the 'span' element.")
    # Step 7: Find the span element GENERIC SEARCH with class "dashOptEnrolementEdit" and click it
    span_element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="div_menuId_1202"]//span[@class="dashOptPopup"]'))
    )
    span_element.click()
    print("Clicked on the 'span' element inside div #div_menuId_1202.")

    # Step 8: Read the Excel file and transfer mobile numbers to the webpage
    excel_path = 'TEMP.xlsx'  # Replace with your Excel file path
    df = pd.read_excel(excel_path)

    for index, row in df.iterrows():
    # Check if the mobile number is NaN (empty) before converting to string
        if pd.isna(row['MOBILE_NO']):
            print(f"Skipping blank mobile number at row {index}")
            continue  # Skip this iteration if the mobile number is blank
        else:
            mobile_number = str(row['MOBILE_NO']).strip()  # Convert to string and strip any whitespace

            # Step 7: Enter the mobile number into the input field
            mobile_input = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, 'mobileNo'))
            )
            mobile_input.clear()
            mobile_input.send_keys(mobile_number)
            print(f"Entered mobile number: {mobile_number}")

            # Step 8: Click on the "Search" button
            search_button = driver.find_element(By.ID, 'btn')
            search_button.click()
            print("Clicked the 'Search' button")
            time.sleep(5)  # Adjust based on the search results loading time


except Exception as e:
    print(f"Error: {e}")

input("Press Enter to exit and close the browser...")
# Remove or comment out the line to close the browser
# driver.quit()  # This line is commented out to keep the browser open
