import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Path to the Excel file
EXCEL_FILE = r"C:\Users\bilal\Downloads\BAJAUR BISC TAINEES DATABASE REPLACED V1.xlsx"

# Website login page URL and certificate URL
LOGIN_URL = "https://lms.digitalkonect.com/Student/login/"
CERTIFICATE_URL = "https://lms.digitalkonect.com/Student/stu_viewCertifcate/Vm0xMGFrMVhVWGhUYmtwT1ZsVndVbFpyVWtKUFVUMDk="

# HTML element identifiers for the login form, dashboard, and certificate process
EMAIL_FIELD_ID = "email"  # Replace with the actual ID or selector for the email input field
PASSWORD_FIELD_ID = "password"  # Replace with the actual ID or selector for the password input field
LOGIN_BUTTON_ID = "submit"  # Replace with the actual ID or selector for the login button
PROFILE_BUTTON_ID = "navbarDropdown"  # Replace with the actual ID or selector for the profile menu button
LOGOUT_LINK_CSS = "i.fa-sign-out"  # CSS selector for logout link
DOWNLOAD_PNG_BUTTON_ID = "downloadPNG"  # ID for the download button
PASSWORD = "hit-01"  # Fixed password for all logins


# Initialize the WebDriver (Chrome)
def init_driver():
    service = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=service)  # Assumes chromedriver is in PATH
    return driver


# Read email from Excel
def read_credentials(file_path):
    df = pd.read_excel(file_path)
    if 'EMAIL FOR LOGIN' not in df.columns:
        raise ValueError("Excel file must contain 'EMAIL FOR LOGIN' column")
    return df['EMAIL FOR LOGIN'].tolist()


# Attempt login, download certificate, and log out
def attempt_login_download_logout(driver, email):
    try:
        driver.get(LOGIN_URL)
        # Wait for the email field to be visible
        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, EMAIL_FIELD_ID))
        )
        password_field = driver.find_element(By.ID, PASSWORD_FIELD_ID)
        login_button = driver.find_element(By.ID, LOGIN_BUTTON_ID)

        # Enter credentials
        email_field.clear()
        email_field.send_keys(email)
        password_field.clear()
        password_field.send_keys(PASSWORD)

        # Click login button
        login_button.click()

        # Wait for dashboard to load
        time.sleep(2)

        # Navigate to certificate URL
        driver.get(CERTIFICATE_URL)
        time.sleep(2)  # Wait for certificate page to load

        # Click Download Certificate
        download_png_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, DOWNLOAD_PNG_BUTTON_ID))
        )
        download_png_button.click()
        time.sleep(2)  # Wait for download to initiate

        # Log out
        profile_button = driver.find_element(By.ID, PROFILE_BUTTON_ID)
        profile_button.click()
        time.sleep(1)  # Wait for menu to appear
        logout_link = driver.find_element(By.CSS_SELECTOR, LOGOUT_LINK_CSS)
        logout_link.click()
        time.sleep(1)  # Wait for logout to complete
        print(f"Certificate downloaded and logged out for {email}")
    except Exception as e:
        print(f"Error during process for {email}: {str(e)}")
        return False
    else:
        return True


def main():
    # Initialize WebDriver
    driver = init_driver()
    try:
        # Read emails from Excel
        emails = read_credentials(EXCEL_FILE)

        # Process each email
        for email in emails:
            print(f"Attempting process for {email}...")
            success = attempt_login_download_logout(driver, email)
            if not success:
                print(f"Process failed for {email}")
            # Delay to avoid overwhelming the server
            time.sleep(1)
    finally:
        # Close the browser
        driver.quit()


if __name__ == "__main__":
    main()