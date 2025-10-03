# RDash Vendor Automation Script v2.1 (Simplified Workflow)
# Clicks "+ Add New Vendor" once, then loops through the Excel file.
# After each entry, it uses the browser's "Back" button to start the next.

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time

# --- 1. CONFIGURATION ---
# !!! IMPORTANT !!!
# UPDATE THIS PATH to where you saved your chromedriver.exe file.
CHROME_DRIVER_PATH = "C:/Users/rando/Desktop/codes/Projects/Vendor Bulk Upload/chromedriver-win64/chromedriver.exe"
EXCEL_FILE_PATH = "Vendors.xlsx"

# --- 2. SETUP SELENIUM & ATTACH TO BROWSER ---
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9211")
try:
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 25) # Wait up to 25 seconds for elements to appear
    print("✅ Script Attached to Browser Successfully.")
except Exception as e:
    print("❌ ERROR: Could not attach to Chrome.")
    print("   Please ensure you have started Chrome in debugging mode using the command provided.")
    print(f"   Error details: {e}")
    exit()

# --- 3. READ DATA FROM EXCEL ---
try:
    df = pd.read_excel(EXCEL_FILE_PATH, dtype={'TRN': str})
    df['TRN'] = df['TRN'].str.replace('.0', '', regex=False)
    print(f"✅ Found {len(df)} vendors in '{EXCEL_FILE_PATH}'. Starting automation...")
except FileNotFoundError:
    print(f"❌ ERROR: The file '{EXCEL_FILE_PATH}' was not found. Please make sure it's in the same folder.")
    exit()

# --- 4. ONE-TIME ACTION: ENTER THE "ADD VENDOR" MODE ---
try:
    print("\n--- Initializing Form ---")
    print("   > Clicking '+ Add New Vendor' once to begin...")
    add_vendor_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='+ Add New Vendor']]")))
    add_vendor_button.click()
except TimeoutException:
    print("❌ ERROR: Could not find the '+ Add New Vendor' button on the page.")
    print("   Please make sure you are on the correct 'Manage Vendor' page before running the script.")
    exit()

# --- 5. MAIN AUTOMATION LOOP ---
for index, row in df.iterrows():
    vendor_name = row['VendorName']
    trn = row['TRN']
    
    print(f"\n--- Processing Vendor: {vendor_name} ---")

    try:
        # Step 1: Click 'Do not have Trade License'
        print("   > 1/5: Clicking 'Do not have Trade License'...")
        no_trade_license_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Do not have Trade License Number']")))
        no_trade_license_button.click()
        
        # Step 2: Enter Company Registration Name
        print(f"   > 2/5: Entering Company Name: '{vendor_name}'...")
        company_name_input = wait.until(EC.visibility_of_element_located((By.NAME, "companyName")))
        company_name_input.send_keys(vendor_name)
        
        # Step 3: Click the first 'Add & Continue'
        print("   > 3/5: Clicking first 'Add & Continue'...")
        add_continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Add & Continue']")))
        add_continue_1.click()
        
        # Step 4: Handle TRN entry or Skip conditionally
        if pd.notna(trn) and str(trn).strip():
            print(f"   > 4/5: TRN found. Entering TRN: '{trn}'...")
            trn_input = wait.until(EC.visibility_of_element_located((By.NAME, "TRN")))
            trn_input.send_keys(str(trn))
            
            print("   > 5/5: Clicking second 'Add & Continue'...")
            add_continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Add & Continue']]")))
            add_continue_2.click()
        else:
            print("   > 4-5/5: No TRN in Excel file. Skipping KYC page...")
            skip_kyc = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='side-panel-footer']//button[.//span[text()='Skip']]")))
            skip_kyc.click()
        
        # Step 5: Wait for success, then navigate back for the next vendor
        print("   > Waiting for success confirmation...")
        # A short pause for the success message to appear and the action to complete.
        time.sleep(3) 

        print("   > Navigating back to the new vendor form...")
        driver.back()

        # Wait for an element on the new form to ensure the 'back' navigation is complete
        wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Do not have Trade License Number']")))
        
        print(f"--- ✅ SUCCESS: Vendor '{vendor_name}' added. Ready for next. ---")

    except Exception as e:
        print(f"   >>> ❌ ERROR processing '{vendor_name}'. Skipping to next vendor. <<<")
        print(f"   Reason: {e}")
        # If an error occurs, we assume the form is in a bad state and try going back
        # to reset for the next vendor in the loop.
        try:
            print("   > Attempting to navigate back to reset the form...")
            driver.back()
            wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Do not have Trade License Number']")))
        except:
            print("   > Could not navigate back. The script might stop here.")
            break # Exit the loop if we can't recover the state
            
# --- 6. CLEANUP ---
print("\n\nAutomation complete. All vendors processed.")
driver.quit()