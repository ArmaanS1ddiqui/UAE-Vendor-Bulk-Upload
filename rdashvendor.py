# RDash Vendor Automation Script v1.5 (Final)
# Reads vendor data from an Excel file and automates adding them to Rdash.
# Includes conditional TRN logic, specific button identifiers, and a JavaScript click for the Bank Details page.

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
import time

# --- 1. CONFIGURATION ---
# !!! IMPORTANT !!!
# UPDATE THIS PATH to where you saved your chromedriver.exe file.
CHROME_DRIVER_PATH = "C:/Users/rando/Desktop/codes/Projects/Vendor Bulk Upload/chromedriver-win64/chromedriver.exe" 
EXCEL_FILE_PATH = "Vendors.xlsx"

# --- 2. SETUP SELENIUM TO ATTACH TO YOUR OPEN BROWSER ---
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
    # Read TRN as a string to preserve leading zeros and prevent scientific notation
    df = pd.read_excel(EXCEL_FILE_PATH, dtype={'TRN': str}) 
    # Clean up '.0' if Excel auto-formats a number
    df['TRN'] = df['TRN'].str.replace('.0', '', regex=False) 
    print(f"✅ Found {len(df)} vendors in '{EXCEL_FILE_PATH}'. Starting automation...")
except FileNotFoundError:
    print(f"❌ ERROR: The file '{EXCEL_FILE_PATH}' was not found. Please make sure it's in the same folder.")
    exit()

# --- 4. MAIN AUTOMATION LOOP ---
for index, row in df.iterrows():
    vendor_name = row['VendorName']
    trn = row['TRN']
    vendor_tag = row['VendorTag']
    
    print(f"\n--- Processing Vendor: {vendor_name} ---")

    try:
        # Step 1: Click '+ Add New Vendor'
        print("   > 1/12: Clicking '+ Add New Vendor'...")
        add_vendor_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='+ Add New Vendor']]")))
        add_vendor_button.click()

        # Step 2: Click 'Do not have Trade License'
        print("   > 2/12: Clicking 'Do not have Trade License'...")
        no_trade_license_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Do not have Trade License Number']")))
        no_trade_license_button.click()
        
        # Step 3: Enter Company Registration Name
        print(f"   > 3/12: Entering Company Name: '{vendor_name}'...")
        company_name_input = wait.until(EC.visibility_of_element_located((By.NAME, "companyName")))
        company_name_input.send_keys(vendor_name)
        
        # Step 4: Click the first 'Add & Continue'
        print("   > 4/12: Clicking first 'Add & Continue'...")
        add_continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Add & Continue']")))
        add_continue_1.click()
        
        # Step 5 & 6: Handle TRN entry or Skip conditionally
        if pd.notna(trn) and str(trn).strip():
            print(f"   > 5/12: TRN found. Entering TRN: '{trn}'...")
            trn_input = wait.until(EC.visibility_of_element_located((By.NAME, "TRN")))
            trn_input.send_keys(str(trn))
            
            print("   > 6/12: Clicking second 'Add & Continue'...")
            add_continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Add & Continue']]")))
            add_continue_2.click()
            
            # Cooldown to allow the next page to fully load and animations to finish
            print("   > Waiting for 3 seconds for Bank Details page to load...")
            time.sleep(3)
        else:
            print("   > 5-6/12: No TRN in Excel file. Skipping KYC page...")
            skip_kyc = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='side-panel-footer']//button[.//span[text()='Skip']]")))
            skip_kyc.click()
        
        # Step 7: Skip Bank Details page (using a direct JavaScript click)
        print("   > 7/12: Skipping Bank Details page with a direct click...")
        skip_bank_button = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='side-panel-footer']//button[.//span[text()='Skip']]")))
        driver.execute_script("arguments[0].click();", skip_bank_button)
        
        # Step 8: Skip Other Details page
        print("   > 8/12: Skipping Other Details page...")
        skip_other = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='side-panel-footer']//button[.//span[text()='Skip']]")))
        skip_other.click()

        # Step 9: Skip Vendor User Details page
        print("   > 9/12: Skipping Vendor User Details page...")
        skip_user = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='side-panel-footer']//button[@title='Skip']")))
        skip_user.click()
        
        # Step 10: Open the Vendor Tags dropdown
        print("   > 10/12: Opening Vendor Tags dropdown...")
        tags_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[div[text()='Select vendor Tags']]/following-sibling::div")))
        tags_dropdown.click()
        
        # Step 11: Select the specific vendor tag from the list
        print(f"   > 11/12: Selecting tag: '{vendor_tag}'...")
        tag_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@title='{vendor_tag}']")))
        tag_option.click()
        
        # Click outside the dropdown to close it
        driver.find_element(By.XPATH, "//body").click()
        time.sleep(0.5) 

        # Step 12: Click the final 'Add' button
        print("   > 12/12: Clicking final 'Add' button...")
        final_add_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Add']")))
        final_add_button.click()
        
        print(f"--- ✅ SUCCESS: Vendor '{vendor_name}' added! ---")
        
        # Wait intelligently for the form's overlay to disappear
        print("   > Waiting for the form to close completely...")
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "MuiModal-backdrop")))
        time.sleep(0.5)

    except Exception as e:
        print(f"   >>> ❌ ERROR processing '{vendor_name}'. Skipping. <<<")
        print(f"   Reason: {e}")
        try:
            # Try to find a generic close/cancel button to reset the form
            cancel_button = driver.find_element(By.XPATH, "//button[.//span[text()='Cancel']] | //button[@data-testid='CloseRoundedIcon']")
            cancel_button.click()
            print("   > Form cancelled, proceeding to next vendor.")
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "MuiModal-backdrop")))
        except:
            print("   > Could not find a cancel button. Refreshing page to be safe.")
            driver.refresh()
            time.sleep(3) 

# --- 5. CLEANUP ---
print("\n\nAutomation complete. All vendors processed.")
driver.quit()