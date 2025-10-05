
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time


CHROME_DRIVER_PATH = "C:/Users/rando/Desktop/codes/Projects/Vendor Bulk Upload/chromedriver-win64/chromedriver.exe"
EXCEL_FILE_PATH = "Vendors.xlsx"


chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9211")
try:
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 25) 
    print("✅ Script Attached to Browser Successfully.")
except Exception as e:
    print("❌ ERROR: Could not attach to Chrome.")
    print("   Please ensure you have started Chrome in debugging mode using the command provided.")
    print(f"   Error details: {e}")
    exit()


try:
    df = pd.read_excel(EXCEL_FILE_PATH, dtype={'TRN': str})
    df['TRN'] = df['TRN'].str.replace('.0', '', regex=False)
    print(f"✅ Found {len(df)} vendors in '{EXCEL_FILE_PATH}'. Starting automation...")
except FileNotFoundError:
    print(f"❌ ERROR: The file '{EXCEL_FILE_PATH}' was not found. Please make sure it's in the same folder.")
    exit()


for index, row in df.iterrows():
    vendor_name = row['VendorName']
    trn = row['TRN']
    
    print(f"\n--- Processing Vendor: {vendor_name} ---")

    try:
        
        print("   > 1/6: Clicking '+ Add New Vendor'...")
        add_vendor_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='+ Add New Vendor']]")))
        add_vendor_button.click()

        
        print("   > 2/6: Clicking 'Do not have Trade License'...")
        no_trade_license_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Do not have Trade License Number']")))
        no_trade_license_button.click()
        
        
        print(f"   > 3/6: Entering Company Name: '{vendor_name}'...")
        company_name_input = wait.until(EC.visibility_of_element_located((By.NAME, "companyName")))
        company_name_input.send_keys(vendor_name)
        
        
        print("   > 4/6: Clicking first 'Add & Continue'...")
        add_continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Add & Continue']")))
        add_continue_1.click()
        
        
        if pd.notna(trn) and str(trn).strip():
            print(f"   > 5/6: TRN found. Entering TRN: '{trn}'...")
            trn_input = wait.until(EC.visibility_of_element_located((By.NAME, "TRN")))
            trn_input.send_keys(str(trn))
            
            print("   > 6/6: Clicking second 'Add & Continue'...")
            add_continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Add & Continue']]")))
            add_continue_2.click()
        else:
            print("   > 5-6/6: No TRN in Excel file. Skipping KYC page...")
            skip_kyc = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='side-panel-footer']//button[.//span[text()='Skip']]")))
            skip_kyc.click()
        
        
        print("   > Waiting for success confirmation...")
        time.sleep(3) 

        print("   > Navigating back to prepare for next vendor...")
        driver.back()

        
        wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='+ Add New Vendor']]")))
        
        print(f"--- ✅ SUCCESS: Vendor '{vendor_name}' added. Ready for next. ---")

    except Exception as e:
        print(f"   >>> ❌ ERROR processing '{vendor_name}'. Skipping to next vendor. <<<")
        print(f"   Reason: {e}")
        try:
            print("   > Attempting to recover by refreshing the page...")
            driver.refresh()
            time.sleep(3) 
        except:
            print("   > Could not recover. The script might stop here.")
            break 
            

print("\n\nAutomation complete. All vendors processed.")
driver.quit()