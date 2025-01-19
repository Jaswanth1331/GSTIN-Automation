from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
import time
import pandas as pd

# Loading the input data
data = pd.read_excel('companies.xlsx')
data.columns = data.columns.str.strip()

# Edge driver setup
edge_driver_path = r"C:\webdrivers\msedgedriver.exe"
service = Service(executable_path=edge_driver_path)
driver = webdriver.Edge(service=service)

# Proposed Solution
with open('gst_status_results.txt', 'w') as file:
    file.write("GSTIN\tLegal Name\tGSTIN Status\n")
    file.write("=" * 50 + "\n")

    max_retries = 3  # Max retries for Error case
    api_call_delay = 1.5  # Delay between searches to reduce server load
    on_result_page = False  # Tracking variable

    #Iterating through excel file
    for index, entry in data.iterrows():
        gst_number = entry.iloc[1]
        company_name = entry.iloc[2]
        success = False
        attempts = 0

        while not success and attempts < max_retries:
            try:
                if not on_result_page:
                    # Navigate to the home page of website
                    driver.get("https://piceapp.com/gst-number-search")
                    print(f"Processing: {company_name} ({gst_number}) [Home Page]")

                    # Locating the input field and making search through GST number
                    input_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="search_gst_input"]'))
                    )
                    input_field.clear()
                    input_field.send_keys(gst_number)

                    # Locating and hitting the search button
                    search_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="search_gst_button"]'))
                    )
                    search_button.click()
                    print(f"Search initiated for GST: {gst_number}")

                else:
                    print(f"Processing: {company_name} ({gst_number}) [Result Page]")

                    #locating the input field in result page of previous search for subsequent searches
                    input_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="input-search-gst"]'))
                    )
                    input_field.clear()
                    input_field.send_keys(gst_number)

                    # Locate and hit the search button
                    search_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-search-gst"]'))
                    )
                    search_button.click()
                    print(f"Search initiated for GST: {gst_number}")

                # Wait for potential fetching of error or results
                time.sleep(3)  # Allow some time for the page to load

                # Check for the dynamic Error during the laod time "Fetching GST details"
                try:
                    fetching_message = driver.find_element(By.XPATH, '//*[@id="text-error-search-gst"]')
                    if "Fetching GST details" in fetching_message.text:
                        print("Message displayed: Fetching details. Retrying...")
                        time.sleep(3)  # Wait and retry if error message occurs
                        continue
                except:
                    pass

                # Locate the GSTIN status element
                status_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/main/section[1]/div[2]/div/div/div[2]/h3/span[1]'))
                )
                status = status_element.text.strip()
                print(f"Status found: {status}")

                # Write into the result file

                file.write(f"{gst_number}\t{company_name}\t{status}\n")
                success = True  # Mark as success and exit retry loop

                # Mark as result page for subsequent searches
                on_result_page = True

                # Add a delay for each successful call
                print(f"Waiting for {api_call_delay} seconds before the next API call...")
                time.sleep(api_call_delay)

            except Exception as e:
                attempts += 1
                print(f"Error while processing {company_name} ({gst_number}): {e}")

                if attempts >= max_retries:
                    print(f"Max retries reached for {company_name} ({gst_number}). Skipping...")
                    file.write(f"{gst_number}\t{company_name}\tError or not found\n")

# Closing the browser
driver.quit()
print("Process complete. Results saved to gst_status_results.txt.")
