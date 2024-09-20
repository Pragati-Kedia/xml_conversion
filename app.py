##code whre all JQ, MQ, DQ, SQ data is being downloaded

# import os
# import time
# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.keys import Keys
# import pyautogui
 
 
# # Initialize WebDriver
# def initialize_driver(download_dir):
#     options = Options()
#     options.add_argument("--start-maximized")
#     options.add_argument("--incognito")
#     prefs = {
#         "download.default_directory": download_dir,
#         "download.prompt_for_download": False,
#         "download.directory_upgrade": True,
#         "safebrowsing.enabled": True
#     }
#     options.add_experimental_option("prefs", prefs)
#     driver = webdriver.Chrome(options=options)
#     return driver
 
# # Read Excel file
# def read_company_details(file_path):
#     df = pd.read_excel(file_path)
#     return df
 
# # Navigate to BSE results page
# def navigate_to_results(driver, company_name):
#     driver.get("https://www.bseindia.com/corporates/Comp_Resultsnew.aspx")
#     try:
#         search_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_SmartSearch_smartSearch"]')
#         search_box.send_keys(company_name)
#         time.sleep(3)  # Wait for the results to load
#         search_box.send_keys(Keys.ARROW_DOWN)  # Select the first result
#         search_box.send_keys(Keys.ENTER)
#         time.sleep(3)
#         return True
#     except Exception as e:
#         print(f"Error navigating to results for {company_name}: {e}")
#         return False
 
# # Set the broadcast period
# def set_broadcast_period(driver):
#     try:
#         broadcast_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]')
#         broadcast_box.click()
#         last_option = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]/option[7]')
#         last_option.click()
#         submit_button = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnSubmit"]')
#         submit_button.click()
#         time.sleep(5)  # Wait for the page to load
#         return True
#     except Exception as e:
#         print(f"Error setting broadcast period: {e}")
#         return False
 
# # Function to click XBRL links based on period starting with JQ, MQ, DQ, SQ
# def click_xbrl_links(driver, company_name, start_date, end_date, download_dir):
#     try:
#         rows = driver.find_elements(By.XPATH, "//tr[td[4] and td[7]//a]")  # Find rows with both period and XBRL link
#         clicked_count = 0
 
#         for row in rows:
#             try:
#                 # Extract the period text from the 4th cell and XBRL link from the 7th cell
#                 period_element = row.find_element(By.XPATH, './td[4]')
#                 period_text = period_element.text.strip()
#                 xbrl_element = row.find_element(By.XPATH, './td[7]//a')
#                 xbrl_script = xbrl_element.get_attribute("href")
 
#                 # Debugging info
#                 print(f"Processing row with period: '{period_text}' and XBRL script: {xbrl_script}")
 
#                 # Check if the first two characters of the period text are JQ, MQ, DQ, or SQ
#                 if period_text[:2] in ["JQ", "MQ", "DQ", "SQ"]:
#                     print(f"Scraping data for period: '{period_text}'")
 
#                     # Use JavaScript to simulate a click
#                     driver.execute_script("arguments[0].click();", xbrl_element)
#                     time.sleep(5)  # Ensure the new content has loaded
 
#                     # Switch to new tab/window if opened
#                     if len(driver.window_handles) > 1:
#                         driver.switch_to.window(driver.window_handles[1])
 
#                         # Save the file
#                         pyautogui.hotkey('ctrl', 's')
#                         time.sleep(3)
#                         custom_file_name = f"{period_text}_{company_name}.xml"
#                         custom_file_path = os.path.join(download_dir, custom_file_name)
#                         pyautogui.typewrite(custom_file_path)
#                         time.sleep(3)
#                         pyautogui.press('enter')
#                         time.sleep(5)
 
#                         driver.close()
#                         driver.switch_to.window(driver.window_handles[0])
 
#                     clicked_count += 1
#                 else:
#                     print(f"Skipping row with period: '{period_text}' because it does not start with JQ, MQ, DQ, or SQ")
 
#             except Exception as e:
#                 print(f"Error processing row: {e}")
 
#         print(f"Total XBRL files downloaded for {company_name}: {clicked_count}")
#     except Exception as e:
#         print(f"Error clicking XBRL links for {company_name}: {e}")
 
# # Main script execution remains unchanged
# if __name__ == "__main__":
#     input_file_path = r'D:\BSE\input\SampleInput.xlsx'  # Replace with actual path for input
#     xml_file_path = r'D:\BSE\xml'  # Replace with actual path for xml
#     company_df = read_company_details(input_file_path)
 
#     for index, row in company_df.iterrows():
#         try:
#             company_name = row['NAME OF COMPANY']
#             start_date = row['START DATE']
#             end_date = row['END DATE']
           
#             driver = initialize_driver(xml_file_path)
#             if navigate_to_results(driver, company_name):
#                 if set_broadcast_period(driver):
#                     click_xbrl_links(driver, company_name, start_date, end_date, xml_file_path)
#             driver.quit()
#         except Exception as e:
#             print(f"An error occurred during the scraping process: {e}")
#         finally:
#             if driver:
#                 driver.quit()



# ##code where selecctive from start and end date data is being dowwnloaded
# import os
# import time
# import pyautogui
# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.keys import Keys

# # Initialize WebDriver
# def initialize_driver(download_dir):
#     options = Options()
#     options.add_argument("--start-maximized")
#     options.add_argument("--incognito")
#     prefs = {
#         "download.default_directory": download_dir,
#         "download.prompt_for_download": False,
#         "download.directory_upgrade": True,
#         "safebrowsing.enabled": True
#     }
#     options.add_experimental_option("prefs", prefs)
#     driver = webdriver.Chrome(options=options)
#     return driver

# # Convert date like 'JQ2024-2025' to a comparable tuple (year, quarter)
# def parse_quarter_date(quarter_date):
#     quarter = quarter_date[:2]  # Extract "JQ", "MQ", etc.
#     year = int(quarter_date[2:6])  # Extract year part (e.g., 2024)
#     return (year, quarter)

# # Map quarters to numbers to compare
# def quarter_to_num(quarter):
#     quarter_map = {"JQ": 1, "MQ": 2, "SQ": 3, "DQ": 4}
#     return quarter_map.get(quarter, 0)

# # Compare if a given period is within the start and end range
# def is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
#     start_quarter_num = quarter_to_num(start_quarter)
#     end_quarter_num = quarter_to_num(end_quarter)
#     period_quarter_num = quarter_to_num(period_quarter)
    
#     if start_year > end_year:  # Case when range is backwards
#         return (start_year >= period_year >= end_year) and (
#             (period_year == start_year and period_quarter_num >= start_quarter_num) or
#             (period_year == end_year and period_quarter_num <= end_quarter_num) or
#             (end_year < period_year < start_year)
#         )
#     else:  # Normal forward range
#         return (start_year <= period_year <= end_year) and (
#             (period_year == start_year and period_quarter_num >= start_quarter_num) or
#             (period_year == end_year and period_quarter_num <= end_quarter_num) or
#             (start_year < period_year < end_year)
#         )

# # Navigate to BSE results page
# def navigate_to_results(driver, company_name):
#     driver.get("https://www.bseindia.com/corporates/Comp_Resultsnew.aspx")
#     try:
#         search_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_SmartSearch_smartSearch"]')
#         search_box.send_keys(company_name)
#         time.sleep(3)  # Wait for the results to load
#         search_box.send_keys(Keys.ARROW_DOWN)  # Select the first result
#         search_box.send_keys(Keys.ENTER)
#         time.sleep(3)  # Wait for the page to load completely
#         return True
#     except Exception as e:
#         print(f"Error navigating to results for {company_name}: {e}")
#         return False

# # Set the broadcast period
# def set_broadcast_period(driver):
#     try:
#         broadcast_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]')
#         broadcast_box.click()
#         last_option = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]/option[7]')
#         last_option.click()
#         submit_button = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnSubmit"]')
#         submit_button.click()
#         time.sleep(5)  # Wait for the page to load after submitting
#         return True
#     except Exception as e:
#         print(f"Error setting broadcast period: {e}")
#         return False

# # Click XBRL links based on period within start and end date
# def click_xbrl_links(driver, company_name, start_date, end_date, download_dir):
#     try:
#         start_year, start_quarter = parse_quarter_date(start_date)
#         end_year, end_quarter = parse_quarter_date(end_date)
        
#         rows = driver.find_elements(By.XPATH, "//tr[td[4] and td[7]//a]")  # Find rows with both period and XBRL link
#         clicked_count = 0

#         for row in rows:
#             try:
#                 # Extract the period text from the 4th cell and XBRL link from the 7th cell
#                 period_element = row.find_element(By.XPATH, './td[4]')
#                 period_text = period_element.text.strip()
#                 xbrl_element = row.find_element(By.XPATH, './td[7]//a')

#                 # Check if the first two characters of the period text are JQ, MQ, DQ, or SQ
#                 if period_text[:2] not in ["JQ", "MQ", "DQ", "SQ"]:
#                     print(f"Skipping row with period: '{period_text}' because it does not start with JQ, MQ, DQ, or SQ")
#                     continue

#                 # Parse the period to get the year and quarter
#                 period_year, period_quarter = parse_quarter_date(period_text)

#                 # Check if the period is within the range of start_date and end_date
#                 if is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
#                     print(f"Scraping data for period: '{period_text}'")

#                     # Highlight the element by adding a red border
#                     driver.execute_script("arguments[0].style.border='3px solid red';", xbrl_element)
#                     time.sleep(2)  # Add time to see the highlight

#                     # Use JavaScript to simulate a click
#                     driver.execute_script("arguments[0].click();", xbrl_element)
#                     time.sleep(10)  # Increase wait time for new page to load or file to download

#                     # Switch to new tab/window if opened
#                     if len(driver.window_handles) > 1:
#                         driver.switch_to.window(driver.window_handles[1])

#                         # Save the file
#                         pyautogui.hotkey('ctrl', 's')
#                         time.sleep(5)  # Increased wait time for save dialog to appear
#                         custom_file_name = f"{period_text}_{company_name}.xml"
#                         custom_file_path = os.path.join(download_dir, custom_file_name)
#                         pyautogui.typewrite(custom_file_path)
#                         time.sleep(5)  # Increased wait time for typing the file name
#                         pyautogui.press('enter')
#                         time.sleep(15)  # Increased wait time to ensure the file is saved
#                         driver.close()  # Close the new tab
#                         driver.switch_to.window(driver.window_handles[0])

#                     clicked_count += 1
#                 else:
#                     print(f"Skipping row with period: '{period_text}' as it is outside the date range.")

#             except Exception as e:
#                 print(f"Error processing row: {e}")

#         print(f"Total XBRL files downloaded for {company_name}: {clicked_count}")
#     except Exception as e:
#         print(f"Error clicking XBRL links for {company_name}: {e}")

# # Main script execution
# if __name__ == "__main__":
#     input_file_path = r'D:\BSE\input\SampleInput.xlsx'  # Replace with actual path for input
#     xml_file_path = r'D:\BSE\xml'  # Replace with actual path for xml
#     company_df = pd.read_excel(input_file_path)

#     for index, row in company_df.iterrows():
#         try:
#             company_name = row['NAME OF COMPANY']
#             start_date = row['START DATE']
#             end_date = row['END DATE']

#             driver = initialize_driver(xml_file_path)
#             if navigate_to_results(driver, company_name):
#                 if set_broadcast_period(driver):
#                     click_xbrl_links(driver, company_name, start_date, end_date, xml_file_path)
#             driver.quit()
#         except Exception as e:
#             print(f"An error occurred during the scraping process: {e}")
#         finally:
#             if driver:
#                 driver.quit()




##code where selecctive from start and end date data is being dowwnloaded
# import os
# import time
# import pyautogui
# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.keys import Keys

# # Initialize WebDriver
# def initialize_driver(download_dir):
#     options = Options()
#     options.add_argument("--start-maximized")
#     options.add_argument("--incognito")
#     prefs = {
#         "download.default_directory": download_dir,
#         "download.prompt_for_download": False,
#         "download.directory_upgrade": True,
#         "safebrowsing.enabled": True
#     }
#     options.add_experimental_option("prefs", prefs)
#     driver = webdriver.Chrome(options=options)
#     return driver

# # Convert date like 'JQ2024-2025' to a comparable tuple (year, quarter)
# def parse_quarter_date(quarter_date):
#     quarter = quarter_date[:2]  # Extract "JQ", "MQ", etc.
#     year = int(quarter_date[2:6])  # Extract year part (e.g., 2024)
#     return (year, quarter)

# # Map quarters to numbers to compare
# def quarter_to_num(quarter):
#     quarter_map = {"JQ": 2, "MQ": 1, "SQ": 3, "DQ": 4}
#     return quarter_map.get(quarter, 0)

# # Compare if a given period is within the start and end range
# def is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
#     start_quarter_num = quarter_to_num(start_quarter)
#     end_quarter_num = quarter_to_num(end_quarter)
#     period_quarter_num = quarter_to_num(period_quarter)
    
#     if start_year == end_year and start_quarter == end_quarter:
#         return (period_year == start_year) and (period_quarter == start_quarter)
    
#     if start_year > end_year:  # Case when range is backwards
#         return (start_year >= period_year >= end_year) and (
#             (period_year == start_year and period_quarter_num >= start_quarter_num) or
#             (period_year == end_year and period_quarter_num <= end_quarter_num) or
#             (end_year < period_year < start_year)
#         )
#     else:  # Normal forward range
#         return (start_year <= period_year <= end_year) and (
#             (period_year == start_year and period_quarter_num >= start_quarter_num) or
#             (period_year == end_year and period_quarter_num <= end_quarter_num) or
#             (start_year < period_year < end_year)
#         )


# # Navigate to BSE results page
# def navigate_to_results(driver, company_name):
#     driver.get("https://www.bseindia.com/corporates/Comp_Resultsnew.aspx")
#     try:
#         search_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_SmartSearch_smartSearch"]')
#         search_box.send_keys(company_name)
#         time.sleep(3)  # Wait for the results to load
#         search_box.send_keys(Keys.ARROW_DOWN)  # Select the first result
#         search_box.send_keys(Keys.ENTER)
#         time.sleep(3)  # Wait for the page to load completely
#         return True
#     except Exception as e:
#         print(f"Error navigating to results for {company_name}: {e}")
#         return False

# # Set the broadcast period
# def set_broadcast_period(driver):
#     try:
#         broadcast_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]')
#         broadcast_box.click()
#         last_option = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]/option[7]')
#         last_option.click()
#         submit_button = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnSubmit"]')
#         submit_button.click()
#         time.sleep(5)  # Wait for the page to load after submitting
#         return True
#     except Exception as e:
#         print(f"Error setting broadcast period: {e}")
#         return False

# # Click XBRL links based on period within start and end date
# def click_xbrl_links(driver, company_name, start_date, end_date, download_dir):
#     try:
#         start_year, start_quarter = parse_quarter_date(start_date)
#         end_year, end_quarter = parse_quarter_date(end_date)
        
#         rows = driver.find_elements(By.XPATH, "//tr[td[4] and td[7]//a]")  # Find rows with both period and XBRL link
#         clicked_count = 0

#         for row in rows:
#             try:
#                 # Extract the period text from the 4th cell and XBRL link from the 7th cell
#                 period_element = row.find_element(By.XPATH, './td[4]')
#                 period_text = period_element.text.strip()
#                 xbrl_element = row.find_element(By.XPATH, './td[7]//a')

#                 # Check if the first two characters of the period text are JQ, MQ, DQ, or SQ
#                 if period_text[:2] not in ["JQ", "MQ", "DQ", "SQ"]:
#                     print(f"Skipping row with period: '{period_text}' because it does not start with JQ, MQ, DQ, or SQ")
#                     continue

#                 # Parse the period to get the year and quarter
#                 period_year, period_quarter = parse_quarter_date(period_text)

#                 # Check if the period is within the range of start_date and end_date
#                 if is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
#                     print(f"Scraping data for period: '{period_text}'")

#                     # Highlight the element by adding a red border
#                     driver.execute_script("arguments[0].style.border='3px solid red';", xbrl_element)
#                     time.sleep(2)  # Add time to see the highlight

#                     # Use JavaScript to simulate a click
#                     driver.execute_script("arguments[0].click();", xbrl_element)
#                     time.sleep(10)  # Increase wait time for new page to load or file to download

#                     # Switch to new tab/window if opened
#                     if len(driver.window_handles) > 1:
#                         driver.switch_to.window(driver.window_handles[1])

#                         # Save the file
#                         pyautogui.hotkey('ctrl', 's')
#                         time.sleep(5)  # Increased wait time for save dialog to appear
#                         custom_file_name = f"{period_text}_{company_name}.xml"
#                         custom_file_path = os.path.join(download_dir, custom_file_name)
#                         pyautogui.typewrite(custom_file_path)
#                         time.sleep(5)  # Increased wait time for typing the file name
#                         pyautogui.press('enter')
#                         time.sleep(15)  # Increased wait time to ensure the file is saved
#                         driver.close()  # Close the new tab
#                         driver.switch_to.window(driver.window_handles[0])

#                     clicked_count += 1
#                 else:
#                     print(f"Skipping row with period: '{period_text}' as it is outside the date range.")

#             except Exception as e:
#                 print(f"Error processing row: {e}")

#         print(f"Total XBRL files downloaded for {company_name}: {clicked_count}")
#     except Exception as e:
#         print(f"Error clicking XBRL links for {company_name}: {e}")

# # Main script execution
# if __name__ == "__main__":
#     input_file_path = r'D:\BSE\input\SampleInput.xlsx'  # Replace with actual path for input
#     xml_file_path = r'D:\BSE\xml'  # Replace with actual path for xml
#     company_df = pd.read_excel(input_file_path)

#     for index, row in company_df.iterrows():
#         try:
#             company_name = row['NAME OF COMPANY']
#             start_date = row['START DATE']
#             end_date = row['END DATE']

#             driver = initialize_driver(xml_file_path)
#             if navigate_to_results(driver, company_name):
#                 if set_broadcast_period(driver):
#                     click_xbrl_links(driver, company_name, start_date, end_date, xml_file_path)
#             driver.quit()
#         except Exception as e:
#             print(f"An error occurred during the scraping process: {e}")
#         finally:
#             if driver:
#                 driver.quit()



##CODE WHERE NAMING OF XML IS YYYY-YYYY-QUATER_COMPANY NAME
import os
import time
import pyautogui
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

# Initialize WebDriver
def initialize_driver(download_dir):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--incognito")
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return driver

# Convert date like 'JQ2024-2025' to a comparable tuple (year, quarter)
def parse_quarter_date(quarter_date):
    quarter = quarter_date[:2]  # Extract "JQ", "MQ", etc.
    year = int(quarter_date[2:6])  # Extract year part (e.g., 2024)
    return (year, quarter)

# Map quarters to numbers to compare
def quarter_to_num(quarter):
    quarter_map = {"JQ": 2, "MQ": 1, "SQ": 3, "DQ": 4}
    return quarter_map.get(quarter, 0)

# Compare if a given period is within the start and end range
def is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
    start_quarter_num = quarter_to_num(start_quarter)
    end_quarter_num = quarter_to_num(end_quarter)
    period_quarter_num = quarter_to_num(period_quarter)
    
    if start_year == end_year and start_quarter == end_quarter:
        return (period_year == start_year) and (period_quarter == start_quarter)
    
    if start_year > end_year:  # Case when range is backwards
        return (start_year >= period_year >= end_year) and (
            (period_year == start_year and period_quarter_num >= start_quarter_num) or
            (period_year == end_year and period_quarter_num <= end_quarter_num) or
            (end_year < period_year < start_year)
        )
    else:  # Normal forward range
        return (start_year <= period_year <= end_year) and (
            (period_year == start_year and period_quarter_num >= start_quarter_num) or
            (period_year == end_year and period_quarter_num <= end_quarter_num) or
            (start_year < period_year < end_year)
        )


# Navigate to BSE results page
def navigate_to_results(driver, company_name):
    driver.get("https://www.bseindia.com/corporates/Comp_Resultsnew.aspx")
    try:
        search_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_SmartSearch_smartSearch"]')
        search_box.send_keys(company_name)
        time.sleep(3)  # Wait for the results to load
        search_box.send_keys(Keys.ARROW_DOWN)  # Select the first result
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)  # Wait for the page to load completely
        return True
    except Exception as e:
        print(f"Error navigating to results for {company_name}: {e}")
        return False

# Set the broadcast period
def set_broadcast_period(driver):
    try:
        broadcast_box = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]')
        broadcast_box.click()
        last_option = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_broadcastdd"]/option[7]')
        last_option.click()
        submit_button = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnSubmit"]')
        submit_button.click()
        time.sleep(5)  # Wait for the page to load after submitting
        return True
    except Exception as e:
        print(f"Error setting broadcast period: {e}")
        return False

# Click XBRL links based on period within start and end date
# Click XBRL links based on period within start and end date
def click_xbrl_links(driver, company_name, start_date, end_date, download_dir):
    try:
        start_year, start_quarter = parse_quarter_date(start_date)
        end_year, end_quarter = parse_quarter_date(end_date)
        
        rows = driver.find_elements(By.XPATH, "//tr[td[4] and td[7]//a]")  # Find rows with both period and XBRL link
        clicked_count = 0

        for row in rows:
            try:
                # Extract the period text from the 4th cell and XBRL link from the 7th cell
                period_element = row.find_element(By.XPATH, './td[4]')
                period_text = period_element.text.strip()
                xbrl_element = row.find_element(By.XPATH, './td[7]//a')

                # Check if the first two characters of the period text are JQ, MQ, DQ, or SQ
                if period_text[:2] not in ["JQ", "MQ", "DQ", "SQ"]:
                    print(f"Skipping row with period: '{period_text}' because it does not start with JQ, MQ, DQ, or SQ")
                    continue

                # Parse the period to get the year and quarter
                period_year, period_quarter = parse_quarter_date(period_text)

                # Check if the period is within the range of start_date and end_date
                if is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
                    print(f"Scraping data for period: '{period_text}'")

                    # Highlight the element by adding a red border
                    driver.execute_script("arguments[0].style.border='3px solid red';", xbrl_element)
                    time.sleep(2)  # Add time to see the highlight

                    # Use JavaScript to simulate a click
                    driver.execute_script("arguments[0].click();", xbrl_element)
                    time.sleep(10)  # Increase wait time for new page to load or file to download

                    # Switch to new tab/window if opened
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[1])

                        # Define the quarter name based on the period abbreviation
                        quarter_name_map = {"JQ": "Q1", "MQ": "Q4", "SQ": "Q2", "DQ": "Q3"}
                        quarter_name = quarter_name_map.get(period_text[:2], "")

                        # Create the custom file name using the new format
                        custom_file_name = f"{period_year}-{period_year+1}_{quarter_name}_{company_name}.xml"
                        custom_file_path = os.path.join(download_dir, custom_file_name)

                        # Save the file
                        pyautogui.hotkey('ctrl', 's')
                        time.sleep(5)  # Increased wait time for save dialog to appear
                        pyautogui.typewrite(custom_file_path)
                        time.sleep(5)  # Increased wait time for typing the file name
                        pyautogui.press('enter')
                        time.sleep(15)  # Increased wait time to ensure the file is saved
                        driver.close()  # Close the new tab
                        driver.switch_to.window(driver.window_handles[0])

                    clicked_count += 1
                else:
                    print(f"Skipping row with period: '{period_text}' as it is outside the date range.")

            except Exception as e:
                print(f"Error processing row: {e}")

        print(f"Total XBRL files downloaded for {company_name}: {clicked_count}")
    except Exception as e:
        print(f"Error clicking XBRL links for {company_name}: {e}")

# Main script execution
if __name__ == "__main__":
    input_file_path = r'D:\BSE\input\SampleInput.xlsx'  # Replace with actual path for input
    xml_file_path = r'D:\BSE\xml'  # Replace with actual path for xml
    company_df = pd.read_excel(input_file_path)

    for index, row in company_df.iterrows():
        try:
            company_name = row['NAME OF COMPANY']
            start_date = row['START DATE']
            end_date = row['END DATE']

            driver = initialize_driver(xml_file_path)
            if navigate_to_results(driver, company_name):
                if set_broadcast_period(driver):
                    click_xbrl_links(driver, company_name, start_date, end_date, xml_file_path)
            driver.quit()
        except Exception as e:
            print(f"An error occurred during the scraping process: {e}")
        finally:
            if driver:
                driver.quit()



##yvthtgyuewm,fgyuasgfioSUGFYUEWIFHHEGCJXTFHJDYJDGYJDTTYJTDJUDTYJDRYKXDRRY5WERTDHREWRTHWWDEWGRWQSEWETSRRYEWQR54654324HGTRGHJL