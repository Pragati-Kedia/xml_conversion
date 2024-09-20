import os
import time
import pyautogui
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import re
import requests
import xml.etree.ElementTree as ET
import traceback
from selenium.webdriver.common.keys import Keys

options = Options()
   
options.add_argument("--headless")

options.add_argument("--start-maximized")
options.add_argument("--incognito")

# Convert date like 'JQ2024-2025' to a comparable tuple (year, quarter)
def parse_quarter_date(period):
    quarter = period[:2]  # Extract "JQ", "MQ", etc.
    year = int(period[2:6])  # Extract year part (e.g., 2024)
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

def XML_extraction (Security_code,Stock_Name,start_period, end_period, Save_Folder):
    Top_URL = "https://www.bseindia.com/corporates/Comp_Resultsnew.aspx"
    driver = webdriver.Chrome(options=options)
    driver.get(Top_URL)
    print(Stock_Name)
    try:
        Security_Search = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID,"ContentPlaceHolder1_SmartSearch_smartSearch")))
        Security_Search.clear()    
        Security_Search.send_keys(Security_code)
    
        li_element = driver.find_element(By.XPATH,f"//li[contains(@onclick, \"'{Security_code}'\")]")
        li_element.click()
    
        #li_element = driver.find_element(By.XPATH,"//li[@onclick=\"liclick('532454','BHARTI AIRTEL LTD')\"]")
        #li_element.click()
    
    
        dropdown = driver.find_element(By.ID,"ContentPlaceHolder1_broadcastdd")
        select = Select(dropdown)
        select.select_by_value("7")
    
        Submit_button = driver.find_element(By.ID,"ContentPlaceHolder1_btnSubmit")
        Submit_button.click()
        try:
            start_year, start_quarter = parse_quarter_date(start_period)
            end_year, end_quarter = parse_quarter_date(end_period)
            xbrl_links = driver.find_elements(By.XPATH, f"//td[text()='{Security_code}']/following-sibling::td[6]//a")
            periods = driver.find_elements(By.XPATH, f"//td[text()='{Security_code}']/following-sibling::td[3]//a")
            time.sleep(1)
            print(f"Total XBRL files located for {Stock_Name}:{len(xbrl_links)}")
            clicked_count = 0
            skipped_count = 0
            for i in range(len(xbrl_links)):
                link = xbrl_links[i]
                period = periods[i].text
                print(period)
                # Parse the period to get the year and quarter
                period_year, period_quarter = parse_quarter_date(period)
                # Check if the first two characters of the period text are JQ, MQ, DQ, or SQ
                if period[:2] not in ["MQ", "JQ", "SQ", "DQ"]:
                    print(f"Skipping row with period: '{period}' because it does not start with JQ, MQ, DQ, or SQ")
                    skipped_count += 1
                    continue
                # Check if the period is within the range of start_period and end_period
                elif is_within_range(period_year, period_quarter, start_year, start_quarter, end_year, end_quarter):
                    print(f"Scraping data for period: '{period}'")
                    # Define the quarter name based on the period abbreviation
                    quarter_name_map = {"JQ": "Q1", "MQ": "Q4", "SQ": "Q2", "DQ": "Q3"}
                    quarter_name = quarter_name_map.get(period[:2], "")
            
                    # Create the custom file name using the new format
                    custom_file_name = f"{period_year}-{period_year+1}_{quarter_name}_{Stock_Name}.xml" 
                    custom_file_path = os.path.join(Save_Folder, custom_file_name)
                    time.sleep(1)
                    # Scroll the link into view before attempting to click
                    driver.execute_script("arguments[0].scrollIntoView(true);", link)
                    time.sleep(1)  # Ensure the scrolling action completes before the click
                    try:
                        link.click()
                    except Exception as e:
                        print(f"Error clicking link: {e}")
                        continue  # If it fails to click, move to the next row
                    # Switch to the new window opened by the click event
                    driver.switch_to.window(driver.window_handles[-1])
                    current_url = driver.current_url
                    print(current_url)
                    
                    # Re-fetch the URL to make sure it's up-to-date
                    driver.get(current_url)
                    
                    # Get the page source and save the XML file
                    xml_content = driver.page_source
            
                    #with open(custom_file_path, 'wb') as file:
                       #file.write(xml_content.encode('utf-8'))
                    
                    with open(custom_file_path, 'w') as file:
                        file.write(xml_content)
                        
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(1)
                    clicked_count += 1
                else:
                        print(f"Skipping row with period: '{period}' as it is outside the date range.")
    
        except Exception as e:
            print(f"Error processing row: {e}")
    except Exception as e:
        print(f"Error clicking XBRL links for {Stock_Name}: {e}")
    finally:
        print(f"Total XBRL files downloaded for {Stock_Name}: {clicked_count}")
        print(f"Total XBRL files not downloaded as they are not quarterly files for {Stock_Name}: {skipped_count}")

# Main script execution
if __name__ == "__main__":
    input_file_path = r'D:\BSE\input\SampleInput.xlsx'  # Replace with actual path for input
    xml_file_path = r"D:\BSE\xml"
    #Top_URL = "https://www.bseindia.com/corporates/Comp_Resultsnew.aspx"
    company_df = pd.read_excel(input_file_path)
 
     
    for index, row in company_df.iterrows():
        try:
            Security_code = str(row['Code'])
            Stock_Name = str(row['SYMBOL'])
            start_period = row['START DATE']
            end_period = row['END DATE']
            folder_name = Stock_Name
            Save_Folder = os.path.join(xml_file_path, folder_name)
            os.makedirs(Save_Folder)

            XML_extraction(Security_code,Stock_Name, start_period, end_period, Save_Folder)
            
        except Exception as e:
            print(f"An error occurred during the scraping process: {e}")
    print("test complete")
            


