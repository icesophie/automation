# All import <package> statements must be at the top.



# If you need to add a new method, the method name should follow the criteria below:
# 1. Must start with 'request' or 'selenium'.
# 2. Must be followed by an underscore and 'SourceAbb' (which can be found in the Excel file).
# 3. Must end with a number.

# Examples:
# selenium_BURSA1
# request_BOJ1
# request_BOJ2


# if method == 'request':
#     pass
# elif method == 'selenium':
#     driver.quit()
# elif method == 'selenium_POM1':
#     driver.get()

import requests
from lxml import html
from urllib.parse import urljoin
import os
import shutil
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from datetime import datetime



#define function for chromedriver
def chromedriver():
    #specify the path to ChromeDriver
    chrome_driver_path = "C:\\Users\\nahmadshaarani\\OneDrive - Internet Securities, LLC\\Desktop\\ChromeDriver\\chromedriver129.exe"
    options = Options()
    options.add_experimental_option("detach", True)

    # Initialize the WebDriver (only when needed)
    driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)
    return driver

directory_name = "Downloaded Files"

os.makedirs(directory_name, exist_ok=True)

files_to_skip = ['.py']

# Ensure the download directory exists
def delete_existing_files(directory_name, target_filenames):
    if os.path.exists(directory_name):
        # Iterate only through target filenames
        for target_filename in target_filenames:
            file_path = os.path.join(directory_name, target_filename)
            if os.path.exists(file_path):  # Check if the target file exists in the directory
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Remove file
                    print(f'Deleted file: {file_path}')
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f'Deleted directory: {file_path}')
            else:
                print(f'File not found, skipping: {file_path}')



def navigation(method, url, ref):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    # Default method
    # You use code from you previous task on request
    if method == 'request': # main request
        responses = requests.get(url, headers=headers, timeout=5)
        if responses.status_code == 200:
            tree = html.fromstring(responses.content)
            #find elements using xpath
            firstresult = tree.xpath(f'{ref}')
            if firstresult:
                result = firstresult[0]
                #Get the text content of the element
                result = result.text_content().strip()

        return result

    # You use code from you previous task on selenium
    elif method == 'selenium':  # main request
        driver = chromedriver()
        driver.get(url)
        time.sleep(3)
        targetedpost = driver.find_element(By.XPATH, f'{ref}')
        result = targetedpost.text  # Retrieve the text content
        driver.quit()
        return result

    elif method == 'request_POM1':
        responses = requests.get(url, headers=headers, timeout=5)
        #print(responses.status_code)
        if responses.status_code == 200:
            tree = html.fromstring(responses.content)
            # find elements using xpath
            firstresult = tree.xpath(f'{ref}')
            result = firstresult[0]
            if result:
                resultresponse = requests.get(result)
                if resultresponse.status_code == 200:
                    secondtree = html.fromstring(resultresponse.content)
                    latestpost = secondtree.xpath('//a[contains(text(), "Monthly Trade Report")]')
                    if latestpost:
                        # Get the text content of the element
                        exactlatestpost = latestpost[0]
                        exactlatestpost = exactlatestpost.text_content().strip()

        return exactlatestpost

    elif method == 'request_JPX1':
        responses = requests.get(url, headers=headers, timeout=5)
        if responses.status_code == 200:
            tree = html.fromstring(responses.content)
            months = [12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]  # List of months from December to January
            for month in months:
                firstresult = tree.xpath(f'{ref}{month:02d}")]/@href')
                if firstresult:
                    result = firstresult[0]
                    resulturl = urljoin(url, result)
                    return resulturl
                    break
                else:
                    continue
                    return None





    elif method == 'selenium_BTEI1':
        driver = chromedriver()
        driver.get(url)
        time.sleep(3)
        months = [12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]  # List of months from December to January
        for month in months:  # Loop from 12 (December) to 1 (January)
            latest_months = driver.find_elements(By.XPATH, f'{ref}{month:02d}")]')
            if latest_months:
                latest_month = latest_months[0]
                result = latest_month.text
                return result
                break
            else:
                continue
                return None

        driver.quit()


    elif method == 'request_JNTO1':
        responses = requests.get(url, headers=headers, timeout=5)
        file_path = 'marketingdata_outbound.pdf'
        if responses.status_code == 200:
            file_path = os.path.join(directory_name, "marketingdata_outbound.pdf")
            with open(file_path, 'wb') as file:
                file.write(responses.content)
                print("Downloaded successfully")
                download = True
                if download == True:
                    reader = PdfReader(file_path)
                    metadata = reader.metadata
                    modification_date_str = metadata.get("/ModDate")
                    if modification_date_str:
                        #Remove D: prefix and parse the date
                        modification_date_str = modification_date_str[2:]
                        if '+' in modification_date_str or '-' in modification_date_str:
                            modification_date_str = modification_date_str.split('+')[0].split('-')[0]

                        #Parse the datetime object
                        modification_date = datetime.strptime(modification_date_str, "%Y%m%d%H%M%S")

        else:
            print("Failed to download the file")

        return f"{modification_date}"

    elif method == 'selenium_JPEA1':
        driver = chromedriver()
        driver.get(url)
        time.sleep(3)
        months = [12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]  # List of months from December to January
        for month in months:  # Loop from 12 (December) to 1 (January)
            latest_months = driver.find_elements(By.XPATH, f'//section[contains(@class, "sokuhou")]//td[2]/a[contains(@href, "24{month:02d}")]')
            if latest_months:
                latest_month = latest_months[0]
                result = latest_month.get_attribute("href")
                driver.quit()
                return result

        driver.quit()
        return None



    elif method == 'request_MOF2':
        responses = requests.get(url, headers=headers, timeout=5)
        if responses.status_code == 200:
            tree = html.fromstring(responses.content)

            # find elements using xpath
            link = tree.xpath('//ol/li[2]/a[contains(text(), "The Final Seasonally")]/@href')
            if link:
                excel = link[0]
                excelurl = urljoin(url, excel)
                excel_response = requests.get(excelurl, headers=headers, timeout=5)
                if excel_response.status_code == 200:
                    os.makedirs(directory_name, exist_ok=True)  # Ensure the directory exists
                    file_path = os.path.join(directory_name, "percent.xlsx")
                    with open(file_path, 'wb') as file:
                        file.write(excel_response.content)
                        print("Downloaded successfully")
                        download_successful = True

                        if download_successful:
                            #load an existing workbook
                            excelworkbook = load_workbook(file_path)

                            #access the properties
                            properties = excelworkbook.properties

                            #get the modification date
                            modification_date = properties.modified

                            #check if the modification date is available
                            if modification_date:
                                formatted_date = modification_date.strftime("%Y-%m-%d-%H:%M:%S")
                else:
                    print("Failed to download the Excel file")
            else:
                print("Failed to retrieve the main page")


        return f"{formatted_date}"


    else:
        print('Method not found')



