from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from modules.helpers import *
from modules.workbook import write_to_workbook
from modules.const.settings import SETTINGS
from modules.const.colors import fore

import time
import json
import xlsxwriter

def scrape(args, workbook_name):
    '''
    Scrapes the results and puts them in the excel spreadsheet.

    Parameters:
            args (object): CLI arguments
    '''
    if args.pages is not None:
        SETTINGS["PAGE_DEPTH"] = args.pages
    SETTINGS["BASE_QUERY"] = args.query
    SETTINGS["PLACES"] = args.places.split(',')


    # Created driver and wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 15)

    # Set main box class name
    BOX_CLASS = "Nv2PK"
    NAME_CLASS = "qBF1Pd"
    URL_CLASS = "hfpxzc"
    BOX_DETAILS_CLASS = "W4Efsd"
    SEARCH_RESULTS_CLASS = "ecceSd"

    # Initialize workbook / worksheet
    workbook = xlsxwriter.Workbook(workbook_name)
    worksheet = workbook.add_worksheet(SETTINGS["WORKSHEET_NAME"])

    # Headers and data
    data = {
        "name": "",
        "phone": "",
        "address": "",
        "website": "",
        "googe_profile_page": "",
        "email": ""
    }
    headers = generate_headers(args, data)
    print_table_headers(worksheet, headers)

    # Start from second row in xlsx, as first one is reserved for headers
    row = 1

    # Remember scraped addresses to skip duplicates
    addresses_scraped = {}

    start_time = time.time()

    for place in SETTINGS["PLACES"]:
        # Go to the index page
        driver.get(SETTINGS["MAPS_INDEX"])

        # Build the query string
        query = "{0} {1}".format(SETTINGS["BASE_QUERY"], place)
        print(f"{fore.GREEN}Moving on to {place}{fore.RESET}")

        # Fill in the input and press enter to search
        q_input = driver.find_element(By.NAME, "q")
        q_input.send_keys(query, Keys.ENTER)
        
        # Wait for the results page to load. If no results load in 10 seconds, continue to next place
        try:
            w = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, BOX_CLASS))
            )
        except:
            continue

        index = 1
        # Loop through pages and results
        for _ in range(0, SETTINGS["PAGE_DEPTH"]):
            
            # Get all the results boxes
            boxes = driver.find_elements(By.CLASS_NAME, BOX_CLASS)

            # Loop through all boxes and get the info from it and store into an excel
            for box in boxes:
                # Just get the values, add only after we determine this is not a duplicate (or duplicates should not be skiped)
                name = box.find_element(By.CLASS_NAME, NAME_CLASS).find_element(By.XPATH, ".//span[1]").text
                box_details = box.find_elements(By.CLASS_NAME, BOX_DETAILS_CLASS)
                address_line = box_details[2].text.split("·")
                if len(address_line) >= 2:
                    address = address_line[1]

                scraped = name in addresses_scraped

                if scraped and args.skip_duplicate_addresses:
                    print(f"{fore.WARNING}Skipping {name} as duplicate by name{fore.RESET}")
                else:

                    if scraped:
                        addresses_scraped[name] += 1
                        print(f"{fore.WARNING}Currently scraping on{fore.RESET}: {name}, for the {addresses_scraped[name]}. time")
                    else:
                        addresses_scraped[name] = 1
                        print(f"{fore.GREEN}Currently scraping on{fore.RESET}: {name}")

                    # Only if user wants to get the URL to, get it
                    if args.scrape_website:
                        url = box.find_element(By.CLASS_NAME, URL_CLASS).get_attribute("href")
                        data["googe_profile_page"] = url

                    data["name"] = name
                    data["address"] = address
                    
                    # If additional output is requested
                    if args.verbose:
                        print(json.dumps(data, indent=1))

                    write_data_row(worksheet, data, row)
                    row += 1

            # Go to next page
            search_results = driver.find_elements(By.CLASS_NAME, SEARCH_RESULTS_CLASS)[1]
            driver.execute_script("arguments[0].scroll(0, {0});".format(index * 900), search_results)
            
            if (is_end_of_list(driver)):
                print(f"{fore.WARNING} end of list after {index} iterations")
                break
            else:
                index += 1
                print("-------"+str(index)+"------------")

    workbook.close()
    driver.close()

    end_time = time.time()
    elapsed = round(end_time-start_time, 2)
    print(f"{fore.GREEN}Done. Time it took was {elapsed}s{fore.RESET}")

def scrape_profile_page(driver, profile_page_url, row, workbook_name):

    URL_CLASS = "AeaXub"
    ICON_CLASS = "Liguzb"
    WEBSITE_ICON = "https://www.gstatic.com/images/icons/material/system_gm/1x/public_gm_blue_24dp.png"

    wait = WebDriverWait(driver, 15)

    # Go to profile_page_url
    driver.get(profile_page_url)

    # Wait for the results page to load. If no results load in 15 seconds, continue to next place
    try:
        w = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, URL_CLASS))
        )
    except:
        print(f"{fore.ERROR} line 158 Error fetching google_profile_page for {profile_page_url}")
        return

    # Get Icons
    icons = driver.find_elements(By.CLASS_NAME, ICON_CLASS)
    index = -1
    if len(icons) > 0:
        icons_src = [index for (index, icon) in enumerate(icons) if icon.get_attribute("src") == WEBSITE_ICON]
        if len(icons_src) > 0:
            index = icons_src[0]
    
    # Get website link
    if index > 0:
        data = driver.find_elements(By.CLASS_NAME, URL_CLASS)[index]
        website = data.text
    else:
        website = "No website found"
        print(f"{fore.ERROR} line 173 Error fetching website url for row {row}")

    # Write url to 4th column of row
    # time.sleep(30)
    write_to_workbook(website, row, 4, workbook_name)