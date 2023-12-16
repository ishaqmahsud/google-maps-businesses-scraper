from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support import expected_conditions as EC

from modules.helpers import *
from modules.const.settings import SETTINGS
from modules.const.colors import fore

import time
import json
import xlsxwriter
def parse_item(el, xpath):
    return el.find_element(By.XPATH,xpath)
def parse_items(el, xpath):
    return el.find_elements(By.XPATH, xpath)
def parse_attr(el, attr):
    return el.get_attribute(attr)
def wait_for_el(wait, xpath):
       return wait.until(EC.visibility_of_element_located((By.XPATH, xpath)) )

def scrape(args):
    '''
    Scrapes the results and puts them in the excel spreadsheet.

    Parameters:
            args (object): CLI arguments
    '''
    if args.pages is not None:
        SETTINGS["PAGE_DEPTH"] = args.pages
        print("pages set to -> ",  SETTINGS['PAGE_DEPTH'])
    SETTINGS["BASE_QUERY"] = args.query
    SETTINGS["PLACES"] = args.places.split(',')
    #supported formats: excel, json
    SETTINGS["FORMAT"] = "excel"
    if args.format is not None:
        SETTINGS["FORMAT"] = args.format
    FILE_NAME = "googlemaps_scraped_data"
    if args.filename is not None:
        FILE_NAME = args.filename.split(".")[0]
    # Created driver and wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 15)

    # Set main box class name
    BOX_CLASS = "//div[starts-with(@jsaction,'mouseover:pane')]"
    DETAIL_BOX = "//div[@data-omit-title]/../.././.."

    if SETTINGS['FORMAT'] == "excel":
        # Initialize workbook / worksheet
        workbook = xlsxwriter.Workbook(f'{FILE_NAME}.xlsx')
        worksheet = workbook.add_worksheet()

        # Headers and data
        data = {
            "name": "",
            "phone": "",
            "address": "",
            "website": "",
            "email": ""
        }
        headers = generate_headers(args, data)
        print_table_headers(worksheet, headers)
     
  
    # Start from second row in xlsx, as first one is reserved for headers
    row = 1

    # Remember scraped addresses to skip duplicates
    addresses_scraped = {}

    start_time = time.time()
    last_item_index = 0
    for place in SETTINGS["PLACES"]:
        # Go to the index page
        driver.get(SETTINGS["MAPS_INDEX"])

        # Build the query string
        query = f'{SETTINGS["BASE_QUERY"]} {place}'
        print(f"{fore.GREEN} Moving on to {place} {fore.RESET}")

        # Fill in the input and press enter to search
        q_input = driver.find_element(By.NAME,"q")
        q_input.send_keys(query, Keys.ENTER)
        
        # Wait for the results page to load. If no results load in 10 seconds, continue to next place
        try:
           wait_for_el(wait, BOX_CLASS)
        except:
            continue

        # Loop through pages and results
        scraped_data = []
        for pageNo in range(0, SETTINGS["PAGE_DEPTH"]):
             # Headers and data
            
            # Get all the results boxes
            boxes = driver.find_elements(By.XPATH,BOX_CLASS)
            last_item_index = len(boxes)-1
            last_item_item =  boxes[last_item_index]
           
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);",last_item_item)
                if (pageNo+ 1) < SETTINGS["PAGE_DEPTH"]:
                    continue
            except Exception as e:
                    pass
            # Loop through all boxes and get the info from it and store into an excel
            for box in boxes:
                    idata = {
                "name": "",
                "phone": "",
                "address": "",
                "website": "",
                "email": ""
                        }
                    # Just get the values, add only after we determine this is not a duplicate (or duplicates should not be skiped)
                    el_name = parse_item(box, ".//a")
                    name = parse_attr(el_name, "aria-label")
                   
                    try:
                        el_name.click()
                    except:
                         driver.execute_script("arguments[0].scrollIntoView(true);",box)
                         el_name.click()
                    wait_for_el(wait, DETAIL_BOX)
                    el_detail = parse_item(driver, DETAIL_BOX)
                    website = ""
                    try:
                        web_xpath = "//a[starts-with(@aria-label, 'Website: ')]"
                        wait_for_el(wait, adress_xpath)
                        el_website = parse_item(box, web_xpath)#parse_item(box, ".//a[@data-value='Website']")
                        website = parse_attr(el_website, "aria-label").replace('Website: ', '')
                    except: 
                        pass
                    address = ""
                    try:
                        adress_xpath = ".//button[@data-item-id='address']"
                        wait_for_el(wait, adress_xpath)
                        el_address = parse_item(driver,adress_xpath )
                        address = parse_attr(el_address,"aria-label")
                    except Exception as e:
                        print("address error -> " )
                    phone = ""
                    try:
                        phone_xpath = ".//button[starts-with(@data-item-id,'phone:tel:')]"
                        wait_for_el(wait, phone_xpath)
                        el_phone = parse_item(driver,phone_xpath )
                        phone = parse_attr(el_phone,"aria-label")
                    except Exception as e:
                      print("phone error -> ")
                    email = ""
                    web, email = get_website_data(website.strip())
                    if email is not None:
                        email = ""
                    idata["name"] = name
                    if args.scrape_website:
                        idata['website'] = website
                    idata["address"] = address.replace("Address: ", "")
                    idata["phone"] = phone.replace("Phone: ", "")
                    idata['email'] = email
                    print(idata["name"])
                    # If additional output is requested
                    if args.verbose:
                        print(json.dumps(idata, indent=1))
                    if SETTINGS['FORMAT'] == "excel":
                        write_data_row(worksheet, idata, row)
                    row += 1
                    if SETTINGS['FORMAT'] == "json":
                        scraped_data.append(idata)
     
            time.sleep(1)
        # Writing JSON data to a file
        if SETTINGS['FORMAT'] == "json":
            with open(f'{FILE_NAME}.json', 'w') as json_file:
                json.dump(scraped_data, json_file, indent=4)  
        print("-------------------")
    if SETTINGS['FORMAT'] == "excel":
        workbook.close()
    driver.close()

    end_time = time.time()
    elapsed = round(end_time-start_time, 2)
    print(f"{fore.GREEN}Done. Time it took was {elapsed}s{fore.RESET}")