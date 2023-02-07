import requests
import re
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By

def generate_headers(args, example_dict):
    '''
    Generates headeers from the data dictionary by capitalizing it's keys.

    Parameters:
            args (object): Object containging CLI arguments passed as they can affect which columns are needed
            example_dict (dict): Data dictionary with keys

    Returns:
            list (list): List of capitalized strings representing headers
    '''
    headers = example_dict.keys()
    if not args.scrape_website:
        del example_dict["website"]

    headers = example_dict.keys()
    
    return [header.capitalize() for header in headers]

def print_table_headers(worksheet, headers):
    '''
    Writes headers to the worksheet.

    Parameters:
            worksheet (worksheet object): Worksheet where headsers should be written
            headers (list): List of headers to vrite
    '''
    col = 0
    for header in headers:
        worksheet.write(0, col, header)
        col += 1

def write_data_row(worksheet, data, row):
    '''
    Writes data dictionary to row.

    Parameters:
            worksheet (worksheet object): Worksheet where data should be written
            data (dict): Dictionary containing data to write
            row (int): No. of row to write to
    '''
    col = 0
    for key in data:
        worksheet.write(row, col, data[key])
        # write_data_row_col(worksheet, data[key], row, col)
        col += 1

def write_data_row_col(worksheet, data, row, col):
    '''
    Writes data to row-col.

    Parameters:
            worksheet (worksheet object): Worksheet where data should be written
            data (dict): Data to write
            row (int): No. of row to write to
            col (int): No. of col to write to
    '''
    worksheet.write(row, col, data)

def get_website_data(url):
    '''
    Returns the website URL and email address found in HTML
    code got from the URL.

    Parameters:
            url (string): URL to send the request to
    '''
    try:
        if url is not None:
            response = requests.get(url, allow_redirects=True, timeout=10)

            # Get the url
            url_retrieved = response.url
            content = response.content.decode("utf-8")
            soup = BeautifulSoup(content, 'html.parser')
            
            # Get emails recursively
            emails = []
            if url_retrieved is not None:
                q = ["contact","about"]
                print(f"Looking for emails in {url_retrieved}")
                emails = find_emails(content, soup, 0, q, [])
                emails = list(dict.fromkeys(emails))

            return url_retrieved, emails
        else:
            return None, None
    except:
        return None, None

def find_emails(content, base_soup, i, queries=[], found=[]):
    '''

    '''
    if i < len(queries) and content is not None:
        # Get the emails with regex
        soup = BeautifulSoup(content, 'html.parser')
        body = soup.find('body')
        html_text_only = body.get_text()
        match = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', html_text_only)

        # Removes duplicate values
        if match is not None:
            found = found + match

        # Advance to next page
        links = base_soup.find_all('a')
        next_page_url = None
        for link in links:
            curr_url = link.get("href")
            if curr_url is not None and queries[i] in curr_url:
                next_page_url = curr_url
                print(f"NPU found {next_page_url}")
                break

        cont = None
        if next_page_url is not None:
            try:
                response = requests.get(next_page_url, allow_redirects=True, timeout=10)
                cont = response.content.decode("utf-8")
                print(f"NPU: Looking for emails in {next_page_url}")
            except:
                print("Error occurred while looking for emails in NPU")
                cont = None

        return find_emails(cont, base_soup, i + 1, queries, found)
    else:
        return found

def is_end_of_list(driver):
    END_OF_LIST_CLASS = "PbZDve"
    END_OF_LIST_DIV_TEXT = "You've reached the end of the list."

    try:
        end_of_list_div = driver.find_element(By.CLASS_NAME, END_OF_LIST_CLASS)
    except:
        # end_of_list_div is not present
        end_of_list_div = None
    finally:
        # end_of_list_div is present
        return end_of_list_div and (end_of_list_div.text == END_OF_LIST_DIV_TEXT)