from modules.scraper import scrape
from modules.scraper import scrape_profile_page
from modules.workbook import count_entries_in_workbook
from modules.workbook import extract_column_from_row
from modules.cliargs import parse_cliargs

from selenium import webdriver

if __name__ == "__main__":
    args = parse_cliargs()
    workbook_name = "{}-{}.xlsx".format(args.places, args.query)
    scrape(args, workbook_name)
    num_rows = count_entries_in_workbook(workbook_name)
    if num_rows > 0:
        driver = webdriver.Chrome()
        for row in range(2, num_rows+1):
            google_profile_page = extract_column_from_row(row, 5, workbook_name)
            scrape_profile_page(driver, google_profile_page, row, workbook_name)
        driver.close()


