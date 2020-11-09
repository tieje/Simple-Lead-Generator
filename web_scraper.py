import re
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import sys
import logging
import xlwings as xw
import time
from gui import 

""" Values
{
    "skip review checkbox": False,
    "emails checkbox": True,
    "phone numbers checkbox": True,
    "contact urls checkbox": True,
    "excel_chosen_listbox": [],
    "excel column": "A",
    "addendum": "",
}
"""


class Scraper:
    def contact_scraper(conn, **settings):
        print("reached contact scraper")
        # logging what actually happens
        logging.basicConfig(level=logging.INFO)
        logging.debug("Debugging is on")
        print(settings["emails checkbox"])

        # initializing chromedriver
        chromedriver_args = [
            "--window-size=1920x1080",
            "start-maximized",
            "disable-popup-blocking",
        ]

        chrome_options = Options()
        chrome_options.headless = False
        for i in chromedriver_args:
            chrome_options.add_argument(i)
        contact_list = [
            "contact us",
            "Contact Us",
            "CONTACT US",
            "contact",
            "CONTACT",
            "Contact",
        ]

        # initializing excel connection

        wb = xw.books["".join(settings["excel_chosen_listbox"])].sheets[0]
        # begin scraping on the second row
        counter = 2
        # establish the next three open columns. Assumes that the next three columns are open after the first blank
        free_column = 65
        cell_value = wb.range(chr(free_column) + "1").value
        while bool(cell_value):
            free_column += 1
            cell_value = wb.range(chr(free_column) + "1").value
        # Begin web scraping based on excel cell values in the chosen column
        driver = webdriver.Chrome(options=chrome_options)
        while bool(wb.range(settings["excel column"] + str(counter)).value):

            # defining variable cells
            email_cell = wb.range(chr(free_column) + str(counter))
            phone_cell = wb.range(chr(free_column) + str(counter + 1))
            contact_url_cell = wb.range(chr(free_column) + str(counter + 2))
            # stop scraping mechanism #1

            if conn.poll():
                break

            # connecting search engine

            driver.get("https://duckduckgo.com/")
            search_box = driver.find_element_by_xpath(
                '//*[@id="search_form_input_homepage"]'
            )

            search_box.send_keys(
                wb.range(settings["excel column"] + str(counter)).value
            )

            # main standard scraping procedure.
            search_button = driver.find_element_by_xpath(
                '//*[@id="search_button_homepage"]'
            )
            search_button.click()
            first_result = driver.find_element_by_xpath('//*[@id="r1-0"]')
            first_result.click()

            for i in contact_list:
                try:
                    print(i)
                    contact_us = driver.find_element_by_xpath('//*[text()="%s"]' % i)
                    print(contact_us)
                    contact_us.click()
                except:
                    try:
                        contact_us = driver.find_element_by_link_text(i)
                        contact_us.click()
                    except:
                        pass
                    contact_us = 0
                if contact_us != 0:
                    break

            if settings["phone numbers checkbox"]:
                phone_list = re.findall(
                    "(\d{3}[-\.\s]\d{3}[-\.\s]\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]\d{4})",
                    driver.page_source,
                )
                phone_cell.value = "\n".join(set([d.lower() for d in phone_list]))
            if settings["emails checkbox"]:
                email_list = re.findall(
                    "[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", driver.page_source
                )
                email_cell.value = "\n".join(set([d.lower() for d in email_list]))
            if settings["contact urls checkbox"]:
                contact_url = driver.current_url
                contact_url_cell = "\n".join(contact_url)

            if settings["skip review checkbox"] == False:
                if phone_list == [] and email_list == []:
                    print("entered review")
                    conn.send("Manual Review")
                    while conn.poll() == False:
                        time.sleep(0.1)
                    receipt = conn.recv()
                    print("The receipt is: " + receipt)
                    if receipt:
                        pass
                    else:
                        break

            # stop scraping mechanism #2
            if conn.poll():
                break

            print(counter)
            counter += 1
        conn.send(True)
        return None
