import re
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import sys

""" Values
{
    "skip review checkbox": False,
    "emails checkbox": True,
    "phone numbers checkbox": True,
    "contact urls checkbox": True,
    "excel_chosen_listbox": [],
    "excel column": "A",
    "top results": "1",
    "addendum": "",
}
"""


class Scraper:
    def contact_scraper():
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

        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://duckduckgo.com/")
        search_box = driver.find_element_by_xpath(
            '//*[@id="search_form_input_homepage"]'
        )
        search_box.send_keys("JEDER VALUATION CONSULTANTS INC connecticut")
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
        phone_list = re.findall(
            "(\d{3}[-\.\s]\d{3}[-\.\s]\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]\d{4})",
            driver.page_source,
        )
        email_list = re.findall(
            "[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", driver.page_source
        )
        contact_url = driver.current_url

        if phone_list == [] and email_list == []:
            print(
                "The script could not scrape anything. Manual website review is required. E.g. there might be a formal request form that needs to be filled out manually.\n"
            )
            print(
                "\nTo continue running the script after manual review, type 'c' and press Enter. To stop the script, press Enter"
            )
        return True
        print(phone_list)
        print(email_list)
        print(contact_url)
