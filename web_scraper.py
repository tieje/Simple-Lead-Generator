#scratch file
driver.find_element(By.ID, "cheese")
cheese = driver.find_element(By.ID, "cheese")
cheddar = cheese.find_elements_by_id("cheddar")

connect to search engine

from selenium import webdriver
driver = webdriver.Chrome()
driver.get('https://duckduckgo.com/')

driver.get("https://google.co.in / search?q = geeksforgeeks")

# duckduckgo search xpath

//*[@id="search_form_input_homepage"]

import re
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import sys

chromedriver_args = [
    "--window-size=1920x1080",
    "start-maximized",
    "disable-popup-blocking",
]


chrome_options = Options()
chrome_options.headless=False
for i in chromedriver_args:
    chrome_options.add_argument(i)

contact_list = [
    'contact us',
    'Contact Us',
    'CONTACT US',
    'contact',
    'CONTACT',
    'Contact']

driver = webdriver.Chrome(options=chrome_options)
driver.get("https://duckduckgo.com/")
search_box = driver.find_element_by_xpath('//*[@id="search_form_input_homepage"]')
search_box.send_keys("JEDER VALUATION CONSULTANTS INC connecticut")
search_button = driver.find_element_by_xpath('//*[@id="search_button_homepage"]')
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
    "(\d{3}[-\.\s]\d{3}[-\.\s]\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]\d{4})", driver.page_source
)
email_list = re.findall(
    "[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", driver.page_source
)
contact_url = driver.current_url

if phone_list == [] and email_list == []:
    print("The script could not scrape anything. Manual website review is required. E.g. there might be a formal request form that needs to be filled out manually.\n")
    print("\nTo continue running the script after manual review, type 'c' and press Enter. To stop the script, press Enter")
    continue_script = input('>')
    if continue_script == 'c':
        pass
    else:
        sys.exit()

print(phone_list)
print(email_list)
print(contact_url)

