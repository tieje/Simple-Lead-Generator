"""
The main search logic has been completed. Time to figure out pysimpleGUI again. I'll be referring to the documentation and my old code.

Modules that need to be built:

- Detect open excel sheets module
    + button
    - wiring
    - connection to feedback window
- choose which detected excel sheet to use from a list of excel sheets
    + scrollable column
    + listbox containing list of detected excel sheets
    + wiring
    - feedback window to confirm choice
- input text of which column to perform searches for
    + inputtext()
    - wiring
- addendum/localization window
    + input text window
    - wiring
- number of top search results to scrape for
    + inputtext() with a default of 1
        + specify "Max of '5'"
    - wiring
- checkmark of what to scrape for
    - note that these will be placed in the first open column header on row 1 of the excel sheet
    + email
    + phone numbers
    + contact URL
    - wiring
- "Begin Scraping" button
    + created button
    - calls on the web scraping module
        + will print which business is being searched and the URLs that are being scraped
        + will input the scraped data into the excel sheet
        + I'll need to run this as a subprocess so that the user can maintain control of the gui.
    - 
- "Stop Scraping" button
    + created button
    - cuts the scraping session
- Manual review
    + created Feedback window
    - feedback window shows if manual review is needed in red text
    - press "Continue Scraping" button when manual review is finished
    x pause scraping must not be visible when Continue Scraping is visible
        # pause button has been removed because create and write abilities are not always available.
    - Continue Scraping must only be visible when manual review is triggered


End-User experience
x checkmark that says show "help" text
+ feedback window
+ choose theme
- choose font
- organize logically
    - organize in steps
    - clearly marked areas
        - use frames to guide the user's eyes
            - it might look clunkier but it will be more obvious
- pyinstaller executable

Future
- could be used to scrape data for anything, not just contacts
- right now it's just a contact scraper

The left is going to be reserved for settings.
The right is going to be reserved for the feedback and status conditions.

1. Help text checkmark
    + no need for show help because the Frame has tooltips option
2. Detect Excel sheets button
3. Choose which excel sheet to use.
4. Input text about which column to perform search on.

                                [
                                    sg.Text(
                                        'Number of Top Results to Scrape (Max is "5")',
                                        pad=pad_text_association,
                                    ),
                                ],
                                [
                                    sg.InputText(
                                        default_text="1",
                                        do_not_clear=True,
                                        size=(2, 1),
                                        justification="center",
                                        key=top_results_inputtext,
                                    )
                                ],

"""

import PySimpleGUI as sg
import detect_excels as des
import multiprocessing as mp
import re
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import logging
import xlwings as xw
import time

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
# message receiver from web scraper


def start_comm():
    print("entered comms")
    global begin_scrape_child
    begin_scrape_child.send("Manual Review")
    print("parent reception " + str(begin_scrape_parent.poll()))
    while True:
        if begin_scrape_child.poll():
            break


def contact_scraper(**settings):
    global begin_scrape_child
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

        if begin_scrape_child.poll():
            break

        # connecting search engine

        driver.get("https://duckduckgo.com/")
        search_box = driver.find_element_by_xpath(
            '//*[@id="search_form_input_homepage"]'
        )

        search_box.send_keys(wb.range(settings["excel column"] + str(counter)).value)

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
                start_comm()
                receipt = begin_scrape_child.recv()
                print("The receipt is: " + str(receipt))
                if receipt:
                    pass
                else:
                    break

        # stop scraping mechanism #2
        if begin_scrape_child.poll():
            break

        print(counter)
        counter += 1
    begin_scrape_child.send(True)
    return None


pad_text_association = ((0, 0), (10, 5))
pad_button = ((0, 0), (10, 10))

# Key values
#   Button/Event Keys
detect_excel_button = "detect excel button"
begin_button = "begin button"
pause_button = "pause button"
continue_button = "continue button"
stop_button = "stop button"
#   Values Keys
skip_manual_review_checkbox = "skip review checkbox"
emails_checkbox = "emails checkbox"
phone_numbers_checkbox = "phone numbers checkbox"
contact_urls_checkbox = "contact urls checkbox"
excel_chosen_listbox = "excel_chosen_listbox"
excel_column_inputtext = "excel column"
# top_results_inputtext = "top results"
addendum_inputtext = "addendum"
choose_this_file = "chosen file"

# Initializing variables
detected_excels = des.DetectExcels.detect()

# Initializing pipes

begin_scrape_parent, begin_scrape_child = mp.Pipe()

sg.theme("DarkAmber")

layout = [
    [
        sg.Column(
            layout=[
                [
                    sg.Frame(
                        title="Settings",
                        layout=[
                            [
                                sg.Checkbox(
                                    "Skip Manual Review",
                                    default=False,
                                    key=skip_manual_review_checkbox,
                                ),
                            ],
                            [
                                sg.Text(
                                    "Scrape for the checkmarked boxes:",
                                    pad=pad_text_association,
                                )
                            ],
                            [
                                sg.Checkbox(
                                    "Emails", default=True, key=emails_checkbox
                                ),
                                sg.Checkbox(
                                    "Phone #s",
                                    default=True,
                                    key=phone_numbers_checkbox,
                                ),
                                sg.Checkbox(
                                    "Contact URLs",
                                    default=True,
                                    key=contact_urls_checkbox,
                                ),
                            ],
                            [
                                sg.Button(
                                    button_text="Detect Open Excel WorkBooks",
                                    pad=pad_button,
                                    key=detect_excel_button,
                                )
                            ],
                            [
                                sg.Text(
                                    "Pick a document to scrape",
                                    pad=pad_text_association,
                                )
                            ],
                            [
                                sg.Listbox(
                                    values=detected_excels,
                                    auto_size_text=True,
                                    size=(40, 10),
                                    key=excel_chosen_listbox,
                                )
                            ],
                            [
                                sg.Button(
                                    button_text="Choose This File",
                                    pad=pad_button,
                                    key=choose_this_file,
                                )
                            ],
                            [
                                sg.Text(
                                    "Excel Column to Scrape Data",
                                    pad=pad_text_association,
                                )
                            ],
                            [
                                sg.InputText(
                                    default_text="A",
                                    do_not_clear=True,
                                    size=(3, 1),
                                    justification="center",
                                    key=excel_column_inputtext,
                                )
                            ],
                            [
                                sg.Text(
                                    "Add to End of Each Search (e.g. state/county)",
                                    pad=pad_text_association,
                                )
                            ],
                            [
                                sg.InputText(
                                    default_text="",
                                    do_not_clear=True,
                                    size=(40, 1),
                                    justification="center",
                                    key=addendum_inputtext,
                                )
                            ],
                            [
                                sg.Button(
                                    button_text="Begin Scraping",
                                    pad=pad_button,
                                    key=begin_button,
                                    visible=False,
                                ),
                            ],
                            [
                                sg.Button(
                                    button_text="Stop All Scraping",
                                    pad=((0, 0), (62, 5)),
                                    visible=False,
                                    key=stop_button,
                                ),
                            ],
                        ],
                        element_justification="center",
                    )
                ],
            ],
            element_justification="right",
            expand_x=True,
            expand_y=True,
        ),
        sg.Column(
            layout=[
                [
                    sg.Frame(
                        title="Script Feedback",
                        layout=[
                            [
                                sg.Output(
                                    echo_stdout_stderr=True,
                                    size=(100, 30),
                                )
                            ],
                            [
                                sg.Text(
                                    "Scraping Progress",
                                    pad=pad_text_association,
                                )
                            ],
                            [
                                sg.ProgressBar(
                                    max_value=100,
                                    orientation="horizontal",
                                    size=(30, 15),
                                    style="default",
                                    relief="RELIEF_RIDGE RELIEF_GROOVE RELIEF_SOLID",
                                    pad=((10, 10), (0, 10)),
                                    bar_color=("#329932", "#979797"),
                                )
                            ],
                            [
                                sg.Button(
                                    button_text="Continue Scraping",
                                    pad=pad_button,
                                    visible=False,
                                    key=continue_button,
                                ),
                            ],
                        ],
                        element_justification="center",
                    )
                ]
            ],
            expand_x=True,
            expand_y=True,
        ),
    ],
]

window = sg.Window(
    "Simple Lead Generator by Francis IT Consulting",
    layout=layout,
    resizable=True,
    # finalize=True,
)
""" Values
settings = {
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


def communicado():
    print("entered communicado")
    global begin_scrape_parent
    global begin_scrape_child
    global window
    print("parent is: " + str(begin_scrape_parent.poll()))
    while True:
        if begin_scrape_parent.poll():
            scraper_message = begin_scrape_parent.recv()
            if scraper_message == "Manual Review":
                window[continue_button].update(visible=True)
                print(
                    "Scraping has been paused because no phone numbers or emails were scraped. Please manually review the website for contact forms. Press the continue button when ready."
                )
                break
            if scraper_message == True:
                print("Web Scraping Completed")
                window[begin_button].update(visible=True)
                window[stop_button].update(visible=False)
                p2.join()
                break
        event, values = window.read()
        print("The event is: " + str(event))
        print("The value is: " + str(values))
        if event == sg.WINDOW_CLOSED or event == "Quit":
            break
        if event == choose_this_file:
            print("You have chosen " + "".join(values["excel_chosen_listbox"]))
        if values["excel_chosen_listbox"] != []:
            window[begin_button].update(visible=True)
        if event == detect_excel_button:
            print("Detecting open Excel docs")
            detected_excels = des.DetectExcels.detect()
            print(detected_excels)
            window[excel_chosen_listbox].update(values=detected_excels)
        if event == begin_button:
            print("Scraping has begun. Do not close the Excel doc.")
            begin_scraping = mp.Process(
                target=contact_scraper,
                kwargs=values,
            )
            begin_scraping.start()
            window[begin_button].update(visible=False)
            window[stop_button].update(visible=True)
        if event == continue_button:
            print("Continue Scraping")
            begin_scrape_parent.send(True)
            window[continue_button].update(visible=False)
        if event == stop_button:
            print("Stop All Scraping.")
            begin_scrape_parent.send(False)
            window[begin_button].update(visible=True)
            window[stop_button].update(visible=False)
            begin_scraping.join()


if __name__ == "__main__":
    communicado()
    window.close()

"""
I'm going to give up on pipes. 
The reason for pipe failure is probably because of the framework I'm using for the gui.
Well I guess also the language to some extent.

remove manual review. it's too difficult for a simple gui application because there needs to be a better controller and the gui needs to be more complex to handle something like pausing without crashing.
It's supposed to be a simple lead generator anyways.
The manual work can be done by hand if needed.
If anything, I should focus on an email exporter because that's the main idea anyways.


"""
