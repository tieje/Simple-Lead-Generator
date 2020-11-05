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
        - will print which business is being searched and the URLs that are being scraped
        - will input the scraped data into the excel sheet
        - I'll need to run this as a subprocess so that the user can maintain control of the gui.
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
- choose theme
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
"""

import PySimpleGUI as sg
import detect_excels as des
import web_scraper as ws
import subprocess

sg.theme("DarkAmber")

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
top_results_inputtext = "top results"
addendum_inputtext = "addendum"

# Initializing variables
detected_excels = des.DetectExcels.detect()
"""
sanitizing functions
case_not_integer = "An integer was not entered. Please Enter an integer."
def sanitize_integer_input(tester, max_numb):
    while tester.isdigit() == False:
        print("\n" + case_1_not_integer + "\n")
        tester = input(">")
    while int(tester) > max_numb or int(tester) < 1:
        print("\n" + case_2_not_acceptable + "\n")
        tester = input(">")
    return tester

"""

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
                                    "Phone #s", default=True, key=phone_numbers_checkbox
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
                                sg.Listbox(
                                    values=detected_excels,
                                    auto_size_text=True,
                                    size=(40, 10),
                                    key=excel_chosen_listbox,
                                )
                            ],
                            [
                                sg.Text(
                                    "Excel Column to Scrape Data",
                                    pad=((0, 0), (10, 5)),
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
                                ),
                            ],
                            [
                                sg.Button(
                                    button_text="Stop All Scraping",
                                    pad=((0, 0), (62, 5)),
                                    visible=True,
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
                                    visible=True,
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

while True:
    event, values = window.read()
    print("The event is: " + str(event))
    print("The value is: " + str(values))
    if event == sg.WINDOW_CLOSED or event == "Quit":
        break
    if event == detect_excel_button:
        print("Triggerring excel detection function. ")
        detected_excels = des.DetectExcels.detect()
        print(detected_excels)
        window[excel_chosen_listbox].update(values=detected_excels)
    if event == begin_button:
        print("Triggerring Begin Scraping function.")

    if event == continue_button:
        print("Triggerring Continue function.")
        window[continue_button].update(visible=False)
    if event == stop_button:
        print("Triggerring Stop Scraping.")
window.close()
