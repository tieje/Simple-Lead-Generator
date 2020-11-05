import selenium
import xlwings as xw
import os
import re
import sys

class ScrapeEmails:
    def __init__(chosen_book, search_addend, search_top):
        self.chosen_book = chosen_book
        self.search_addend = search_addend
        self.search_top = search_top
    def controlla(self):
        self.setup_browser()
    def setup_browser(self):
        browser = webdriver.Chrome()
        b
if __name__ == "__main__":
    # messages
    step_1_which_instance = 'The following excel files are currently open. Type the corresponding integer and press Enter for the file that you would like to use to scrape.'
    step_1a_no_excel_instance = 'Zero instances of excel are open. Please open an excel sheet and run the script again.'
    step_2_localization = 'Type the abbreviation of the state we want to localize the search for, then press Enter. This part is actually added at the end of each search request so you could essentially type anything. If no localizations or addendums are desired, press Enter.'
    step_3_how_many = 'How many of the top search results would you like to scrape from? Type an integer from 1 to 5 and press Enter. OR press Enter to search only the top result.'
    case_1_not_integer = 'An integer was not entered. Please Enter an integer.'
    case_2_not_acceptable = 'An unacceptable integer was entered. Please Enter acceptable integer.'

    # input logic for sanitizing integer input

    def sanitize_integer_input(tester,max_numb):
        while tester.isdigit() == False:
            print('\n'+case_1_not_integer+'\n')
            tester = input('>')
        while int(tester) > max_numb or int(tester) < 1:
            print('\n'+case_2_not_acceptable+'\n')
            tester = input('>')
        return tester

    # connecting to excel sheet
    if str(xw.apps) == 'Apps([])':
        print('\n'+step_1a_no_excel_instance+'\n')
        sys.exit()

    print(step_1_which_instance + '\n')

    files = str(xw.books)

    listit = files.split(',')
    if len(listit) > 1:

        excel_files = []
        for i in range(len(listit)):
            excel_files.append(str(i+1) + ' ===> ' + re.search('<(Book) \[.*\]>', listit[i], re.I).group().lstrip('<Book [').rstrip(']>'))
        for x in excel_files:
            print(x)

        unsanitary_chosen_book = input('>')

        chosen_book = excel_files[int(sanitize_integer_input(unsanitary_chosen_book,len(excel_files)))-1]
    else:
        chosen_book = re.search('<(Book) \[.*\]>', files, re.I).group().lstrip('<Book [').rstrip(']>')
    # step 2: addendums for localized searches
    print('\n'+step_2_localization+'\n')
    search_addend = input('>')
    # step 3: Number of top searches
    print('\n'+step_3_how_many+'\n')
    unsanitary_search_top = input('>')
    search_top = sanitize_integer_input(unsanitary_search_top,5)
    # max is top five searches

    ScrapeEmails(chosen_book, search_addend, search_top).controlla()