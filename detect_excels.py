import xlwings as xw
import re


class DetectExcels:
    # connecting to excel sheet
    def detect():
        case_no_excel_instance = "Zero instances of excel are open."
        if str(xw.apps) == "Apps([])":
            return [case_no_excel_instance]

        files = str(xw.books)

        listit = files.split(",")
        if len(listit) > 1:

            excel_files = []
            for i in range(len(listit)):
                excel_files.append(
                    re.search("<(Book) \[.*\]>", listit[i], re.I)
                    .group()
                    .lstrip("<Book [")
                    .rstrip("]>")
                )
        else:
            excel_files = [
                (
                    re.search("<(Book) \[.*\]>", files, re.I)
                    .group()
                    .lstrip("<Book [")
                    .rstrip("]>")
                )
            ]

        return excel_files  # list of excel_files
