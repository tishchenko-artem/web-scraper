"""Grab information about the agencies. Save the information about agencies as Excel files and download PDF files."""

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
import time
import os


browser_lib = Selenium()
lib = Files()
filesystem = FileSystem()
pdf = PDF()


def open_the_website(url):
    path = os.path.dirname(os.path.abspath(__file__))
    browser_lib.open_chrome_browser(url=url, preferences={
        "download.default_directory": path + "\output",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True})


def click_divein_button(locator):
    browser_lib.click_element("css:" + locator)


def grab_data_about_agencies(t_locator_template, a_locator_template):
    '''Function receives two locator templates: for titles and for amounts. Function grabs title, amount of spending from the main page.'''

    title_locators = []
    amount_locators = []

    for x in range(1, 10):
        template_first_part = "css:" + t_locator_template[:43] + str(x)
        for y in range(1, 4):
            template_second_part = t_locator_template[43:61] + str(y) + t_locator_template[61:]
            title_locators.append(template_first_part + template_second_part)
            if len(title_locators) >= 26:
                break

    for x in range(1, 10):
        template_first_part = "css:" + a_locator_template[:43] + str(x)
        for y in range(1, 4):
            template_second_part = a_locator_template[43:61] + str(y) + a_locator_template[61:]
            amount_locators.append(template_first_part + template_second_part)
            if len(amount_locators) >= 26:
                break

    browser_lib.wait_until_page_contains_element(title_locators[25])
    browser_lib.wait_until_page_contains_element(amount_locators[25])

    titles = [browser_lib.get_text(i) for i in title_locators]
    amounts = [browser_lib.get_text(i) for i in amount_locators]

    return titles, amounts


def write_agencies_data_to_excel(workbook_name, worksheet_name):
    '''Function calls grab_data_about_agencies function, takes data and writes to excel file.'''
    titles, amounts = grab_data_about_agencies(
        "#agency-tiles-widget > div > div:nth-child() > div:nth-child() > div > div > div > div:nth-child(2) > a > span.h4.w200",
        "#agency-tiles-widget > div > div:nth-child() > div:nth-child() > div > div > div > div:nth-child(2) > a > span.h1.w900")
    try:
        lib.create_workbook('output\{}'.format(workbook_name), fmt='xls')
        lib.rename_worksheet("Sheet", worksheet_name)
        lib.append_rows_to_worksheet([titles, amounts], worksheet_name)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def choose_one_of_agencie_and_scrape_a_table(agency_locator, file_name):
    """Function goes to the agency page scrapes a table with all "Individual Investments" 
    and writes it to a new sheet in excel."""
    browser_lib.click_link("css:" + agency_locator)
    browser_lib.wait_until_page_contains_element("css:#investments-table-object_length > label", 15)
    browser_lib.click_element_when_visible("css:#investments-table-object_length > label > select")
    browser_lib.click_element_when_visible(
        "css:#investments-table-object_length > label > select > option:nth-child(4)")
    time.sleep(10)
    сontent_of_table = []

    for i in range(1, 8):
        content = browser_lib.get_table_cell("css:.dataTables_scrollHeadInner > table:nth-child(1)", 2, i)
        сontent_of_table.append(content)
        
    for x in range(3, 161):
        for y in range(1, 8):
            content = browser_lib.get_table_cell("css:#investments-table-object", x, y)
            сontent_of_table.append(content)

    rows_of_table = list(map(list, list(zip(*[iter(сontent_of_table)] * 7))))

    try:
        lib.open_workbook('output\{}'.format(file_name))
        lib.create_worksheet("Department of Commerce")
        lib.append_rows_to_worksheet(rows_of_table, "Department of Commerce")
    finally:
        lib.save_workbook()
        lib.close_workbook()


def open_link_and_download_pdf(number_of_rows_in_the_table):
    '''Function traverses all "UII" elements, if element contains a link opens it, presses a button "Download Business Case PDF", downloads PDF 
    and store it in the output folder.'''

    for x in range(1, number_of_rows_in_the_table + 1):

        time.sleep(20)
        browser_lib.click_element_when_visible("css:#investments-table-object_length > label > select")
        browser_lib.click_element_when_visible(
            "css:#investments-table-object_length > label > select > option:nth-child(4)")
        time.sleep(15)

        if x % 2 == 0:
            time.sleep(15)
            if browser_lib.does_page_contain_link(
                    "css:tr.even:nth-child({}) > td:nth-child(1) > a:nth-child(1)".format(x)):
                browser_lib.click_link("css:tr.even:nth-child({}) > td:nth-child(1) > a:nth-child(1)".format(x))
            else:
                break

        else:
            time.sleep(15)
            if browser_lib.does_page_contain_link(
                    "css:tr.odd:nth-child({}) > td:nth-child(1) > a:nth-child(1)".format(x)):
                browser_lib.click_link("css:tr.odd:nth-child({}) > td:nth-child(1) > a:nth-child(1)".format(x))
            else:
                break

        browser_lib.wait_until_page_contains_element("css:#business-case-pdf > a", 15)
        browser_lib.click_link("css:#business-case-pdf > a")
        time.sleep(10)
        browser_lib.go_back()
        browser_lib.go_back()


def get_data_from_pdf(path_to_pdf_files):
    '''Function extracts data from 'Section A' in each file.'''
    if filesystem.is_directory_not_empty(path_to_pdf_files):
        list_of_files = filesystem.list_files_in_directory(path_to_pdf_files)
        list_of_pdf_files = []
        for file in list_of_files:
            file_name = filesystem.get_file_name(file)
            if file_name.endswith(".pdf"):
                list_of_pdf_files.append(file_name)

    section_a_in_each_file = {}
    list_name_of_investment_uii = []
    dict_name_of_investment_uii = {}

    for file in list_of_pdf_files:
        text = pdf.get_text_from_pdf(path_to_pdf_files + file, pages='1')[1]
        section_a_in_each_file[file] = text[text.find("Section A:"):text.find("Section B:")]

    for data in section_a_in_each_file.values():
        name_of_this_investment = data[data.find("Name of this Investment:"):data.find(
            "2. Unique Investment Identifier (UII)")][25:].replace("\n", " ")
        unique_investment_identifier = data[data.find("2. Unique Investment Identifier (UII)"):][39:]
        list_name_of_investment_uii.append({"Name of this Investment": name_of_this_investment,
                                            "Unique Investment Identifier (UII)": unique_investment_identifier})

    for x in range(len(list_of_pdf_files)):
        dict_name_of_investment_uii[x] = list_name_of_investment_uii[x]

    return dict_name_of_investment_uii


def compare_pdf_and_table_data(path_to_pdf_files, table_locator):
    '''Function gets data from pdf files and compares it with table data.'''
    name_of_investment_uii = get_data_from_pdf(path_to_pdf_files)
    time.sleep(15)
    browser_lib.click_element_when_visible("css:#investments-table-object_length > label > select")
    browser_lib.click_element_when_visible(
        "css:#investments-table-object_length > label > select > option:nth-child(4)")
    time.sleep(15)

    for x in range(0, len(name_of_investment_uii)):
        if name_of_investment_uii[x].get("Unique Investment Identifier (UII)") == browser_lib.get_table_cell("css:" + table_locator, x+3, 1):
            print(True)
        else:
            print(False)

    print()

    for x in range(len(name_of_investment_uii)):
       if name_of_investment_uii[x].get("Name of this Investment") == browser_lib.get_table_cell("css:" + table_locator, x+3, 3):
          print(True)
       else:
          print(False)


def main():
    try:
        open_the_website("https://itdashboard.gov/")
        click_divein_button("#node-23 > div > div > div > div > div > div > div > a")
        write_agencies_data_to_excel('agencies.xls', 'Agencies')
        choose_one_of_agencie_and_scrape_a_table("#agency-tiles-widget > div > div:nth-child(1) > div:nth-child(2) > div > div > div > div:nth-child(3) > a",
        'agencies.xls')
        open_link_and_download_pdf(158)
        compare_pdf_and_table_data(r"output\\", "#investments-table-object")
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
