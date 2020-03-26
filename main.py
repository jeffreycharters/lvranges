import time
import random
import datas
import outputs
from lv import *

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys


def main():

    order = datas.get_elements_order()

    excel_filename = "xlfile.xlsx"
    xl_sheetname = "Nick - Reference Ranges"

    ranges = datas.load_data(filename=excel_filename,
                             sheetname=xl_sheetname, flags=True)

    print("Data Loaded OK")

    # Set up the web driver using Chrome since LV8 only really works with Chrome (boo).
    # Make this a global variable to make code less ornery.
    driver = Chrome()
    driver.maximize_window()

    # Set the default wait time of 10 seconds for an element to load before the script bails.
    driver.implicitly_wait(10)

    # Open LabVantage login page and make sure it exists based on the page title.
    driver.get("http://sapphire.lsd.uoguelph.ca:8080/labservices/logon.jsp")
    assert "LabVantage Logon" in driver.title

    # Call the login function. See lv.py for clarification.
    # Result should be a successful login to LV8.
    login(driver)

    # # # # # # # # # # # # # # # # # # # # # # # # # #
    # User should now be logged in, interaction with LabVantage goes below
    # # # # # # # # # # # # # # # # # # # # # # # # # #

    # Load the submission then add a new specification to it.
    bring_up_submission(driver, "18-074980")

    wb, ws = outputs.create_workbook()

    title_font = Font(size=20, bold=True)
    species_font = Font(size=16, bold=True)
    tissue_font = Font(size=14, bold=True)

    ws.merge_cells("A1:G1")
    ws['A1'] = "LabVantage - Checking Status of Reference Ranges"
    ws['A1'].font = title_font

    base_row = 3

    for species in ranges:

        ws['A'+str(base_row)] = species
        ws['A'+str(base_row)].font = species_font

        base_row += 1

        for type in ranges[species]:

            # Add the new specification
            select_top_sample(driver)
            clear_specifications_and_add(driver, species, type)

            input_string = datas.get_input_string(ranges, species, type, order)

            # Put the input string (data) into the Data Entry page.
            enter_data_for(driver, "18-074980", exit=False)
            flags_list = check_data_flags(driver)
            save_and_exit_data_entry(driver)

            tissue_headers = False
            serum_headers = False
            if type in ["liver", "kidney"]:
                if not tissue_headers:
                    outputs.write_tissue_headers(ws, base_row)
                    base_row += 1
                    outputs.write_tissue_element_names(ws, base_row, order)
                    tissue_headers = True
                    base_row += 1
                outputs.write_tissue_row(ws, flags_list, base_row)
                base_row += 1
            else:
                if not serum_headers:
                    outputs.write_serum_headers(ws, base_row)
                    base_row += 1
                    outputs.write_serum_element_names(ws, base_row, order)
                    serum_headers = True
                    base_row += 1
                outputs.write_serum_row(ws, flags_list, base_row)
                base_row += 1

            base_row += 1

    for i, flag in enumerate(flags_list):
        column = get_column_letter(i+1)
        ws[column + str(base_row)] = order[i]
        ws[column + str(base_row+1)] = flag

    outputs.autoresize_columns(ws)

    outputs.save_workbook(wb)

    time.sleep(5)

    # # # # # #
    # Interactions in LabVantage should all be above this, teardown only past this point.
    # # # # # #

    # Teardown.

    logout(driver)
    driver.close()
    driver.quit()

    # Shut out the lights and turn the heat down on your way out.
    print("done")


if __name__ == "__main__":
    main()
