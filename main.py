import time
import random
import datas
from lv import *

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys


def main():

    ranges = {"bovine": {"kidney": {"selenium": {
        "flag_low_value": 0.9, "flag_ok_value": 1.25, "flag_high_value": 1.75}}}}

    '''
    excel_filename = "xlfile.xlsx"
    xl_sheetname = "Nick - Reference Ranges"

    ranges = datas.load_data(filename=excel_filename,
                             sheetname=xl_sheetname, flags=True)

    print("Data Loaded OK")

    '''

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

    # # # # # #
    # User should now be logged in, interaction with LabVantage goes below
    # # # # # #

    # For future use
    hmsc_order = ["Antimony", "Arsenic‎", "Beryllium‎", "Boron‎", "Cadmium‎", "Chromium‎", "Cobalt‎",
                  "Copper‎", "Iron‎", "Lead‎", "Magnesium‎", "Manganese‎", "Mercury‎", "Molybdenum‎", "Nickel‎",
                  "Selenium", "Thellium", "Tin", "Zinc"]

    icpse_order = ["manganese", "iron", "cobalt",
                   "copper", "zinc", "selenium", "molybdenum"]

    # Load the submission then add a new specification to it.
    bring_up_submission(driver, "18-074980")
    clear_specifications_and_add(driver, "bovine", "kidney")
    # TODO: finish this function. Make sure it loads the correct specification, not old!!

    # TODO: enter correct data for each specification.
    # TODO: create xl spreadsheet with results.

    time.sleep(5)

    '''
    data = []
    for j in range(19):
        data.append(float(random.randint(1, 1000000))/1000000)

    enter_data_for(driver, "18-074980", data)
    '''

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
