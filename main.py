import time
import random
from lv import *

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys


def main():
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

    data = []
    for j in range(19):
        data.append(float(random.randint(1, 1000000))/1000000)

    enter_data_for(driver, "18-074980", data)

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
