import time
from lv import *

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys


def main():
    # Set up the web driver using Chrome since LV8 only really works with Chrome (boo).
    # Make this a global variable to make code less ornery.
    driver = Chrome()

    # Set the default wait time of 10 seconds for an element to load before the script bails.
    driver.implicitly_wait(10)

    # Open LabVantage login page and make sure it exists based on the page title.
    driver.get("http://sapphire.lsd.uoguelph.ca:8080/labservices/logon.jsp")
    assert "LabVantage" in driver.title

    # Call the login function. See lv.py for clarification.
    # Result should be a successful login to LV8.
    login(driver)

    bring_up_submission(driver, "18-074980")
    enter_data_entry(driver)

    time.sleep(3)

    # Teardown.
    driver.close()
    driver.quit()

    # Shut out the lights and turn the heat down on your way out.
    print("done")


if __name__ == "__main__":
    main()
