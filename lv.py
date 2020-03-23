import time
from selenium.webdriver.common.keys import Keys

from selenium.webdriver import ActionChains


# From 'Manage Samples' main screen, bring a particular submission into list_iframe.
def bring_up_submission(driver, submission):
    go_to_nav_iframe(driver)
    driver.find_element_by_link_text("SampleBySubmission").click()
    submission_input = driver.find_element_by_id("SampleBySubmission_arg1")
    submission_input.clear()
    submission_input.send_keys(
        submission, Keys.RETURN)
    go_to_list_iframe(driver)
    driver.find_element_by_name("selector").click()


# From 'Manage Samples' main screen, move to data entry - Fast Grid.
def enter_data_entry(driver):
    go_to_nav_iframe(driver)
    link_buttons = driver.find_elements_by_class_name("gwt-HTML")
    hover = ActionChains(driver).move_to_element(link_buttons[1])
    hover.perform()
    driver.find_element_by_id("DataEntryGrid2").click()


# Go from 'Manage samples' main screen to the iframe where samples are.
def go_to_list_iframe(driver):
    driver.switch_to.default_content()
    driver.switch_to.frame("_nav_frame1")
    time.sleep(0.5)
    driver.switch_to.frame("list_iframe")


# Go from 'Manage samples' main screen to iframe with navigation buttons.
def go_to_nav_iframe(driver):
    driver.switch_to.default_content()
    driver.switch_to.frame("_nav_frame1")


# Gets username and password as strings and log into LabVantage.
# Gets these from a file in the same directory called "credentials.txt".
# First line of the file is the username, second is password.
def login(driver, filename="credentials.txt"):
    # Open the file and extract the first two lines into variables.
    file = open(filename, "r")
    username = file.readline()
    password = file.readline()

    # Find the username field and populate it.
    username_field = driver.find_element_by_id("databaseusername")
    username_field.send_keys(username[:-1])

    # Find the password field, populate it and submit.
    password_field = driver.find_element_by_id("databasepassword")
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    # Make sure credentials were accepted.
    assert "Invalid username or password specified" not in driver.page_source
