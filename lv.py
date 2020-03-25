import time
import pyperclip

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


def clear_specifications_and_add(driver, species, tissue):
    go_to_nav_iframe(driver)
    details_button = driver.find_elements_by_class_name("gwt-HTML")[7]
    hover = ActionChains(driver).move_to_element(details_button)
    hover.perform()
    specifications_button = driver.find_element_by_id("Specifications")
    specifications_button.click()

    go_to_maint_iframe(driver)
    spec_check_box = driver.find_element_by_id("selector_spec_row_0")
    spec_check_box.click()

    spec_remove_button = driver.find_element_by_id("spec_button_3")
    spec_remove_button.click()

    driver.switch_to.default_content()
    driver.find_element_by_id("dlgBtn1_0").click()

    go_to_maint_iframe(driver)
    add_spec_button = driver.find_element_by_id("spec_button_2")
    add_spec_button.click()

    # TODO: add new specification to submission.
    # TODO: return to manage screen.


# From 'Manage Samples' main screen, move to data entry - Fast Grid.
def enter_data_entry(driver):
    go_to_nav_iframe(driver)
    link_buttons = driver.find_elements_by_class_name("gwt-HTML")
    hover = ActionChains(driver).move_to_element(link_buttons[1])
    hover.perform()
    driver.find_element_by_id("DataEntryGrid2").click()


# From the 'Manage samples page, will enter data for Submission ID until list of data is empty.
# List of data should correspond with number of data items for the test code.
def enter_data_for(driver, submission, data, exit=True, input_class="dataentry2-gridentry"):
    bring_up_submission(driver, submission)
    enter_data_entry(driver)
    inputs = get_input_boxes(driver)
    tabbed_data = ""
    for d in data:
        tabbed_data += str(d)+"\t"
    pyperclip.copy(tabbed_data[:-1])
    inputs[0].send_keys(Keys.CONTROL, 'v')
    time.sleep(3)
    if exit:
        save_and_exit_data_entry(driver)


# Return to list from Data Entry screen.
def exit_data_entry(driver):
    go_to_nav_iframe(driver)
    driver.find_elements_by_class_name("gwt-HTML")[0].click()
    # This next section moves the mouse when returning to Manage screen
    # Without it the sample manage menu will block some functions.
    driver.switch_to.default_content()
    action = ActionChains(driver)
    some_button = driver.find_element_by_id("ws_sortable_top")
    action.move_to_element(some_button).perform()


def save_and_exit_data_entry(driver):
    go_to_nav_iframe(driver)
    save_button = driver.find_elements_by_class_name("gwt-HTML")[1]
    save_button.click()
    time.sleep(3)
    exit_data_entry(driver)


# From the Data Entry Screen, get the input boxes as a list.
def get_input_boxes(driver, input_id="dataentry2-gridentry"):
    go_to_right_iframe(driver)
    return driver.find_elements_by_class_name(input_id)


# Go from 'Manage samples' main screen to iframe with navigation buttons.
def go_to_nav_iframe(driver):
    driver.switch_to.default_content()
    driver.switch_to.frame("_nav_frame1")


# Go from 'Manage samples' main screen to the iframe where samples are.
def go_to_list_iframe(driver):
    go_to_nav_iframe(driver)
    time.sleep(0.5)
    driver.switch_to.frame("list_iframe")


# Go from 'Manage samples' main screen to the iframe where samples are.
def go_to_right_iframe(driver):
    go_to_nav_iframe(driver)
    time.sleep(0.5)
    driver.switch_to.frame("rightframe")


def go_to_maint_iframe(driver):
    driver.switch_to.default_content()
    time.sleep(0.5)
    driver.switch_to.frame("dlg_frame0")
    time.sleep(0.5)
    driver.switch_to.frame("maint_iframe")

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


def logout(driver):
    driver.switch_to.default_content()
    menu_button = driver.find_element_by_id("e_profilemenu")
    menu_button.click()
    driver.find_element_by_partial_link_text("Log Off").click()
    time.sleep(0.1)
    # driver.switch_to.alert.accept()
    assert "LabVantage Logon" in driver.title
