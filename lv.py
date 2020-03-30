import time
import pyperclip

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains


# From 'Manage Samples' main screen, bring a particular submission into list_iframe.
def bring_up_submission(driver, submission):
    go_to_nav_iframe(driver)
    driver.find_element_by_id("td_SampleBySubmission").click()
    submission_input = driver.find_element_by_id("SampleBySubmission_arg1")
    submission_input.clear()
    submission_input.send_keys(
        submission, Keys.RETURN)
    time.sleep(0.5)


# This must be called within the Data Entry page.
# Will return a list of green, black, red for each input box.
def check_data_flags(driver):
    raw_flags_list = []
    inputs = get_input_boxes(driver)
    for i in range(len(inputs)):
        if len(inputs[i].get_attribute("value")) > 0:
            raw_flags_list.append(inputs[i].value_of_css_property("color"))
        else:
            raw_flags_list.append(" ")

    flags_list = []
    for flag in raw_flags_list:
        if flag == "rgba(0, 0, 0, 1)":
            flags_list.append("No Range")
        elif flag == "rgba(0, 128, 0, 1)":
            flags_list.append("OK")
        elif flag == "rgba(255, 0, 0, 1)":
            flags_list.append("Flagged")
        else:
            flags_list.append("No Input")
    return flags_list


def clear_specifications_and_add(driver, species, tissue):
    main_window = driver.current_window_handle
    go_to_nav_iframe(driver)
    details_button = driver.find_elements_by_class_name("gwt-HTML")[7]
    hover = ActionChains(driver).move_to_element(details_button)
    hover.perform()
    specifications_button = driver.find_element_by_id("Specifications")
    try:
        specifications_button.click()
    except:
        print("\t\tRetrying to find specifications_button")
        go_to_nav_iframe(driver)
        details_button = driver.find_elements_by_class_name("gwt-HTML")[7]
        hover = ActionChains(driver).move_to_element(details_button)
        hover.perform()
        specifications_button = driver.find_element_by_id("Specifications")
        specifications_button.click()

    time.sleep(1)

    go_to_maint_iframe(driver)
    spec_check_box = driver.find_element_by_id("selector_spec_row_0")
    spec_check_box.click()

    spec_remove_button = driver.find_element_by_id("spec_button_3")
    spec_remove_button.click()

    driver.switch_to.default_content()
    dlg_button = driver.find_element_by_class_name("dialog_contents_btn")
    dlg_button.click()
    time.sleep(1)

    ''' Don't think we really need to save until new spec is added.
    dlg_frame = driver.find_elements_by_tag_name("iframe")
    driver.switch_to.frame(dlg_frame[3])
    save_button = driver.find_element_by_id("Save")
    save_button.click()
    '''

    try:
        find_and_click_add_spec_button(driver)
    except:
        print("\t\tRetrying to find add_spec_button")
        time.sleep(3)
        find_and_click_add_spec_button(driver)

    driver.switch_to.window("spec")
    spec_search_box = driver.find_element_by_id("searchtext")
    spec_search_box.send_keys(species+"-"+tissue)
    spec_search_box.send_keys(Keys.RETURN)

    driver.switch_to.frame("list_iframe")
    highest = 0
    maxxed = False

    id_string = species + "-" + tissue + "|"
    while not maxxed:
        search_id_string = id_string + str(highest+1)
        try:
            driver.find_element_by_id(search_id_string)
        except:
            try:
                driver.find_element_by_id(id_string + str(highest+2))
            except:
                maxxed = True
                continue
        highest += 1

    most_recent_checkbox = driver.find_element_by_id(id_string + str(highest))
    most_recent_checkbox.click()

    driver.switch_to.default_content()
    return_button = driver.find_element_by_id("SelectReturn")
    return_button.click()

    driver.switch_to.window(main_window)
    driver.switch_to.default_content()

    time.sleep(1)

    dlg_frame = driver.find_elements_by_tag_name("iframe")
    driver.switch_to.frame(dlg_frame[3])
    save_button = driver.find_element_by_id("Save")
    save_button.click()

    time.sleep(2)

    exit_specifications_window(driver)


# From 'Manage Samples' main screen, move to data entry - Fast Grid.
def enter_data_entry(driver):
    go_to_nav_iframe(driver)
    link_buttons = driver.find_elements_by_class_name("gwt-HTML")
    hover = ActionChains(driver).move_to_element(link_buttons[1])
    hover.perform()
    driver.find_element_by_id("DataEntryGrid2").click()


# In the Data Entry screen, clears all current data and sends paste command.
def clear_inputs_and_paste_new(driver):
    go_to_nav_iframe(driver)
    inputs = get_input_boxes(driver)
    for input in inputs:
        input.clear()
    time.sleep(3)
    inputs[0].send_keys(Keys.CONTROL, 'v')
    time.sleep(3)


# From the 'Manage samples page, will enter data for Submission ID until list of data is empty.
# List of data should correspond with number of data items for the test code.
def enter_data_for(driver, submission, exit=True, input_class="dataentry2-gridentry"):
    bring_up_submission(driver, submission)
    select_top_sample(driver)
    enter_data_entry(driver)
    clear_inputs_and_paste_new(driver)
    if exit:
        save_and_exit_data_entry(driver)


def exit_specifications_window(driver):
    return_button = driver.find_element_by_id("Close")
    try:
        return_button.click()
    except:
        print("\t\tRetrying to exit specifications window.")
        time.sleep(5)
        return_button.click()

# Return to list from Data Entry screen.


def exit_data_entry(driver):
    go_to_nav_iframe(driver)
    time.sleep(2)
    return_button = driver.find_element_by_class_name("gwt-HTML")
    return_button.click()

    for i in range(2):
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present(),
                                           'Timed out waiting for PA creation ' +
                                           'confirmation popup to appear.')

            alert = driver.switch_to.alert
            alert.accept()
        except:
            print("no alert")
            exit_data_entry(driver)
        time.sleep(0.5)

    # This next section moves the mouse when returning to Manage screen
    # Without it the sample manage menu will block some functions.
    driver.switch_to.default_content()
    action = ActionChains(driver)
    # The next line moves the pointer off a button which would otherwise hover and break things.
    some_button = driver.find_element_by_id("ws_sortable_top")
    action.move_to_element(some_button).perform()


def find_and_click_add_spec_button(driver):
    go_to_maint_iframe(driver)
    add_spec_button = driver.find_element_by_id("spec_button_2")
    add_spec_button.click()


def save_and_exit_data_entry(driver):
    go_to_nav_iframe(driver)
    save_button = driver.find_elements_by_class_name("gwt-HTML")[1]
    save_button.click()
    time.sleep(7)
    exit_data_entry(driver)


# From the Data Entry Screen, get the input boxes as a list.
def get_input_boxes(driver, input_id="dataentry2-gridentry"):
    time.sleep(0.5)
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
    dlg_frame = driver.find_elements_by_tag_name("iframe")
    driver.switch_to.frame(dlg_frame[3])
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
    username_field.send_keys(username[: -1])

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


def select_top_sample(driver):
    go_to_list_iframe(driver)
    driver.find_element_by_name("selector").click()
    time.sleep(0.5)
