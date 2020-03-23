from selenium.webdriver.common.keys import Keys


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
    username_field.send_keys(username)

    # Find the password field, populate it and submit.
    password_field = driver.find_element_by_id("databasepassword")
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    # Make sure credentials were accepted.
    assert "Invalid username or password specified" not in driver.page_source
