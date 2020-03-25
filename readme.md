# LabVantage Reference Range Testing

This software has been written in order to automate testing using LabVantage 8. Our lab needs to add a lot of reference ranges and test to ensure that they are entered and behaving as expected.

## Getting started

In order to use this software you need a UoG LabVantage Account. Put a text file named "credentials.txt" in the same directory as these files. The first line should contain only your username, the second line should contain only your password. Feel free to hard code them if you're a crazy person.

## Commands
### Entering Data for a particular submission.

In order to enter data for a submission, use the enter_data_for() function:

``` python
enter_data_for(driver, submission_id, data)
```
The driver is your WebDriver object, the submission ID is a string containing only the submission ID (i.e "YY-NNNNNN"). The data is a list containing the data to be written.