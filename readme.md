
# LabVantage Reference Range Testing 

This software has been written in order to automate testing using LabVantage 8. Our lab needs to add a lot of reference ranges and test to ensure that they are entered and behaving as expected.

  ## Getting started

In order to use this software you need a UoG LabVantage Account. Put a text file named "credentials.txt" in the same directory as these files. The first line should contain only your username, the second line should contain only your password. Feel free to hard code them if you're a crazy person.

### Files and Purpose

 ```main.py``` will run the webdriver in order to verify reference range data in LIMS. Expects your credentials in a file called ```credentials.txt```, username and password on consecutive lines. The input file is what IT gives us when we request our reference ranges, see ```xlfile.xlsx``` for an example.
 
```lv.py``` is a module containing the functions used to drive the webdriver. Edit very carefully, no testing exists to ensure you won't break anything.

```outputs.py``` is a module containing functions for exporting ```main.py``` into an excel file.
 
```buildnewjson.py``` does what it sounds like. Have a look in the code, will convert the same excel file as used in ```main.py``` into a json file which can then be used in the webapp. This requires ```ranges_blank.json``` as a template, anything without results will end up as ```NaN```. 


```webapp.py``` along with files in the ```/static``` and ```/templates``` folders run the web app. Launch this as a flask app and navigate to ```/localhost:5000``` to see the reference range app.