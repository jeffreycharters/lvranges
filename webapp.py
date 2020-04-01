from flask import Flask
from flask import render_template
from flask import request

import datas
import openpyxl

import json

app = Flask(__name__)


def add_specs(species, tissue_type, element, range1, range2, species_dict):
    range_dict = {"range1": range1, "range2": range2}

    if species in species_dict:
        if tissue_type in species_dict[species]:
            if element in species_dict[species][tissue_type]:
                species_dict[species][tissue_type][element] = range_dict

    return species_dict


@app.route('/')
def web_interface():

    f = open("ranges.json", "r")
    ranges = json.loads(f.read())
    f.close()

    excel_filename = "xlfile.xlsx"
    xl_sheetname = "Nick - Reference Ranges"

    wb = openpyxl.load_workbook(excel_filename)
    sheet = wb[xl_sheetname]

    for row in range(3, sheet.max_row + 1):
        species = sheet['A' + str(row)].value.lower()
        tissue_type = sheet['B' + str(row)].value.lower()
        element = sheet['D' + str(row)].value.lower()
        range1 = sheet['E'+str(row)].value
        range2 = sheet['F' + str(row)].value
        ranges = add_specs(species, tissue_type, element,
                           range1, range2, ranges)

    wb.close()

    return render_template("index.html", ranges=ranges)


@app.route('/save_ranges', methods=["POST"])
def save_ranges():
    ranges = request.form['ranges_object']
    try:
        file = open("ref_ranges.json", "w")
        file.write(ranges)
        file.close()
        return "Changes saved"
    except:
        return "Something broke, go tell Jeffrey"
