import openpyxl
import pprint
import pyperclip


# Adds a row from the spreadsheet into the dictionary.
# Returns the same dictionary with information added.
def add_specs(species, matrix, element, spec_version_id,
              param_list_id, param_id, range1_qualifier, range2_qualifier,
              range1_value, range2_value, species_dict):
    range_dict = {
        "range1_qualifier": range1_qualifier,
        "range1_value": range1_value,
        "range2_qualifier": range2_qualifier,
        "range2_value": range2_value,
        "spec_version_id": spec_version_id,
        "param_list_id": param_list_id,
        "param_id": param_id
    }

    # These if statements make sure keys don't already exist. If they do, we won't overwrite them.
    if species in species_dict:
        if matrix in species_dict[species]:
            species_dict[species][matrix][element] = range_dict
        else:
            species_dict[species][matrix] = {element: range_dict}
    else:
        species_dict[species] = {matrix: {element: range_dict}}

    return species_dict


# Gets the floating point number from a cell containing a string like "<= 3.2"
def extract_float(cell_content):
    range_value = ""
    for t in cell_content.split():
        try:
            range_value = float(t)
        except ValueError:
            pass
    return range_value


def get_qualifier(string):
    list = string.split(" ", 1)
    return list[0]


def get_element(param_id):
    converter = {
        'Lead blood': "lead",
        'Magnesium': "magnesium",
        'Selenium tissue': "selenium",
        'tmo': "molybdenum",
        'Iron': "iron",
        'tcu': "copper",
        'tfe': "iron",
        'Magnesium serum': "magnesium",
        'Selenium': "selenium",
        'tse': "selenium",
        'Calcium': "calcium",
        'Lead': "lead",
        'tmn': "manganese",
        'Zinc-Serum': "zinc",
        'Calcium serum': "calcium",
        'Copper': "copper",
        'Manganese': "manganese",
        'tco': "cobalt",
        'Arsenic': "aresenic",
        'tzn': "zinc",
        'Copper serum': "copper",
        'Cadmium': "cadmium",
        'Ca': "calcium",
        'Hg': "mercury",
        'Zinc': "zinc"
    }
    # Find the param_id in the dictionary, if not present will return unknown.
    return converter.get(param_id, "unknown")


def get_input_string(ranges, species, type, test_type, order=""):
    input_string = ""
    if order == "":
        order = get_elements_order()

    for element in order:
        if element in ranges[species][type].keys():
            if test_type in ranges[species][type][element].keys():
                flag_value = ranges[species][type][element][test_type]
                input_string += str(flag_value) + "\t"
            else:
                input_string += "-444\t"
        else:
            if test_type == "flag_low_value":
                input_string += "0.01\t"
            elif test_type == "flag_ok_value":
                input_string += "1\t"
            elif test_type == "flag_high_value":
                input_string += "1000\t"

    pyperclip.copy(input_string)
    return input_string


def get_elements_order():
    hmsc_order = ["antimony", "arsenic", "beryllium", "boron", "cadmium", "chromium", "cobalt",
                  "copper", "iron", "lead", "magnesium", "manganese", "mercury", "molybdenum", "nickel",
                  "selenium", "thallium", "tin", "zinc"]

    icpti_order = ["calcium", "cobalt", "copper", "iron", "magnesium", "manganese",
                   "molybdenum", "phosphorus", "potassium", "selenium", "sodium", "zinc"]

    icpse_order = ["manganese", "iron", "cobalt",
                   "copper", "zinc", "selenium", "molybdenum"]

    elements_in_order = hmsc_order + ["copper", "copper", "lead", "selenium"] + \
        icpti_order + ["copper", "lead", "selenium",
                       "selenium"] + icpse_order + ["zinc"]

    return elements_in_order


# This function takes a filename and sheet name and extracts the relevant data from it.
# Designed to take specifically the sheet with our current reference ranges.
# Elements column may need modified to specificallt display the written element name, e.g. Lead
def load_data(filename="xlfile.xlsx", sheetname="Nick - Reference Ranges", flags=True):

    species_dict = {}
    # TODO: Load xl file, start populating ranges.
    wb = openpyxl.load_workbook(filename)
    sheet = wb[sheetname]

    for row in range(2, sheet.max_row + 1):
        column_a = sheet['A' + str(row)].value
        species, matrix = column_a.split("-", 1)
        spec_version_id = sheet['B' + str(row)].value
        param_list_id = sheet['C' + str(row)].value
        param_id = sheet['D' + str(row)].value
        range1 = sheet['E' + str(row)].value
        range2 = sheet['F' + str(row)].value

        range1_qualifier = get_qualifier(range1)
        range2_qualifier = get_qualifier(range2)
        range1_value = extract_float(range1)
        range2_value = extract_float(range2)

        element = get_element(param_id)

        species_dict = add_specs(species, matrix, element, spec_version_id,
                                 param_list_id, param_id, range1_qualifier, range2_qualifier,
                                 range1_value, range2_value, species_dict)

    wb.close()

    if flags:
        species_dict = add_flagging_values(species_dict)

    return species_dict


# Takes the data loaded from the spreadsheet and adds a few more fields
# To the data_dict dictionary to make testing easier.
# These are required for automated reference range testing.
def add_flagging_values(data):
    for species in data:
        for tissue in data[species]:
            for element in data[species][tissue]:
                range_dict = data[species][tissue][element]

                if species == "bovine" and element == "cobalt" and tissue == "serum":
                    range_dict["flag_low_value"] = round(
                        range_dict["range1_value"] * 0.9, 3)
                    range_dict["flag_ok_value"] = round(
                        range_dict["range1_value"] * 1.1, 3)

                elif not range_dict["range2_value"]:
                    range_dict["flag_low_value"] = "-666"
                    range_dict["flag_ok_value"] = round(
                        range_dict["range1_value"] * 0.9, 3)
                    range_dict["flag_high_value"] = round(
                        range_dict["range1_value"] * 1.1, 3)

                else:
                    range_dict["flag_low_value"] = round(
                        range_dict["range1_value"] * 0.9, 3)
                    range_dict["flag_ok_value"] = round((
                        range_dict["range1_value"] + range_dict["range2_value"]) / 2, 3)
                    range_dict["flag_high_value"] = round(
                        range_dict["range2_value"] * 1.1, 3)

    return data


def main():
    ranges = load_data()


if __name__ == "__main__":
    main()
