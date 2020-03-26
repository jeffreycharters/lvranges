import openpyxl
import pprint
import pyperclip


# Adds a row from the spreadsheet into the dictionary.
# Returns the same dictionary with information added.
def add_specs(species, tissue_type, element, range1, range2, species_dict):
    range_dict = {"range1": range1, "range2": range2}

    # These if statements make sure keys don't already exist. If they do, we won't overwrite them.
    if species in species_dict:
        if tissue_type in species_dict[species]:
            species_dict[species][tissue_type][element] = range_dict
        else:
            species_dict[species][tissue_type] = {element: range_dict}
    else:
        species_dict[species] = {tissue_type: {element: range_dict}}

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


def get_input_string(ranges, species, type, order=""):
    input_string = ""
    if order == "":
        order = get_elements_order()

    for element in order:
        if element in ranges[species][type].keys():
            flag_ok_value = ranges[species][type][element]["flag_ok_value"]
            input_string += str(flag_ok_value) + "\t"
        else:
            input_string += "\t"

    pyperclip.copy(input_string)
    return


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

    species_dict = dict()
    # TODO: Load xl file, start populating ranges.
    wb = openpyxl.load_workbook(filename)
    sheet = wb["Nick - Reference Ranges"]

    for row in range(3, sheet.max_row + 1):
        species = sheet['A' + str(row)].value
        tissue_type = sheet['B' + str(row)].value
        element = sheet['D' + str(row)].value.lower()
        range1 = extract_float(sheet['E'+str(row)].value)
        range2 = extract_float(sheet['F' + str(row)].value)
        species_dict = add_specs(species, tissue_type, element,
                                 range1, range2, species_dict)

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

                if species == "bovine" and element == "cobalt":
                    range_dict["flag_low_value"] = round(
                        range_dict["range1"] * 0.9, 3)
                    range_dict["flag_ok_value"] = round(
                        range_dict["range1"] * 1.1, 3)

                elif not range_dict["range2"]:
                    range_dict["flag_ok_value"] = round(
                        range_dict["range1"] * 0.9, 3)
                    range_dict["flag_high_value"] = round(
                        range_dict["range1"] * 1.1, 3)

                else:
                    range_dict["flag_low_value"] = round(
                        range_dict["range1"] * 0.9, 3)
                    range_dict["flag_ok_value"] = round((
                        range_dict["range1"] + range_dict["range2"]) / 2, 3)
                    range_dict["flag_high_value"] = round(
                        range_dict["range2"] * 1.1, 3)

    return data


def main():
    ranges = load_data()


if __name__ == "__main__":
    main()
