import openpyxl
import datas
import json

''' Run this to set up a new JSON file with base ranges and all possible hmsc etc elements '''


def main():

    f = open("ranges_blank.json", "r")
    ranges = json.loads(f.read())
    f.close()

    excel_filename = "xlfile.xlsx"
    xl_sheetname = "Export Worksheet"

    ranges = datas.load_data(excel_filename, xl_sheetname,
                             ranges, flags=False, pared_down=True)

    json_ranges = json.dumps(ranges)
    try:
        file = open("ref_ranges_new.json", "w")
        file.write(json_ranges)
        file.close()
        print("Changes saved")
    except:
        print("Something broke, go tell Jeffrey")


if __name__ == "__main__":
    main()
