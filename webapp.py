from flask import Flask
from flask import render_template
from flask import request

import datas
import openpyxl

import json

app = Flask(__name__)


@app.route('/')
def web_interface():

    f = open("ref_ranges_current.json", "r")
    ranges = json.loads(f.read())
    f.close()

    return render_template("index.html", ranges=ranges)


@app.route('/save_ranges', methods=["POST"])
def save_ranges():
    ranges = request.form['ranges_object']
    try:
        file = open("ref_ranges_current.json", "w")
        file.write(ranges)
        file.close()
        return "Changes saved"
    except:
        return "Something broke, go tell Jeffrey"
