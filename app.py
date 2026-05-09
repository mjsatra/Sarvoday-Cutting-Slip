from flask import Flask, render_template, request, redirect
import json
import datetime
import os
from flask import jsonify

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(BASE_DIR, "slip_history.json")


def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE) as f:
            return json.load(f)
    return []


def save_history(data):
    history = load_history()
    history.append(data)
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)

from openpyxl import load_workbook

def load_party_data():

    file_path = os.path.join(BASE_DIR, "party_master.json")

    party_list = []

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:
                data = json.load(f)

                for row in data:

                    party = row.get("party", "").strip()

                    if party:
                        party_list.append(party)

            except:
                pass

    return sorted(list(set(party_list)))

def load_item_data():

    file_path = os.path.join(BASE_DIR, "item_master.json")

    items = []

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:

                data = json.load(f)

                for row in data:

                    item = row.get("item", "").strip()

                    if item:
                        items.append(item)

            except:
                pass

    return sorted(list(set(items)))

def get_next_slip_no():
    history = load_history()
    return len(history) + 1

from openpyxl import load_workbook

import pandas as pd

def load_party_details():

    file_path = os.path.join(BASE_DIR, "party_master.json")

    party_data = {}

    school_file = os.path.join(BASE_DIR, "school_master.json")

    if os.path.exists(school_file):

        with open(school_file, "r") as f:

            try:
                school_list = json.load(f)

            except:
                school_list = []

    cloth_list = []

    cloth_file = os.path.join(BASE_DIR, "cloth_master.json")

    if os.path.exists(cloth_file):

        with open(cloth_file, "r") as f:

            try:
                cloth_list = json.load(f)

            except:
                cloth_list = []

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:

                data = json.load(f)

                for row in data:

                    party = str(
                        row.get("party", "")
                    ).strip()

                    address = str(
                        row.get("address", "")
                    ).strip()

                    if party:
                        party_data[party] = address

            except:
                pass

    return {
        "party_data": party_data,
        "school_list": school_list,
        "cloth_list": cloth_list
    }

def load_size_map():

    file_path = os.path.join(BASE_DIR, "item_master.json")

    size_map = {}

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:

                data = json.load(f)

                for row in data:

                    item = row.get("item", "").strip()

                    sizes = row.get("sizes", [])

                    if item:
                        size_map[item] = sizes

            except:
                pass

    return size_map

@app.route("/", methods=["GET", "POST"])
def home():

    if request.method == "POST":
        slip_no = request.form.get("slip_no")
        
        # Ensure slip_no is stored properly
        try:
            slip_no = int(slip_no) if slip_no else get_next_slip_no()
        except:
            slip_no = get_next_slip_no()
        
        data = {
            "slip_no": str(slip_no),
            "date": datetime.datetime.strptime(
                request.form.get("date"),
                "%Y-%m-%d"
            ).strftime("%d-%m-%Y"),
            "party": request.form.get("party"),
            "address": request.form.get("address"),
            "cloth": request.form.get("cloth"),
            "school": request.form.get("school"),
            "item": request.form.get("item"),
            "sizes": request.form.getlist("size[]"),
            "qtys": request.form.getlist("qty[]"),
            "meters": request.form.getlist("meter[]"),
            "total_qty": request.form.get("total_qty"),
            "total_meter": request.form.get("total_meter")
        }

        save_history(data)

        return redirect("/")

    history = load_history()
    
    # Reverse for display
    history_display = list(reversed(history))

    party_list = load_party_data()

    item_list = load_item_data()

    master_data = load_party_details()

    size_map = load_size_map()

    slip_no = get_next_slip_no()

    return render_template(
        "index.html",
        history=history_display,
        party_list=party_list,
        item_list=item_list,
        slip_no=slip_no,
        party_data=master_data["party_data"],
        school_list=master_data["school_list"],
        cloth_list=master_data["cloth_list"],
        size_map=size_map
    )

@app.route("/save_party", methods=["POST"])
def save_party():

    data = request.json

    party_name = data.get("party")
    address = data.get("address")

    file_path = os.path.join(BASE_DIR, "party_master.json")

    if os.path.exists(file_path):

        with open(file_path, "r") as f:
            parties = json.load(f)

    else:
        parties = []

    parties.append({
        "party": party_name,
        "address": address
    })

    with open(file_path, "w") as f:
        json.dump(parties, f, indent=2)

    return jsonify({
        "status": "success"
    })   

@app.route("/save_cloth", methods=["POST"])
def save_cloth():

    data = request.get_json()

    cloth = data.get("cloth")

    file_path = os.path.join(BASE_DIR, "cloth_master.json")

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:
                cloths = json.load(f)

            except:
                cloths = []

    else:
        cloths = []

    cloths.append(cloth)

    with open(file_path, "w") as f:
        json.dump(cloths, f, indent=2)

    return jsonify({
        "status":"success"
    })

@app.route("/save_school", methods=["POST"])
def save_school():

    data = request.get_json()

    school = data.get("school")

    file_path = os.path.join(BASE_DIR, "school_master.json")

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:
                schools = json.load(f)

            except:
                schools = []

    else:
        schools = []

    schools.append(school)

    with open(file_path, "w") as f:
        json.dump(schools, f, indent=2)

    return jsonify({
        "status":"success"
    })

@app.route("/save_item", methods=["POST"])
def save_item():

    data = request.get_json()

    item = data.get("item")

    sizes = data.get("sizes")

    size_list = [
        s.strip() for s in sizes.split(",")
        if s.strip()
    ]

    file_path = os.path.join(BASE_DIR, "item_master.json")

    if os.path.exists(file_path):

        with open(file_path, "r") as f:

            try:
                items = json.load(f)

            except:
                items = []

    else:
        items = []

    items.append({
        "item": item,
        "sizes": size_list
    })

    with open(file_path, "w") as f:
        json.dump(items, f, indent=2)

    return jsonify({
        "status":"success"
    })

if __name__ == "__main__":
    app.run(debug=True)
