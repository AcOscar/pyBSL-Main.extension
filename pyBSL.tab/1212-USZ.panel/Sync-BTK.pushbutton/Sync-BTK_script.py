# -*- coding: utf-8 -*-
# pyRevit: liest JSON aus Project Information -> Parameter "ExcelReader"
# Struktur:
# {
#   "SheetSync": {...},
#   "ParameterSync": [ {...}, {...} ]
# }

import json
import re
import os
import sys
import csv
import codecs

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

PARAM_NAME = "ExcelReader"

# ---------------------------
# Utils
# ---------------------------
def get_project_info(doc):
    return FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_ProjectInformation) \
        .FirstElement()

def read_json_from_project_param(doc, param_name):
    pi = get_project_info(doc)
    if not pi:
        raise Exception("Project Information Element nicht gefunden.")

    p = pi.LookupParameter(param_name)
    if not p:
        raise Exception("Projektparameter '{}' nicht gefunden.".format(param_name))

    raw = p.AsString() or p.AsValueString()
    if not raw or not raw.strip():
        raise Exception("Projektparameter '{}' ist leer.".format(param_name))

    try:
        return json.loads(raw)
    except Exception as e:
        preview = (raw[:300] + ("..." if len(raw) > 300 else "")).replace("\n", " ")
        raise Exception("Ungültige JSON in '{}': {}.\nVorschau: {}".format(param_name, e, preview))

def ensure_list(value):
    """Wenn value ein Dict ist -> [value], wenn None -> [], sonst value zurück."""
    if value is None:
        return []
    if isinstance(value, dict):
        return [value]
    if isinstance(value, list):
        return value
    raise Exception("Unerwarteter Typ für ParameterSync: {}.".format(type(value)))

def is_probably_url(s):
    # Sehr einfache Plausibilitätsprüfung (kein echter URL-Parser)
    if not s or not isinstance(s, basestring):
        return False
    return bool(re.match(r"^https?://", s))

def validate_sheet_sync(node, path):
    required = ["Order", "FileName", "DataName", "KeyName"]
    for k in required:
        if k not in node:
            raise Exception("{}: Pflichtfeld '{}' fehlt.".format(path, k))
    if not is_probably_url(node.get("FileName", "")):
        # Hinweis statt harter Fehler – je nach Use-Case hier Exception werfen
        print("WARN: {}.FileName sieht nicht wie eine URL aus: {}".format(path, node.get("FileName")))

def validate_parameter_sync_item(node, path):
    required = ["Order", "FileName", "DataName", "KeyName"]
    for k in required:
        if k not in node:
            raise Exception("{}: Pflichtfeld '{}' fehlt.".format(path, k))
    if not is_probably_url(node.get("FileName", "")):
        print("WARN: {}.FileName sieht nicht wie eine URL aus: {}".format(path, node.get("FileName")))

def validate_config(cfg):
    if not isinstance(cfg, dict):
        raise Exception("Top-Level JSON muss ein Objekt sein.")

    if "SheetSync" not in cfg:
        raise Exception("Top-Level 'SheetSync' fehlt.")

    validate_sheet_sync(cfg["SheetSync"], "SheetSync")

    ps_list = ensure_list(cfg.get("ParameterSync"))
    for i, item in enumerate(ps_list):
        if not isinstance(item, dict):
            raise Exception("ParameterSync[{}] ist kein Objekt.".format(i))
        validate_parameter_sync_item(item, "ParameterSync[{}]".format(i))

    return {
        "SheetSync": cfg["SheetSync"],
        "ParameterSync": ps_list
    }

def read_excel(batch_file, excel_filename, excel_rangename, fileCsvlist):
    # Convert xls to temp csv files.
    os.system('"{}" "{}" "{}" "{}"'.format(batch_file, excel_filename, excel_rangename, fileCsvlist))

    if not os.path.exists(fileCsvlist):
        print('Creating the CSV file failed.')
        sys.exit()

    # Get data from csv file.
    data = []
            
    with codecs.open(fileCsvlist, "r", "utf-16") as f:
        clean_lines = (line.replace(u'\x00', u'') for line in f)
        reader = csv.DictReader(clean_lines, delimiter=",", quotechar='"')
        for row in reader:
            data.append(row)
            
    # if all(excel_keyname in row for row in data):
    #     key_row = next((row for row in data if row[excel_keyname] == 0), None)
    # else:
    #     print("The key '{}' not exist in csv data.".format (excel_keyname))
    #     sys.exit()

    print(data)

    return data


# ---------------------------
# Hauptlogik
# ---------------------------
def main():
    cfg_raw = read_json_from_project_param(doc, PARAM_NAME)
    cfg = validate_config(cfg_raw)

    sheet = cfg["SheetSync"]
    params = cfg["ParameterSync"]

    # Beispiele: Zugriff & Ausgabe
    print("=== SheetSync ===")
    print("Order     :", sheet.get("Order"))
    print("FileName  :", sheet.get("FileName"))
    print("DataName  :", sheet.get("DataName"))
    print("KeyName   :", sheet.get("KeyName"))

    print("\n=== ParameterSync ({} Einträge) ===".format(len(params)))
    # Nach Order numerisch sortieren (Order ist String -> in int wandeln, fallback 0)
    def order_key(x):
        try:
            return int(x.get("Order", "0"))
        except:
            return 0

    temp_dir = os.environ.get('TEMP')

    fileCsvlist = os.path.join(temp_dir,'temp-drawing-list.csv')
    print (fileCsvlist)
    batch_file = '{}'.format(os.path.dirname(__file__) + '\convert3.bat')

    for item in sorted(params, key=order_key):
        excel_filename = item.get("FileName")
        excel_rangename = item.get("DataName")
        excel_keyname = item.get("KeyName")

        print("- Order    :", item.get("Order"))
        print("  FileName :", excel_filename)
        print("  DataName :", excel_rangename)
        print("  KeyName  :", excel_keyname)

        # Ab hier: Übergib die Werte an deinen Excel-Reader/Importer
        # z.B.:
        # process_sheet_sync(sheet["FileName"], sheet["DataName"], sheet["KeyName"])
        # for item in sorted(params, key=order_key):
        #     process_parameter_sync(item["FileName"], item["DataName"], item["KeyName"])

        data = read_excel(batch_file, excel_filename, excel_rangename, fileCsvlist)

        print(data)
        os.remove(fileCsvlist)


# ---------------------------
# Start
# ---------------------------
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("FEHLER:", e)
        print(traceback.format_exc())
