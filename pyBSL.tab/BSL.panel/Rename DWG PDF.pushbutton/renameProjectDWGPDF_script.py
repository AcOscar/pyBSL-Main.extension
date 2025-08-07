import os
from collections import defaultdict
import clr
clr.AddReference("RevitAPI")
import Autodesk.Revit.UI
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory
from System.Diagnostics import Stopwatch
from rph import prm
import sys
#todo: check print_path is rigth formated (an \ and the end)

stopwatch = Stopwatch()
stopwatch.Start()

doc = __revit__.ActiveUIDocument.Document

def prj_param_exists(param_exists):
    return doc.ProjectInformation.LookupParameter(param_exists)

if not prj_param_exists("Print_Path"):
    print("The parameter Print_Path is necessary but does not exist. Please create them first.")
    sys.exit()
    
if not prj_param_exists("Print_Config"):
    print("The parameter Print_Config is necessary but does not exist. Please create them first.")
    sys.exit()


print_path = doc.ProjectInformation.LookupParameter("Print_Path").AsString()
print_config = doc.ProjectInformation.LookupParameter("Print_Config").AsString()

sheets = Fec(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
rename_dict = defaultdict(str)

search_types = ["pdf", "dwg"]
found_files = {"pdf": [], "dwg": []}


def get_str_param(element, param_name):
    if element.LookupParameter(param_name).AsString():
        return element.LookupParameter(param_name).AsString()
    else:
        return ""


def get_int_param(element, param_name):
    if element.LookupParameter(param_name).AsInteger():
        return element.LookupParameter(param_name).AsInteger()
    else:
        return 0


for sheet in sheets:
    sheet_nr = sheet.SheetNumber
    if prm.build_str_from_pattern(sheet, print_config, False):
        rename_dict[sheet_nr] += prm.build_str_from_pattern(sheet, print_config, False)

print("\n" + 35 * "_" + "following rename mapping will be used:")
for sheet_nr in rename_dict.keys():
    print([sheet_nr, rename_dict[sheet_nr]])

print("\n" + 35*"_")

if print_config:
    if os.path.exists(print_path):
        for _file in os.listdir(print_path):
            print(_file)
            for file_type in search_types:
                if _file.endswith(file_type):
                    found_files[file_type].append(_file)
                    file_suffix = _file.split(".")[-1]
                    #print (rename_dict.keys())
                    for sheet_nr in rename_dict.keys():
                        if sheet_nr in _file:
                        
                            if os.path.exists(print_path + _file):
                                new_file_name = print_path + rename_dict[sheet_nr] + "." + file_suffix
                                os.rename(print_path + _file, new_file_name)
                                print(_file + " renamed to: " + rename_dict[sheet_nr] + "." + file_suffix)
                            else:
                                print("File {} not exist.".format(_file))

        for file_type in found_files.keys():
            print("\n" + 35 * "_" + "in directory {0} files found of type: {1}".format(print_path, file_type))
            for _file in found_files[file_type]:
                print(_file)
    else:
        print("Required path: {0} does not exist!".format(print_path))
        print("Please create it before using this script. Thank you.")
else:
    print("Required path: print_config does not exist!")
    print("Please set it up before using this script. Thank you.")

print("HdM_pyRevit rename PDFs and DWGs run in:")

stopwatch.Stop()
timespan = stopwatch.Elapsed
print(timespan)
