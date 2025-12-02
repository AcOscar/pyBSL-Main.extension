# -*- coding: utf-8 -*-
#from pyrevit import revit, DB
from pyrevit import script
#from Autodesk.Revit.DB import ElementMulticategoryFilter
#from Autodesk.Revit.DB import ElementId
#from Autodesk.Revit.DB import BuiltInParameter
#from Autodesk.Revit.DB import TransactionGroup, Transaction
#from Autodesk.Revit.DB import SpatialElementGeometryCalculator 
#from Autodesk.Revit.DB.Architecture import Room
from System.Diagnostics import Stopwatch


import os, time, csv, re, sys, collections, datetime
import Autodesk.Revit.DB as rdb
from Autodesk.Revit.DB import FilteredElementCollector as fec
from Autodesk.Revit.DB import BuiltInCategory as bic

__title__ = 'Sync drawing list with sheet informations'

__doc__ = 'Writes parameters '\
          'from an excel file '\
          'to all sheets'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def getDateObject(string):
    if re.search('\d{2}\.\d{2}\.\d{4}', string) is not None:
        return datetime.datetime.strptime(string, '%d.%m.%Y')
    return False

def extractDate(string):
    dates = re.findall(r'(\d{2}.\d{2}.\d{4})', string, re.MULTILINE)
    if dates:
        return dates[0]
    else:
        return ''

def has_value(element, param_name):
    if exists(element, param_name):
        return (element.LookupParameter(param_name).HasValue)

def exists(element, param_name):
    return (element.LookupParameter(param_name) != None)

    
def get_str_param(element, param_name):
    """
    :param element: Element that holds the parameter.
    :param param_name: Name of the parameter to look up.
    :return: Retrieves the string parameter if the value is not None.
             If it is None it returns an empty string.
    """
    value = ''
    if has_value(element, param_name):
        value = element.LookupParameter(param_name).AsString()
        if not isinstance(value, basestring):
            value = ''
    return value

def prmset(element, param_name, value):
    param = element.LookupParameter(param_name)
    if param != None and not param.IsReadOnly:
        param.Set(value)

def getSheetProp(data,key,keyname,propname):
    #print (data)
    #print (key)
    #print (keyname)
    #print (propname)
    key_row = next((row for row in data if row[keyname] == key), None)
    #print (key_row)
    if key_row:
        # Access to other column values
        key_value = key_row[propname] 
        #print(key_value)
        return key_value
    else:
        #print("No result found.")
        return false

if __shiftclick__:
    sheets = fec(doc).OfCategory(bic.OST_Sheets).ToElements()
    EleNums = len(sheets)
else:
    active_view = uidoc.ActiveView
    sheets = [active_view] if active_view.Category.Id.IntegerValue == int(bic.OST_Sheets) else []
    EleNums = len(sheets)
    
    if EleNums == 0:
        print("This is not a sheet. Open a sheet or hold Shift-Key to process.")
        sys.exit()


output = script.get_output()
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

#####################################

# Projektparameter auslesen
#doc = revit.doc
project_info = doc.ProjectInformation

excel_filename = get_str_param(project_info, 'ExcelReader-FileName')
excel_rangename = get_str_param(project_info, 'ExcelReader-DataName')
excel_keyname = get_str_param(project_info, 'ExcelReader-KeyName')

if not excel_filename or not excel_rangename or not excel_keyname:
    print("Project parameters missing!")
    sys.exit()
#print (excel_filename)
#print (excel_rangename)
#print (excel_keyname)
print ("File: {}\n Range:{} KeyName:{}".format(excel_filename, excel_rangename, excel_keyname))

if not excel_filename or not os.path.exists(excel_filename):
    print("Excel file not found or no path specified!")
    sys.exit()

if not excel_rangename or not excel_keyname:
    print("Data range name or key parameter not defined!")
    sys.exit()

fileCsvDrawinglist = os.path.dirname(excel_filename) + r'\temp-drawing-list.csv'

batch_file = '{}'.format(os.path.dirname(__file__) + '\convert3.bat')

# Convert xls to temp csv files.
os.system('"{}" "{}" "{}" "{}"'.format(batch_file, excel_filename, excel_rangename, fileCsvDrawinglist))

if not os.path.exists(fileCsvDrawinglist):
    print('Creating the CSV file failed.')
    sys.exit()

# Get sheet data from csv file.
data = []
with open(fileCsvDrawinglist) as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        data.append(row)

if all(excel_keyname in row for row in data):
    key_row = next((row for row in data if row[excel_keyname] == 0), None)
else:
    print("Der key '{}' not exist in csv data".format (excel_keyname))
    sys.exit()



t = rdb.Transaction(doc, 'Sync Drawinglist')
t.Start()

print('Syncing Drawinglist ...')
EleNums = len(fec(doc).OfCategory(bic.OST_Sheets).ToElements())-1

#for sheet in fec(doc).OfCategory(bic.OST_Sheets).ToElements():
for sheet in sheets:
    idx+=1
    nr = get_str_param(sheet, 'Sheet Number')
    
    if str(rdb.WorksharingUtils.GetCheckoutStatus(doc, sheet.Id)) == 'OwnedByOtherUser':
        #print('Sheet ' + nr + ' is used by another user and will not be updated!')
        print output.linkify(sheet.Id) + " " + nr + " is used by another user and will not be updated!" 

    else:
        print output.linkify(sheet.Id) + " " + nr 
        for header in data[0]:
            prmset(sheet, header, getSheetProp(data,nr,excel_keyname,header))
            
    output.update_progress(idx, EleNums)
    
t.Commit() 

# Delete temp csv.
time.sleep(2)
os.remove(fileCsvDrawinglist)

output.reset_progress()

#happy to be here
print('Done')
stopwatch.Stop()

timespan = stopwatch.Elapsed

print(timespan)