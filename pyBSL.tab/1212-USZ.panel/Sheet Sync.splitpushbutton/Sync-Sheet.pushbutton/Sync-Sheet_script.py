# -*- coding: utf-8 -*-
from pyrevit import script
from System.Diagnostics import Stopwatch
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms

import codecs
import os, time, csv, re, sys, collections, datetime
import Autodesk.Revit.DB as rdb

__title__ = 'Sync sheet'

__doc__ = 'Sync sheet informations'\
          ' with the drawing list'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Mapping deutsch/englisch zu BuiltInParameter
PARAMETER_MAPPING = {
    # SHEET_NUMBER
    "Plannummer": BuiltInParameter.SHEET_NUMBER,
    "Sheet Number": BuiltInParameter.SHEET_NUMBER,

    # SHEET_NAME
    "Planname": BuiltInParameter.SHEET_NAME,
    "Sheet Name": BuiltInParameter.SHEET_NAME,

    # SHEET_APPROVED_BY
    "Bestätigt von": BuiltInParameter.SHEET_APPROVED_BY,
    "Approved By": BuiltInParameter.SHEET_APPROVED_BY,

    # SHEET_CHECKED_BY
    "Geprüft von": BuiltInParameter.SHEET_CHECKED_BY,
    "Checked By": BuiltInParameter.SHEET_CHECKED_BY,

    # SHEET_DRAWN_BY
    "Gezeichnet von": BuiltInParameter.SHEET_DRAWN_BY,
    "Drawn By": BuiltInParameter.SHEET_DRAWN_BY,

    # SHEET_DESIGNED_BY
    "Entworfen von": BuiltInParameter.SHEET_DESIGNED_BY,
    "Designed By": BuiltInParameter.SHEET_DESIGNED_BY,

    # SHEET_SCALE
    #is READONLY
    #"Maßstab": BuiltInParameter.SHEET_SCALE,
    #"Scale": BuiltInParameter.SHEET_SCALE,

    # SHEET_REVISION
    "Revisionsnummer": BuiltInParameter.SHEET_REVISIONS_ON_SHEET,
    "Revision": BuiltInParameter.SHEET_REVISIONS_ON_SHEET,

    # SHEET_FILE_PATH
    "Dateipfad": BuiltInParameter.SHEET_FILE_PATH,
    "File Path": BuiltInParameter.SHEET_FILE_PATH,

    # SHEET_GUIDE_GRID
    "Hilfslinienraster": BuiltInParameter.SHEET_GUIDE_GRID,
    "Guide Grid": BuiltInParameter.SHEET_GUIDE_GRID,

    # SHEET_ISSUE_DATE
    "Planausgabedatum": BuiltInParameter.SHEET_ISSUE_DATE,
    "Sheet Issue Date": BuiltInParameter.SHEET_ISSUE_DATE,
}

def file_exists(path):
    """
    Prüft, ob der gegebene Pfad existiert.
    - Lokale oder UNC-Pfade mit os.path.exists
    - HTTP/HTTPS-URLs (z.B. SharePoint) per HEAD-Request (DefaultNetworkCredentials)
    """
    # 1. Lokale Datei
    if not path.lower().startswith(("http://", "https://")):
        return os.path.exists(path)

    # 2. SharePoint-URL
    try:
        print ("sharepoint")
        # IronPython: .NET–Klassen importieren
        import clr
        clr.AddReference("System")
        from System import Uri
        from System.Net import HttpWebRequest, CredentialCache

        uri = Uri(path)
        req = HttpWebRequest.Create(uri)
        req.Method = "HEAD"
        # Anmeldung mit den aktuellen Windows-Credentials
        req.Credentials = CredentialCache.DefaultNetworkCredentials

        resp = req.GetResponse()
        resp.Close()
        return True

    except Exception as e:
        # z.B. 404, Auth-Fehler, Timeout …
        # Debug-Ausgabe optional:
        print("SharePoint check failed:", e)
        return False

def has_value(element, param_name):
    if exists(element, param_name):
        return (element.LookupParameter(param_name).HasValue)

def exists(element, param_name):
    return (element.LookupParameter(param_name) != None)
    
#def get_str_param(element, param_name):
#    """
#    :param element: Element that holds the parameter.
#    :param param_name: Name of the parameter to look up.
#    :return: Retrieves the string parameter if the value is not None.
#             If it is None it returns an empty string.
#    """
#    value = ''
#    if has_value(element, param_name):
#        value = element.LookupParameter(param_name).AsString()
#        if not isinstance(value, basestring):
#            value = ''
#    return value
def get_str_param(element, param_name):
    """
    Holt den Stringwert eines Parameters mit Unterstützung für BuiltInParameter-Mapping.
    Gibt '' zurück, falls kein Wert vorhanden ist.
    """
    bip = PARAMETER_MAPPING.get(param_name)
    param = None

    if bip:
        param = element.get_Parameter(bip)
        if __forceddebugmode__:
            print("Mapping '{}' to {}".format(param_name, bip))
    else:
        param = element.LookupParameter(param_name)

    if param and param.HasValue:
        try:
            value = param.AsString()
            if isinstance(value, basestring):
                return value
        except Exception as e:
            if __forceddebugmode__:
                print("Fehler beim Lesen des Parameters '{}': {}".format(param_name, e))
    return ''
    
def prmset(element, param_name, value):
    """
    Setzt einen Parameter auf ein Element, entweder per BuiltInParameter-Mapping oder via LookupParameter.
    """
    # Versuche über BuiltInParameter
    bip = PARAMETER_MAPPING.get(param_name)
    param = None

    if bip:
        param = element.get_Parameter(bip)
        if __forceddebugmode__:
            print("Maping '{}' to {}".format (param_name,bip))
            
        #if bip == BuiltInParameter.SHEET_NAME:
        #    if re.search(r"[\\:{}\[\]¦;<>?`~]", value):
        #    print("⚠: Ungültige Zeichen gefunden ->", value)
    else:
        # Fallback zu LookupParameter, falls kein Mapping existiert
        param = element.LookupParameter(param_name)

    if param and not param.IsReadOnly:
        if __forceddebugmode__:
            print("Set parameter '{}' to value: {}".format (param.Definition.Name,value))
            
       
        try:
            param.Set(value)
        except Exception as e:
            print("Fehler beim Setzen des Parameters '{}': {}".format (param_name,e))

    else:
        if __forceddebugmode__:
            print("Parameter '{param_name}' not found or read-only.")
            
def getSheetProp(data,key,keyname,propname):
    """
    :param data: eine Liste von Dictionaries (aus der CSV bzw Excel Sheet)
    :param key: der Wert, nach dem gesucht wird.
    :param keyname: der Name des Keys (Spaltennamens), in dem key gesucht wird.
    :param propname: der Name des Keys (Spaltennamens), dessen Wert zurückgegeben werden soll.
    """
    #print (data)
    if __forceddebugmode__:
        print ("key: " + key)
        print ("keyname: " + keyname)
        print ("propname: " + propname)
    
    #erstes Dictionary (row), bei dem der Wert unter keyname gleich key ist
    key_row = next((row for row in data if row[keyname] == key), None)
    if key_row:

        # Access to other column values
        key_value = key_row[propname] 
        if __forceddebugmode__:
            print ("key_row: " , key_row)
            print ("key_value: " + key_value)
        return key_value
    else:
        if __forceddebugmode__:
            print("No key_value found.")
        return False
        
################################################ 
       
if __shiftclick__:
    active_view = uidoc.ActiveView
    sheets = []
    if active_view is not None and active_view.Category is not None:
        if active_view.Category.Id.IntegerValue == int(BuiltInCategory.OST_Sheets):
            sheets = [active_view]
    sheets = [active_view] if active_view.Category.Id.IntegerValue == int(BuiltInCategory.OST_Sheets) else []

else:
    sheets = forms.select_sheets(title="Select Sheets to sync from Excel", button_name="Excel Sync")
    if not sheets:
        sys.exit()
    
    #sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
    
EleNums = len(sheets)

if EleNums == 0:
    print("No sheets selected!")
    sys.exit()

output = script.get_output()
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

#####################################

# Projektparameter auslesen
project_info = doc.ProjectInformation

excel_filename = get_str_param(project_info, 'ExcelReader-FileName')
excel_rangename = get_str_param(project_info, 'ExcelReader-DataName')
excel_keyname = get_str_param(project_info, 'ExcelReader-KeyName')

if not excel_filename or not excel_rangename or not excel_keyname:
    print("Project parameters missing! Please add the paramters ExcelReader-FileName, ExcelReader-DataName and ExcelReader-KeyName as text parameters to the projectinformations.")
    sys.exit()
print ("File: {}\nRange: {}\nKeyName: {}".format(excel_filename, excel_rangename, excel_keyname))

#if not os.path.exists(excel_filename):
#    print("Excel file not found!")
#    sys.exit()
#if not file_exists(excel_filename):
#    print("Excel file not found!")
#    sys.exit()
    
#fileCsvDrawinglist = os.path.dirname(excel_filename) + r'\temp-drawing-list.csv'

temp_dir = os.environ.get('TEMP')

fileCsvDrawinglist = os.path.join(temp_dir,'temp-drawing-list.csv')
print fileCsvDrawinglist
batch_file = '{}'.format(os.path.dirname(__file__) + '\convert3.bat')

# Convert xls to temp csv files.
os.system('"{}" "{}" "{}" "{}"'.format(batch_file, excel_filename, excel_rangename, fileCsvDrawinglist))

if not os.path.exists(fileCsvDrawinglist):
    print('Creating the CSV file failed.')
    sys.exit()

# Get sheet data from csv file.
data = []
        
with codecs.open(fileCsvDrawinglist, "r", "utf-16") as f:
    clean_lines = (line.replace(u'\x00', u'') for line in f)
    reader = csv.DictReader(clean_lines, delimiter=",", quotechar='"')
    for row in reader:
        data.append(row)
        
if all(excel_keyname in row for row in data):
    key_row = next((row for row in data if row[excel_keyname] == 0), None)
else:
    print("The key '{}' not exist in csv data.".format (excel_keyname))
    sys.exit()

t = rdb.Transaction(doc, 'Sync Drawinglist')
t.Start()

print('Syncing Drawinglist ...')
EleNums = len(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements())-1

#for sheet in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements():
for sheet in sheets:
    idx+=1
    #nr = get_str_param(sheet, 'Sheet Number')
    #print (excel_keyname)
    
    nr = get_str_param(sheet, excel_keyname)
    #print (nr)

    key_value = get_str_param(sheet, excel_keyname)
    #print (key_value)

    if str(rdb.WorksharingUtils.GetCheckoutStatus(doc, sheet.Id)) == 'OwnedByOtherUser':
        #print('Sheet ' + nr + ' is used by another user and will not be updated!')
        print output.linkify(sheet.Id) + " " + nr + " is used by another user and will not be updated!" 

    else:
        print output.linkify(sheet.Id) + " " + nr 
        # erste Zeile 
        for header in data[0]:
            sheetprop = getSheetProp(data,key_value,excel_keyname,header)
            if sheetprop:
                prmset(sheet, header, sheetprop)
            #else:
            #    print ("no ", header)
            
    output.update_progress(idx, EleNums)
    
t.Commit() 

if not __forceddebugmode__:
    # Delete temp csv.
    time.sleep(2)
    os.remove(fileCsvDrawinglist)

output.reset_progress()

#happy to be here
print('Done')
stopwatch.Stop()

timespan = stopwatch.Elapsed

print(timespan)