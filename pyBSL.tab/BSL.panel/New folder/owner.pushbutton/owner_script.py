import os
from collections import defaultdict
import clr
clr.AddReference("RevitAPI")
import Autodesk.Revit.UI
from System.Diagnostics import Stopwatch
from rph import prm
import sys
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import WorksharingUtils
#from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, DB, forms

stopwatch = Stopwatch()
stopwatch.Start()

doc = __revit__.ActiveUIDocument.Document

output = script.get_output()

# Get all elements in the model
elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

for element in elements:
    # In a Worksharing-enabled model, the "WorksharingTooltipInfo" provides information about creator, owner and last changed by user
    tooltip_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
    
    # If the element has an owner
    if tooltip_info.Owner != "":

        # Check if the element is owned by the current user
        if tooltip_info.Owner == doc.Application.Username:
            print("Revit", "Element ID: {0} is owned by you.".format(element.Id))
        else:
            print("Revit", "Element ID: {0} is owned by {1}.".format(element.Id, tooltip_info.Owner)+output.linkify(element.Id))
            #print(" Room number: " + room.Number  + " " + output.linkify(room.Id) + " -> manual")

print("456HdM_pyRevit rename PDFs and DWGs run in:")

stopwatch.Stop()
timespan = stopwatch.Elapsed
print(timespan)
