import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import AreaVolumeSettings 
from Autodesk.Revit.DB import SpatialElementBoundaryOptions
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils
from Autodesk.Revit.DB import BuiltInParameter, HostObjectUtils, HostObject
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, SpatialElementType

from System.Diagnostics import Stopwatch
from System.Collections.Generic import List

from rpw import db, doc, uidoc
from pyrevit import script, DB, forms, revit

#we need two parameters
#one as length to write the roomheigth
#the second as yes/no to prevent the first one to overwrite it with this script
#so we have the opportunity to write the height value manually

# this Yes/ No parameter protect a room to overwrite by this script
#manual_room_materials_parameter_name = "11xxx Materialis. manuell"

#this is the source parameter name
room_source_parameter_name = "Raumnummer"

__title__ = 'Copy\nRaumnummer'

__doc__ = "Runs through all rooms or the preselcted one,\n"\
          "copy the ""Raumnummer"" valu to the Nummer \n"\
          "of a room.\n"\

#set to True for more detailed messages
DEBUG = False
if __shiftclick__:
    DEBUG = True

def param_exists_by_name(any_room, parameter_name):
    return any_room.LookupParameter(parameter_name)

#######################################

def owned_by_other_user(elem):
#from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

    # Checkout Status of the element
    checkout_status = WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
    if checkout_status == CheckoutStatus.OwnedByOtherUser:
        return True
    else:
        return False

########################################################################  
    
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

output = script.get_output()

rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()

selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

#start transaction
with db.Transaction("write room data"):
    #if there was an slection we use them instead
    if selection:
        rooms = selection

    #for the process bar on top
    EleNums = rooms.Count         

    if EleNums ==0:
        forms.alert('There are no rooms to work with.')
        script.exit()
    
    #just on room to check if the parameters attached
    #TODO: check an other way to determine parameters
    room = rooms[0]
    
    if not param_exists_by_name(room, room_source_parameter_name):
        forms.alert("The parameter " + room_source_parameter_name + " is necessary but does not exist. Please create them first.")
        script.exit()
            
            
                
    #room by room
    for room in rooms:
        
        #check if someone else has the room, otherwise an error will thrown
        #TODO maybe this throw an error in a non worksharing model
        if owned_by_other_user(room):
            print(" Room number: " + room.Number + output.linkify(room.Id) + " -> OwnedByOtherUser") 
            continue
    
        #rooms where not placed as 0 area
        if not room.Area > 0:
            continue
        
    
        #there is a parameter that protects the room from being overwritten by this script
        #manual = room.LookupParameter(manual_room_materials_parameter_name).AsInteger()

        #if manual:
        #    print(" Room number: " + room.Number  + " " + output.linkify(room.Id) + " -> manual")
        #    continue

        #else:

        source_parameter_value = room.LookupParameter(room_source_parameter_name).AsString()
        if DEBUG:
        
            room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
            print(" Room number: " + room.Number + " " + output.linkify(room.Id) + " " + room_name)

        if source_parameter_value:
            # Kopieren Sie den Wert in die Raum-Nummer
            room.Number = source_parameter_value

        output.update_progress(idx, EleNums)
        idx+=1   

output.reset_progress()

print(53 * "=")
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
