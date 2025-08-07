from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from System.Diagnostics import Stopwatch
from rpw import doc
from pyrevit import script
from pyrevit import revit, DB

#the door family needs a type parameter with the standard operation swing

#
#
__doc__ = 'the script will go through all door instances '\
          'and will check if the mirrored or not  '\
          'the mirrored doors will set to the opposite  '\
          'of the standard direction, otherwise '\
          'it get the standard value '
def get_str_type_param(element, param_name):
    value = element.LookupParameter(param_name).AsString()
    if not value:
        value = ""
    return value

stopwatch = Stopwatch()
stopwatch.Start()

#parameter name for the standard opening direction, should a type parameter inside of the door family definition
default_type_param = "door_opening_direction"

#parameter name for the opening direction, should a instance parameter of the door family definition
din_param = "Operation_I"


opening_side =     {"links": "SINGLE_SWING_RIGHT", "rechts": "SINGLE_SWING_LEFT", "DOUBLE_DOOR_DOUBLE_SWING": "Beide", "oben": "Oben", "ROLLINGUP": "NOTDEFINED"}
opening_side_mir = {"links": "SINGLE_SWING_LEFT", "rechts": "SINGLE_SWING_RIGHT", "beide": "DOUBLE_DOOR_DOUBLE_SWING", "oben": "Oben", "ROLLINGUP": "NOTDEFINED"}


output = script.get_output()
# get last phase
selection = revit.get_selection()

#print selection.element_ids.Count
# get last phase
phases = doc.Phases

#print (phases.Size)
phase = phases[phases.Size - 1]

doors = []
if revit.active_view:

    #if len(selection) == 1:
    if not selection.is_empty:
        for selel in selection.elements:
        
            if isinstance (selel, DB.FamilyInstance):
                #print selel.Category.Name
                if selel.Category.Name  == "Doors":
                    #print "HAH"
                    doors.append(selel)
    else:
        print "nothing preselected"
                
    if len(doors) == 0:
        print "get the whole model"
        doors = Fec(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()

    doornums = len(doors) 
    idx = 0
    #with rpw_Transaction('door aufschlag din'):
    with revit.Transaction("door aufschlag din"):

        for door in doors:
            # check for opening side
            door_type = doc.GetElement(door.GetTypeId())
            #print(door.Id.IntegerValue)
            try:
                default = get_str_type_param(door_type, default_type_param)
                if door.Mirrored:
                    door.LookupParameter(din_param).Set(opening_side_mir[default])
                    print('{} -> {}'.format(output.linkify(door.Id),opening_side_mir[default]))

                else:
                    door.LookupParameter(din_param).Set(opening_side[default])
                    print('{} -> {}'.format(output.linkify(door.Id),opening_side[default]))

            except:
                pass
    
            output.update_progress(idx, doornums)
            idx+=1                           
                            
output.reset_progress()

print(" updated in: ")

stopwatch.Stop()
timespan = stopwatch.Elapsed
print(timespan)
