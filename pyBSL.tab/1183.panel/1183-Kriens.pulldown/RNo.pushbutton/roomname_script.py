#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter

from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInParameter

__doc__ = 'Set the Parameter "Number" of a room '\
            'in the format like "B4.3.02.-.TH"'\
          'where "B4" is builduing name from the project indformatio'\
          '3 is the name of the an enclosing scopebox with an name like TRH_B4.3 '\
          'LI is an abbreviation given from LM'
def get_kuerzel(Raumname):
    Raumname = Raumname[0:4]
    #print(Raumname)
    if Raumname == "":
        return "-"
    elif Raumname == "Kell":
        return "K.."
    elif Raumname == "Lift":
            return "LI"
    elif Raumname == "Redu":
            return "RA"
    elif Raumname == "TH B":
            return "TH"
    elif Raumname == "Trep":
            return "TH"
    elif Raumname == "Vorr":
            return "RA"
    else:
        return "-"

output = script.get_output()

#Names of the Scopboxes that are used 
scopeboxesnameprefix = 'TRH_'

#parameter to write
parametername = 'Number'
parametername2 = 'Name'

projectInfo = revit.doc.ProjectInformation

biparam = DB.BuiltInParameter.PROJECT_BUILDING_NAME
Gebbezeichnung = projectInfo.Parameter[biparam].AsString()

print(Gebbezeichnung)

if revit.active_view:

    scopeboxes = DB.FilteredElementCollector(revit.doc)\
                 .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)

    if scopeboxes.GetElementCount() == 0:
        print('There are no scopeboxes.')
        sys.exit()

    roomids = DB.FilteredElementCollector(revit.doc)\
             .OfCategory(DB.BuiltInCategory.OST_Rooms)\
             .WhereElementIsNotElementType()\
             .ToElementIds()
    # doorids = DB.FilteredElementCollector(revit.doc)\
              # .OfCategory(DB.BuiltInCategory.OST_Doors)\
              # .OfCategory(DB.BuiltInCategory.OST_Doors)\
              # .WhereElementIsNotElementType()\
              # .ToElementIds()

    if roomids.Count == 0:
        print('There are no doors.')
        sys.exit()

    for sc in scopeboxes:
        #if sc.Name in namedscopeboxes:
        if sc.Name.startswith(scopeboxesnameprefix):
            #print('\tScopebox: {}'.format(sc.Name.ljust(30)))

            bb = sc.get_BoundingBox(revit.active_view)
            outline = DB.Outline(bb.Min, bb.Max)
            filter = DB.BoundingBoxIsInsideFilter(outline)
            roomcollector = DB.FilteredElementCollector(revit.doc, roomids).WherePasses(filter)

            with revit.Transaction("Set Parameter by Scopebox"):

                for room in roomcollector:
                        dparam = room.LookupParameter(parametername)
                        dparam2 = room.LookupParameter(parametername2)
                        level = room.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL)

                        #print(level.AsValueString()[0:2])
                        #print(dparam2.AsString())
                        if dparam.StorageType == DB.StorageType.String:
                            #dparam.Set(sc.Name )
                            parametervalue = Gebbezeichnung + "." + sc.Name[6:7] + "." + level.AsValueString()[0:2] + ".-." + get_kuerzel(dparam2.AsString())
                            print(Gebbezeichnung + "." + sc.Name[6:7] + "." + level.AsValueString()[0:2] + ".-." + get_kuerzel(dparam2.AsString()))
                            #print(output.linkify(room.Id))
                            #parametervalue += '_'
                            #parametervalue += room.Level.Name.split('_')[1]
                            
                            dparam.Set(parametervalue)

    #                    dparam = rm.LookupParameter('Comments')
    #                    if dparam.StorageType == DB.StorageType.String:
    #                        dparam.Set('')
    #                print(rm.Level.Name.split('_')[1])

print('Done')
