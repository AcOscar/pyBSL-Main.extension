#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import ElementId

from pyrevit import revit, DB
from pyrevit import script

from rpw import doc

import sys

__doc__ = 'Set the Parameter "Einbauort" of a door '\
          'to the value of the Name of the Room (fromRoom) '\
          'continued by a , and a whitespace and the Name of the Room (toRoom) '\
          'finaly it attached the doornumber.'\
          'The doornumber is the number of the room (toRoom) and TuernummerImRaum. If '\
          'the TuernummerImRaum is empty, it will by used T1.'



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
        print "nothing selected"
                
    if len(doors) == 0:
        print "get the whole model"
        doors = Fec(doc).OfCategory(Bic.OST_Doors).WhereElementIsNotElementType().ToElements()
    #idx=0
    #for door in doors:
    #    idx +=1
    #    door_id = door.Id.IntegerValue
    #    print (door_id)
    #    print('{}: {}'.format(idx, output.linkify(door.Id)))



    #doors = DB.FilteredElementCollector(revit.doc)\
    #    .OfCategory(DB.BuiltInCategory.OST_Doors)\
    #    .WhereElementIsNotElementType()

    #doors.GetElementCount()
    #doornums = doors.GetElementCount() 
    doornums = len(doors) 
    idx = 0
    for door in doors:

        if door.HasSpatialElementFromToCalculationPoints == True:
            if door.ToRoom == None:
                print('failed')
                #if door.Room == None:
                #    print('failed again')
                #    doorFailed = door
                #    doorLevel = prm.get_valstr_param(door, 'Level')
                #    print('Level: ' + doorLevel + ' - Door ID could not find an adjacent room: {}'.format(out.linkify(door.Id)) + '. Not useful for dRofus')
            else: 
                doorToRoom = door.ToRoom[phase]
                doorFromRoom = door.FromRoom[phase]
                
                #if doorToRoom == None:
                #    doorFromRoom = door.FromRoom[phase]
                    #print('failed')
                #else:
                #print doorFromRoom.Number
                with revit.Transaction("Set Parameter"):
                    dparam = door.LookupParameter("Positionsnummer")
                    dtnrparam = door.LookupParameter("TuernummerImRaum")
                    EinbauortParam = door.LookupParameter("Einbauort")
                    
                    
                    #Einbauortstr = doorFromRoom.Name()
                    #Einbauortstr += ", " + doorToRoom.Name
                    if doorFromRoom is None:
                        Einbauortstr = ""
                    else:
                    
                        Einbauortstr = doorFromRoom.LookupParameter("Name").AsString()
                        
                        
                    if doorToRoom is not None:
                        if len(Einbauortstr) != 0:
                            Einbauortstr += ", "
                        Einbauortstr += doorToRoom.LookupParameter("Name").AsString()
                        doorNumber = doorToRoom.Number
                    else:
                        doorNumber = ""

                    if dparam.StorageType == DB.StorageType.String:
                        
                        #doorNumber = doorToRoom.Number

                        if dtnrparam is not None:
                            if dtnrparam.AsString() is None:
                                doorNumber += ".T1"
                            else:
                                doorNumber += ".T"
                                doorNumber += dtnrparam.AsString()
                        else:
                            doorNumber += ".T1"
                        print('{} : {} : {}'.format(output.linkify(door.Id),doorNumber,Einbauortstr))
                        #print doorNumber
                        
                        dparam.Set(doorNumber)
                        #if EinbauortParam is not None:
                        
                        EinbauortParam.Set(Einbauortstr)
                                               
        output.update_progress(idx, doornums)
        idx+=1                           
                            
output.reset_progress()
print('Done')
