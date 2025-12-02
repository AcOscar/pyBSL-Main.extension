# -*- coding: utf-8 -*-

from pyrevit import revit
from pyrevit import script
from Autodesk.Revit.DB import TransactionGroup
from Autodesk.Revit.DB import XYZ, ElementTransformUtils
from rpw import db, doc, uidoc
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic

        
__doc__ = 'Replace '\
          'rooms '\
          'on her level'

output = script.get_output()

def main():

    idx = 0

    transGroup = TransactionGroup(revit.doc, "pyScript replace Rooms")
    transGroup.Start()

    selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds() if doc.GetElement(elId).Category.BuiltInCategory == Bic.OST_Rooms]

    """if was a selection, use that, otherwise get all rooms"""
    if selection:
        rooms = selection
    else:
        rooms = list(Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements())
    
    """check if we have rooms and the necessary parameters"""
    if not rooms:
        print("No rooms found.")
        return

    EleNums =   rooms.Count         

    print (EleNums)

    for room in rooms:
        print (output.linkify(room.Id))
        rmPT = room.Location
        roomPoint = rmPT.Point
        level = room.get_Parameter(Bic.ROOM_UPPER_LEVEL)
        pinned = room.Pinned 
        SourceElement = room.Document.GetElement(level.AsElementId())
        
        offset = room.LimitOffset
        
        with db.Transaction("Unplace Room"):
            if pinned:
                room.Pinned = False
            room.Unplace()

        with db.Transaction("New Room"):

            topo = revit.doc.get_PlanTopology(SourceElement)
            circuits = topo.Circuits

            for circuit in circuits:
            
                if not circuit.IsRoomLocated:
                    newRoom = revit.doc.Create.NewRoom(room, circuit)
                    newRoom.get_Parameter(Bic.ROOM_UPPER_LEVEL).Set(SourceElement.Id)
                    newRoom.LimitOffset = offset
                    move = XYZ(roomPoint.X - rmPT.Point.X, roomPoint.Y - rmPT.Point.Y, rmPT.Point.Z)
                    ElementTransformUtils.MoveElement(revit.doc, newRoom.Id, move)
                    if pinned:
                        newRoom.Pinned = pinned
                    break
                
        output.update_progress(idx, EleNums)
        idx+=1   

    transGroup.Assimilate()

    output.reset_progress()

    print('Done')

if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:        
        import traceback
        print("ERROR:", e)
        print(traceback.format_exc())