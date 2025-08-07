# -*- coding: utf-8 -*-

#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter
from System.Collections.Generic import List
from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ElementMulticategoryFilter
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import IFailuresPreprocessor
from Autodesk.Revit.DB import TransactionGroup, Transaction
from Autodesk.Revit.DB import XYZ, ElementTransformUtils
from Autodesk.Revit.DB import BuiltInFailures, FailureSeverity, FailuresAccessor 

class failureProcess(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        fail_acc_list = failuresAccessor.GetFailureMessages().GetEnumerator()
        for failure in fail_acc_list:
            #print failure.ToString()
			fSeverity = failuresAccessor.GetSeverity()
			if fSeverity == fSeverity.Warning:
				failuresAccessor.DeleteWarning(failure)
			else:
				failuresAccessor.ResolveFailure(failure)
				return FailureProcessingResult.ProceedWithCommit
			
        return FailureProcessingResult.Continue
        
__doc__ = 'Replace '\
          'rooms '\
          'on her level'

output = script.get_output()

idx = 0

AllElements = []

cat_list = [BuiltInCategory.OST_Rooms]

typed_list = List[BuiltInCategory](cat_list)

filterCategories = ElementMulticategoryFilter(typed_list)


AllElements =   FilteredElementCollector(revit.doc)\
            .WherePasses(filterCategories)\
            .WhereElementIsNotElementType()
                    
EleNums =   AllElements.ToElementIds().Count         

print EleNums

#try:
transGroup = TransactionGroup(revit.doc, "pyScript replace Rooms")
transGroup.Start()

for room in AllElements:
    print output.linkify(room.Id)
    rmPT = room.Location
    roomPoint = rmPT.Point
    level = room.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL)
    pinned = room.Pinned 
    SourceElement = room.Document.GetElement(level.AsElementId())
    
    offset = room.LimitOffset
    
    trans1 = Transaction(revit.doc, "Unplace Room")
    trans1.Start()
    
    #failOptions = trans1.GetFailureHandlingOptions()
    #failOptions.SetFailuresPreprocessor(failureProcess())
    #failOptions.SetFailuresPreprocessor(UnplacedRoomWarning())
    #trans1.SetFailureHandlingOptions(failOptions)
    room.Pinned = False
    room.Unplace()
    #print "unplaced"

    trans1.Commit()

    trans2 = Transaction(revit.doc, "New Room")
    trans2.Start()

    #level = cboLevel.SelectedValue
    topo = revit.doc.get_PlanTopology(SourceElement)
    circuits = topo.Circuits

    for circuit in circuits:
        #print "new room"
    
        if not circuit.IsRoomLocated:
            #print "place"
            #print room.Location
            newRoom = revit.doc.Create.NewRoom(room, circuit)
            newRoom.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).Set(SourceElement.Id)
            newRoom.LimitOffset = offset
            move = XYZ(roomPoint.X - rmPT.Point.X, roomPoint.Y - rmPT.Point.Y, rmPT.Point.Z)
            ElementTransformUtils.MoveElement(revit.doc, newRoom.Id, move)
            newRoom.Pinned = pinned
            break
        
    trans2.Commit()
    
    output.update_progress(idx, EleNums)
    idx+=1   

transGroup.Assimilate()

#except Exception:
#    raise Exception('Critical Error')

output.reset_progress()

#happy to be here
print('Done')
