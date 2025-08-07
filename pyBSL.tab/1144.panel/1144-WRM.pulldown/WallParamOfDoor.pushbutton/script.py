"""Sets the value of Wandkonstruktion_T parameter from walls to the Wandkonstruktion parameters of the door"""
#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from rpw import doc
from System.Diagnostics import Stopwatch

from Autodesk.Revit.DB import Element
from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInCategory as Bic
stopwatch = Stopwatch()
stopwatch.Start()

output = script.get_output()

#parameter to write
#wallthicknesparametername = 'Wanddicke'
ParameterNameToWrite = 'Wandkonstruktion'
ParameterNameToRead = 'Wandkonstruktion_T'

selection = revit.get_selection()
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
    doornums = len(doors) 
    idx = 0
    for door in doors:

		with revit.Transaction("Set Parameter"):

			wallhost = door.Host
			#print('{}'.format(output.linkify(door.Id)))

			if wallhost is not None:
			
				if isinstance(wallhost, DB.Wall):  
					# Ihr bisheriger Code geht hier weiter

					wallType = wallhost.WallType
					#width = wallType.Width
					#parametervaluestr = door.Host.Name
					
					dparamread = wallType.LookupParameter(ParameterNameToRead)
					writevalue = dparamread.AsString()
					
					#dparam = door.LookupParameter(wallthicknesparametername)
					
					#if dparam.StorageType == DB.StorageType.Double:
						
						#dparam.Set(width * 304.8)
					
					dparam = door.LookupParameter(ParameterNameToWrite)
					
					if dparam.StorageType == DB.StorageType.String:
						#dparam.SetValueString(writevalue)
						if writevalue is not None:

							dparam.Set(writevalue.ToString())
						
						#print (door.GUID)
						#print (writevalue)
						print('{} -> {}'.format(output.linkify(door.Id),writevalue))
						
		output.update_progress(idx, doornums)
		idx+=1                           
								
output.reset_progress()
					

print('Done')
stopwatch.Stop()
timespan = stopwatch.Elapsed
print(timespan)