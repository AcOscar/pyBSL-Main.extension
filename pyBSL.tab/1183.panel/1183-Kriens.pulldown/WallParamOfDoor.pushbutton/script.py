"""Sets the value of ART_Material parameter from walls to the Wandtyp parameters of the door"""
#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter

from Autodesk.Revit.DB import Element
from pyrevit import revit, DB
from pyrevit import script
import sys

output = script.get_output()

#parameter to write
#wallthicknesparametername = 'Wanddicke'
ParameterNameToWrite = 'Wandtyp'
ParameterNameToRead = 'ART_Material'


if revit.active_view:

	doors = DB.FilteredElementCollector(revit.doc)\
		.OfCategory(DB.BuiltInCategory.OST_Doors)\
		.WhereElementIsNotElementType()

	for dr in doors:

		with revit.Transaction("Set Parameter"):

			wallhost = dr.Host

			if wallhost is not None:
				
				wallType = wallhost.WallType
				#width = wallType.Width
				#parametervaluestr = dr.Host.Name
				
				dparamread = wallType.LookupParameter(ParameterNameToRead)
				writevalue = dparamread.AsString()
				
				#dparam = dr.LookupParameter(wallthicknesparametername)
				
				#if dparam.StorageType == DB.StorageType.Double:
					
					#dparam.Set(width * 304.8)
				
				dparam = dr.LookupParameter(ParameterNameToWrite)
				
				if dparam.StorageType == DB.StorageType.String:
					#dparam.SetValueString(writevalue)
					if writevalue is not None:

						dparam.Set(writevalue.ToString())
                    
					#print (dr.GUID)
					print (writevalue)
print('Done')
