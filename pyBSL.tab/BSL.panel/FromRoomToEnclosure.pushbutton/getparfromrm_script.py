# -*- coding: utf-8 -*-

#pylint: disable=invalid-name,import-error,superfluous-parens
from System.Collections.Generic import List
from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ElementMulticategoryFilter
#from Autodesk.Revit.DB import ElementId
#from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import TransactionGroup, Transaction
from Autodesk.Revit.DB import SpatialElementGeometryCalculator 
from Autodesk.Revit.DB.Architecture import Room

__title__ = 'Get Paramters\nfrom Room'

__doc__ = 'Writes parameters '\
          'from a room '\
          'to all enclosing elements'

output = script.get_output()

idx = 0

AllElements = []

cat_list = [BuiltInCategory.OST_Rooms]

typed_list = List[BuiltInCategory](cat_list)

filterCategories = ElementMulticategoryFilter(typed_list)

selection = revit.get_selection()

#check if was something pre selected
if not selection.is_empty:
                
    for selel in selection.elements:
        if selel.GetType() == Room:
            AllElements.append(selel)
   
    EleNums = AllElements.Count 
    
    print "Working with pre selected element(s)."
    
else:
    print "There was nothing preselected."
    print "Try to work with all elements"
    
    #ok catch all what we get
    AllElements =   FilteredElementCollector(revit.doc)\
                    .WherePasses(filterCategories)\
                    .WhereElementIsNotElementType()
                    
    EleNums = AllElements.ToElementIds().Count         

print EleNums

#try:
calculator = SpatialElementGeometryCalculator(revit.doc)

transGroup = TransactionGroup(revit.doc, "pyScript")
transGroup.Start()

for room in AllElements:
    
    if room.Area == 0 :
        print "room not placed"
        continue
        
    Geschoss = room.LookupParameter("Geschoss").AsString() 
    if Geschoss is None:
        Geschoss = ""
        print "Geschoss is None!"
    Teilobjekt = room.LookupParameter("Teilobjekt").AsString()
      
    if Teilobjekt is None:
        Teilobjekt = ""
        print "Teilobjekt is None!"

    print output.linkify(room.Id) + " " + room.Number.ToString() + " "  + Geschoss + " " + Teilobjekt

    ### Calculate a room's geometry and find its boundary faces
    results = calculator.CalculateSpatialElementGeometry(room)## compute the room geometry 
    roomSolid = results.GetGeometry()##get the solid representing the room's geometry
    myelemids = []
    
    for face in roomSolid.Faces:
    
        subfaceList = results.GetBoundaryFaceInfo(face)# get the sub-faces for the face of the room
        
        for subface in subfaceList:
            
            elemid = subface.SpatialBoundaryElement.HostElementId

            myelemids.append(elemid)
            
    #remove all duplicates
    myelemids = list(set(myelemids))
    
    trans2 = Transaction(revit.doc, "Room")
    trans2.Start()

    for elemid in myelemids:

        elem = revit.doc.GetElement(elemid)

        elemparamGeschoss = elem.LookupParameter("Geschoss")
        elemparamTeilobjekt = elem.LookupParameter("Teilobjekt")

            

        print output.linkify(elemid) + " " + elem.Name 

        elemparamGeschoss.Set(Geschoss)
        
        if elemparamTeilobjekt.AsString() == "MIT1":
            print "keep MIT1"
            continue
        elemparamTeilobjekt.Set(Teilobjekt)
        
    trans2.Commit()

    output.update_progress(idx, EleNums)
    idx+=1   

transGroup.Assimilate()

#except Exception:
#    raise Exception('Critical Error')

output.reset_progress()

#happy to be here
print('Done')
