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

__doc__ = 'Set the Parameter Gebäudebezeichnung from Project information '\
          'to the value of Gebäudebezeichnung to every element and '\
          'the value of Teilobjekt and Geschoss from the hosted level paramter '\
          'to the element paramters with the same name'

def writeParamtersFromLevel(TargetElement, SourceElementID):
    #printvalue = output.linkify(TargetElement.Id) 
    print output.linkify(TargetElement.Id) 
     #get the host level
    SourceElement = TargetElement.Document.GetElement(SourceElementID)

    #print SourceElement
    #get the value for the Teilobjekt and Geschoss
    Teilobjekt = SourceElement.LookupParameter('Teilobjekt').AsString()
    Geschoss = SourceElement.LookupParameter('Geschoss').AsString()
    #Gebbezeichnung = SourceElement.LookupParameter('Gebäude').AsString()
     
    #get the target parameter
    Teilobjektparameter = TargetElement.LookupParameter('Teilobjekt')
    Geschossparameter = TargetElement.LookupParameter('Geschoss')
    Gebbezeichnungparameter = TargetElement.LookupParameter('Gebäude')

    #print('{}'.format(output.linkify(TargetElement.Id)))
    #write them only if it is not the right values


    if Teilobjektparameter.AsString() != Teilobjekt:
        print('{} Set Teilobjekt from {} to {}'.format(output.linkify(TargetElement.Id),Teilobjektparameter.AsString(),Teilobjekt))
        #print('Set Teilobjekt from {} to {}'.format(Teilobjektparameter.AsString(),Teilobjekt))

        Teilobjektparameter.Set(Teilobjekt)
    if Geschossparameter.AsString() != Geschoss:
        print('{} Set Geschoss from {} to {}'.format(output.linkify(TargetElement.Id),Geschossparameter.AsString(),Geschoss))
        #print('Set Geschoss from {} to {}'.format(Geschossparameter.AsString(),Geschoss))

        Geschossparameter.Set(Geschoss)
    if Gebbezeichnungparameter.AsString() != Gebbezeichnung:
        print('{} Set Gebäude from {} to {}'.format(output.linkify(TargetElement.Id),Gebbezeichnungparameter.AsString(),Gebbezeichnung))
        #print('Set Gebäude from {} to {}'.format(Gebbezeichnungparameter.AsString(),Gebbezeichnung))

        Gebbezeichnungparameter.Set(Gebbezeichnung)
    #print printvalue    


printvalue =''

output = script.get_output()

selection = revit.get_selection()

idx = 0

AllElements = []

#list with related BuiltInCategories
cat_list = [
BuiltInCategory.OST_Areas,\
BuiltInCategory.OST_Ceilings,\
BuiltInCategory.OST_Columns,\
BuiltInCategory.OST_CurtainWallMullionsCut,\
BuiltInCategory.OST_CurtainWallMullions,\
BuiltInCategory.OST_CurtainWallPanels,\
BuiltInCategory.OST_Doors,\
BuiltInCategory.OST_Floors,\
BuiltInCategory.OST_GenericModel,\
BuiltInCategory.OST_Joist,\
BuiltInCategory.OST_MechanicalEquipment,\
BuiltInCategory.OST_Parking,\
BuiltInCategory.OST_Railings,\
BuiltInCategory.OST_Ramps,\
BuiltInCategory.OST_Roofs,\
BuiltInCategory.OST_Rooms,\
BuiltInCategory.OST_Stairs,\
BuiltInCategory.OST_StructuralColumns,\
BuiltInCategory.OST_ShaftOpening,\
BuiltInCategory.OST_Topography,\
BuiltInCategory.OST_Walls,\
BuiltInCategory.OST_Windows]


typed_list = List[BuiltInCategory](cat_list)

filterCategories = ElementMulticategoryFilter(typed_list)

projectInfo = revit.doc.ProjectInformation

#Gebbezeichnung = projectInfo.LookupParameter('Gebäude').AsString()

biparam = DB.BuiltInParameter.PROJECT_BUILDING_NAME
Gebbezeichnung = projectInfo.Parameter[biparam].AsString()

print Gebbezeichnung


if revit.active_view:

    #check if was somthing pre selcted
    if not selection.is_empty:
                    
        AllElements =  selection 
        EleNums = len(selection)
        
        print "Working pre selected element(s)."
        
    else:
        print "There was nothing selected."
        print "Try to work with all elements"
        
        #ok catch all what we get
        AllElements =   FilteredElementCollector(revit.doc)\
                        .WherePasses(filterCategories)\
                        .WhereElementIsNotElementType()
                        
        EleNums =   AllElements.ToElementIds().Count         
    
    #start the transaction here, faster than permanetly open and close later
    with revit.Transaction("pyScript IFC-Parameter"):
        
        for element in AllElements:
            printvalue =''
            #print output.linkify(element.Id)
            #print element.Name
            #print element.Category.Name
            # if the elemente is not based on a level,we try different other things and and give a hint to the user
            if element.LevelId.Equals(ElementId.InvalidElementId):
                print "no level"
                #for parking elemnts we use the host level
                if element.Category.Name == "Parking":
                
                    writeParamtersFromLevel(element,element.Host.Id)
                    
                # for generic model we check if a schedle level exist
                elif element.Category.Name == "Generic Models":

                    binp = BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
                    Teilobjektparameter1 = element.get_Parameter(binp)
                    if Teilobjektparameter1 is not None:
                        writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())

                elif element.Category.Name == "Structural Trusses":
                    binp = BuiltInParameter.TRUSS_ELEMENT_REFERENCE_LEVEL_PARAM
                    Teilobjektparameter1 = element.get_Parameter(binp)
                    if Teilobjektparameter1 is not None:
                        writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
                        
                elif element.Category.Name == "Structural Framing":
                    binp = BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
                    Teilobjektparameter1 = element.get_Parameter(binp)
                    if Teilobjektparameter1 is not None:
                        writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())

                elif element.Category.Name == "Railings":
                    binp = BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
                    Teilobjektparameter1 = element.get_Parameter(binp)
                    if Teilobjektparameter1 is not None:
                        writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
                        
                    #if there is no references level we looking at the host element
                    elif not element.HostId.Equals(ElementId.InvalidElementId): 

                        stairhost = element.Document.GetElement(element.HostId)
                        
                        binp = BuiltInParameter.STAIRS_BASE_LEVEL_PARAM
                        Teilobjektparameter1 = stairhost.get_Parameter(binp)

                        if Teilobjektparameter1 is not None:

                            writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
                        
                elif element.Category.Name == "Stairs":
                    binp = BuiltInParameter.STAIRS_BASE_LEVEL_PARAM
                    Teilobjektparameter1 = element.get_Parameter(binp)

                    if Teilobjektparameter1 is not None:

                        writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
                
                        
                elif element.Category.Name == "Floor opening cut":
                    print "not posible for", element.Category.Name
                        
                elif element.Category.Name == "Scope Boxes":
                    print "not posible for", element.Category.Name
                elif element.Category.Name == "Curtain Wall Grids":
                    print "not posible for", element.Category.Name
                        
                        
                else:
                    printvalue +=('The element {} is not based on a level. '.format(element.Name))
                    
                    Teilobjektparameter = element.LookupParameter('Teilobjekt')
                    printvalue +=('Teilobjekt is: {} '.format(Teilobjektparameter.AsString()))
                    
                    Geschossparameter = element.LookupParameter('Geschoss')
                    printvalue +=('Geschoss is: {} '.format(Geschossparameter.AsString()))
                    
                    Gebbezeichnungparameter = element.LookupParameter('Gebäude')
                    printvalue +=('Building Name is: {}'.format(Gebbezeichnungparameter.AsString()))
                    print printvalue
                
            else:
                
                if element.Category.Name == "Rectangular Arc Wall Opening":
                    print "not posible for", element.Name
            
                elif element.Category.Name == "Rectangular Straight Wall Opening":
                    print "not posible for", element.Name
                elif element.Category.Name == "<Room Separation>":
                    print "not posible for", element.Name
                elif element.Category.Name == "Curtain Panels":
                    #print"hallo1"
                    #curtainwall = element.Document.GetElement(element.Host)
                    
                    #binp = BuiltInParameter.CURTAIN_WALL_PANEL_HOST_ID
                    #Teilobjektparameter1 = element.Host.get_Parameter(binp)

                    #if Teilobjektparameter1 is not None:
                    #print"hallo"
                    writeParamtersFromLevel(element,element.Host.Id)

                    
                    
                    
                else:
                    print element.Category.Name
                    writeParamtersFromLevel(element,element.LevelId)

            #the stylish orange progress bar
            output.update_progress(idx, EleNums)
            idx+=1   
        #for all elements
    #end transaction

output.reset_progress()

#happy to be here
print('Done')
