# -*- coding: utf-8 -*-

#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter
from System.Collections.Generic import List
from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInCategory, Group
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ElementMulticategoryFilter
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import BuiltInParameter
from System.Diagnostics import Stopwatch
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

__doc__ = 'Set the Parameter Gebäudebezeichnung from Project information '\
          'to the value of Gebäudebezeichnung to every element and '\
          'the value of Teilobjekt and Geschoss from the hosted level paramter '\
          'to the element paramters with the same name'
          
          
skipped = 0
newval = 0
nochange = 0
#changed = 0

def OwnedByOtherUser(doc, elem):
#from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

    # Checkout Status of the element
    checkout_status = WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
    #message += ("\nCheckout status : " + str(checkout_status))
    if checkout_status == CheckoutStatus.OwnedByOtherUser:
        return True
    else:
        return False
  
  
def writeParamtersFromLevel(TargetElement, SourceElementID):
    global skipped 
    global newval 
    global nochange 
    #global changed
    
    #get the host level
    SourceElement = TargetElement.Document.GetElement(SourceElementID)
    #info = TargetElement.Document.GetWorksharingTooltipInfo(TargetElement.Document, TargetElement.Id)
    #info =  revit.doc.GetWorksharingTooltipInfo(TargetElement.Document, TargetElement.Id)
    
    OwnedByOtherUser(TargetElement.Document, TargetElement)
    
    printvalue =''
    doprint = False
    if SourceElement is not None:

        #get the value for the Teilobjekt and Geschoss
        Teilobjekt = SourceElement.LookupParameter('Teilobjekt').AsString()
        if Teilobjekt is None:
            Teilobjekt = ""
        Geschoss = SourceElement.LookupParameter('Geschoss').AsString()
        if Geschoss is None:
            Geschoss = ""
        #print('{}'.format(output.linkify(TargetElement.Id)))
        #write them only if it is not the right values
        printvalue += 'Element {}'.format(output.linkify(TargetElement.Id))


        if OwnedByOtherUser(TargetElement.Document, TargetElement):
            doprint = True
            skipped += 1
            printvalue +=  " skipped; OwnedByOtherUser"
            
        else:
        
            #get the target parameter
            Teilobjektparameter = TargetElement.LookupParameter('Teilobjekt')
            Geschossparameter = TargetElement.LookupParameter('Geschoss')
            Gebbezeichnungparameter = TargetElement.LookupParameter('Gebäude')
            
            wastouched = False
            if Teilobjektparameter is not None:
            
                if Teilobjektparameter.AsString() != Teilobjekt:
                    #print('{} Set Teilobjekt from {} to {}'.format(output.linkify(TargetElement.Id),Teilobjektparameter.AsString(),Teilobjekt))
                    printvalue += ' Set Teilobjekt from {} to {}'.format(Teilobjektparameter.AsString(),Teilobjekt)
                    #print('Set Teilobjekt from {} to {}'.format(Teilobjektparameter.AsString(),Teilobjekt))
                    doprint = True
                    Teilobjektparameter.Set(Teilobjekt)
                    wastouched = True
                if Geschossparameter.AsString() != Geschoss:
                    #print('{} Set Geschoss from {} to {}'.format(output.linkify(TargetElement.Id),Geschossparameter.AsString(),Geschoss))
                    printvalue +=' Set Geschoss from {} to {}'.format(Geschossparameter.AsString(),Geschoss)
                    #print('Set Geschoss from {} to {}'.format(Geschossparameter.AsString(),Geschoss))
                    doprint = True
                    wastouched = True
                    Geschossparameter.Set(Geschoss)
                    
                if Gebbezeichnungparameter.AsString() != Gebbezeichnung:
                    #print('{} Set Gebäude from {} to {}'.format(output.linkify(TargetElement.Id),Gebbezeichnungparameter.AsString(),Gebbezeichnung))
                    printvalue +=' Set Gebäude from {} to {}'.format(Gebbezeichnungparameter.AsString(),Gebbezeichnung)
                    #print('Set Gebäude from {} to {}'.format(Gebbezeichnungparameter.AsString(),Gebbezeichnung))
                    doprint = True
                    wastouched = True
                    Gebbezeichnungparameter.Set(Gebbezeichnung)
                
            if wastouched:
                newval +=1
            else:
                nochange +=1
        #print printvalue    
        if doprint:
            print printvalue

def process_element(element):
    # Funktion zum Schreiben der Parameter (Vorhandene Logik)
    # ...
    #print (element)
    printvalue =''
    #print output.linkify(element.Id)
    #print element.Name
    #print element.Category.Name
    # if the elemente is not based on a level,we try different other things and and give a hint to the user
    if element.LevelId.Equals(ElementId.InvalidElementId):
        #print output.linkify(element.Id)


        #print "no level"
        #for parking elemnts we use the host level
        #if element.Category.Name == "Parking":
        if element.Category.Id.IntegerValue == -2001180:
        
            writeParamtersFromLevel(element,element.Host.Id)
            
        # for generic model we check if a schedle level exist
        #elif element.Category.Name == "Generic Models":
        elif element.Category.Id.IntegerValue == -2000151:
            binp = BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
            Teilobjektparameter1 = element.get_Parameter(binp)
            if Teilobjektparameter1 is not None:
                writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())

        #elif element.Category.Name == "Structural Trusses":
        elif element.Category.Id.IntegerValue == -2001336:
            binp = BuiltInParameter.TRUSS_ELEMENT_REFERENCE_LEVEL_PARAM
            Teilobjektparameter1 = element.get_Parameter(binp)
            if Teilobjektparameter1 is not None:
                writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
                
        #elif element.Category.Name == "Structural Framing":
        elif element.Category.Id.IntegerValue == -2001320:
            binp = BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
            Teilobjektparameter1 = element.get_Parameter(binp)
            if Teilobjektparameter1 is not None:
                writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())

        #elif element.Category.Name == "Railings":
        elif element.Category.Id.IntegerValue == -2000126:
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
                
        #elif element.Category.Name == "Stairs":
        elif element.Category.Id.IntegerValue == -2000120:
            binp = BuiltInParameter.STAIRS_BASE_LEVEL_PARAM
            Teilobjektparameter1 = element.get_Parameter(binp)

            if Teilobjektparameter1 is not None:

                writeParamtersFromLevel(element,Teilobjektparameter1.AsElementId())
        
                
        elif element.Category.Name == "Floor opening cut":
            print "not posible for", element.Category.Name
                
        #elif element.Category.Name == "Scope Boxes":
        elif element.Category.Id.IntegerValue == -2006000:
            print "not posible for", element.Category.Name
        #elif element.Category.Name == "Curtain Wall Grids":
        elif element.Category.Id.IntegerValue == -2000321:
            print "not posible for", element.Category.Name
                
                
        else:
            print output.linkify(element.Id)
            printvalue +=('The element {} is not based on a level. '.format(element.Name))
            
            Teilobjektparameter = element.LookupParameter('Teilobjekt')
            
            if Teilobjektparameter is not None:
                printvalue +=('Teilobjekt is: {} '.format(Teilobjektparameter.AsString()))
            
            Geschossparameter = element.LookupParameter('Geschoss')
            if Geschossparameter is not None:
                printvalue +=('Geschoss is: {} '.format(Geschossparameter.AsString()))
            Gebbezeichnungparameter = element.LookupParameter('Gebäude')
            
            if Geschossparameter is not None:
                printvalue +=('Building Name is: {}'.format(Gebbezeichnungparameter.AsString()))
                
            print printvalue
        
    else:
        #print "hallo3"
        if element.Category.Name == "Rectangular Arc Wall Opening":
            print "not posible for", element.Name
    
        elif element.Category.Name == "Rectangular Straight Wall Opening":
            print "not posible for", element.Name
        elif element.Category.Name == "<Room Separation>":
            print "not posible for", element.Name
            
        elif element.Category.Id.IntegerValue  == -2000170:
        #elif element.Category.Name == "Curtain Panels":
            #print output.linkify(element.Id)
            if hasattr(element,'Host'):
                writeParamtersFromLevel(element,element.Host.Id)
        #elif element.Category.Id.IntegerValue  == -2000035:
        #elif element.Category.Name == "Roofs":
            
        #   writeParamtersFromLevel(element,element.Host.Id)



        #elif element.Category.Name == "Windows":
        elif element.Category.Id.IntegerValue == -2000014:
            #print output.linkify(element.Id)
            #print "22  "
            
            if element.SuperComponent:
            
            #if not element.SuperComponent.Equals(None):
            
                #print "2"
                writeParamtersFromLevel(element,element.SuperComponent.Host.Id)
            else:
                #print "3"
                writeParamtersFromLevel(element,element.LevelId)
            
        else:
            writeParamtersFromLevel(element,element.LevelId)


def process_group(group):
    # Iteriere durch alle Elemente in der Gruppe und wende `process_element` an
    for member_id in group.GetMemberIds():
        member = revit.doc.GetElement(member_id)
        if member.Category is not None:
        
            #print ( member.Category.BuiltInCategory )
            #print (typed_list)
            if member.Category.BuiltInCategory in typed_list:
                process_element(member)

stopwatch = Stopwatch()
stopwatch.Start()


printvalue =''

output = script.get_output()

selection = revit.get_selection()

idx = 0

AllElements = []

#list with related BuiltInCategories
cat_list = [
#BuiltInCategory.OST_Areas,\
BuiltInCategory.OST_Ceilings,\
BuiltInCategory.OST_Columns,\
BuiltInCategory.OST_CurtainWallMullionsCut,\
BuiltInCategory.OST_CurtainWallMullions,\
BuiltInCategory.OST_CurtainWallPanels,\
BuiltInCategory.OST_Doors,\
BuiltInCategory.OST_Floors,\
BuiltInCategory.OST_Furniture,\
BuiltInCategory.OST_FurnitureSystems,\
BuiltInCategory.OST_GenericModel,\
BuiltInCategory.OST_Joist,\
BuiltInCategory.OST_MechanicalEquipment,\
BuiltInCategory.OST_Parts,\
BuiltInCategory.OST_PlumbingEquipment,\
BuiltInCategory.OST_PlumbingFixtures,\
BuiltInCategory.OST_Parking,\
BuiltInCategory.OST_Railings,\
BuiltInCategory.OST_Ramps,\
BuiltInCategory.OST_Roofs,\
#BuiltInCategory.OST_Rooms,\
BuiltInCategory.OST_SpecialityEquipment,\
BuiltInCategory.OST_Stairs,\
BuiltInCategory.OST_StructuralColumns,\
BuiltInCategory.OST_ShaftOpening,\
BuiltInCategory.OST_Topography,\
BuiltInCategory.OST_Walls,\
BuiltInCategory.OST_Windows,\
]

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
        categoryElements =   FilteredElementCollector(revit.doc)\
                        .WherePasses(filterCategories)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
                        
        
        groupElements = FilteredElementCollector(revit.doc)\
                        .OfClass(Group)\
                        .ToElements()
                
       
        AllElements.extend(categoryElements)
        AllElements.extend(groupElements)
        
        EleNums = len(AllElements)   
    
    #start the transaction here, faster than permanetly open and close later
    with revit.Transaction("pyScript IFC-Parameter"):
        
        for element in AllElements:
            # Überprüfen, ob das Element Teil einer Gruppe ist
            if isinstance(element, Group):
                # Behandlung für Gruppenelemente
                process_group(element)
            else:
                # Behandlung für nicht gruppierte Elemente
                process_element(element)        
            
        

            #the stylish orange progress bar
            output.update_progress(idx, EleNums)
            idx+=1   
        #for all elements
    #end transaction

output.reset_progress()

#happy to be here
if EleNums == 1:
    print('Done with 1 element.')
else:
    print('Done with {} elements.'.format(EleNums))
    print('{} elements has a new or a changed value.'.format(newval))
    print('{} elements was skipped because they owned by someone else.'.format(skipped))
    print('{} elements was not touched.'.format(nochange))
    
    
stopwatch.Stop()

timespan = stopwatch.Elapsed

print(timespan)