# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import AreaVolumeSettings 
from Autodesk.Revit.DB import SpatialElementBoundaryOptions
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils
from Autodesk.Revit.DB import BuiltInParameter, HostObjectUtils, HostObject
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, SpatialElementType

from System.Diagnostics import Stopwatch
from System.Collections.Generic import List

from rpw import db, doc, uidoc
from pyrevit import script, DB, forms
from Autodesk.Revit.DB.Architecture import Room

#from pyrevit import revit,
#import sys
#from Autodesk.Revit.DB import ElementMulticategoryFilter
#from Autodesk.Revit.DB import TransactionGroup 
#from Autodesk.Revit.DB import Line, SketchPlane, XYZ, Plane

#this is the destination parameter name
#room_parameter_name_wall_materials =    "11200 Materialis. Wand/Oberfl."
#room_parameter_name_ceiling_materials = "11300 Typ Bauteilkatalog"
#room_parameter_name_floor_materials =   "11140 Typ Bauteilkatalog"
#room_parameter_name_floor_materials_face =   "11130 Materialis. Bodenbelag"
room_parameter_name_wall_materials =    "BuiltInCategory.OST_Walls"
room_parameter_name_ceiling_materials = "BuiltInCategory.OST_Roofs"
room_parameter_name_floor_materials =   "BuiltInCategory.OST_"
room_parameter_name_floor_materials_face =   "BuiltInCategory.OST_"

#this is the source parameter name
wall_material_parameter_name = "Typ Bauteilkatalog"

#this Yes/ No parameter protect a values on a room to overwrite by this script, keep it empty "" to work without this mechanism
#manual_room_materials_parameter_name = "11xxx Materialis. manuell"
manual_room_materials_parameter_name = ""

#switches to work on
SetWalls = False    
SetFloors = True
SetCeilings = True

#################################

__title__ = 'Materials from\nbounding Room Faces'

__doc__ = "Runs through all rooms or the preselcted one, and write the materials of the enclosing walls, \n"\
          "and floors into parameters of that room \n"\
          "It will be write into the a parameter.\n"\
          "It requiers this paramters as text paramters on the rooms:\n"\
          "11300 Typ Bauteilkatalog\n"\
          "11200 Materialis. Wand/Oberfl.\n"\
          "11140 Typ Bauteilkatalog\n"\
          "and this as Yes/No Paratmeter:\n"\
          "11xxx Materialis. manuell\n"\
          "and this on the wall and floor elements as source:\n"\
          "Typ Bauteilkatalog"

#set to True or hold shift while click for more detailed messages
DEBUG = False
if __shiftclick__:
    DEBUG = True

def param_exists_by_name(any_room, parameter_name):
    return any_room.LookupParameter(parameter_name)

def join_mat(materials):
    joined = ', '.join(materials)
    return joined

def room_finishes (the_room):

    calculator = SpatialElementGeometryCalculator(doc)
    options = SpatialElementBoundaryOptions()
    # get boundary location from area computation settings
    boundloc = AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
    options.SpatialElementBoundaryLocation = boundloc
    material_list = []
    type_list = []
    element_list = []
    area_list = []
    face_list = []
    try:
        results = calculator.CalculateSpatialElementGeometry(the_room)
        for face in results.GetGeometry().Faces:
            for bface in results.GetBoundaryFaceInfo(face):
                type_list.append(str(bface.SubfaceType))
                if bface.GetBoundingElementFace().MaterialElementId.IntegerValue == -1:
                    material_list.append(None)
                else:
                    material_list.append(doc.GetElement(bface.GetBoundingElementFace().MaterialElementId))
                element_list.append(doc.GetElement(bface.SpatialBoundaryElement.HostElementId))
                area_list.append(bface.GetSubface().Area)
                face_list.append(bface.GetBoundingElementFace())
    except:
        pass	
        
    return(type_list,material_list,area_list,face_list,element_list)

def owned_by_other_user(elem):
    # Checkout Status of the element
    checkout_status = WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
    if checkout_status == CheckoutStatus.OwnedByOtherUser:
        return True
    else:
        return False

########################################################################  
    
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

output = script.get_output()



# Prepare a collector for GenericModels and Roofs
#categories = [BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_Roofs]
#filterCategories = ElementMulticategoryFilter(List[BuiltInCategory](categories))

# Collect all elements of type GenericModel and Roof
#generic_models_and_roofs = FilteredElementCollector(doc).WherePasses(filterCategories).WhereElementIsNotElementType().ToElements()
generic_models_Collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
Roofs_Collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements()
#Roofs_Collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElements()


# Collect all group instances
group_instances = FilteredElementCollector(doc).OfClass(DB.Group).ToElements()

# Create a geometry calculator to analyze room boundaries
#geo_calculator = SpatialElementGeometryCalculator(doc)

#rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()
rooms = []
selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

#start transaction
with db.Transaction("write room data"):
    #if there was an slection we use them instead
    if selection:
        for selel in selection:
            if selel.GetType() == Room:
                rooms.append(selel)
    else:
        rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

    #for the process bar on top
    EleNums = rooms.Count         

    if EleNums ==0:
        forms.alert('There are no rooms to work with.')
        script.exit()
    
    print ("Rooms: " + str(EleNums))
    #just on room to check if the parameters attached
    #TODO: check an other way to determine parameters
    room = rooms[0]
    
    if SetWalls:
        if not param_exists_by_name(room, room_parameter_name_wall_materials):
            forms.alert("The parameter " + room_parameter_name_wall_materials + " is necessary but does not exist. Please create them first.")
            script.exit()
            
    if SetCeilings: 
        if not room_parameter_name_ceiling_materials.startswith("BuiltInCategory.OST"):
            if not param_exists_by_name(room, room_parameter_name_ceiling_materials):
                forms.alert("The parameter " + room_parameter_name_ceiling_materials + " is necessary but does not exist. Please create them first.")
                script.exit()
            
    if SetFloors: 
        if not room_parameter_name_floor_materials.startswith("BuiltInCategory.OST"):
            if not param_exists_by_name(room, room_parameter_name_floor_materials):
                forms.alert("The parameter " + room_parameter_name_floor_materials + " is necessary but does not exist. Please create them first.")
                script.exit()
        if not room_parameter_name_floor_materials_face.startswith("BuiltInCategory.OST"):
            if not param_exists_by_name(room, room_parameter_name_floor_materials_face):
                forms.alert("The parameter " + room_parameter_name_floor_materials_face + " is necessary but does not exist. Please create them first.")
                script.exit()
    
    if manual_room_materials_parameter_name <> "":
        if not param_exists_by_name(room, manual_room_materials_parameter_name):
            forms.alert("The parameter " + manual_room_materials_parameter_name + " is necessary but does not exist. Please create them first.")
            script.exit()
        
    wall_mat_str = set()
    up_mat_str = set()
    down_mat_str = set()
    floor_mat = set()
    
    filtered_roofs = []

    for element in Roofs_Collection:
        element_location = element.Location
        if not element_location:
            continue  # Skip if there's no location
        
        element_bb = element.get_BoundingBox(None)
        if not element_bb:
            continue  # Skip if there's no bounding box
        
        bb_center = (element_bb.Min + element_bb.Max) / 2
        
        # Das Element und sein bb_center zum Dictionary oder Array hinzufügen
        filtered_roofs.append((element, bb_center))
        
    filtered_group_instances = []
    for group_instance in group_instances:
        if group_instance.Location:
            gi_loc = group_instance.Location
            gi_locpoint = gi_loc.Point
            
            filtered_group_instances.append((group_instance, gi_locpoint))    
    
    #room by room
    for room in rooms:
        wall_mat_str.clear()
        up_mat_str.clear()
        down_mat_str.clear()
        floor_mat.clear()
        
        #check if someone else has the room, otherwise an error will thrown
        #TODO maybe this throw an error in a non worksharing model
        if owned_by_other_user(room):
            print(" Room number: " + room.Number + output.linkify(room.Id) + " -> OwnedByOtherUser") 
            continue
    
        #rooms where not placed as 0 area
        if not room.Area > 0:
            continue
        
        #if DEBUG:
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        print(" Room number: " + room.Number + " " + output.linkify(room.Id) + " " + room_name)
    
        #there is a parameter that protects the room from being overwritten by this script
        if manual_room_materials_parameter_name <> "":
            manual = room.LookupParameter(manual_room_materials_parameter_name).AsInteger()
        else:
            manual = False
            
        if manual:
            print(" Room number: " + room.Number  + " " + output.linkify(room.Id) + " -> manual")
            continue

        else:

            myfinish = room_finishes(room)

            for i, item in enumerate(myfinish[0]):

                if i < len(myfinish[4]):
                    myfinishobj = myfinish[4][i]
                else:
                    continue
    
                myfinishobj = myfinish[4][i]
                if myfinishobj is None: 
                    continue
                
                if SetWalls and myfinish[0][i] == "Side":
                    if DEBUG: print myfinishobj.ElementID
                    wall_mat = myfinishobj.LookupParameter(wall_material_parameter_name)
                    if wall_mat.HasValue:
                        wall_mat_str.add(wall_mat.AsString())
                        
                if SetCeilings and myfinish[0][i] == "Top":
                
                    #if DEBUG: print ("myfinishobj Ceiling" + output.linkify(myfinishobj.Id))
                    mat_ids = myfinishobj.GetMaterialIds(False)
                    material_name = ""
                    if mat_ids:
                        #print (len(mat_ids))
                        material = doc.GetElement(mat_ids[len(mat_ids)-1])
                        material_name = material.Name if material else ""
                    #if DEBUG: print (material_name)
                    
                    up_mat_str.add(material_name)
                  
                    #up_mat = myfinishobj.LookupParameter(wall_material_parameter_name)
                    #if up_mat.HasValue:
                    #    up_mat_str.add(up_mat.AsString())

                if SetFloors and myfinish[0][i] == "Bottom":
                    #if DEBUG: print ("myfinishobj Floors" + output.linkify(myfinishobj.Id))
                    mat_ids = myfinishobj.GetMaterialIds(False)
                    material_name = "Unknown"
                    if mat_ids:
                        material = doc.GetElement(mat_ids[0])
                        material_name = material.Name if material else "Unknown"
                    #if DEBUG: print (material_name)
                    #down_mat = myfinishobj.LookupParameter(wall_material_parameter_name)
                    #down_mat = myfinishobj.get_Parameter(DB.BuiltInParameter.ROOM_FINISH_FLOOR)
                    
                    
                    #if down_mat:  
                    #    if down_mat.HasValue:
                    #        down_mat_str.add(down_mat.AsString())
                    down_mat_str.add(material_name)
                    
                    #try:
                    ##if isinstance(myfinishobj,HostObject):
                    #    floor_face_refs = HostObjectUtils.GetTopFaces(myfinishobj)
                    #        
                    #    for ffr in floor_face_refs:
                    #        floor_face = doc.GetElement(ffr).GetGeometryObjectFromReference(ffr)
                    #        if not floor_face == None:
                    #            matr = doc.GetElement(floor_face.MaterialElementId)
                    #            floor_mat.add(matr.Name[-3:])
                    #            if DEBUG: print "floor material: " + matr.Name[-3:]
                    #except:
                    #    print "no floor material available "

            for element, bb_center in filtered_roofs:
            
                if not room.IsPointInRoom(bb_center):
                    continue  # Skip if the center is not inside the room

                thegridset = element.CurtainGrids
                if thegridset is None:
                    if DEBUG: print("CurtainGrids is None or empty")
                    continue  # Skip if there's no curtain grid

                for curtaingrid in thegridset:

                    if curtaingrid.NumPanels <= 1:
                        continue  # Skip if there aren't enough panels

                    thepanelIds = curtaingrid.GetPanelIds()
                    for pId in thepanelIds:
                        mypanel = doc.GetElement(pId)
                        mymaterialsIds = mypanel.GetMaterialIds(False)
                        for mymatid in mymaterialsIds:
                            mymat = doc.GetElement(mymatid)
                            up_mat_str.add(mymat.Name)

                if DEBUG: print("Element " + element.Name + " is inside. " + output.linkify( element.Id) )    

            for group_instance, gi_locpoint in filtered_group_instances:
                if room.IsPointInRoom(gi_locpoint):
                    if group_instance.Name == "DCK_NPS_HKS_PLH":
                        if DEBUG: print("Element " + group_instance.Name + " is inside. " + output.linkify( group_instance.Id) )        
                        up_mat_str.add("Heiz-kühl Segelsystem") 
                
            if SetWalls:
                room_walls_param = room.LookupParameter(room_parameter_name_wall_materials)
                room_walls_param.Set(join_mat(wall_mat_str))
                if DEBUG: print (room_parameter_name_wall_materials + " :" + str(join_mat(wall_mat_str)))

            if SetCeilings: 
                room_ceiling_param = room.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING)
            
                #room_ceiling_param = room.LookupParameter(room_parameter_name_ceiling_materials)
                room_ceiling_param.Set(join_mat(up_mat_str))
                if DEBUG: print ("ROOM_FINISH_CEILING: " + str(join_mat(up_mat_str)))

            if SetFloors: 
                #room_floor_param = room.LookupParameter(room_parameter_name_floor_materials)
                room_floor_param = room.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR)

                room_floor_param.Set(join_mat(down_mat_str))
                if DEBUG: print("ROOM_FINISH_FLOOR: " + str(join_mat(down_mat_str)))
                
                #room_floor_param_mat = room.LookupParameter(room_parameter_name_floor_materials_face)
                #room_floor_param_mat.Set(join_mat(floor_mat))
                #if DEBUG: print (room_parameter_name_floor_materials_face + ": " + str(join_mat(floor_mat)))

        output.update_progress(idx, EleNums)
        idx+=1   

output.reset_progress()

group_instance
output.insert_divider()
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
