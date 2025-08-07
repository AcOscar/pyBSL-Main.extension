import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import AreaVolumeSettings 
from Autodesk.Revit.DB import SpatialElementBoundaryOptions
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils
from Autodesk.Revit.DB import BuiltInParameter, HostObjectUtils, HostObject
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, SpatialElementType

from System.Diagnostics import Stopwatch
from System.Collections.Generic import List

from rpw import db, doc, uidoc
from pyrevit import script, DB, forms
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB.Architecture import Room
#we need two parameters
#one as length to write the roomheigth
#the second as yes/no to prevent the first one to overwrite it with this script
#so we have the opportunity to write the height value manually

#this is the destination parameter name
room_parameter_name_wall_materials =    "11200 Materialis. Wand/Oberfl."
room_parameter_name_ceiling_materials = "11300 Typ Bauteilkatalog"
room_parameter_name_floor_materials =   "11140 Typ Bauteilkatalog"
room_parameter_name_floor_materials_face =   "11130 Materialis. Bodenbelag"

#this Yes/ No parameter protect a room to overwrite by this script
manual_room_materials_parameter_name = "11xxx Materialis. manuell"

SetWalls = False    
SetFloors = True
SetCeilings = False 


#this is the source parameter name
wall_material_parameter_name = "Typ Bauteilkatalog"

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

#set to True for more detailed messages
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

#######################################

def owned_by_other_user(elem):
#from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

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

#rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()
rooms = []
selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

#start transaction
with db.Transaction("write room data"):
    #if there was an slection we use them instead
    if selection:
        #rooms = selection


        for selel in selection:
        
            if selel.GetType() == Room:
                rooms.append(selel)
       
        #EleNums = AllElements.Count
    else:
        rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()

    #for the process bar on top
    EleNums = rooms.Count         

    if EleNums ==0:
        forms.alert('There are no rooms to work with.')
        script.exit()
    
    #just on room to check if the parameters attached
    #TODO: check an other way to determine parameters
    room = rooms[0]
    
    print ("Rooms: " + str(EleNums))
    if SetWalls:
        if not param_exists_by_name(room, room_parameter_name_wall_materials):
            forms.alert("The parameter " + room_parameter_name_wall_materials + " is necessary but does not exist. Please create them first.")
            script.exit()
            
    #if SetCeilings: 
    #    if not param_exists_by_name(room, room_parameter_name_ceiling_materials):
    #        forms.alert("The parameter " + room_parameter_name_ceiling_materials + " is necessary but does not exist. Please create them first.")
    #        script.exit()
            
    #if SetFloors: 
    #    if not param_exists_by_name(room, room_parameter_name_floor_materials):
    #        forms.alert("The parameter " + room_parameter_name_floor_materials + " is necessary but does not exist. Please create them first.")
    #        script.exit()
    #    if not param_exists_by_name(room, room_parameter_name_floor_materials_face):
    #        forms.alert("The parameter " + room_parameter_name_floor_materials_face + " is necessary but does not exist. Please create them first.")
    #        script.exit()
        
    #if not param_exists_by_name(room, manual_room_materials_parameter_name):
    #    forms.alert("The parameter " + manual_room_materials_parameter_name + " is necessary but does not exist. Please create them first.")
    #    script.exit()
        
    wall_mat_str = set()
    up_mat_str = set()
    down_mat_str = set()
    floor_mat = set()
    
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
        
        if DEBUG:
            room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
            print(" Room number: " + room.Number + " " + output.linkify(room.Id) + " " + room_name)
    
        #there is a parameter that protects the room from being overwritten by this script
        #manual = room.LookupParameter(manual_room_materials_parameter_name).AsInteger()
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
    
            
            if SetWalls:
                room_walls_param = room.LookupParameter(room_parameter_name_wall_materials)
                room_walls_param.Set(join_mat(wall_mat_str))
                if DEBUG: print (room_parameter_name_wall_materials + " :" + str(join_mat(wall_mat_str)))

            if SetCeilings: 
                room_ceiling_param = room.get_Parameter(DB.BuiltInParameter.ROOM_FINISH_CEILING)
            
                #room_ceiling_param = room.LookupParameter(room_parameter_name_ceiling_materials)
                room_ceiling_param.Set(join_mat(up_mat_str))
                if DEBUG: print ("ROOM_FINISH_CEILING: " + str(join_mat(up_mat_str)))

            if SetFloors: 
                #room_floor_param = room.LookupParameter(room_parameter_name_floor_materials)
                room_floor_param = room.get_Parameter(DB.BuiltInParameter.ROOM_FINISH_FLOOR)

                room_floor_param.Set(join_mat(down_mat_str))
                if DEBUG: print("ROOM_FINISH_FLOOR: " + str(join_mat(down_mat_str)))
                
                #room_floor_param_mat = room.LookupParameter(room_parameter_name_floor_materials_face)
                #room_floor_param_mat.Set(join_mat(floor_mat))
                #if DEBUG: print (room_parameter_name_floor_materials_face + ": " + str(join_mat(floor_mat)))
                
                

        output.update_progress(idx, EleNums)
        idx+=1   

output.reset_progress()


output.insert_divider()
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
