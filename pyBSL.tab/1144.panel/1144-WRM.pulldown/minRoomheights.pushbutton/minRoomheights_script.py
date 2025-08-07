# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, XYZ, UV, AreaVolumeSettings
from System.Diagnostics import Stopwatch
from collections import namedtuple
from rpw import db, doc, uidoc
from pyrevit import script

# Parameterdefinitionen
heigth_parameter_name = "Room Height Maximum"
manual_parameter_name = "Height_Minimum_manual"
height_min_parameter_name = "Room Height Minimum"

__title__ = 'min Raumhöhe'

__doc__ = "Runs through all rooms and calculate the room height between "\
          "the largest lower surface and the largest upper surface of a room "

DEBUG = False

def clear_height_param_exists(any_room):
    return any_room.LookupParameter(heigth_parameter_name)

def get_face_normal_and_mid_z(solid_face):
    mid_uv = UV(0.5, 0.5)
    normal = solid_face.ComputeNormal(mid_uv)
    mid_z = solid_face.Evaluate(mid_uv).Z
    return normal, mid_z

stopwatch = Stopwatch()
stopwatch.Start()

if __shiftclick__:
    DEBUG = True
    
output = script.get_output()
ft_mm = 304.8
sqft_sqm = 0.092903
UP = XYZ.BasisZ
DOWN = UP.Negate()

rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()
ceilings = Fec(doc).OfCategory(Bic.OST_Ceilings).WhereElementIsNotElementType().ToElements()

space_calc = SpatialElementGeometryCalculator(doc)
volume_calculation = AreaVolumeSettings.GetAreaVolumeSettings(doc).ComputeVolumes
selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

BoundFace = namedtuple("BoundFace", "area height face")

with db.Transaction("write room data"):
    if selection:
        rooms = selection

    for room in rooms:
        if DEBUG:

            print(" Room number: " + room.Number + output.linkify(room.Id) )

        if not volume_calculation:
            print("Volume computations is off (Areas only). Please switch Areas and Volumes Computation on and re-run this script.")
            continue

        if not clear_height_param_exists(room):
            print("The parameter " + heigth_parameter_name + " is necessary but does not exist. Please create them first.")
            continue

        if not room.Area > 0:
            continue

        manual = room.LookupParameter(manual_parameter_name).AsInteger()

        if manual:
            if DEBUG:
                print("  room height is set to manual: {}".format(manual))
            print(" Room number: " + room.Number + output.linkify(room.Id) + " -> manual")
            continue

        elif not manual:
            room_name = room.LookupParameter("Name").AsString()
            space_results = space_calc.CalculateSpatialElementGeometry(room)
            room_solid = space_results.GetGeometry()

            floor_faces = []
            ceiling_faces = []
            room_heights = []
            
            for face in room_solid.Faces:
                face_normal, face_mid_z = get_face_normal_and_mid_z(face)

                if face_normal.IsAlmostEqualTo(DOWN):
                    height_sample = round(face_mid_z * ft_mm, 2)
                    floor_faces.append(BoundFace(face.Area, height_sample, face))

                elif face_normal.IsAlmostEqualTo(UP):
                    height_sample = round(face_mid_z * ft_mm, 2)
                    ceiling_faces.append(BoundFace(face.Area, height_sample, face))

            floor_major_face = max(floor_faces)
            okff = floor_major_face.height
            
            if DEBUG:
                print (" 'ceiling_faces':", len(ceiling_faces))
                print (" 'floor_faces':", len(floor_faces))
            
            if len(ceiling_faces)>0:
            
                for cf in ceiling_faces:
                    room_heights.append(cf.height - okff)
                    if DEBUG:
                        print (cf.height - okff)
            
                room_clear_height = max(room_heights)
            
                room.LookupParameter(heigth_parameter_name).Set(room_clear_height / ft_mm)
            
                room_clear_height_min = min (room_heights)

                if room_clear_height_min < room_clear_height:
                
                    # Speichern der niedrigsten Deckenhöhe im Parameter
                    room.LookupParameter(height_min_parameter_name).Set(room_clear_height_min / ft_mm)

                print(" Room number: " + room.Number + output.linkify(room.Id) + 
                      " -> " + heigth_parameter_name + ": {}".format(room_clear_height / 1000) +
                      ", " + height_min_parameter_name + ": {}".format(room_clear_height_min / 1000))
            else:
                print(" Room number: " + room.Number + output.linkify(room.Id) + 
                      " could not measure height ")

print(53 * "=")
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
