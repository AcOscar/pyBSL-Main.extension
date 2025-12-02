import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, XYZ, UV, AreaVolumeSettings
from System.Diagnostics import Stopwatch
from collections import namedtuple
from rpw import db, doc, uidoc
from pyrevit import script
#we need two parameters
#one as length to write the roomheigth
#the second as yes/no to prevent the first one to overwrite it with this script
#so we have the opportunity to write the height value manually

heigth_parameter_name = "Lichte_Hoehe"
manual_parameter_name = "Lichte_Hoehe_manuell"

__title__ = 'Room heights: weighted room srfs'

__doc__ = "Runs through all rooms and calculate the room height between "\
          "the largest lower surface and the largest upper surface of a room "\
          "It will be write into the the parameter Lichte_Hoehe."\
          "If there is a true/false parameter Lichte_Hoehe_manuell the value of Lichte_Hoehe will not be touched."\


#set to True for more detailed messages
debug = True

def clear_height_param_exists(any_room):
    return any_room.LookupParameter(heigth_parameter_name)


def get_face_normal_and_mid_z(solid_face):
    mid_uv = UV(0.5, 0.5)
    normal = solid_face.ComputeNormal(mid_uv)
    mid_z = solid_face.Evaluate(mid_uv).Z
    return normal, mid_z


stopwatch = Stopwatch()
stopwatch.Start()

output = script.get_output()
ft_mm = 304.8
sqft_sqm = 0.092903
UP = XYZ.BasisZ
DOWN = UP.Negate()

rooms    = Fec(doc).OfCategory(Bic.OST_Rooms   ).WhereElementIsNotElementType().ToElements()
ceilings = Fec(doc).OfCategory(Bic.OST_Ceilings).WhereElementIsNotElementType().ToElements()

space_calc = SpatialElementGeometryCalculator(doc)
volume_calculation = AreaVolumeSettings.GetAreaVolumeSettings(doc).ComputeVolumes
selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

BoundFace    = namedtuple("BoundFace",    "area height face")
BoundSubFace = namedtuple("BoundSubFace", "area bound_elem material subface face thickness is_raw is_coating")

with db.Transaction("write room data"):
    if selection:
        rooms = selection

    for room in rooms:
        if not volume_calculation:
            print("Volume computations is off (Areas only). Please switch Areas and Volumes Computation on and re-run this script.")
            continue

        if not clear_height_param_exists(room):
            print("The parameter " + heigth_parameter_name + " is necessary but does not exist. Please create them first.")
            continue

        if not room.Area > 0:
            continue

        manual = room.LookupParameter(manual_parameter_name).AsInteger()
        #print(" Room number: " + room.Number + output.linkify(room.Id))

        if manual:
            if debug:
                print("  room height is set to manual: {}".format(manual))
            print(" Room number: " + room.Number + output.linkify(room.Id) + " -> manual")

            continue

        elif not manual:
            room_name = room.LookupParameter("Name").AsString()
            space_results = space_calc.CalculateSpatialElementGeometry(room)
            room_solid = space_results.GetGeometry()

            floor_faces = []
            ceiling_faces = []

            for face in room_solid.Faces:
                face_normal, face_mid_z = get_face_normal_and_mid_z(face)

                if face_normal.IsAlmostEqualTo(DOWN):
                    if debug:
                        print("FLOOR".ljust(10) + 25 * "-")
                    height_sample = round(face_mid_z * ft_mm, 2)
                    floor_faces.append(BoundFace(face.Area, height_sample, face))
                    if debug:
                        print("area: {} face_mid_z: {}".format(face.Area * sqft_sqm, height_sample / 1000))

                elif face_normal.IsAlmostEqualTo(UP):
                    if debug:
                        print("CEILING".ljust(10) + 25 * "-")
                    height_sample = round(face_mid_z * ft_mm, 2)
                    ceiling_faces.append(BoundFace(face.Area, height_sample, face))
                    if debug:
                        print("area: {} face_mid_z: {}".format(face.Area * sqft_sqm, height_sample / 1000))

            if debug:
                print(6 * "-" + " -> major FLOOR data extract:")
            floor_major_face = max(floor_faces)
            okff = floor_major_face.height

            if debug:
                print(" -> OKFF: {}".format(floor_major_face.height / 1000))
            # room.LookupParameter("OKFF").Set(floor_major_face.height / 1000)  # inconsistent up param

            if debug:
                print(6 * "-" + " -> major CEILING data extract:")
            ceiling_major_face = max(ceiling_faces)
            ukfd = ceiling_major_face.height

            if debug:
                print(" -> UKFD: {}".format(ceiling_major_face.height / 1000))
            # room.LookupParameter("UKFD").Set(ceiling_major_face.height / ft_mm)

            room_clear_height = (ukfd - okff)
            room.LookupParameter(heigth_parameter_name).Set(room_clear_height / ft_mm)
            print(" Room number: " + room.Number + output.linkify(room.Id) + " -> " + heigth_parameter_name + ": {}".format(room_clear_height / 1000))
            
            #print(" -> " + heigth_parameter_name + ": {}".format(room_clear_height))

print(53 * "=")
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
