# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import SpatialElementGeometryCalculator, XYZ, UV, AreaVolumeSettings
from Autodesk.Revit.DB import UnitFormatUtils, SpecTypeId
from System.Diagnostics import Stopwatch
from collections import namedtuple
from rpw import db, doc, uidoc
from pyrevit import script


"""
we need two parameters
one in which we write the room height
the second as Yes/No to prevent the first from overwriting it with this script
this gives us the option of writing the height value manually
"""

heigth_parameter_name = "clear height"
manual_parameter_name = "clear height manually"

__title__ = 'Room heights: weighted room srfs'

__doc__ = "Calculates the clear room height between the height of the"\
          "largest lower surface and the lowest upper surface of a room. "\
          "The result is written to a parameter."\
          "Rooms for which a yes/no parameter is set are ignored."


stopwatch = Stopwatch()
UP = XYZ.BasisZ
DOWN = UP.Negate()

BoundFace = namedtuple("BoundFace", "area height face")

def clear_height_param_exists(any_room):
    return any_room.LookupParameter(heigth_parameter_name)

def clear_height_manual_param_exists(any_room):
    return any_room.LookupParameter(manual_parameter_name)

def get_face_normal_and_mid_z(solid_face):
    mid_uv = UV(0.5, 0.5)
    normal = solid_face.ComputeNormal(mid_uv)
    mid_z = solid_face.Evaluate(mid_uv).Z
    return normal, mid_z

def get_room_clear_height(calculator, room):
    """
    Get the clear height of a room by analyzing its bounding faces.
    """
    geometry = calculator.CalculateSpatialElementGeometry(room)
    solid = geometry.GetGeometry()
    
    lower_faces = []
    upper_faces = []
    
    for face in solid.Faces:
        normal, mid_z = get_face_normal_and_mid_z(face)
        if normal.IsAlmostEqualTo(UP):
            upper_faces.append((face, mid_z))
        elif normal.IsAlmostEqualTo(DOWN):
            lower_faces.append((face, mid_z))
    
    if not lower_faces or not upper_faces:
        return 0.0  # No valid bounding faces found
    
    max_lower_z = max(mid_z for face, mid_z in lower_faces)
    min_upper_z = min(mid_z for face, mid_z in upper_faces)
    
    clear_height = min_upper_z - max_lower_z
    return clear_height

def convert_and_format(length_in_internal_units, Format=True):
    """Convert length from internal units to project units and format as string."""
 
    units = doc.GetUnits()
    
    formatted_string = UnitFormatUtils.Format(
        units,
        SpecTypeId.Length,
        length_in_internal_units,
        Format
    )
    
    return formatted_string

def set_length_to_parameter(parameter, length_in_internal_units):
    """Set the length value to the given parameter, handling different data types.
    """
    if parameter.IsReadOnly:
        return False
    
    param_type = parameter.Definition.GetDataType()
    
    if param_type == SpecTypeId.Length:
        parameter.Set(length_in_internal_units)
        return True
    
    elif param_type == SpecTypeId.Number:
        formatted_string = convert_and_format(length_in_internal_units, Format=True)
        parameter.SetValueString (formatted_string)
        return True
    
    elif param_type == SpecTypeId.String.Text:
        formatted_string = convert_and_format(length_in_internal_units, Format=True)
        parameter.Set(formatted_string)
        return True
    
    return False

def main():
    """Main function"""
    stopwatch.Start()

    output = script.get_output()

    space_calc = SpatialElementGeometryCalculator(doc)
    volume_calculation = AreaVolumeSettings.GetAreaVolumeSettings(doc).ComputeVolumes

    """check if volume computations is on"""
    if not volume_calculation:
        print("Volume computations is off (Areas only). Please switch Areas and Volumes Computation on and re-run this script.")
        return

    selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]

    """if was a selection, use that, otherwise get all rooms"""
    if selection:
        rooms = selection
    else:
        rooms = Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()

    """check if we have rooms and the necessary parameters"""
    if not rooms:
        print("No rooms found.")
        return
    elif not clear_height_param_exists(rooms[0]):
        print("The parameter " + heigth_parameter_name + " is necessary but does not exist. Please create them first.")
        return
    elif not clear_height_manual_param_exists(rooms[0]):
        print("The parameter " + manual_parameter_name + " is necessary but does not exist. Please create them first.")
        return
    
    print("Processing {} rooms...".format(len(rooms)))

    with db.Transaction("write room data"):
        for room in rooms:
            """skip rooms with no area (usually unplaced)"""
            if not room.Area > 0:
                continue

            """Check if manual flag is set"""
            manual = room.LookupParameter(manual_parameter_name).AsInteger()

            if manual:
                print(" Room number: " + room.Number + output.linkify(room.Id) + " -> manual")
                continue

            else:
               room_clear_height = get_room_clear_height(space_calc, room)
               #room.LookupParameter(heigth_parameter_name).Set(room_clear_height)

               heigth_paramter = room.LookupParameter(heigth_parameter_name)
               success = set_length_to_parameter(heigth_paramter, room_clear_height)
               if not success:
                   print(" Room number: " + room.Number + output.linkify(room.Id) + " -> could not write to parameter.")
               else:
                   print(" Room number: " + room.Number + output.linkify(room.Id) + " -> " + heigth_parameter_name + ": {}".format(convert_and_format(room_clear_height)))
                
    stopwatch.Stop()
    timespan = stopwatch.Elapsed
    print("Run in: {}".format(timespan))

if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:
        import traceback
        print("ERROR:", e)
        print(traceback.format_exc())