from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, UnitUtils, SpecTypeId, BuiltInParameter
from rpw import db, doc, uidoc

finish_floor_elevation_param_name = "Raum_OKFB"

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

def room_param_exists(any_room, param_name):
    return any_room.LookupParameter(param_name)

def get_room_finish_floor_elevation(level, elevation_offset):
    level_elevation = level.Elevation  # Get level elevation in internal units (feet)
    level_elevation = UnitUtils.ConvertFromInternalUnits(level_elevation, SpecTypeId.Length)  
    elevation_relative_to_pbp = level_elevation - elevation_offset  # Adjust relative to project base point
    return elevation_relative_to_pbp

def main():

    # Get the project base point
    base_points = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
    if not base_points:
        return
    project_base_point = base_points[0]
    pbp_elevation = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION).AsDouble()

    selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds() if doc.GetElement(elId).Category.BuiltInCategory == Bic.OST_Rooms]

    """if was a selection, use that, otherwise get all rooms"""
    if selection:
        rooms = selection
    else:
        rooms = list(Fec(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements())
    
    """check if we have rooms and the necessary parameters"""
    if not rooms:
        print("No rooms found.")
        return
    elif not room_param_exists(rooms[0], finish_floor_elevation_param_name):
        print("The parameter " + finish_floor_elevation_param_name + " is necessary but does not exist. Please create them first.")
        return

    # Start a transaction
    with db.Transaction("write room data"):

        for room in rooms:
            level = doc.GetElement(room.LevelId) 
            if not level:
                continue

            elevation_relative_to_pbp = get_room_finish_floor_elevation(level, pbp_elevation)  # Adjust relative to project base point

            param = room.LookupParameter(finish_floor_elevation_param_name)

            set_length_to_parameter(param, elevation_relative_to_pbp)

    return

if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:        
        import traceback
        print("ERROR:", e)
        print(traceback.format_exc())