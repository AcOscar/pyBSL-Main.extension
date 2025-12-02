import clr
import sys
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
clr.AddReference('RevitNodes')
clr.AddReference('RevitAPIUI')

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, UnitUtils, SpecTypeId, BuiltInParameter
from Autodesk.Revit.DB import Level, SpatialElement

# Get the current document
doc = DocumentManager.Instance.CurrentDBDocument
if doc is None:
    raise Exception("Error: Could not retrieve the active Revit document.")

# Get the project base point
base_points = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
if not base_points:
    raise Exception("Error: No project base point found.")

project_base_point = base_points[0]
pbp_elevation = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION).AsDouble()

# Get all rooms in the project
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
if not rooms:
    raise Exception("Error: No rooms found in the project.")

# Start a transaction
TransactionManager.Instance.EnsureInTransaction(doc)

for room in rooms:
    level = doc.GetElement(room.LevelId)  # Get the level associated with the room
    if level:
        level_elevation = level.Elevation  # Get level elevation in internal units (feet)
        level_elevation_mm = UnitUtils.ConvertFromInternalUnits(level_elevation, SpecTypeId.Length)  # Convert to mm
        elevation_relative_to_pbp = level_elevation_mm - pbp_elevation  # Adjust relative to project base point
        
        # Set the value to the room parameter "Raum_OKFB"
        param = room.LookupParameter("Raum_OKFB")
        if param and not param.IsReadOnly:
            param.Set(elevation_relative_to_pbp)

# Commit the transaction
TransactionManager.Instance.TransactionTaskDone()

# Output success message
OUT = "Parameter 'Raum_OKFB' updated successfully for all rooms."
