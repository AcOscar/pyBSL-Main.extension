import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementCategoryFilter, \
    Transaction, SpatialElementBoundaryOptions, Wall, Floor, ElementId, Material, BuiltInParameter, \
    Structure, Group, GeometryInstance

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

def get_bdn_floors(doc):
    """Get all floor elements whose type name starts with 'BDN'."""
    floors = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Floors) \
        .WhereElementIsNotElementType() \
        .ToElements()

    # Filter floors whose type name starts with 'BDN'
    bdn_floors = [floor for floor in floors if floor.Name.startswith("BDN")]
    return bdn_floors

def get_load_bearing_layer(floor):
    """Retrieve the material of the load-bearing layer in a floor."""
    floor_type = doc.GetElement(floor.GetTypeId())
    if hasattr(floor_type, "GetCompoundStructure"):
        structure = floor_type.GetCompoundStructure()
        if structure:
            # Find the structural material layer
            for i in range(structure.LayerCount):
                layer = structure.GetLayer(i)
                if layer.Function == Structure.StructuralLayerFunction.Structure:
                    material_id = layer.MaterialId
                    if material_id != ElementId.InvalidElementId:
                        material = doc.GetElement(material_id)
                        return material
    return None

def get_selected_rooms(uidoc):
    """Get all selected rooms."""
    selection = uidoc.Selection.GetElementIds()
    selected_elements = [doc.GetElement(id) for id in selection]
    
    # Filter only the selected rooms
    selected_rooms = [elem for elem in selected_elements if elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms)]
    return selected_rooms

def is_element_in_group(element):
    """Check if an element is part of a group."""
    return element.GroupId != ElementId.InvalidElementId

def get_group_elements(group):
    """Retrieve all elements in a group."""
    group_elements = []
    geo_instance = group.get_Geometry(SpatialElementBoundaryOptions())
    for geom in geo_instance:
        if isinstance(geom, GeometryInstance):
            group_elements.extend([doc.GetElement(gid) for gid in geom.GetSymbolGeometry().GetInstanceReferenceIds()])
    return group_elements

def update_room_floor_material(rooms, bdn_floors):
    """Update the 'Fußboden' parameter of selected rooms with the floor material that touches them."""
    sbo = SpatialElementBoundaryOptions()

    for room in rooms:
        room_boundaries = room.GetBoundarySegments(sbo)
        floor_materials = set()

        for boundary_list in room_boundaries:
            for boundary_segment in boundary_list:
                boundary_element = doc.GetElement(boundary_segment.ElementId)

                # Check if the boundary element is in a group
                if is_element_in_group(boundary_element):
                    group = doc.GetElement(boundary_element.GroupId)
                    group_elements = get_group_elements(group)
                    boundary_element = [el for el in group_elements if isinstance(el, Floor)][0]

                if isinstance(boundary_element, Floor) and boundary_element in bdn_floors:
                    # Get load-bearing material
                    material = get_load_bearing_layer(boundary_element)
                    if material:
                        floor_materials.add(material.Name)

        if floor_materials:
            # Sort the materials alphabetically and join into a string
            material_text = ', '.join(sorted(floor_materials))

            # Update the room's "Fußboden" parameter
            footboden_param = room.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR)
            if footboden_param and not footboden_param.IsReadOnly:
                footboden_param.Set(material_text)

# Start the transaction to modify the model
TransactionManager.Instance.EnsureInTransaction(doc)

bdn_floors = get_bdn_floors(doc)
selected_rooms = get_selected_rooms(uidoc)
update_room_floor_material(selected_rooms, bdn_floors)

TransactionManager.Instance.TransactionTaskDone()
