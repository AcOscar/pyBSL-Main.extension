# -*- coding: utf-8 -*-
from pyrevit import forms
from Autodesk.Revit.DB import *


__title__ = 'ViewFilterLegend'
__doc__ = 'Create a legend to all wall type filters from a view'

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


region_width_mm = 200
region_height_mm = 50
region_spacing_mm = 10
text_type_name = "3.5mm Arial"


def mm_to_feet(mm):
    return float(mm) / 304.8


def get_textnotetype(doc, preferred_name="3.5mm Arial"):
    # Collect all TextNoteTypes
    types = list(FilteredElementCollector(doc).OfClass(TextNoteType))
    if not types:
        raise Exception("There is no TextNoteType in the project.")

    # 1) Exact match (case insensitive) by name
    want = preferred_name.strip().lower()
    for t in types:
        tname = (t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "").strip().lower()
        if tname == want:
            return t

    # 2) Fallback: Project setting (Default TextNoteType), if available
    did = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
    if did and did != ElementId.InvalidElementId:
        dt = doc.GetElement(did)
        if isinstance(dt, TextNoteType):
            return dt

    # 3) last fallback: first available
    return types[0]


# ============================
# GET ALL VIEWS WITH FILTERS
# ============================
all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
views_with_filters = []
for v in all_views:
    try:
        if not v.IsTemplate and v.GetFilters():
            views_with_filters.append(v)
    except:
        continue

if not views_with_filters:
    forms.alert("There are no filtered views in the project.", exitscript=True)

selected_views = forms.SelectFromList.show(
    views_with_filters,
    multiselect=True,
    name_attr="Name",
    title="Select View Name"
)

if not selected_views:
    forms.alert("No views were selected. Canceling...", exitscript=True)

# ============================
# SELECT TEXT STYLE
# ============================
# text_types = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()

# class TextTypeWrapper(object):
#     def __init__(self, text_type):
#         self.text_type = text_type
#         self.Name = text_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()

# text_wrappers = [TextTypeWrapper(tt) for tt in text_types]

# selected_wrapper = forms.SelectFromList.show(
#     text_wrappers,
#     name_attr="Name",
#     title="Select Text Type"
# )

# if not selected_wrapper:
#     forms.alert("No text style was selected. Canceling...", exitscript=True)

# text_type = selected_wrapper.text_type


text_type = get_textnotetype(doc, text_type_name)

# ============================
# SELECT VIEW LEGEND 
# ============================
legend_views = [v for v in all_views if v.ViewType == ViewType.Legend and not v.IsTemplate]

if not legend_views:
    forms.alert("You must have at least one legend view in the project.", exitscript=True)

legend_view_selected = forms.SelectFromList.show(
    legend_views,
    name_attr="Name",
    title="Select Legend View"
)

if not legend_view_selected:
    forms.alert("You did not select any legend views. Canceling...", exitscript=True)

# ============================
# REQUEST DIMENSIONS IN mm
# ============================
# region_width_mm = forms.ask_for_string(default="200", prompt="Width FilledRegion (mm):")
# region_height_mm = forms.ask_for_string(default="50", prompt="Heigth FilledRegion (mm):")
# region_spacing_mm = forms.ask_for_string(default="10", prompt="Spacing FilledRegion (mm):")

try:
    region_width = mm_to_feet(float(region_width_mm))
    region_height = mm_to_feet(float(region_height_mm))
    region_spacing = mm_to_feet(float(region_spacing_mm))
except:
    forms.alert("Error in dimensions. Make sure you enter only numbers.", exitscript=True)

# ============================
# START TRANSACTION
# ============================
t = Transaction(doc, "Generate Legend with Filters")
t.Start()

for view in selected_views:
    new_legend_id = legend_view_selected.Duplicate(ViewDuplicateOption.WithDetailing)
    new_legend = doc.GetElement(new_legend_id)

    # check if the name already exist and count up in this case
    base_name = "Legend_{}".format(view.Name)
    name = base_name
    counter = 1

    # check if name already exists
    existing_names = [v.Name for v in legend_views] + \
        [v.Name for v in FilteredElementCollector(doc).OfClass(View).ToElements() if v.ViewType == ViewType.Legend]

    while name in existing_names:
        name = "{} ({})".format(base_name, counter)
        counter += 1

    print("New legend'{}': ".format(name))
    new_legend.Name = name

    # List visible walls
    walls_collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    wall_names = []
    for w in walls_collector:
        try:
            name = w.Name
            if name and name.strip():
                wall_names.append(name)

        except:
            continue

    wall_names = sorted(set(wall_names))

    print("Walls visible in the view'{}':".format(view.Name))
    if wall_names:
        for name in wall_names:
            print(" - {}".format(name))
    else:
        print(" - (no visible walls)")
        continue
    # Here you can add logic to create FilledRegions with colors from the filter if you want.
    x_origin = mm_to_feet(20)
    y_origin = mm_to_feet(20)

    region_type = FilteredElementCollector(doc)\
        .OfClass(FilledRegionType)\
        .FirstElement()

    if not region_type:
        forms.alert("No FilledRegionType found!", exitscript=True)

    filters = view.GetFilters()
    y_offset = 0
    wall_cat_id = ElementId(BuiltInCategory.OST_Walls)

    #print("View '{}' has {} Filters".format(view.Name, len(filters)))
    for f_id in filters:
        # Filterelement holen
        filt = doc.GetElement(f_id)
        override = view.GetFilterOverrides(f_id)

        if not filt or not isinstance(filt, ParameterFilterElement):
            continue  # not a parameter filter

        # Make sure that walls are affected by the filter.
        categories = filt.GetCategories()

        if wall_cat_id not in categories:
            continue
        print(" - {}".format(filt.Name))

        # rectangle position
        p1 = XYZ(x_origin, y_origin - y_offset, 0)
        p2 = XYZ(x_origin + region_width, y_origin - y_offset, 0)
        p3 = XYZ(x_origin + region_width, y_origin - y_offset + region_height, 0)
        p4 = XYZ(x_origin, y_origin - y_offset + region_height, 0)

        loop = CurveLoop()
        loop.Append(Line.CreateBound(p1, p2))
        loop.Append(Line.CreateBound(p2, p3))
        loop.Append(Line.CreateBound(p3, p4))
        loop.Append(Line.CreateBound(p4, p1))

        # FilledRegion
        region = FilledRegion.Create(doc, region_type.Id, new_legend.Id, [loop])

        cfg_pattern_id = override.CutForegroundPatternId if override else None
        cfg_color = override.CutForegroundPatternColor if override else None
        cbg_pattern_id = override.CutBackgroundPatternId if override else None
        cbg_color = override.CutBackgroundPatternColor if override else None

        g_override = OverrideGraphicSettings()

        if cfg_pattern_id != ElementId.InvalidElementId:
            g_override.SetSurfaceForegroundPatternId (cfg_pattern_id)
        if cfg_color != ElementId.InvalidElementId:
            g_override.SetSurfaceForegroundPatternColor (cfg_color)
        if cbg_pattern_id != ElementId.InvalidElementId:
            g_override.SetSurfaceBackgroundPatternId (cbg_pattern_id)
        if cbg_color != ElementId.InvalidElementId:
            g_override.SetSurfaceBackgroundPatternColor (cbg_color)

        new_legend.SetElementOverrides(region.Id, g_override)

        # Textnote next to the area
        text_point = XYZ(p2.X + mm_to_feet(10), (p2.Y + p3.Y) / 2, 0)
        label_text = filt.Name
        TextNote.Create(doc, new_legend.Id, text_point, label_text, text_type.Id)

        y_offset += region_height + region_spacing
t.Commit()

#forms.alert(" Legends generated and lists of elements displayed in console.")
