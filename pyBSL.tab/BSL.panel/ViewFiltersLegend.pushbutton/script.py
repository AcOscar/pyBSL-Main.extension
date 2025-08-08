# -*- coding: utf-8 -*-
from pyrevit import forms
from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def mm_to_feet(mm):
    return float(mm) / 304.8

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
# SELECT TEXT STYLEXTO
# ============================
text_types = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()

class TextTypeWrapper(object):
    def __init__(self, text_type):
        self.text_type = text_type
        self.Name = text_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()

text_wrappers = [TextTypeWrapper(tt) for tt in text_types]

selected_wrapper = forms.SelectFromList.show(
    text_wrappers,
    name_attr="Name",
    title="Select Text Type"
)

if not selected_wrapper:
    forms.alert("No text style was selected. Canceling...", exitscript=True)

text_type = selected_wrapper.text_type

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
region_width_mm = forms.ask_for_string(default="200", prompt="Width FilledRegion (mm):")
region_height_mm = forms.ask_for_string(default="50", prompt="Heigth FilledRegion (mm):")
region_spacing_mm = forms.ask_for_string(default="10", prompt="Spacing FilledRegion (mm):")

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
    new_legend_id = legend_view_selected.Duplicate(ViewDuplicateOption.Duplicate)
    new_legend = doc.GetElement(new_legend_id)
    new_legend.Name = "Legend_{}".format(view.Name)

    # Show visible elements in console
    print("\n🔍 View: {}".format(view.Name))
    print("Visible elements:")

    collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    element_names = []

    for el in collector:
        try:
            name = el.Name
            if name and name.strip():
                element_names.append(name)
        except:
            continue

    if not element_names:
        print(" - (nno visible elements with name)")
    else:
        for name in element_names:
            print(" - {}".format(name))

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

    print("Walls visible in the view'{}':".format(view.Name))
    if wall_names:
        for name in wall_names:
            print(" - {}".format(name))
    else:
        print(" - (no visible walls)")

    # Here you can add logic to create FilledRegions with colors from the filter if you want.

t.Commit()

forms.alert(" Legends generated and lists of elements displayed in console.")
