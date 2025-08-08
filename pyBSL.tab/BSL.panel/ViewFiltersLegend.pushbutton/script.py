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
# OBTENER TODAS LAS VISTAS CON FILTROS
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
    forms.alert("No hay vistas con filtros en el proyecto.", exitscript=True)

selected_views = forms.SelectFromList.show(
    views_with_filters,
    multiselect=True,
    name_attr="Name",
    title="Select View Name"
)

if not selected_views:
    forms.alert("No se seleccionaron vistas. Cancelando...", exitscript=True)

# ============================
# SELECCIONAR ESTILO DE TEXTO
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
    forms.alert("No se seleccionó estilo de texto. Cancelando...", exitscript=True)

text_type = selected_wrapper.text_type

# ============================
# SELECCIONAR VISTA LEYENDA BASE
# ============================
legend_views = [v for v in all_views if v.ViewType == ViewType.Legend and not v.IsTemplate]

if not legend_views:
    forms.alert("Debes tener al menos una vista de leyenda en el proyecto.", exitscript=True)

legend_view_selected = forms.SelectFromList.show(
    legend_views,
    name_attr="Name",
    title="Select Legend View"
)

if not legend_view_selected:
    forms.alert("No seleccionaste ninguna vista de leyenda. Cancelando...", exitscript=True)

# ============================
# PEDIR DIMENSIONES EN mm
# ============================
region_width_mm = forms.ask_for_string(default="200", prompt="Width FilledRegion (mm):")
region_height_mm = forms.ask_for_string(default="50", prompt="Heigth FilledRegion (mm):")
region_spacing_mm = forms.ask_for_string(default="10", prompt="Spacing FilledRegion (mm):")

try:
    region_width = mm_to_feet(float(region_width_mm))
    region_height = mm_to_feet(float(region_height_mm))
    region_spacing = mm_to_feet(float(region_spacing_mm))
except:
    forms.alert("Error en las dimensiones. Asegúrate de ingresar solo números.", exitscript=True)

# ============================
# COMENZAR TRANSACCIÓN
# ============================
t = Transaction(doc, "Generar Leyenda con Filtros")
t.Start()

for view in selected_views:
    new_legend_id = legend_view_selected.Duplicate(ViewDuplicateOption.Duplicate)
    new_legend = doc.GetElement(new_legend_id)
    new_legend.Name = "Legend_{}".format(view.Name)

    # Mostrar elementos visibles en consola
    print("\n🔍 Vista: {}".format(view.Name))
    print("Elementos visibles:")

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
        print(" - (ningún elemento visible con nombre)")
    else:
        for name in element_names:
            print(" - {}".format(name))

    # Listar muros visibles
    walls_collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    wall_names = []
    for w in walls_collector:
        try:
            name = w.Name
            if name and name.strip():
                wall_names.append(name)
        except:
            continue

    print("Muros visibles en la vista '{}':".format(view.Name))
    if wall_names:
        for name in wall_names:
            print(" - {}".format(name))
    else:
        print(" - (ningún muro visible)")

    # Aquí puedes agregar la lógica para crear FilledRegions con colores del filtro si quieres

t.Commit()

forms.alert(" Leyendas generadas y listas de elementos mostradas en consola.")
