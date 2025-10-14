# -*- coding: utf-8 -*-
from pyrevit import forms

# Revit Core
from Autodesk.Revit.DB import (
    BuiltInCategory, 
    BuiltInParameter,
    ElementId,
    ElementMulticategoryFilter,
    ElementTypeGroup,
    FilteredElementCollector,
    ParameterFilterElement,
    SelectionFilterElement,
    Transaction,
    View,
    ViewDuplicateOption,
)
# Revit Geometry / Graphics
from Autodesk.Revit.DB import (
    CurveLoop,
    FilledRegion,
    Line,
    OverrideGraphicSettings,
    XYZ,
)
# Revit Types
from Autodesk.Revit.DB import (
    CeilingType,
    FilledRegionType,
    FloorType,
    StorageType,
    TextNote,
    TextNoteType,
    ViewType,
    WallType,
)

from System import Enum
from System.Collections.Generic import List

__title__ = 'ViewFilterLegend Floor'
__doc__ = 'Create a legend to all type filters from a view'

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# P4  region_width_mm  P3
# +----------------------+
# ¦                      ¦ region_height_mm    "Arial_normal_schwarz_1.5mm"                  "Arial_normal_schwarz_1.5mm"
# ¦                      ¦                      text_distance_y_mm
# +----------------------+ text_distance_x_mm   ¦             text2_distance_mm = 4000       ¦
# P1                    P2 

region_width_mm = 390
region_height_mm = 200
region_spacing_mm = 200
text_distance_x_mm = 1100
text_distance_y_mm = 175
text2_distance_mm = 4000
text_type_name_default = "Arial_normal_schwarz_1.5mm"

unsupported = set([
    ViewType.ProjectBrowser,
    ViewType.DrawingSheet,
    ViewType.Schedule,
    ViewType.Report,
    ViewType.SystemBrowser,
    ViewType.Internal,
    ViewType.Undefined,
    ViewType.DraftingView  
])

# sources for categories
cats_py = [
BuiltInCategory.OST_Floors, 
BuiltInCategory.OST_Walls, 
BuiltInCategory.OST_Ceilings
]

paramter_name_toread = "Materialbeschreibung"

# ElementMulticategoryFilter
cats_net = List[BuiltInCategory](cats_py)
multi_filter = ElementMulticategoryFilter(cats_net)

# Allowed-Set of IntegerValue der Category.Id)
allowed = {int(c) for c in cats_py}



def mm_to_feet(mm):
    return float(mm) / 304.8

def get_textnotetypenames(doc):
    text_note_types_names = []
    types = list(FilteredElementCollector(doc).OfClass(TextNoteType))
    for t in types:
        tname = (t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "").strip()
        if tname:
            text_note_types_names.append(tname)
    return text_note_types_names

def get_text_typename(doc,texttypesn):
    text_type_name = []
    text_type_name = forms.SelectFromList.show(
        texttypesn,
        multiselect=True,
        title="Select text type"
    )

    if not text_type_name:
        text_type_name = []
        text_type_name.append(text_type_name_default)

    return text_type_name[0]

def select_texttype(doc):
    texttypesn = get_textnotetypenames(doc)

    text_typename = get_text_typename(doc, texttypesn)

    text_type = get_textnotetype(doc, text_typename)

    return text_type

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

def get_parameter_value_safely(element, param_name):
    """
    Safely get parameter value from element
    
    Args:
        element: Revit element
        param_name (str): Parameter name
        
    Returns:
        str: Parameter value or empty string if not found
    """
    try:
        param = element.LookupParameter(param_name)
        if param:
            if param.StorageType == StorageType.String:
                return param.AsString() or ""
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                return str(param.AsDouble())
            elif param.StorageType == StorageType.ElementId:
                return str(param.AsElementId().IntegerValue)
    except:
        pass
    return ""

def get_type_names_hit_by_filter(doc, view, filt, max_names=None,only_existing=True):
    """
    Returns sorted, unique Type names that are actually affected by the filter filt in the view view
    
    Args:
        doc:
        view:
        filt:
        max_names:
        only_existing: check if the filterd elemts realy exist in the view
        
    Returns:
        list: List of str typenames
    """
    print("++++++++++")
    print("Filter: " + filt.Name)
    
    names = set()
    # try:
    # Case A: parameter filter
    if isinstance(filt, ParameterFilterElement):
        filt_cat = filt.GetCategories()
        
        ef = filt.GetElementFilter()  # Autodesk.Revit.DB.ElementFilter
        
        if only_existing:
            
            ele_col = (FilteredElementCollector(doc, view.Id)
                         .WherePasses(multi_filter)
                         .WhereElementIsNotElementType()
                         .WherePasses(ef))        
        
        for c in filt_cat:
            if only_existing:

                for ele in ele_col:
                    elet = doc.GetElement(ele.GetTypeId())
                    
                    if c != ele.Category.Id:
                        continue
                    # para_val =elet.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    # para_val = elet.LookupParameter(paramter_name_toread).AsString()
                    
                    para_val = get_parameter_value_safely(elet, paramter_name_toread)
                    # print(para_val)
                    
                    if elet.Category.Id.IntegerValue not in allowed:
                        continue
                        
                    if isinstance(para_val, str) and para_val.strip():
                        names.add(para_val.strip())


    # Case B: SelectionFilterElement
    elif isinstance(filt, SelectionFilterElement):
        # col = (FilteredElementCollector(doc, view.Id)
                     # .OfCategory(BuiltInCategory.OST_Walls)
                     # .WhereElementIsNotElementType()
                     # .ToElementIds())
        col = (FilteredElementCollector(doc, view.Id)
                     .WherePasses(multi_filter)
                     .WhereElementIsNotElementType()
                     .ToElementIds())     
         
        sel_ids = set(eid.IntegerValue for eid in filt.GetElementIds())
        for eid in col:
            if eid.IntegerValue in sel_ids:
                try:
                    w = doc.GetElement(eid)
                    wt = doc.GetElement(w.GetTypeId())
                    # if isinstance(wt, WallType) and wt.Name:
                        # names.add(wt.Name)
                    print(wt_name)
                    if isinstance(wt, FloorType) and wt.Name:
                        names.add(wt.Name)
                    if isinstance(wt, CeilingType) and wt.Name:
                        names.add(wt.Name)
                except:
                    pass

    # Fallback: if another filter element nevertheless offers GetElementFilter()
    else:
        getef = getattr(filt, "GetElementFilter", None)
        if getef:
            ef = getef()
            # col = (FilteredElementCollector(doc, view.Id)
                   # .OfCategory(BuiltInCategory.OST_Walls)
                   # .WhereElementIsNotElementType()
                   # .WherePasses(ef))
            col = (FilteredElementCollector(doc, view.Id)
                   .WherePasses(multi_filter)
                   .WhereElementIsNotElementType()
                   .ToElementIds())
            for w in col:
                try:
                    wt = doc.GetElement(w.GetTypeId())
                    print(wt_name)
                    # if isinstance(wt, WallType) and wt.Name:
                        # names.add(wt.Name)
                    if isinstance(wt, FloorType) and wt.Name:
                        names.add(wt.Name)
                    if isinstance(wt, CeilingType) and wt.Name:
                        names.add(wt.Name)
                except:
                    pass
    # except:
        # pass
    if not names:
        return []
    res = sorted(names, key=lambda s: s.lower())
    if max_names and len(res) > max_names:
        return res[:max_names] 
    print(res)
    return res

def get_safe_view(doc):
    views = []

    for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType():
        try:
            if v.IsTemplate:
                continue
            if v.ViewType in unsupported:
                continue
            # keine GetFilters() hier!
            views.append(v)
        except:
            continue

    if not views:
        forms.alert("No eligible views found.", exitscript=True)
        
    return views

def get_view_names(doc,safe_views):
    # Aktuelle View ermitteln
    current_view = doc.ActiveView
    current_view_label = None

    # schlanke Items: Label -> ElementId
    items = []
    id_by_label = {}
    for v in safe_views:
        if v.GetFilters():

            label = u"{name}  [{vt}]  (Id:{id})".format(
                name=v.Name,
                vt=Enum.GetName(ViewType, v.ViewType),
                id=v.Id.IntegerValue
            )
            items.append(label)
            id_by_label[label] = v.Id

            # Prüfen ob dies die aktuelle View ist
            if v.Id == current_view.Id:
                current_view_label = label

    return items,id_by_label

def select_legend(safe_views):
    legend_views = [v for v in safe_views if v.ViewType == ViewType.Legend and not v.IsTemplate]

    if not legend_views:
        forms.alert("You must have at least one legend view in the project.", exitscript=True)

    legend_view_selected = forms.SelectFromList.show(
        legend_views,
        name_attr="Name",
        title="Select Legend View"
    )

    if not legend_view_selected:
        forms.alert("You did not select any legend views. Canceling...", exitscript=True)

    return legend_views,legend_view_selected

def main():
    """
    Main function
    """
    # ==========================================
    # GET ALL VIEWS WITH FILTERS AND A SELECTION
    # ==========================================
    selected_views,views = select_views(doc)

    # ==========================================
    # SELECT TEXT TYPE
    # ==========================================
    text_type = select_texttype(doc)
    
    # ==========================================
    # SELECT VIEW LEGEND 
    # ==========================================
    legend_views,legend_view_selected = select_legend(views)

    # ==========================================
    # REQUEST DIMENSIONS IN mm
    # ==========================================
    region_width_mmm = forms.ask_for_string(default=str(region_width_mm), prompt="Width FilledRegion (mm):")
    region_height_mmm = forms.ask_for_string(default=(str(region_height_mm)), prompt="Heigth FilledRegion (mm):")
    region_spacing_mmm = forms.ask_for_string(default=str(region_spacing_mm), prompt="Spacing FilledRegion (mm):")
    
    # ==========================================
    # REQUEST OPTIONS
    # ==========================================
    ops = ['only Filter', 'wtith Parameter Materialbeschreibung']
    opp = forms.CommandSwitchWindow.show(ops, message='Select Option')

    if opp == 'only Filter':
        filters_only = True
    else:
        filters_only = False

    try:
        region_width = mm_to_feet(float(region_width_mmm))
        region_height = mm_to_feet(float(region_height_mmm))
        region_spacing = mm_to_feet(float(region_spacing_mmm))
    except:
        forms.alert("Error in dimensions. Make sure you enter only numbers.", exitscript=True)
        
    # ==========================================
    # GET A REGIONTYPE TO COPY AND OVERWRITE
    # ==========================================
    region_type = get_region_type(doc)   
        
    # ============================
    # START TRANSACTION
    # ============================
    t = Transaction(doc, "Generate Legend with Filters")
    t.Start()

    for view in selected_views:
        
        new_legend = get_new_legend(doc, legend_view_selected, legend_views, view.Name)
        
        x_origin = mm_to_feet(20)
        y_origin = mm_to_feet(20)
        
        p = XYZ(x_origin, y_origin,0)

        filters = view.GetFilters()

        #print("View '{}' has {} Filters".format(view.Name, len(filters)))
        for f_id in filters:
            # Filterelement holen
            filt = doc.GetElement(f_id)
            override = view.GetFilterOverrides(f_id)

            if not filt or not isinstance(filt, ParameterFilterElement):
                continue  # not a parameter filter
            if filters_only:
                g_override = get_graphic_override(override)
                p = XYZ(p.X, p.Y + region_height + region_spacing, 0)
                draw_row(doc, new_legend, p,region_type, g_override, region_height, region_width,text_type.Id,text_distance_x_mm, filt.Name)

            else:
                type_names = get_type_names_hit_by_filter(doc, view, filt)
                for t_name in type_names:
                    g_override = get_graphic_override(override)
                    p = XYZ(p.X, p.Y + region_height + region_spacing, 0)
                    draw_row(doc, new_legend, p, region_type, g_override, region_height, region_width,text_type.Id,text_distance_x_mm, t_name,text2_distance_mm,filt.Name)
    t.Commit()
 
def select_views(doc):
    
    safe_views = []
    
    safe_views = get_safe_view(doc)

    view_names,id_by_label = get_view_names(doc, safe_views)

    selected_labels = forms.SelectFromList.show(
        view_names,
        multiselect=True,
        title="Select Views (filters will be checked after selection)",
       )

    if not selected_labels:
        forms.alert("No views were selected. Canceling...", exitscript=True)

    selected_views = [doc.GetElement(id_by_label[lbl]) for lbl in selected_labels]

    views_with_filters = []
    for v in selected_views:
        try:
            flt_ids = v.GetFilters()  # jetzt sicher(er)
            if flt_ids and len(flt_ids) > 0:
                views_with_filters.append(v)
        except:
            continue

    if not views_with_filters:
        forms.alert("Selected views have no filters.", exitscript=True)

    return selected_views,safe_views

def get_region_type(doc):

    region_type = FilteredElementCollector(doc)\
        .OfClass(FilledRegionType)\
        .FirstElement()

    if not region_type:
        forms.alert("No FilledRegionType found!", exitscript=True)
    
    return region_type

def get_new_legend(doc, legend_view_selected, legend_views, new_name):
    
    new_legend_id = legend_view_selected.Duplicate(ViewDuplicateOption.WithDetailing)
    new_legend = doc.GetElement(new_legend_id)

    # check if the name already exist and count up in this case
    base_name = "Legend_{}".format(new_name)
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
    
    return new_legend

def get_graphic_override(override):
    
    g_override = OverrideGraphicSettings()
    
    cfg_pattern_id = override.CutForegroundPatternId if override else None
    cfg_color = override.CutForegroundPatternColor if override else None
    cbg_pattern_id = override.CutBackgroundPatternId if override else None
    cbg_color = override.CutBackgroundPatternColor if override else None

    sfg_pattern_id = override.SurfaceForegroundPatternId if override else None
    sfg_color = override.SurfaceForegroundPatternColor if override else None
    sbg_pattern_id = override.SurfaceBackgroundPatternId if override else None
    sbg_color = override.SurfaceBackgroundPatternColor  if override else None

    if cfg_pattern_id != ElementId.InvalidElementId:
        g_override.SetSurfaceForegroundPatternId (cfg_pattern_id)
    if cfg_color != ElementId.InvalidElementId:
        g_override.SetSurfaceForegroundPatternColor (cfg_color)
    if cbg_pattern_id != ElementId.InvalidElementId:
        g_override.SetSurfaceBackgroundPatternId (cbg_pattern_id)
    if cbg_color != ElementId.InvalidElementId:
        g_override.SetSurfaceBackgroundPatternColor (cbg_color)
        
    if sfg_pattern_id != ElementId.InvalidElementId:
        g_override.SetSurfaceForegroundPatternId (sfg_pattern_id)
    if sfg_color != ElementId.InvalidElementId:
        g_override.SetSurfaceForegroundPatternColor (sfg_color)
    if sbg_pattern_id != ElementId.InvalidElementId:
        g_override.SetSurfaceBackgroundPatternId (sbg_pattern_id)
    if sbg_color != ElementId.InvalidElementId:
        g_override.SetSurfaceBackgroundPatternColor (sbg_color)

    return g_override

def get_loop(point,height,width):
    
    p1 = XYZ(point.X, point.Y , 0)
    p2 = XYZ(point.X + width, point.Y  , 0)
    p3 = XYZ(point.X + width, point.Y + height, 0)
    p4 = XYZ(point.X, point.Y + height, 0)

    loop = CurveLoop()
    
    loop.Append(Line.CreateBound(p1, p2))
    loop.Append(Line.CreateBound(p2, p3))
    loop.Append(Line.CreateBound(p3, p4))
    loop.Append(Line.CreateBound(p4, p1))
    
    return loop

def draw_row (doc, legend, point, region_type, override, height, width, tt_id, distance1, text1, distance2=0, text2=""):
    """
    draw a row wtih a filled region and one ort two text
    
    Args:
        doc:
        legend:
        point: XYZ of the starting point from the row
        override: the graphical override of the filedregion
        tt_id: texttype id for the text to generate
        distance1: the distance along the x axis in mm from point
        text1: the text string
        distance2: the distance along the x axis in mm from text1
        text2: the 2nd text string

    Returns:
        True if scussefull
    """
 
    loop = get_loop(point,height,width)
    
    # FilledRegion
    region = FilledRegion.Create(doc, region_type.Id, legend.Id, [loop])

    legend.SetElementOverrides(region.Id, override)

    # Textnote next to the area
    text_point = XYZ(point.X + mm_to_feet(distance1), point.Y , 0)

    TextNote.Create(doc, legend.Id, text_point, text1, tt_id)

    if text2:
        text_point2 = XYZ(text_point.X + mm_to_feet(distance2), text_point.Y , 0)
        TextNote.Create(doc, legend.Id, text_point2, text2, tt_id)

    return True 

# ---------------------------
# Start
# ---------------------------
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("ERROR:", e)
        print(traceback.format_exc())