# -*- coding: utf-8 -*-

from collections import Counter
from System.Collections.Generic import List
from pyrevit import revit, DB
from pyrevit import script
import sys
from Autodesk.Revit.DB import BuiltInCategory, Group
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ElementMulticategoryFilter
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import BuiltInParameter
from System.Diagnostics import Stopwatch
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

__title__ = 'IFC Parameters\nMODELLGRUPPE'

skipped = 0
newval = 0
nochange = 0

def OwnedByOtherUser(doc, elem):
    checkout_status = WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
    return checkout_status == CheckoutStatus.OwnedByOtherUser

def writeParametersFromLevel(TargetElement, SourceElementID):
    global skipped, newval, nochange 
    try:
        SourceElement = TargetElement.Document.GetElement(SourceElementID)
        if SourceElement is not None:
            Teilobjekt = SourceElement.LookupParameter('Teilobjekt').AsString() or ""
            Geschoss = SourceElement.LookupParameter('Geschoss').AsString() or ""

            if OwnedByOtherUser(TargetElement.Document, TargetElement):
                skipped += 1
            else:
                Teilobjektparameter = TargetElement.LookupParameter('Teilobjekt')
                Geschossparameter = TargetElement.LookupParameter('Geschoss')
                Gebbezeichnungparameter = TargetElement.LookupParameter('Gebäude')
                wastouched = False

                if Teilobjektparameter and not Teilobjektparameter.IsReadOnly and Teilobjektparameter.AsString() != Teilobjekt:
                    Teilobjektparameter.Set(Teilobjekt)
                    wastouched = True

                if Geschossparameter and not Geschossparameter.IsReadOnly and Geschossparameter.AsString() != Geschoss:
                    Geschossparameter.Set(Geschoss)
                    wastouched = True

                if Gebbezeichnungparameter and not Gebbezeichnungparameter.IsReadOnly and Gebbezeichnungparameter.AsString() != Gebbezeichnung:
                    Gebbezeichnungparameter.Set(Gebbezeichnung)
                    wastouched = True

                if wastouched:
                    newval += 1
                else:
                    nochange += 1
    except Exception as e:
        skipped += 1

def writeParametersFromGroup(TargetElement, GroupElement):
    global skipped, newval, nochange 
    try:
        Teilobjekt = GroupElement.LookupParameter('Teilobjekt').AsString() or ""
        Geschoss = GroupElement.LookupParameter('Geschoss').AsString() or ""

        if OwnedByOtherUser(TargetElement.Document, TargetElement):
            skipped += 1
        else:
            Teilobjektparameter = TargetElement.LookupParameter('Teilobjekt')
            Geschossparameter = TargetElement.LookupParameter('Geschoss')
            Gebbezeichnungparameter = TargetElement.LookupParameter('Gebäude')
            wastouched = False

            if Teilobjektparameter and not Teilobjektparameter.IsReadOnly and Teilobjektparameter.AsString() != Teilobjekt:
                Teilobjektparameter.Set(Teilobjekt)
                wastouched = True

            if Geschossparameter and not Geschossparameter.IsReadOnly and Geschossparameter.AsString() != Geschoss:
                Geschossparameter.Set(Geschoss)
                wastouched = True

            if Gebbezeichnungparameter and not Gebbezeichnungparameter.IsReadOnly and Gebbezeichnungparameter.AsString() != Gebbezeichnung:
                Gebbezeichnungparameter.Set(Gebbezeichnung)
                wastouched = True

            if wastouched:
                newval += 1
            else:
                nochange += 1
    except Exception as e:
        skipped += 1

def process_element(element):
    if element is None:
        return

    if element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_StructuralFraming):
        ref_level_param = element.LookupParameter("Reference Level")
        if ref_level_param:
            ref_id = ref_level_param.AsElementId()
            if ref_id and ref_id.IntegerValue != -1:
                writeParametersFromLevel(element, ref_id)
                return

    host = getattr(element, 'Host', None)
    if host:
        writeParametersFromLevel(element, host.Id)
    elif hasattr(element, 'LevelId'):
        level_id = element.LevelId
        if level_id and isinstance(level_id, ElementId) and level_id.IntegerValue != -1:
            writeParametersFromLevel(element, level_id)

def process_group(group):
    for member_id in group.GetMemberIds():
        member = revit.doc.GetElement(member_id)
        if member is None:
            continue

        if member.Category is not None:
            if member.Category.Id.IntegerValue == int(BuiltInCategory.OST_Railings):
                # Sonderfall: Geländer in Gruppe ohne gültige Ebene
                if not hasattr(member, 'LevelId') or not member.LevelId or member.LevelId.IntegerValue == -1:
                    writeParametersFromGroup(member, group)
                    continue

        process_element(member)

stopwatch = Stopwatch()
stopwatch.Start()

output = script.get_output()
selection = revit.get_selection()
idx = 0
AllElements = []

cat_list = [
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_Railings,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_StructuralFraming
]

typed_list = List[BuiltInCategory](cat_list)
filterCategories = ElementMulticategoryFilter(typed_list)
projectInfo = revit.doc.ProjectInformation
biparam = DB.BuiltInParameter.PROJECT_BUILDING_NAME
Gebbezeichnung = projectInfo.Parameter[biparam].AsString()

if revit.active_view:
    if not selection.is_empty:
        AllElements = selection
    else:
        categoryElements = FilteredElementCollector(revit.doc).WherePasses(filterCategories).WhereElementIsNotElementType().ToElements()
        groupElements = FilteredElementCollector(revit.doc).OfClass(Group).ToElements()
        AllElements.extend(categoryElements)
        AllElements.extend(groupElements)

    with revit.Transaction("pyScript IFC-Parameter"):
        for element in AllElements:
            if isinstance(element, Group):
                process_group(element)
            else:
                process_element(element)
            output.update_progress(idx, len(AllElements))
            idx += 1

output.reset_progress()
stopwatch.Stop()
print('Done. Updated:', newval, 'Skipped:', skipped, 'Unchanged:', nochange)
