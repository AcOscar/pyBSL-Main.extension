# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector as Fec,
    BuiltInCategory, RevitLinkInstance
)
from System.Diagnostics import Stopwatch
from rpw import doc
from rpw.db import Transaction as rpw_Transaction
from pyrevit import script

stopwatch = Stopwatch()
stopwatch.Start()
output = script.get_output()

"""settings"""
"""
linked model name without .rvt extension
at least first few letters must match the name in the project browser
if there are multiple links with similar names, the first one found will be used"""
linked_modelname = "Fassadenmodell"  

"""parameter names"""
exclude_param = "Fensterflaeche_Exklusion"
glazingarea_param_name = "Glasflaeche"
window_area_param = "Fensterflaeche_Tag"


room_offset = -0.3  # offset from window center to room in meters

logger = script.get_logger() # Initialisiert den Logger

def get_inked_instance_by_name(name):
    """Get Revit link instance by name."""
    return next(
        (inst for inst in Fec(doc).OfClass(RevitLinkInstance)
         if inst.Name.startswith(name)), None)  


def main():
    filtered_windows = []
    w2calc = []
    fromRoomId = []
    fromRoom = []
    wRoomSet = []

    linked_instance = get_inked_instance_by_name(linked_modelname)

    if not linked_instance:
        print("Linked model ‘{}’ not found.".format(linked_modelname))
        return

    linked_doc = linked_instance.GetLinkDocument()
    transform = linked_instance.GetTotalTransform()

    windows = Fec(linked_doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()

    if not windows:
        print("No windows found in the linked model.")
        return

    """filter windows based on exclusion parameter"""
    for window in windows:
        try:
            if not window.LookupParameter(exclude_param).AsInteger():
                filtered_windows.append(window)
        except:
            continue

    if not filtered_windows:
        print("No evaluable windows found.")
        return

    """ Room detection with virtual room point (offset from center of bounding box)"""
    for filtered_window in filtered_windows:
        try:
            bbox = filtered_window.get_BoundingBox(linked_doc.ActiveView)
            if not bbox:
                continue
            midpt = (bbox.Min + bbox.Max) * 0.5

            """Use host wall orientation, if available"""
            try:
                host = filtered_window.Host
                facing = host.Orientation
                adjusted_pt = midpt + (facing * room_offset)
            except:
                adjusted_pt = midpt
            
            """adjust point to project coordinates"""
            pt = transform.OfPoint(adjusted_pt)

            """get room at point"""
            from_room = doc.GetRoomAtPoint(pt)

            if from_room:
                fromRoom.append(from_room)
                fromRoomId.append(from_room.Id)
                w2calc.append(filtered_window)
        except:
            continue

    if not fromRoomId:
        print("No rooms found for the windows.")
        return

    if not w2calc:
        print("No windows to calculate.")
        return

    """Sorting and grouping"""
    fromRoomId, w2calc = (list(x) for x in zip(*sorted(zip(fromRoomId, w2calc), reverse=False)))
    fromRoomSet = list(set(fromRoomId))

    for d, r in zip(w2calc, fromRoomId):
        wRoomSet[fromRoomSet.index(r)].append(d)

    """Calculate and set window areas per room"""
    with rpw_Transaction("Window_area_per_room_update"):
        for si, wrs in enumerate(wRoomSet):
            room = doc.GetElement(fromRoomSet[si])
            roomSum = 0

            for window in wrs:
                try:
                    glazingarea_param = window.LookupParameter(glazingarea_param_name)
                    if not glazingarea_param:
                        glazingarea_param = window.Symbol.LookupParameter(glazingarea_param_name)

                    if glazingarea_param and glazingarea_param.HasValue:
                        glasflaeche = glazingarea_param.AsDouble()
                        roomSum += glasflaeche
                    else:
                        logger.info("Window {} does not have a valid '{}'".format(output.linkify(window.Id), glazingarea_param_name))
                except Exception as e:
                    logger.error("Error at window {}: {}".format(window.Id,e))
            sumStr = str(round(roomSum, 2)) 
            param = room.LookupParameter(window_area_param)
            
            if param:
                param.Set(roomSum)
            else:
                logger.info("Room {} has no parameter '{}'".format(room.Id,window_area_param))

    stopwatch.Stop()
    logger.info("WindowAreaPerRoom run in: {}" .format (stopwatch.Elapsed))

if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:
        import traceback
        logger.error("ERROR:", e)
        logger.error(traceback.format_exc())