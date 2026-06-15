# -*- coding: utf-8 -*-

import clr
from pyrevit import revit
from pyrevit import script
from pyrevit.coreutils import logger
from Autodesk.Revit.DB import TransactionGroup
from Autodesk.Revit.DB import XYZ, ElementTransformUtils
from rpw import db, doc, uidoc
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import BuiltInParameter as Bip
from Autodesk.Revit.DB import Transaction, IFailuresPreprocessor, FailureProcessingResult, FailureSeverity
from Autodesk.Revit.DB import UV, LinkElementId, ElementId, Line
from System.Collections.Generic import List

class SuppressWarnings(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for failure in failuresAccessor.GetFailureMessages():
            if failure.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue

        
__doc__ = 'Unplace all rooms from current view'\
          'and replace them with the associated level of the current view.'
         
output = script.get_output()
mlogger = logger.get_logger('Room Replace')

def get_all_room_tags(all_tags=None):
    
    room_tags = {}
    mlogger.info("Caching room tags...")
    mlogger.info("Total tags found: {}".format(len(all_tags)))
    for tag in all_tags:
        mlogger.info("Checking tag: {} in view {}".format(tag.Id, doc.GetElement(tag.OwnerViewId).Name))
        
        tagged_room = None
        
        if hasattr(tag, "GetTaggedLocalElements"):
            elems = tag.GetTaggedLocalElements()
            if elems:
                tagged_room = elems[0]
        else:
            try:
                for prop in clr.GetClrType(type(tag)).GetProperties():
                    if prop.Name == "Room" and prop.DeclaringType.Name == "RoomTag":
                        tagged_room = prop.GetValue(tag, None)
                        break
            except Exception:
                pass
                
        if tagged_room and hasattr(tagged_room, "Id"):
            mlogger.info("Found tag for room: {}".format(tagged_room.Id))
            r_id = tagged_room.Id.IntegerValue
            if r_id not in room_tags:
                room_tags[r_id] = []
            room_tags[r_id].append(tag)
    return room_tags


def get_rooms_to_process(doc, uidoc, active_view):
    """Holt die auszuwertenden Raeume (Auswahl oder alle in der Ansicht)."""
    selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds() if doc.GetElement(elId).Category.BuiltInCategory == Bic.OST_Rooms]
    if selection:
        return selection
    return list(Fec(doc, active_view.Id).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements())


def extract_tag_info(tag):
    """Liest alle relevanten Informationen aus einem bestehenden Tag aus."""
    pt = None
    if hasattr(tag, "TagHeadPosition"):
        pt = tag.TagHeadPosition
    elif tag.Location and hasattr(tag.Location, "Point"):
        pt = tag.Location.Point
    
    rot = 0.0
    if tag.Location and hasattr(tag.Location, "Rotation"):
        rot = tag.Location.Rotation
    
    elbow_pt = None
    if getattr(tag, "HasLeader", False) and hasattr(tag, "LeaderElbow"):
        elbow_pt = tag.LeaderElbow

    elbowend_pt = None
    if getattr(tag, "HasLeader", False) and hasattr(tag, "LeaderEnd"):
        elbowend_pt = tag.LeaderEnd
    
    if pt:
        return {
            'view_id': tag.OwnerViewId,
            'point': pt,
            'type_id': tag.GetTypeId(),
            'has_leader': getattr(tag, "HasLeader", False),
            'orientation': getattr(tag, 'TagOrientation', None),
            'rotation': rot,
            'elbow': elbow_pt,
            'elbow_end': elbowend_pt
        }
    return None


def unplace_room_safely(doc, room, pinned):
    """Entplatziert den Raum sicher und unterdrueckt Warnungen in einer eigenen Transaktion."""
    t_unplace = Transaction(doc, "Unplace Room")
    t_unplace.Start()
    options = t_unplace.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(SuppressWarnings())
    t_unplace.SetFailureHandlingOptions(options)
    
    if pinned:
        room.Pinned = False 
        mlogger.info("Unpinned room: {}".format(room.Id))

    room.Unplace()
    mlogger.info("Unplaced room: {}".format(room.Id))
    t_unplace.Commit()


def restore_tags(doc, newRoom, room_tags_info):
    """Stellt die zuvor gesammelten Raumbeschriftungen samt Position, Rotation und Knickpunkten wieder her."""
    for t_info in room_tags_info:
        view_id = t_info['view_id']
        view = doc.GetElement(view_id)
        if not view:
            continue
        
        uv_pt = UV(t_info['point'].X, t_info['point'].Y)
        link_id = LinkElementId(newRoom.Id)
        new_tag = doc.Create.NewRoomTag(link_id, uv_pt, view_id)
        
        if new_tag:
            if t_info['type_id']:
                try:
                    new_tag.ChangeTypeId(t_info['type_id'])
                except Exception:
                    pass
            saved_rot = t_info.get('rotation', 0.0)
            current_rot = 0.0
            if new_tag.Location and hasattr(new_tag.Location, "Rotation"):
                current_rot = new_tag.Location.Rotation
            
            rot_diff = saved_rot - current_rot
            if abs(rot_diff) > 0.001:
                try:
                    axis_pt = XYZ(t_info['point'].X, t_info['point'].Y, 0.0)
                    axis = Line.CreateBound(axis_pt, XYZ(axis_pt.X, axis_pt.Y, 1.0))
                    ElementTransformUtils.RotateElement(doc, new_tag.Id, axis, rot_diff)
                except Exception:
                    pass  
                
            if hasattr(new_tag, "HasLeader"):
                new_tag.HasLeader = t_info['has_leader']

            if t_info['has_leader'] and t_info.get('elbow_end') and hasattr(new_tag, "LeaderEnd"):
                try:
                    new_tag.LeaderEnd = t_info['elbow_end']
                except Exception as e:
                    mlogger.warning("Could not set leader end: {}".format(e)) 
            
            if t_info['has_leader'] and t_info.get('elbow') and hasattr(new_tag, "LeaderElbow"):
                try:
                    new_tag.LeaderElbow = t_info['elbow']
                except Exception as e:
                    mlogger.warning("Could not set leader elbow: {}".format(e))

            if t_info['orientation'] is not None and hasattr(new_tag, "TagOrientation"):
                new_tag.TagOrientation = t_info['orientation']

            current_pt = None
            if hasattr(new_tag, "TagHeadPosition"):
                current_pt = new_tag.TagHeadPosition
            elif new_tag.Location and hasattr(new_tag.Location, "Point"):
                current_pt = new_tag.Location.Point
                
            if current_pt:
                tag_move = XYZ(t_info['point'].X - current_pt.X, t_info['point'].Y - current_pt.Y, 0.0)
                if tag_move.GetLength() > 0.001:
                    try:
                        ElementTransformUtils.MoveElement(doc, new_tag.Id, tag_move)
                    except Exception:
                        pass

            mlogger.info("Restored tag in view: {}".format(view.Name))


def main():

    idx = 0

    transGroup = TransactionGroup(revit.doc, "pyScript replace rooms")
    transGroup.Start()

    active_view = uidoc.ActiveView
    base_level = active_view.GenLevel
    
    if not base_level:
        mlogger.error("Fehler: Die aktuelle Ansicht hat keine verknuepfte Ebene (z.B. 3D-Ansicht).")
        return

    rooms = get_rooms_to_process(revit.doc, uidoc, active_view)

    """check if we have rooms and the necessary parameters"""
    if not rooms:
        mlogger.warning("No rooms found.")
        return

    EleNums = len(rooms)         

    # Alle Raumbeschriftungen vorab im Dokument suchen und cachen
    all_tags = Fec(doc).OfCategory(Bic.OST_RoomTags).WhereElementIsNotElementType().ToElements()
    
    tags_by_room = get_all_room_tags(all_tags)


    for room in rooms:
        mlogger.info(output.linkify(room.Id))
        if room.Area == 0:
            mlogger.warning("Room {} has zero area, skipping.".format(room.Id))
            continue
        
        rmPT = room.Location
        roomPoint = rmPT.Point
        upper_level_param = room.get_Parameter(Bip.ROOM_UPPER_LEVEL)
        upper_level_id = upper_level_param.AsElementId() if upper_level_param else None
        pinned = room.Pinned 
        
        offset = room.LimitOffset
        
        # Tags für diesen speziellen Raum sammeln, bevor er entplatziert wird
        room_id_int = room.Id.IntegerValue
        room_tags_info = []
        if room_id_int in tags_by_room:
            for tag in tags_by_room[room_id_int]:
                t_info = extract_tag_info(tag)
                if t_info:
                    room_tags_info.append(t_info)

        unplace_room_safely(revit.doc, room, pinned)

        with db.Transaction("New Room"):

            topo = revit.doc.get_PlanTopology(base_level)
            circuits = topo.Circuits

            for circuit in circuits:
            
                if not circuit.IsRoomLocated:
                    # Zwischenspeichern aller Tag-IDs VOR der Raumerstellung, um automatische Tags zu erkennen
                    tags_before_ids = set(Fec(doc).OfCategory(Bic.OST_RoomTags).WhereElementIsNotElementType().ToElementIds())

                    # Raum erstellen (dies kann automatisch einen Tag erzeugen, wenn "Tag on Placement" aktiv ist)
                    newRoom = revit.doc.Create.NewRoom(room, circuit)

                    # Abrufen aller Tag-IDs NACH der Raumerstellung
                    tags_after_ids = Fec(doc).OfCategory(Bic.OST_RoomTags).WhereElementIsNotElementType().ToElementIds()

                    # Differenz finden, um den automatisch erstellten Tag zu identifizieren und sofort zu löschen
                    auto_tag_ids = [t_id for t_id in tags_after_ids if t_id not in tags_before_ids]
                    if auto_tag_ids:
                        revit.doc.Delete(List[ElementId](auto_tag_ids))
                        mlogger.info("Entferne {} automatisch platzierten Tag(s).".format(len(auto_tag_ids)))

                    if upper_level_id and upper_level_id.IntegerValue != -1:
                        newRoom.get_Parameter(Bip.ROOM_UPPER_LEVEL).Set(upper_level_id)
                    newRoom.LimitOffset = offset
                    move = XYZ(roomPoint.X - rmPT.Point.X, roomPoint.Y - rmPT.Point.Y, 0.0)
                    ElementTransformUtils.MoveElement(revit.doc, newRoom.Id, move)
                    if pinned:
                        newRoom.Pinned = pinned
                    mlogger.info("Created new room: {}".format(newRoom.Id))
                    
                    # Gespeicherte Raumbeschriftungen wiederherstellen
                    restore_tags(revit.doc, newRoom, room_tags_info)
                    
                    break
                
        output.update_progress(idx, EleNums)
        idx+=1   

    transGroup.Assimilate()

    output.reset_progress()

    mlogger.info('Done')

if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:        
        import traceback
        mlogger.error("ERROR: {}".format(e))
        mlogger.error(traceback.format_exc())