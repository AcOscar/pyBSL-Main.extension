# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector as Fec,
    BuiltInCategory, BuiltInParameter, RevitLinkInstance,
    UnitUtils, UnitTypeId, XYZ
)
from System.Diagnostics import Stopwatch
from rpw import doc
from rpw.db import Transaction as rpw_Transaction
import sys
from pyrevit import script

stopwatch = Stopwatch()
stopwatch.Start()
output = script.get_output()

# === Einstellungen ===
verlinkter_modellname = "Fassadenmodell"  # <- Name des Links im Projektbrowser (ohne .rvt)
exclude_param = "Fensterflaeche_Exklusion"
glazingarea_param = "Glasflaeche"
fflaeche_param = "Fensterflaeche_Tag"
cnvrt = 3.28084  # Fuß in Meter
room_offset = -0.3  # 30 cm "nach innen" als Ersatz für Room Calculation Point

ToRoom = []
filtered_windows = []
w2calc = []
fromRoomId = []
fromRoom = []
wRoomSet = []

# Verlinktes Modell holen
linked_instance = next(
    (inst for inst in Fec(doc).OfClass(RevitLinkInstance)
     if inst.Name.startswith(verlinkter_modellname)), None)

if not linked_instance:
    print("Verlinktes Modell '{}' nicht gefunden.".format(verlinkter_modellname))
    sys.exit()

linked_doc = linked_instance.GetLinkDocument()
transform = linked_instance.GetTotalTransform()

# Fenster im verlinkten Modell sammeln
windows = Fec(linked_doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()

if not windows:
    print("Keine Fenster im verlinkten Modell gefunden.")
    sys.exit()

# Filterung: Nur Fenster ohne Ausschluss
for window in windows:
    try:
        if not window.LookupParameter(exclude_param).AsInteger():
            filtered_windows.append(window)
    except:
        continue

if not filtered_windows:
    print("Keine auswertbaren Fenster gefunden.")
    sys.exit()

# Raumermittlung mit künstlichem Raum-Punkt (Offset von BoundingBox-Mitte)
for filtered_window in filtered_windows:
    try:
        bbox = filtered_window.get_BoundingBox(linked_doc.ActiveView)
        if not bbox:
            continue
        midpt = (bbox.Min + bbox.Max) * 0.5

        # Host-Wandrichtung verwenden, falls vorhanden
        try:
            host = filtered_window.Host
            facing = host.Orientation
            adjusted_pt = midpt + (facing * room_offset)
        except:
            adjusted_pt = midpt

        pt = transform.OfPoint(adjusted_pt)
        from_room = doc.GetRoomAtPoint(pt)

        if from_room:
            fromRoom.append(from_room)
            fromRoomId.append(from_room.Id)
            w2calc.append(filtered_window)
    except:
        continue

if not fromRoomId:
    print("Keine Räume für Fenster gefunden.")
    sys.exit()

if not w2calc:
    print("Keine Fenster zu berechnen.")
    sys.exit()

# Sortieren und Gruppieren
fromRoomId, w2calc = (list(x) for x in zip(*sorted(zip(fromRoomId, w2calc), reverse=False)))
fromRoomSet = list(set(fromRoomId))

for rid in fromRoomSet:
    wRoomSet.append([])

for d, r in zip(w2calc, fromRoomId):
    wRoomSet[fromRoomSet.index(r)].append(d)

# Berechnung und Schreiben
with rpw_Transaction("Fensterflaeche_per_room_update"):
    for si, wrs in enumerate(wRoomSet):
        room = doc.GetElement(fromRoomSet[si])
        roomSum = 0

        for window in wrs:
            try:
                glasflaeche_param = window.LookupParameter(glazingarea_param)
                if not glasflaeche_param:
                    glasflaeche_param = window.Symbol.LookupParameter(glazingarea_param)

                if glasflaeche_param and glasflaeche_param.HasValue:
                    glasflaeche = glasflaeche_param.AsDouble()
                    roomSum += glasflaeche
                else:
                    print("Fenster {} hat keine gültige '{}'".format(output.linkify(window.Id), glazingarea_param))
            except Exception as e:
                print("Fehler bei Fenster {}: {}".format(window.Id,e))
        sumStr = str(round(roomSum, 2)) 
        param = room.LookupParameter(fflaeche_param)
        
        if param:
            param.Set(roomSum)
        else:
            print("Raum {} hat keinen Parameter '{}'".format(room.Id,fflaeche_param))

stopwatch.Stop()
print("FensterflaechePerRoom run in: {}" .format (stopwatch.Elapsed))

