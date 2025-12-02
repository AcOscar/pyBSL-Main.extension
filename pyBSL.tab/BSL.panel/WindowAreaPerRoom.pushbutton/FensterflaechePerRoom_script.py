import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter
from System.Diagnostics import Stopwatch
from rpw import doc
from rpw.db import Transaction as rpw_Transaction
import sys
from pyrevit import script
from Autodesk.Revit.DB import UnitUtils, UnitTypeId

stopwatch = Stopwatch()
stopwatch.Start()
output = script.get_output()

ToRoom = []
exclude_param = "Fensterflaeche_Exklusion"
glazingarea_param = "Glasflaeche"
instW = "Width"
fflaeche_param = "Fensterflaeche_Tag"
filtered_windows = []
w2calc = []
fromRoomId = []
fromRoom = []
wRoomSet = []
cnvrt = 3.280

windows = Fec(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()

if not windows:
    print ("No windows in the project.")
    sys.exit()
    
# find phases
for phase in doc.Phases:
    last_phase = phase

for window in windows:
    try:
        if not window.LookupParameter(exclude_param).AsInteger():
            filtered_windows.append(window)
    except:
        continue  
        #check.append(window)
        
        
if not filtered_windows:
    print ("No windows to count.")
    sys.exit()

for filtered_window in filtered_windows:
    try:
        from_room = filtered_window.FromRoom[last_phase]
        if from_room:
            fromRoom.append(from_room)
            fromRoomId.append(from_room.Id)
            w2calc.append(filtered_window)
    except:
        continue
        
if not fromRoomId:
    print ("No rooms to write on.")
    sys.exit()

if not w2calc:
    print ("No windows to calculate.")
    sys.exit()

# sort windows by FromRoom
fromRoomId, w2calc = (list(x) for x in zip(*sorted(zip(fromRoomId, w2calc), reverse=False)))

fromRoomSet = list(set(fromRoomId))

# build empty sublists per FromRoom set
for rid in fromRoomSet:
    wRoomSet.append([])

# loop through windows and sort them to their according FromRoom sets
for d, r in zip(w2calc, fromRoomId):
    wRoomSet[fromRoomSet.index(r)].append(d)



with rpw_Transaction("Fensterflaeche_per_room_update"):
    # Schleife über alle Fenstergruppen pro Raum
    for si, wrs in enumerate(wRoomSet):
        room = doc.GetElement(fromRoomSet[si])
        roomSum = 0

        for window in wrs:
            try:
                # Glasfläche vom Instanzparameter lesen
                glasflaeche_param = window.LookupParameter(glazingarea_param)
                
                if not glasflaeche_param:
                    glasflaeche_param = window.Symbol.LookupParameter(glazingarea_param)
                    
                if glasflaeche_param and glasflaeche_param.HasValue:
                    glasflaeche = glasflaeche_param.AsDouble() / (cnvrt ** 2)  # Fuß² in m² umrechnen
                    roomSum += glasflaeche
                else:
                    print("Fenster {} hat keine gültige {}".format(output.linkify(window.Id), glazingarea_param))
            except Exception as e:
                print("Fehler bei Fenster {}: {}".format(window.Id,e))

        #print([room.Id.IntegerValue, roomSum])
        # Ergebnis in Textform formatieren
        sumStr = str(round(roomSum, 2)) + " m²"

        # Wert in den Raumparameter schreiben
        param = room.LookupParameter(fflaeche_param)
        if param:
            param.Set(sumStr)
        else:
            #print("Raum {room.Id} hat keinen Parameter '{fflaeche_param}'")
            print("Raum {} hat keinen Parameter {}".format(room.Id,fflaeche_param))


print("FensterflaechePerRoom run in: ")
stopwatch.Stop()
timespan = stopwatch.Elapsed

print (timespan)
