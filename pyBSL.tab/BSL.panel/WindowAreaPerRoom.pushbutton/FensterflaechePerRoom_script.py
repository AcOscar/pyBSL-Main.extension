from Autodesk.Revit.DB import FilteredElementCollector as Fec
from Autodesk.Revit.DB import BuiltInCategory
from System.Diagnostics import Stopwatch
from rpw import doc
from rpw.db import Transaction as rpw_Transaction
from pyrevit import script


from Autodesk.Revit.DB import UnitTypeId, UnitUtils

stopwatch = Stopwatch()
stopwatch.Start()
output = script.get_output()
logger = script.get_logger() # Initialisiert den Logger

exclude_param = "Fensterflaeche_Exklusion"
glazingarea_param = "Glasflaeche"
instW = "Width"
fflaeche_param = "Fensterflaeche_Tag"

ToRoom = []
filtered_windows = []
w2calc = []
fromRoomId = []
fromRoom = []
wRoomSet = []

def main():

    windows = Fec(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()

    if not windows:
        print ("No windows in the project.")
        return

    # Explicitly select the last phase
    phases = list(doc.Phases)
    last_phase = phases[-1]

    for window in windows:
        try:
            if not window.LookupParameter(exclude_param).AsInteger():
                filtered_windows.append(window)
        except:
            continue  
            #check.append(window)

    if not filtered_windows:
        print ("No windows to count.")
        return

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
        print ("o rooms found for the windows.")
        return

    if not w2calc:
        print ("No windows to calculate.")
        return

    """Sorting and grouping"""
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
                        glasflaeche = round(UnitUtils.ConvertFromInternalUnits(glasflaeche_param.AsDouble(), UnitTypeId.SquareMeters), 2) # Konvertieren Squarefeet in Quadratmeter nötig
                        roomSum += glasflaeche
                    else:
                        logger.info("Fenster {} hat keine gültige {}".format(output.linkify(window.Id), glazingarea_param))
                except Exception as e:
                    logger.error("Fehler bei Fenster {}: {}".format(window.Id,e))

    
            # Ergebnis in Textform formatieren
            sumStr = str(round(roomSum, 2)) + " m²"

            # Wert in den Raumparameter schreiben
            param = room.LookupParameter(fflaeche_param)
            if param:
                param.Set(sumStr)
            else:
                logger.info("Raum {} hat keinen Parameter {}".format(room.Id,fflaeche_param))

    stopwatch.Stop()
    timespan = stopwatch.Elapsed
    logger.info("FensterflaechePerRoom run in:  {} ".format(timespan))


if __name__ == "__main__":
    """Run main and catch exceptions to print them in the output window."""
    try:
        main()
    except Exception as e:
        import traceback
        logger.error("ERROR:", e)
        logger.error(traceback.format_exc())