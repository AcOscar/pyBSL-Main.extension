# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import script
from Autodesk.Revit.DB import *
from System.Diagnostics import Stopwatch

__title__ = 'WallTag Distance'
__doc__ = 'Determines the shortest distance between a wall tag and the referenced wall in the active view.'

output = script.get_output()
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveView

collector = FilteredElementCollector(doc, view.Id)\
            .OfClass(IndependentTag)\
            .ToElements()

print("Number of wall tags in view: {}".format(len(collector)))

transGroup = TransactionGroup(doc, "WallTag Distance")
transGroup.Start()
results = []
for tag in collector:
    prop = tag.GetTaggedReferences()[0]
    if prop is None:
        continue

    tagged_elem = doc.GetElement(prop)
    if tagged_elem is None or not isinstance(tagged_elem, Wall):
        continue

    tag_point = tag.TagHeadPosition
    if tag_point is None:
        continue

    wall_curve = tagged_elem.Location.Curve
    if wall_curve is None:
        continue

    ref_result = wall_curve.Project(tag_point)
    if ref_result is None:
        continue

    closest_point = ref_result.XYZPoint
    distance = tag_point.DistanceTo(closest_point)
    distance_mm = round(distance * 304.8, 2)

    #taglink = output.linkify(tag.Id)
    #print("WallTag ID " + taglink + " → Distance to wall: " + str(distance_mm) + " mm")
    results.append((distance_mm, tag.Id))  # Speichere mm + ID für spätere Ausgabe

    idx += 1
    output.update_progress(idx, len(collector))

transGroup.Assimilate()
output.reset_progress()


# Ergebnisse sortieren und ausgeben
for distance_mm, tag_id in sorted(results):
    taglink = output.linkify(tag_id)
    print("WallTag ID " + taglink + " → Distance to wall: " + str(distance_mm) + " mm")
    
    
print('Done')
stopwatch.Stop()
print("Time: {}".format(stopwatch.Elapsed))
