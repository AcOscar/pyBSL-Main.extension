# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, ElementIntersectsSolidFilter, Transaction
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import Solid 
from Autodesk.Revit.DB import Options,GeometryInstance
from pyrevit import script
from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

from Autodesk.Revit.UI import UIApplication

from System.Diagnostics import Stopwatch

__title__ = 'Funktionseinheit'

__doc__ = "Writes the value of the parameter Funktionseinheit \n"\
          "of a Generic Model \n"\
          "into the Funktionseinheit paramter of\n"\
          "its enclosing geometry."

#set to True for more detailed messages
DEBUG = False
if __shiftclick__:
    DEBUG = True

          
          
# Kategorien, die überprüft werden sollen
categories_to_check = [
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_StructuralFraming,  # Geschossdecken
    # Weitere Kategorien je nach Bedarf hinzufügen
]
stopwatch = Stopwatch()
stopwatch.Start()
idx = 0

output = script.get_output()
       
# Zugriff auf das aktuelle Revit-Dokument und die Anwendung
uiapp = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

def get_solids(element):
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    geometry = element.get_Geometry(opt)
    
    solids = []

    def process_geometry(geom):
        for geom_obj in geom:
            if isinstance(geom_obj, Solid):
                if geom_obj.Volume > 0:
                    solids.append(geom_obj)
            elif isinstance(geom_obj, GeometryInstance):
                instance_geometry = geom_obj.GetInstanceGeometry()
                process_geometry(instance_geometry)
    
    process_geometry(geometry)
    
    return solids

def owned_by_other_user(elem):
#from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils

    # Checkout Status of the element
    checkout_status = WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
    if checkout_status == CheckoutStatus.OwnedByOtherUser:
        return True
    else:
        return False

# Workset-Name, nach dem gefiltert werden soll
workset_name = "Funktionseinheiten"

# Alle Generic Models sammeln
generic_models = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

# Dictionary für die Generic Models im gewünschten Workset
models_dict = {}

# Durch alle Generic Models iterieren
for gm in generic_models:
    # Workset-ID des Elements erhalten
    workset_id = gm.WorksetId
    # Workset-Element anhand der ID erhalten
    workset = doc.GetWorksetTable().GetWorkset(workset_id)
    # Überprüfen, ob der Name des Worksets mit dem gewünschten Workset übereinstimmt
    if workset.Name == workset_name:
        funktionseinheit = gm.LookupParameter('Funktionseinheit').AsString()
        solids = get_solids(gm)
        #print solids
        models_dict[gm.Id] = {'funktionseinheit': funktionseinheit, 'solids': solids}

print("Anzahl der Generic Models im Workset '{}': {}".format(workset_name, len(models_dict)))
# Ergebnis ausgeben
for gm_id, data in models_dict.items():
    #print("Element ID: {} - Funktionseinheit: {}".format(gm_id, data['funktionseinheit']))
    print output.linkify(gm_id) + " " + data['funktionseinheit'] 

EleNums = 0    
# Hauptlogik
with Transaction(doc, "Select Elements in Solid") as t:
    t.Start()
    #example_solid = create_example_solid()

    for gm_id, data in models_dict.items():
        funktionseinheit = data['funktionseinheit']
        solids = data['solids']
        
        for solid in solids:
            #print (funktionseinheit)
            print(" Solid: " + funktionseinheit) 

            # Verwendung des ElementIntersectsSolidFilter
            intersect_filter = ElementIntersectsSolidFilter(solid)

            elements_in_solid = []
            for category in categories_to_check:
                collector = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().WherePasses(intersect_filter)
                elements_in_solid.extend(collector.ToElements())
                
            idx = EleNums
            
            EleNums += elements_in_solid.Count 

            for elem in elements_in_solid:
                #check if someone else has the element, otherwise an error will thrown
                #TODO maybe this throw an error in a non worksharing model
                if owned_by_other_user(elem):
                    print(" Element: " + output.linkify(elem.Id) + " -> OwnedByOtherUser") 
                    continue
                
                if elem:
                    elem.LookupParameter('Funktionseinheit').Set(funktionseinheit)
                    
                    print ("  " + output.linkify(elem.Id))

                    idx+=1     
                    output.update_progress(idx, EleNums)
                    t.Commit()

    #t.Commit()
    output.reset_progress()
    
print(50 * "=")
stopwatch.Stop()
timespan = stopwatch.Elapsed
print("Run in: {}".format(timespan))
print ("Done")