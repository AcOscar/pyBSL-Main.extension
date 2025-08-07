#pylint: disable=invalid-name,import-error,superfluous-parens
from collections import Counter

from pyrevit import revit, DB
from pyrevit import script
import sys

__doc__ = 'Set the Parameter "TRH" of a door the the enclosing scopebox'\
          'The name of the scopebox should started with TRH_ '\
          'The value for the TRH paramer gehts two characters followed by TRH'\
          'example: a door is inside of a scopbox with name "TRH_AB", the value of the TRH paramer will be AB'\
          'if a door is inside of a scopbox with name "TRH_AB-part2", the value of the TRH paramer will be also AB'

output = script.get_output()

#Names of the Scopboxes that are used 
scopeboxesnameprefix = 'TRH_'

#parameter to write
parametername = 'TRH'


if revit.active_view:

    scopeboxes = DB.FilteredElementCollector(revit.doc)\
                 .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)

    if scopeboxes.GetElementCount() == 0:
        print('There are no scopeboxes.')
        sys.exit()

    #roomids = DB.FilteredElementCollector(revit.doc)\
    #          .OfCategory(DB.BuiltInCategory.OST_Rooms)\
    #          .WhereElementIsNotElementType()\
    #          .ToElementIds()
    doorids = DB.FilteredElementCollector(revit.doc)\
              .OfCategory(DB.BuiltInCategory.OST_Doors)\
              .WhereElementIsNotElementType()\
              .ToElementIds()

    if doorids.Count == 0:
        print('There are no doors.')
        sys.exit()

    for sc in scopeboxes:
        #if sc.Name in namedscopeboxes:
        if sc.Name.startswith(scopeboxesnameprefix):
            print('\tScopebox: {}'.format(sc.Name.ljust(30)))

            bb = sc.get_BoundingBox(revit.active_view)
            outline = DB.Outline(bb.Min, bb.Max)
            filter = DB.BoundingBoxIsInsideFilter(outline)
            doorcollector = DB.FilteredElementCollector(revit.doc, doorids).WherePasses(filter)

            with revit.Transaction("Set Parameter by Scopebox"):

                for dr in doorcollector:
                        dparam = dr.LookupParameter(parametername)

                        if dparam.StorageType == DB.StorageType.String:
                            #dparam.Set(sc.Name )
                            parametervalue = sc.Name[4:5]
                            #parametervalue += '_'
                            #parametervalue += dr.Level.Name.split('_')[1]
                            
                            dparam.Set(parametervalue)

    #                    dparam = rm.LookupParameter('Comments')
    #                    if dparam.StorageType == DB.StorageType.String:
    #                        dparam.Set('')
    #                print(rm.Level.Name.split('_')[1])

print('Done')
