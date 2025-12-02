'''
Copyright (c) 2014-2016 Ehsan Iran-Nejad
Python scripts for Autodesk Revit

This file is part of pyRevit repository at https://github.com/eirannejad/pyRevit

pyRevit is a free set of scripts for Autodesk Revit: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3, as published by
the Free Software Foundation.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

See this link for a copy of the GNU General Public License protecting this package.
https://github.com/eirannejad/pyRevit/blob/master/LICENSE
'''

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, LinePatternElement

doc = __revit__.ActiveUIDocument.Document

line_patterns = FilteredElementCollector(doc).OfClass(LinePatternElement).ToElements()

t = Transaction(doc, 'Remove IMPORT Patterns')
t.Start()

for line_pattern in line_patterns:
    if line_pattern.Name.lower().startswith('import'):
        print('\nIMPORTED LINETYPE FOUND:\n{0}'.format(line_pattern.Name))
        doc.Delete(line_pattern.Id)
        print('--- DELETED ---')

t.Commit()
