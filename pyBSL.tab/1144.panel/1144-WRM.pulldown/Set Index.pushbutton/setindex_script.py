#-*- coding: UTF-8 -*-

from pyrevit import forms
from pyrevit import revit
import revitron
from revitron import _
from pyrevit import script
import os


doc = __revit__.ActiveUIDocument.Document

# define a list of all parameters
params_to_check = [
    "LPH3-1-Index",
    "LPH3-2-Index",
    "LPH3-3-Index",
    "LPH3-4-Index",
    "LPH3-5-Index",
    "LPH3-6-Index",
    "LPH3-7-Index",
    "LPH3-8-Index",
    "LPH3-9-Index",
    "LPH3-10-Index",
    "LPH3-11-Index",
    "LPH3-12-Index",
    "LPH3-13-Index",
    "LPH3-14-Index"
]

sheets = []

sheet = revitron.ACTIVE_VIEW

if _(sheet).getClassName() == 'ViewSheet':
	sheets.append(sheet)

if __shiftclick__:
	# select sheets
	sheets = forms.select_sheets()




with revit.Transaction("Set Parameter"):

	for sheet in sheets:
		# get parameter
		filename_param = sheet.LookupParameter('Dateiname')
		last_value = None
		# run throug the list and get the last not empty value
		for param_name in params_to_check:
			param = sheet.LookupParameter(param_name)
			if param and param.AsString():  # check parameter and value exists
				last_value = param.AsString()
				
		status_param = sheet.LookupParameter('Status')
				
		print('{} Index is: {} Status is: {}'.format(sheet.Name,last_value,status_param.AsString()))

		if last_value and filename_param:
			index_param = sheet.LookupParameter('Index')
			
			
			#set the big index the the last index
			if index_param:
				index_param.Set(last_value) 

			# fix dateiname
			if len(filename_param.AsString()) >= 33:
				modified_filename = filename_param.AsString()[:32] + last_value + filename_param.AsString()[33:]
				
				modified_filename = modified_filename[:34] + status_param.AsString() + modified_filename[35:]
				
				filename_param.Set(modified_filename) 

