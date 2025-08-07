import revitron
from revitron import _
from rpw.ui.forms import FlexForm, TextBox, Button, Label, Separator, ComboBox
from collections import defaultdict
from pyrevit import script
import System.Windows
from pyrevit import forms

def alertReopen():
    forms.alert(
        'Note that changes won\'t take effect until the current file is closed and reopened again.'
    )

def openFile(sender, e):
    sqliteFile = forms.save_file(
        file_ext='sqlite',
        default_name='{}.sqlite'.format(revitron.DOC.Title),
        unc_paths=False
    )
    if sqliteFile:
        file_textbox.Text = sqliteFile
        
components = []
file_textbox = TextBox

if not revitron.Document().isFamily():

    config = revitron.DocumentConfigStorage().get('revitron.history', defaultdict())
                           
    components.append(Label('Forge Project ID'))   
    components.append(TextBox('forgeprojectid', Text=config.get('forgeprojectid')))
    
    components.append(Label('Item URN'))   
    components.append(TextBox('itemurn', Text=config.get('itemurn')))
    
    components.append(Label('File'))  
    file_textbox = TextBox('file', Text=config.get('file'))
    components.append(file_textbox)
    
    components.append(Button('Choose file', on_click=openFile))

    components.append(Label(''))
    components.append(Button('Save'))

    form = FlexForm('History Settings', components)
    form.show()

    if form.values:
        print  (form.values)
        alertReopen()
        revitron.DocumentConfigStorage().set('revitron.history', form.values)
