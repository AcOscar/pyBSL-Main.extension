#-*- coding: UTF-8 -*-

import re

from pyrevit import coreutils
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script


logger = script.get_logger()


class BatchSheetMakerWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self._sheet_dict = {}
        self._titleblock_id = None
        self.sheets_tb.Focus()

    def _process_sheet_code(self):
        for sheet_code in str(self.sheets_tb.Text).split('\n'):
            # Leere oder nur whitespace-Zeilen überspringen
            if coreutils.is_blank(sheet_code):
                continue

            # Zeile bereinigen: mehrere Tabs reduzieren, Zeilenumbrüche entfernen
            code = re.sub("\t+", "\t", sheet_code.strip())

            # Prüfen, ob ein Tab-Trenner vorhanden ist
            if '\t' not in code:
                # Nur ein Wert vorhanden: als Nummer und Name verwenden
                num = code
                name = code
            else:
                # Tab-getrennte Nummer und Name
                num, name = code.split('\t', 1)

            try:
                # Extrahiere alle Nummern (inkl. Bereiche)
                for single_num in coreutils.extract_range(num):
                    self._sheet_dict[single_num] = name
            except Exception as err:
                logger.error('Fehler beim Verarbeiten des Sheets {}: {}'.format (code,err))
                return False
        return True

    def _ask_for_titleblock(self):
        tblock = forms.select_titleblocks(doc=revit.doc)
        if tblock is not None:
            self._titleblock_id = tblock
            return True

        return False

    @staticmethod
    def _create_placeholder(sheet_num, sheet_name):
        with DB.Transaction(revit.doc, 'Create Placeholder') as t:
            try:
                t.Start()
                new_phsheet = DB.ViewSheet.CreatePlaceholder(revit.doc)
                new_phsheet.Name = sheet_name
                new_phsheet.SheetNumber = sheet_num
                t.Commit()
            except Exception as create_err:
                t.RollBack()
                logger.error('Error creating placeholder sheet {}:{} | {}'
                             .format(sheet_num, sheet_name, create_err))

    def _create_sheet(self, sheet_num, sheet_name):
        with DB.Transaction(revit.doc, 'Create Sheet') as t:
            try:
                t.Start()
                new_phsheet = DB.ViewSheet.Create(revit.doc,
                                                  self._titleblock_id)
                new_phsheet.Name = sheet_name
                new_phsheet.SheetNumber = sheet_num
                if self.idparam_cb.IsChecked:
                    self._prmset(new_phsheet, "Plan_ID", sheet_num)
                t.Commit()
                
            except Exception as create_err:
                t.RollBack()
                logger.error('Error creating sheet sheet {}:{} | {}'
                             .format(sheet_num, sheet_name, create_err))
                             
    def _prmset(self, element, param_name, value):
        param = element.LookupParameter(param_name)
        if param != None and not param.IsReadOnly:
            param.Set(value)
            
    def create_sheets(self, sender, args):
        self.Close()

        if self._process_sheet_code():
            if self.sheet_cb.IsChecked:
                create_func = self._create_sheet
                transaction_msg = 'Batch Create Sheets'
                if not self._ask_for_titleblock():
                    script.exit()
            else:
                create_func = self._create_placeholder
                transaction_msg = 'Batch Create Placeholders'

            with DB.TransactionGroup(revit.doc, transaction_msg) as tg:
                tg.Start()
                for sheet_num, sheet_name in self._sheet_dict.items():
                    logger.debug('Creating Sheet: {}:{}'.format(sheet_num,
                                                                sheet_name))
                    create_func(sheet_num, sheet_name)
                tg.Assimilate()
        else:
            logger.error('Aborted with errors.')


BatchSheetMakerWindow('BatchSheetMakerWindow.xaml').ShowDialog()
