#! python3
# -*- coding: utf-8 -*-
__title__ = "Mein Tool"
__doc__ = "WPF ohne pyrevit.forms"

from pyrevit import script
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Xaml')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox
from System.Windows.Markup import XamlReader
from System.IO import StreamReader
import os


class SyncSheetRevisionsWindow(object):
    def __init__(self):
        self._load_xaml()
        self._get_controls()
        self._wire_events()

        # interner Zustand
        self.regex_enabled = False

    # ---------------------------
    # XAML laden
    # ---------------------------
    def _load_xaml(self):
        script_dir = os.path.dirname(__file__)
        xaml_file = os.path.join(script_dir, 'test.xaml')

        stream = StreamReader(xaml_file)
        self.window = XamlReader.Load(stream.BaseStream)
        stream.Close()

    # ---------------------------
    # Controls holen
    # ---------------------------
    def _get_controls(self):
        w = self.window
        self.header_grid     = w.FindName('header_grid')
        self.input_box       = w.FindName('input_textbox')
        self.regexToggle_btn = w.FindName('regexToggle_b')
        self.ok_btn          = w.FindName('ok_button')
        self.cancel_btn      = w.FindName('cancel_button')
        self.results_listbox = w.FindName('results_listbox')

    # ---------------------------
    # Events verbinden
    # ---------------------------
    def _wire_events(self):
        if self.header_grid:
            self.header_grid.MouseDown += self.header_drag

        if self.regexToggle_btn:
            self.regexToggle_btn.Click += self.toggle_regex

        if self.ok_btn:
            self.ok_btn.Click += self.on_ok_click

        if self.cancel_btn:
            self.cancel_btn.Click += self.on_cancel_click

    # ---------------------------
    # Eventhandler
    # ---------------------------
    def header_drag(self, sender, args):
        # Fenster verschieben
        try:
            self.window.DragMove()
        except:
            pass

    def toggle_regex(self, sender, args):
        # Toggle-Status auslesen
        try:
            self.regex_enabled = bool(self.regexToggle_btn.IsChecked)
        except:
            self.regex_enabled = not self.regex_enabled

        state = "aktiv" if self.regex_enabled else "inaktiv"
        # Debug / Info – später durch echte Logik ersetzen
        MessageBox.Show(
            "Regex ist jetzt {}".format(state),
            "Regex"
        )

    def on_ok_click(self, sender, args):
        text = self.input_box.Text if self.input_box else ""
        msg = u"Du hast eingegeben: {}\nRegex: {}".format(
            text,
            "an" if self.regex_enabled else "aus"
        )
        MessageBox.Show(msg, "Info")
        self.window.Close()

    def on_cancel_click(self, sender, args):
        self.window.Close()

    # ---------------------------
    # Public API
    # ---------------------------
    def show_dialog(self):
        self.window.ShowDialog()


# Entry Point für pyRevit
if __name__ == '__main__':
    ui = SyncSheetRevisionsWindow()
    ui.show_dialog()