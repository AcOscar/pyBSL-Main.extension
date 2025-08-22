# -*- coding: utf-8 -*-
#pyrevit python3: true

import clr
import os
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore") 
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.Windows.Interop import WindowInteropHelper
from System.IO import FileStream, FileMode
import System

def set_owner_to_revit(window):
    """Setzt Revit als Owner-Fenster"""
    try:
        try:
            uiapp = __revit__.CurrentUIApplication
        except NameError:
            import Autodesk.Revit.UI as ui
            uiapp = ui.UIApplication.GetActiveUIApplication()
        
        hwnd = uiapp.MainWindowHandle
        helper = WindowInteropHelper(window)
        helper.Owner = hwnd
    except Exception:
        pass

class XamlWindow:
    def __init__(self, xaml_path):
        """
        Lädt ein WPF-Fenster aus einer XAML-Datei
        
        Args:
            xaml_path: Pfad zur XAML-Datei
        """
        self.xaml_path = xaml_path
        self.window = None
        self._load_xaml()
        
    def _load_xaml(self):
        """Lädt die XAML-Datei und erstellt das Window"""
        try:
            # XAML-Datei laden
            with open(self.xaml_path, 'r', encoding='utf-8') as f:
                xaml_content = f.read()
            
            # Alternative Methode mit XamlReader.Parse (für String-Content)
            self.window = XamlReader.Parse(xaml_content)
            
            # Oder mit FileStream (für direkte Datei-Ladung):
            # stream = FileStream(self.xaml_path, FileMode.Open)
            # self.window = XamlReader.Load(stream)
            # stream.Close()
            
        except Exception as e:
            raise Exception(f"Fehler beim Laden der XAML-Datei: {e}")
    
    def find_element(self, name):
        """Findet ein benanntes Element in der XAML"""
        return self.window.FindName(name)
    
    def show_dialog(self):
        """Zeigt das Fenster als Dialog an"""
        set_owner_to_revit(self.window)
        return self.window.ShowDialog()
    
    def show(self):
        """Zeigt das Fenster an (nicht-modal)"""
        set_owner_to_revit(self.window)
        self.window.Show()
    
    def close(self):
        """Schließt das Fenster"""
        if self.window:
            self.window.Close()

# Beispiel-Klasse für eine spezifische XAML-Form
class SelectFromListXamlWindow(XamlWindow):
    def __init__(self, xaml_path, items, title="Auswahl", multiselect=False, display_fn=None):
        super().__init__(xaml_path)
        
        self._all_items = list(items)
        self._display_fn = display_fn or (lambda x: str(x))
        self.selected_items = None
        
        # UI-Elemente finden (Namen müssen in XAML definiert sein)
        self.search_box = self.find_element("SearchTextBox")
        self.listbox = self.find_element("ItemsListBox") 
        self.ok_button = self.find_element("OkButton")
        self.cancel_button = self.find_element("CancelButton")
        
        # Eigenschaften setzen
        if self.window:
            self.window.Title = title
        
        if self.listbox:
            self.listbox.SelectionMode = SelectionMode.Multiple if multiselect else SelectionMode.Single
        
        # Event-Handler verbinden
        self._setup_events()
        
        # Liste initial befüllen
        self._refresh_list("")
    
    def _setup_events(self):
        """Verbindet Event-Handler"""
        if self.search_box:
            self.search_box.KeyUp += self._on_search
        
        if self.ok_button:
            self.ok_button.Click += self._on_ok
            
        if self.cancel_button:
            self.cancel_button.Click += self._on_cancel
    
    def _refresh_list(self, query):
        """Aktualisiert die Liste basierend auf Suchkriterien"""
        if not self.listbox:
            return
            
        self.listbox.Items.Clear()
        q = (query or "").strip().lower()
        
        for obj in self._all_items:
            disp = self._display_fn(obj)
            if not q or q in disp.lower():
                # Tupel aus Anzeige-String und Original-Objekt
                self.listbox.Items.Add((disp, obj))
        
        # Ersten Eintrag auswählen wenn Single-Select
        if (self.listbox.SelectionMode == SelectionMode.Single and 
            self.listbox.Items.Count > 0):
            self.listbox.SelectedIndex = 0
    
    def _on_search(self, sender, args):
        """Handler für Suche"""
        if self.search_box:
            self._refresh_list(self.search_box.Text)
    
    def _on_ok(self, sender, args):
        """Handler für OK-Button"""
        if self.listbox and self.listbox.SelectedItems.Count > 0:
            # Extrahiere die Original-Objekte aus den Tupeln
            self.selected_items = [item[1] for item in self.listbox.SelectedItems]
            self.window.DialogResult = True
            self.close()
    
    def _on_cancel(self, sender, args):
        """Handler für Cancel-Button"""
        self.selected_items = None
        self.window.DialogResult = False
        self.close()




def select_from_list_xaml(xaml_path, items, title="Auswahl", multiselect=False, name_attr=None, map_fn=None):
    """
    Zeigt eine Auswahlliste basierend auf XAML-Datei an.
    
    Args:
        xaml_path: Pfad zur XAML-Datei
        items: Liste der auswählbaren Objekte
        title: Fenstertitel
        multiselect: True für Mehrfachauswahl
        name_attr: Attributname für die Anzeige
        map_fn: Funktion zur Anzeige-Transformation
    
    Returns:
        Ausgewählte(s) Objekt(e) oder None/[]
    """
    if not items:
        return [] if multiselect else None
    
    # Display-Funktion bestimmen
    if map_fn:
        display_fn = map_fn
    elif name_attr:
        display_fn = lambda x: getattr(x, name_attr, str(x))
    else:
        display_fn = lambda x: str(x)
    
    # Fenster erstellen und anzeigen
    win = SelectFromListXamlWindow(xaml_path, items, title=title, 
                                   multiselect=multiselect, display_fn=display_fn)
    
    result = win.show_dialog()
    
    if not result or win.selected_items is None:
        return [] if multiselect else None
    
    return win.selected_items if multiselect else win.selected_items[0]

# Einfache XAML-Loader Funktion (für beliebige XAML-Dateien)
def load_xaml_window(xaml_path):
    """
    Lädt eine beliebige XAML-Datei als Window
    
    Args:
        xaml_path: Pfad zur XAML-Datei
        
    Returns:
        XamlWindow-Instanz
    """
    return XamlWindow(xaml_path)