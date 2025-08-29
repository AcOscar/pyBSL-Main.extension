#! python3
# -*- coding: utf-8 -*-
"""PyRevit Script: Excel Named Range Reader
Reading BTK list from excel
"""

__title__ = "Excel BTK Reader"
__author__ = "PyRevit User"

import sys
import os
import json

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.DB import StorageType, ViewType, View

from urllib.parse import urlparse, unquote

# sys.path.append(os.path.join(os.path.dirname(__file__), "..", "lib"))

# from wpf_select import select_from_list_xaml

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
PARAM_NAME = "ExcelReader"

# Simplified version without pyrevit.script import
class SimpleOutput:
    def print_md(self, text):
        print(text.replace("#", "").replace("*", "").replace("`", ""))
    
    def __call__(self):
        return self
    
# CPython-Module for Excel
try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl is not installed.")
    sys.exit(1)

# ---------------------------
# Utils
# ---------------------------
def get_project_info(doc):
    return FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_ProjectInformation) \
        .FirstElement()

def get_onedrive_root(tenant, team, folder_list):
    """
    Ermittelt den lokalen OneDrive-Root.
    """ 
    user_profile = os.environ.get("USERPROFILE")
    if not user_profile:
        print("ERROR: USERPROFILE environment variable not found")
        return None

    ondr = ["","OneDrive - "]
    for dr in ondr:
        for folder in folder_list:
            rel = folder.lstrip("\\/")
            if dr == "":
                 rel = f"{team} - {rel}"
                 test_root = os.path.join(user_profile,f"{tenant}",rel)
            else:
                rel = f"{rel} - {team}"
                test_root = os.path.join(user_profile,f"{dr}{tenant}",rel)

            if  os.path.isdir(test_root):

                return test_root

    return None

def read_json_from_project_param(doc, param_name):
    pi = get_project_info(doc)
    if not pi:
        raise Exception("Project Information Element not found.")

    p = pi.LookupParameter(param_name)
    if not p:
        raise Exception("Parameter '{}' not found.".format(param_name))

    raw = p.AsString() or p.AsValueString()
    if not raw or not raw.strip():
        raise Exception("Parameter '{}' is empty.".format(param_name))

    try:
        return json.loads(raw)
    except Exception as e:
        preview = (raw[:300] + ("..." if len(raw) > 300 else "")).replace("\n", " ")
        raise Exception("Invalid JSON in '{}': {}.\nPreview: {}".format(param_name, e, preview))

def get_onedrive_config(cfg):
    """
    Get OneDrive configuration from config, with default values
    
    Args:
        cfg (dict): Configuration dictionary
        
    Returns:
        tuple: (tenant, team, local_paths)
    """
    onedrive_config = cfg.get("OneDriveConfig", {})
    
    # Default values (current hardcoded values)
    # default_team = ""
    # default_library = ""
    # team_name = onedrive_config.get( "TeamsLocalPaths", "")

    # library_name = onedrive_config.get("LibraryName", default_library)
    Tenant = onedrive_config.get("Tenant","")
    Team = onedrive_config.get("Team","")
    LocalPaths = onedrive_config.get("LocalPaths",[])
    
    return Tenant,Team,LocalPaths

def ensure_list(value):
    """if value is a Dict  -> [value], if None -> [], else value return."""
    if value is None:
        return []
    if isinstance(value, dict):
        return [value]
    if isinstance(value, list):
        return value
    raise Exception("Unexpected type for ParameterSync: {}.".format(type(value)))

def validate_parameter_sync_item(node, path):
    required = ["Order", "FileName", "DataName", "KeyName"]
    for k in required:
        if k not in node:
            raise Exception("{}: mandatory field '{}' missing.".format(path, k))

def validate_config(cfg):
    if not isinstance(cfg, dict):
        raise Exception("Top-level JSON must be an object.")

    ps_list = ensure_list(cfg.get("ParameterSync"))
    for i, item in enumerate(ps_list):
        if not isinstance(item, dict):
            raise Exception("ParameterSync[{}] is not an object.".format(i))
        validate_parameter_sync_item(item, "ParameterSync[{}]".format(i))

    return {
        "ParameterSync": ps_list
    }

def parse_builtin_categories(filter_string):
    """
    Parse comma-separated BuiltInCategory names and return actual category objects
    
    Args:
        filter_string (str): Comma-separated category names like "OST_Walls,OST_Doors"
        
    Returns:
        list: List of BuiltInCategory objects
    """
    categories = []
    category_names = [name.strip() for name in filter_string.split(",")]
    
    for cat_name in category_names:
        try:
            # Get the BuiltInCategory enum value
            builtin_cat = getattr(BuiltInCategory, cat_name)
            categories.append(builtin_cat)
            print("  Added category filter: {}".format(cat_name))
        except AttributeError:
            print("  WARNING: Unknown category '{}' - skipping".format(cat_name))
    
    return categories

def get_selected_elements(uidoc, categories):
    """
    Get currently selected elements that match the specified categories
    
    Args:
        uidoc: UI Document
        categories (list): List of BuiltInCategory objects
        
    Returns:
        list: List of matching selected elements
    """
    selection = uidoc.Selection
    selected_element_ids = selection.GetElementIds()
    
    if not selected_element_ids:
        return []
    
    matching_elements = []
    
    for element_id in selected_element_ids:
        element = uidoc.Document.GetElement(element_id)
        if element and hasattr(element, 'Category') and element.Category:
            element_category = element.Category.Id.IntegerValue
            
            # Check if element category matches any of our target categories
            for target_category in categories:
                if element_category == target_category.value__:
                    matching_elements.append(element)
                    break
    
    return matching_elements

def collect_elements_by_categories_in_views(doc, categories, views=None):
    """
    Collect all elements from specified categories, optionally filtered by views
    
    Args:
        doc: Revit document
        categories (list): List of BuiltInCategory objects
        views (list): Optional list of views to filter by
        
    Returns:
        list: List of collected elements
    """
    all_elements = []
    
    for category in categories:
        try:
            collector = FilteredElementCollector(doc) \
                .OfCategory(category) \
                .WhereElementIsNotElementType()
            
            # If views are specified, filter by views
            if views:
                view_elements = set()
                for view in views:
                    try:
                        view_collector = FilteredElementCollector(doc, view.Id) \
                            .OfCategory(category) \
                            .WhereElementIsNotElementType()
                        view_elements.update(view_collector.ToElementIds())
                    except Exception as e:
                        print("  WARNING: Error collecting from view '{}': {}".format(view.Name, str(e)))
                
                # Filter main collector by view elements
                if view_elements:
                    elements = [doc.GetElement(eid) for eid in view_elements if doc.GetElement(eid)]
                else:
                    elements = []
            else:
                elements = collector.ToElements()
            
            element_count = len(list(elements))
            print("  Found {} elements in category {}".format(element_count, category))
            all_elements.extend(elements)
            
        except Exception as e:
            print("  ERROR collecting from category {}: {}".format(category, str(e)))
    
    return all_elements

def get_all_views(doc):
    """
    Get all views from the document
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of View objects
    """
    return FilteredElementCollector(doc) \
        .OfClass(View) \
        .WhereElementIsNotElementType() \
        .ToElements()

def select_views_for_processing(doc, uidoc):
    """
    Let user select views for processing using pyrevit.forms
    
    Args:
        doc: Revit document
        uidoc: UI Document
        
    Returns:
        list: List of selected View objects
    """
    PYREVIT_AVAILABLE = False
    if not PYREVIT_AVAILABLE:
        print("PyRevit forms not available - using active view only")
        return [uidoc.ActiveView] if uidoc.ActiveView else []
    
    # Get all views
    all_views = get_all_views(doc)
    
    # Filter to meaningful views (not templates, schedules, etc.)
    selectable_views = []
    current_view = uidoc.ActiveView
    current_view_in_list = False
    
    for view in all_views:
        # Skip view templates and certain view types
        if (view.IsTemplate or 
            view.ViewType == ViewType.ProjectBrowser or
            view.ViewType == ViewType.SystemBrowser or
            view.ViewType == ViewType.Undefined):
            continue
            
        selectable_views.append(view)
        if current_view and view.Id == current_view.Id:
            current_view_in_list = True
    
    if not selectable_views:
        print("No selectable views found")
        return []
    
    # Create view selection list with names
    view_options = []
    preselected_indices = []
    
    for i, view in enumerate(selectable_views):
        view_name = "{} [{}]".format(view.Name, view.ViewType.ToString())
        view_options.append(view_name)
        
        # Preselect current view if it's in the list
        if current_view and view.Id == current_view.Id:
            preselected_indices.append(i)
    
    # Show selection dialog
    try:
        selected_indices = pyrevit.forms.SelectFromList.show(
            view_options,
            title="Select Views for Parameter Sync",
            width=500,
            height=400,
            multiselect=True,
            checked=preselected_indices
        )
        
        if not selected_indices:
            print("No views selected - operation cancelled")
            return []
        
        # Return selected views
        return [selectable_views[i] for i in selected_indices]
        
    except Exception as e:
        print("Error in view selection: {} - using active view".format(str(e)))
        return [current_view] if current_view else []

def collect_elements_by_categories(doc, categories):
    """
    Collect all elements from specified categories
    
    Args:
        doc: Revit document
        categories (list): List of BuiltInCategory objects
        
    Returns:
        list: List of collected elements
    """
    all_elements = []
    
    for category in categories:
        try:
            elements = FilteredElementCollector(doc) \
                .OfCategory(category) \
                .WhereElementIsNotElementType() \
                .ToElements()
            
            element_count = len(list(elements))
            print("  Found {} elements in category {}".format(element_count, category))
            all_elements.extend(elements)
            
        except Exception as e:
            print("  ERROR collecting from category {}: {}".format(category, str(e)))
    
    return all_elements

def read_excel_named_range(file_path, range_name):
    """
    Reads a named range from an Excel file
    
    Args:
        file_path (str): Path to the Excel file
        range_name (str): Name of the named range
        
    Returns:
        list: List of lists containing cell data
    """
    try:
        # open Excel file
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        
        # Check whether the specified range exists
        if range_name not in workbook.defined_names:
            raise ValueError(f"Named range '{range_name}' not found in Excel file")
        
        # Retrieve named range
        defined_name = workbook.defined_names[range_name]
        
        # Parse range definition
        range_def = str(defined_name.attr_text)
        
        # Extract worksheet and cell range
        if '!' in range_def:
            sheet_name, cell_range = range_def.split('!')
            sheet_name = sheet_name.replace("'", "")  # Remove quotation marks
            cell_range = cell_range.replace('$', '')  # Remove dollar sign
        else:
            # Fallback: use first worksheet
            sheet_name = workbook.sheetnames[0]
            cell_range = range_def.replace('$', '')
        
        # Select worksheet
        worksheet = workbook[sheet_name]
        
        # Read data from the range
        data = []
        for row in worksheet[cell_range]:
            row_data = []
            for cell in row:
                # Retrieve cell value (None becomes empty string)
                value = cell.value if cell.value is not None else ""
                row_data.append(str(value))
            data.append(row_data)
        
        workbook.close()
        return data
        
    except FileNotFoundError:
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    except Exception as e:
        raise Exception(f"Errors when reading the Excel file: {str(e)}")
    
def convert_excel_data_to_dict(data, key_column_name):
    """
    Convert Excel data to dictionary for easier lookup
    
    Args:
        data (list): List of lists from Excel (first row = headers)
        key_column_name (str): Name of the key column
        
    Returns:
        dict: Dictionary with key column values as keys
    """
    if not data or len(data) < 2:
        return {}
    
    headers = data[0]
    
    # Find key column index
    try:
        key_index = headers.index(key_column_name)
    except ValueError:
        raise Exception("Key column '{}' not found in headers: {}".format(key_column_name, headers))
    
    # Convert to dictionary
    lookup_dict = {}
    for row in data[1:]:  # Skip header row
        if len(row) > key_index and row[key_index]:  # Ensure key exists and is not empty
            key_value = str(row[key_index]).strip()
            if key_value:  # Only add non-empty keys
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row):
                        row_dict[header] = str(row[i]).strip() if row[i] else ""
                    else:
                        row_dict[header] = ""
                lookup_dict[key_value] = row_dict
    
    return lookup_dict

def get_parameter_value_safely(element, param_name):
    """
    Safely get parameter value from element
    
    Args:
        element: Revit element
        param_name (str): Parameter name
        
    Returns:
        str: Parameter value or empty string if not found
    """
    try:
        param = element.LookupParameter(param_name)
        if param:
            if param.StorageType == StorageType.String:
                return param.AsString() or ""
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                return str(param.AsDouble())
            elif param.StorageType == StorageType.ElementId:
                return str(param.AsElementId().IntegerValue)
    except:
        pass
    return ""

def set_parameter_value_safely(element, param_name, value, transaction):
    """
    Safely set parameter value on element
    
    Args:
        element: Revit element
        param_name (str): Parameter name
        value (str): Value to set
        transaction: Active transaction
        
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return False, "Parameter '{}' not found".format(param_name)
        
        if param.IsReadOnly:
            return False, "Parameter '{}' is read-only".format(param_name)
        
        # Convert value based on parameter storage type
        if param.StorageType == StorageType.String:
            param.Set(str(value))
        elif param.StorageType == StorageType.Integer:
            try:
                param.Set(int(float(value)))  # float first to handle decimal strings
            except ValueError:
                return False, "Cannot convert '{}' to integer for parameter '{}'".format(value, param_name)
        elif param.StorageType == StorageType.Double:
            try:
                param.Set(float(value))
            except ValueError:
                return False, "Cannot convert '{}' to double for parameter '{}'".format(value, param_name)
        else:
            return False, "Unsupported parameter type for '{}'".format(param_name)
        
        return True, "Success"
        
    except Exception as e:
        return False, "Error setting parameter '{}': {}".format(param_name, str(e))

def sync_parameters_to_elements(doc, elements, excel_data_dict, key_param_name, output):
    """
    Sync parameters from Excel data to Revit elements
    
    Args:
        doc: Revit document
        elements (list): List of Revit elements
        excel_data_dict (dict): Dictionary with Excel data
        key_param_name (str): Name of the key parameter
        output: Output object for printing
    """
    if not excel_data_dict:
        output.print_md("No Excel data available for synchronization")
        return
    
    # Statistics
    stats = {
        'processed': 0,
        'matched': 0,
        'updated': 0,
        'errors': 0,
        'warnings': 0
    }
    
    # Get available parameter names from Excel data (excluding the key column)
    sample_row = next(iter(excel_data_dict.values()))
    available_params = [k for k in sample_row.keys() if k != key_param_name]
    
    output.print_md("Available parameters from Excel: {}".format(", ".join(available_params)))
    output.print_md("Key parameter: {}".format(key_param_name))
    output.print_md("-" * 50)
    
    # Start transaction for parameter updates
    transaction = Transaction(doc, "Sync Parameters from Excel")
    transaction.Start()
    
    try:
        for element in elements:
            stats['processed'] += 1
            
            # Get the key parameter value from the element
            key_value = get_parameter_value_safely(element, key_param_name)
            
            if not key_value:
                print("  Element {} has no value for key parameter '{}'".format(
                    element.Id, key_param_name))
                stats['warnings'] += 1
                continue
            
            # Look up the element in Excel data
            if key_value not in excel_data_dict:
                print("  Key '{}' not found in Excel data".format(key_value))
                stats['warnings'] += 1
                continue
            
            stats['matched'] += 1
            excel_row = excel_data_dict[key_value]
            
            print("  Processing element {} with key '{}'".format(element.Id, key_value))
            
            # Update parameters for this element
            element_updated = False
            for param_name in available_params:
                if param_name == key_param_name:
                    continue  # Skip the key parameter itself
                
                excel_value = excel_row.get(param_name, "")
                if not excel_value:  # Skip empty values
                    continue
                
                success, message = set_parameter_value_safely(element, param_name, excel_value, transaction)
                
                if success:
                    print("    Set {}: '{}'".format(param_name, excel_value))
                    element_updated = True
                else:
                    print("    FAILED {}: {}".format(param_name, message))
                    stats['errors'] += 1
            
            if element_updated:
                stats['updated'] += 1
        
        # Commit transaction
        transaction.Commit()
        
    except Exception as e:
        transaction.RollBack()
        raise Exception("Transaction failed: {}".format(str(e)))
    
    # Print statistics
    output.print_md("-" * 50)
    output.print_md("SYNCHRONIZATION COMPLETE")
    output.print_md("Elements processed: {}".format(stats['processed']))
    output.print_md("Elements matched: {}".format(stats['matched']))
    output.print_md("Elements updated: {}".format(stats['updated']))
    output.print_md("Parameter errors: {}".format(stats['errors']))
    output.print_md("Warnings: {}".format(stats['warnings']))

    """
    Formats the data for output in the PyRevit window.
    
    Args:
        data (list): List of lists containing cell data
        
    Returns:
        str: Formatted output
    """
    if not data:
        return "No data found."
    
    # Calculate maximum column widths
    max_widths = []
    for row in data:
        for i, cell in enumerate(row):
            if i >= len(max_widths):
                max_widths.append(0)
            max_widths[i] = max(max_widths[i], len(str(cell)))
    
    # Create formatted output
    output_lines = []
    
    # Header separator line
    separator = "+" + "+".join(["-" * (width + 2) for width in max_widths]) + "+"
    output_lines.append(separator)
    
    # data rows
    for i, row in enumerate(data):
        formatted_row = "|"
        for j, cell in enumerate(row):
            width = max_widths[j] if j < len(max_widths) else 10
            formatted_row += f" {str(cell):<{width}} |"
        output_lines.append(formatted_row)
        
        # Separation line after the first line (header)
        if i == 0 and len(data) > 1:
            output_lines.append(separator)
    
    # Final dividing line
    output_lines.append(separator)
    
    return "\n".join(output_lines)

def sp_url_to_local_path(sp_url, onedrive_base):
    """
    Convert a SharePoint URL to a local OneDrive path

    Args:
        sp_url (str):
        onedrive_base (str): 
    Returns:
        str: Formatted output
    """
    parsed = urlparse(sp_url)
    parts = unquote(parsed.path).split("/Shared Documents/", 1)
    if len(parts) != 2:
        raise ValueError("Invalid SharePoint URL: {}".format(sp_url))
    relative_path = parts[1].replace("/", os.sep)
    return os.path.join(onedrive_base, relative_path)
    
def main():
    """
    Main function
    """
    
    output = SimpleOutput()

    # xaml_path = os.path.join(os.path.dirname(__file__), "SelectFromList.xaml")
   
    # fruits = ["Apfel", "Banane", "Mango", "Orange"]
    # choice = select_from_list_xaml(xaml_path,fruits, title="Frucht auswählen")
    # output.print_md("Auswahl: {}".format (choice))

    try:
        cfg_raw = read_json_from_project_param(doc, PARAM_NAME)
        cfg = validate_config(cfg_raw)
    except Exception as e:
        output.print_md("ERROR reading configuration: {}".format(str(e)))
        return

    # Get OneDrive configuration
    # teams_root = get_onedrive_config(cfg_raw)
    tenant, team, local_paths = get_onedrive_config(cfg_raw)

    output.print_md("OneDrive Configuration:")
    output.print_md("  Tenant: {}".format(tenant))
    output.print_md("  Team: {}".format(team))
    # output.print_md("  local_paths: {}".format(local_paths))
    
    onedrive_root = get_onedrive_root(tenant, team, local_paths)
    output.print_md("  Local Root: {}".format(onedrive_root))
    output.print_md("-" * 50)
    params = cfg["ParameterSync"]

    output.print_md("=== Excel Parameter Sync ({} Entries) ===".format(len(params)))
    
    # Sort numerically by order (order is string -> convert to int, fallback 0)
    def order_key(x):
        try:
            return int(x.get("Order", "0"))
        except:
            return 0

    for item in sorted(params, key=order_key):
        excel_filename = item.get("FileName")
        excel_rangename = item.get("DataName")
        excel_keyname = item.get("KeyName")
        filter_categories = item.get("BuiltInCategory", "")
        sync_type = item.get("Type", "Instance")

        output.print_md("-" * 50)
        output.print_md(" Order\t: {}".format(item.get("Order")))
        output.print_md("-  FileName\t: {}".format(excel_filename))
        output.print_md("-  DataName\t: {}".format(excel_rangename))
        output.print_md("-  KeyName \t: {}".format(excel_keyname))
        output.print_md("-  Category\t: {}".format(filter_categories))
        output.print_md("-  Type\t\t\t: {}".format(sync_type))
        output.print_md("-" * 50)

        # Parse categories from filter string
        if not filter_categories:
            output.print_md("  WARNING: No filter categories specified - skipping")
            continue
            
        categories = parse_builtin_categories(filter_categories)
        if not categories:
            output.print_md("  WARNING: No valid categories found - skipping")
            continue

        # Check for pre-selected elements first
        output.print_md("  Checking for selected elements...")
        selected_elements = get_selected_elements(uidoc, categories)
        
        if selected_elements:
            output.print_md("  Using {} pre-selected elements".format(len(selected_elements)))
            elements = selected_elements
            selected_views = None
        else:
            output.print_md("Nothing selected, please selct elements before")
            """
             output.print_md("  No relevant elements selected - proceeding with view selection")
            
            # Select views for processing
            # selected_views = select_views_for_processing(doc, uidoc)
            selected_views  = [uidoc.ActiveView]

            if not selected_views:
                output.print_md("  No views selected - skipping this sync item")
                continue
            
            output.print_md("  Selected {} views for processing".format(len(selected_views)))
            for view in selected_views:
                output.print_md("    - {}".format(view.Name))
            
            # Collect elements from selected views
            output.print_md("  Collecting elements from selected views...")
            elements = collect_elements_by_categories_in_views(doc, categories, selected_views) 
            """

        output.print_md("Total elements for processing: {}".format(len(elements)))

        if not elements:
            output.print_md("WARNING: No elements found for the specified categories")
            continue

        # convert SharePoint URL to local path
        try:
            excel_file_local_path = sp_url_to_local_path(excel_filename, onedrive_root)
        except Exception as e:
            output.print_md("ERROR converting SharePoint URL: {}".format(str(e)))
            continue
                
        try:
            if not os.path.exists(excel_file_local_path):
                output.print_md("ERROR: Excel file not found: {}".format(excel_file_local_path))
                output.print_md("Please ensure that the file exists and that the path is correct.")
                continue
            
            # Header printing
            output.print_md("Excel Parameter Synchronization")
            output.print_md("(Local)File: {}".format(excel_file_local_path))
            output.print_md("Named Range: {}".format(excel_rangename))
            output.print_md("Key Parameter: {}".format(excel_keyname))
            output.print_md("-" * 50)
            
            # reading data
            output.print_md("Reading Excel data...")
            data = read_excel_named_range(excel_file_local_path, excel_rangename)
            
            if not data:
                output.print_md("WARNING: The specified range is empty or contains no data.")
                continue
            
            # Display Excel data info
            output.print_md("Successfully read: {} Row(s), {} Column(s)".format(
                len(data), len(data[0]) if data else 0))
            
            # Convert Excel data to lookup dictionary
            excel_lookup = convert_excel_data_to_dict(data, excel_keyname)
            output.print_md("Excel lookup dictionary created with {} entries".format(len(excel_lookup)))
            
            # Synchronize parameters
            output.print_md("Starting parameter synchronization...")
            sync_parameters_to_elements(doc, elements, excel_lookup, excel_keyname, output)
            
        except Exception as e:
            output.print_md("ERROR processing the Excel file or synchronizing parameters")
            output.print_md("Error message: {}".format(str(e)))
            output.print_md("-" * 50)
            output.print_md("Possible solutions:")
            output.print_md("1. Check whether the Excel file exists and is accessible")
            output.print_md("2. Make sure that the named range is defined in the Excel file")
            output.print_md("3. Check whether the Excel file is being used by another program")
            output.print_md("4. For online files, check whether the file is synced locally")
            output.print_md("5. The file should be closed locally")
            output.print_md("6. Verify that the specified categories exist in the model")
            output.print_md("7. Check that the key parameter exists on the elements")
            
# ---------------------------
# Start
# ---------------------------
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("ERROR:", e)
        print(traceback.format_exc())