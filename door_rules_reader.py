# -*- coding: utf-8 -*-
import clr
import System

def read_door_direction_rules(excel_path):
    """Read door direction rules from Excel file using late binding."""
    # Create Excel application using late binding
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    excel = System.Activator.CreateInstance(excel_type)
    
    try:
        # Use reflection to set properties and call methods
        excel_type.InvokeMember("Visible", 
            System.Reflection.BindingFlags.SetProperty, 
            None, excel, System.Array[object]([False]))
        
        excel_type.InvokeMember("DisplayAlerts", 
            System.Reflection.BindingFlags.SetProperty, 
            None, excel, System.Array[object]([False]))
        
        # Get Workbooks collection
        workbooks = excel_type.InvokeMember("Workbooks", 
            System.Reflection.BindingFlags.GetProperty, 
            None, excel, None)
        
        # Open workbook
        workbook = workbooks.GetType().InvokeMember("Open", 
            System.Reflection.BindingFlags.InvokeMethod, 
            None, workbooks, System.Array[object]([excel_path]))
        
        # Get Sheets collection
        sheets = workbook.GetType().InvokeMember("Sheets", 
            System.Reflection.BindingFlags.GetProperty, 
            None, workbook, None)
        
        # Get first sheet
        sheet = sheets.GetType().InvokeMember("Item", 
            System.Reflection.BindingFlags.GetProperty, 
            None, sheets, System.Array[object]([1]))
        
        # Get Cells collection
        cells = sheet.GetType().InvokeMember("Cells", 
            System.Reflection.BindingFlags.GetProperty, 
            None, sheet, None)
        
        rules = {
            "flip_contains": [], 
            "flip_search_contains": [], 
            "block_flip_equals": []
        }
        
        row = 2  # assuming headers in row 1
        while True:
            # Read cell values
            cell1 = cells.GetType().InvokeMember("Item", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cells, System.Array[object]([row, 1]))
            c1 = cell1.GetType().InvokeMember("Value2", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cell1, None)
            
            cell2 = cells.GetType().InvokeMember("Item", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cells, System.Array[object]([row, 2]))
            c2 = cell2.GetType().InvokeMember("Value2", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cell2, None)
            
            cell3 = cells.GetType().InvokeMember("Item", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cells, System.Array[object]([row, 3]))
            c3 = cell3.GetType().InvokeMember("Value2", 
                System.Reflection.BindingFlags.GetProperty, 
                None, cell3, None)
            
            if not c1 and not c2 and not c3:
                break
                
            if c1:
                rules["flip_contains"].append(str(c1).strip().upper())
            if c2:
                rules["flip_search_contains"].append(str(c2).strip().upper())
            if c3:
                rules["block_flip_equals"].append(str(c3).strip().upper())
            
            row += 1
        
        # Close workbook without saving
        workbook.GetType().InvokeMember("Close", 
            System.Reflection.BindingFlags.InvokeMethod, 
            None, workbook, System.Array[object]([False]))
        
        return rules
        
    finally:
        # Quit Excel
        try:
            excel_type.InvokeMember("Quit", 
                System.Reflection.BindingFlags.InvokeMethod, 
                None, excel, None)
        except:
            pass
        
        # Release COM object
        try:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        except:
            pass