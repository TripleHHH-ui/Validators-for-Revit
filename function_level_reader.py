# -*- coding: utf-8 -*-
import clr
import os
import System

def read_excel_sheet(file_path):
    """Read Excel data as a list of rows using late binding COM Excel."""
    # Create Excel application using late binding
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    xl = System.Activator.CreateInstance(excel_type)
    
    try:
        # Use reflection to set properties
        excel_type.InvokeMember("Visible", 
            System.Reflection.BindingFlags.SetProperty, 
            None, xl, System.Array[object]([False]))
        
        excel_type.InvokeMember("DisplayAlerts", 
            System.Reflection.BindingFlags.SetProperty, 
            None, xl, System.Array[object]([False]))
        
        # Get Workbooks collection
        workbooks = excel_type.InvokeMember("Workbooks", 
            System.Reflection.BindingFlags.GetProperty, 
            None, xl, None)
        
        # Open workbook
        wb = workbooks.GetType().InvokeMember("Open", 
            System.Reflection.BindingFlags.InvokeMethod, 
            None, workbooks, System.Array[object]([file_path]))
        
        # Get ActiveSheet
        ws = wb.GetType().InvokeMember("ActiveSheet", 
            System.Reflection.BindingFlags.GetProperty, 
            None, wb, None)
        
        # Get UsedRange
        used_range = ws.GetType().InvokeMember("UsedRange", 
            System.Reflection.BindingFlags.GetProperty, 
            None, ws, None)
        
        # Get Columns collection from UsedRange
        columns = used_range.GetType().InvokeMember("Columns", 
            System.Reflection.BindingFlags.GetProperty, 
            None, used_range, None)
        
        # Get column count
        col_count = columns.GetType().InvokeMember("Count", 
            System.Reflection.BindingFlags.GetProperty, 
            None, columns, None)
        
        # Get Cells collection
        cells = ws.GetType().InvokeMember("Cells", 
            System.Reflection.BindingFlags.GetProperty, 
            None, ws, None)
        
        data = []
        row = 1
        
        while True:
            row_values = []
            empty_row = True
            
            for col in range(1, col_count + 1):
                cell = cells.GetType().InvokeMember("Item", 
                    System.Reflection.BindingFlags.GetProperty, 
                    None, cells, System.Array[object]([row, col]))
                val = cell.GetType().InvokeMember("Value2", 
                    System.Reflection.BindingFlags.GetProperty, 
                    None, cell, None)
                
                if val is not None:
                    empty_row = False
                row_values.append(val)
            
            if empty_row:
                break
            data.append(row_values)
            row += 1
        
        # Close workbook
        wb.GetType().InvokeMember("Close", 
            System.Reflection.BindingFlags.InvokeMethod, 
            None, wb, System.Array[object]([False]))
        
        return data
        
    finally:
        # Quit Excel
        try:
            excel_type.InvokeMember("Quit", 
                System.Reflection.BindingFlags.InvokeMethod, 
                None, xl, None)
        except:
            pass
        
        # Release COM object
        try:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xl)
        except:
            pass

def read_function_map(xlsx_path):
    """Read GIFA NAME -> FUNCTION ID mapping."""
    if not os.path.exists(xlsx_path):
        return {}
    data = read_excel_sheet(xlsx_path)
    if not data:
        return {}
    headers = [str(h).strip().upper() if h else "" for h in data[0]]
    try:
        idx_gifa = headers.index("GIFA NAME")
        idx_funcid = headers.index("FUNCTION ID")
    except ValueError:
        return {}
    function_map = {}
    for row in data[1:]:
        try:
            name = row[idx_gifa]
            fid = row[idx_funcid]
        except IndexError:
            continue
        if name and fid not in (None, "", "N", "N/A"):
            function_map[str(name).strip().upper()] = str(fid).strip()
    return function_map

def read_level_map(xlsx_path):
    """Read Elevation -> Level Code mapping."""
    if not os.path.exists(xlsx_path):
        return {}
    data = read_excel_sheet(xlsx_path)
    if not data:
        return {}
    headers = [str(h).strip().upper() if h else "" for h in data[0]]
    try:
        idx_elev = headers.index("ELEVATION")
        idx_code = headers.index("CODE")
    except ValueError:
        return {}
    level_map = {}
    for row in data[1:]:
        try:
            elev = row[idx_elev]
            code = row[idx_code]
        except IndexError:
            continue
        if elev is None or code in (None, "", "N/A"):
            continue
        try:
            level_map[int(round(float(elev)))] = str(code).strip().upper()
        except:
            continue
    return level_map