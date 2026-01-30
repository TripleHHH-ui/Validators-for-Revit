# -*- coding: utf-8 -*-
__title__ = 'Validate Room &\n Door Numbers'
__author__ = 'Huang Yuhan (Revit 2025 compatible version)'
try:
    import os
    import re
    from pyrevit import revit, script, forms
    from Autodesk.Revit.DB import *
    from door_rules_reader import read_door_direction_rules
    from function_level_reader import read_function_map, read_level_map
    from collections import defaultdict
    
    doc = revit.doc
    view = revit.active_view
    output = script.get_output()
    tool_dir = os.path.dirname(__file__)
    
    # --- Load Excel Config Files ---
    door_rule_file = os.path.join(tool_dir, "door_direction_rules.xlsx")
    function_map_file = os.path.join(tool_dir, "function_map.xlsx")
    level_map_file = os.path.join(tool_dir, "level_map.xlsx")
    
    door_rules = read_door_direction_rules(door_rule_file)
    function_map = read_function_map(function_map_file)
    level_map = read_level_map(level_map_file)
    
    output.print_md("### Config Files Loaded")
    output.print_md("- Door rules: `{}`".format(len(door_rules)))
    output.print_md("- Function map: `{}`".format(len(function_map)))
    output.print_md("- Level map: `{}`".format(len(level_map)))
    output.print_md("")
    
    # --- Find "New Construction" phase ---
    new_con_phase = None
    phase_collector = FilteredElementCollector(doc).OfClass(Phase)
    for phase in phase_collector:
        if phase.Name and phase.Name.lower() in ["new construction", "new"]:
            new_con_phase = phase
            break
    
    if not new_con_phase:
        raise Exception("Could not find 'New Construction' phase in model.")
    
    new_con_phase_id = new_con_phase.Id
    
    # --- Ask user which mode to validate ---
    modes = ['ALL', 'FRONT OF HOUSE (FOH)', 'BACK OF HOUSE (BOH)']
    selected_mode = forms.CommandSwitchWindow.show(
        modes,
        message='Select which category of rooms to validate:'
    )
    if not selected_mode:
        script.exit()
    
    MODE_ALIASES = {
        'ALL': 'ALL',
        'FRONT OF HOUSE (FOH)': 'FOH',
        'FOH': 'FOH',
        'BACK OF HOUSE (BOH)': 'BOH',
        'BOH': 'BOH'
    }
    validation_mode = MODE_ALIASES.get(selected_mode, 'ALL')
    output.print_md("### Validation Mode: **{}**".format(validation_mode))
    
    # =============================================================
    # GLOBAL SCOPE BOX COLLECTION + HELPERS (sectors)
    # =============================================================
    def parse_sector_code(text):
        if not text:
            return None
        t = (text or "").strip()
        m = re.match(r'^100_(\d{4})$', t)
        if m:
            return m.group(1)
        m = re.search(r'100_(\d{4})', t)
        return m.group(1) if m else None
    
    all_scope_boxes = []
    scope_box_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType()
    for sb in scope_box_collector:
        try:
            name = (sb.Name or "").strip()
            sector_code = parse_sector_code(name)
            if not sector_code:
                continue
            bbox = sb.get_BoundingBox(None)
            if not bbox:
                continue
            all_scope_boxes.append((sector_code, bbox))
        except:
            pass
    
    output.print_md("### Scope Boxes Loaded: {}".format(len(all_scope_boxes)))
    
    def _room_ref_point(room):
        try:
            bb = room.get_BoundingBox(None) if room else None
            if bb:
                return XYZ((bb.Min.X + bb.Max.X) * 0.5,
                           (bb.Min.Y + bb.Max.Y) * 0.5,
                           0.0)
        except:
            pass
        return None
    
    def _door_ref_point(door):
        try:
            loc = door.Location
            if isinstance(loc, LocationPoint):
                return loc.Point
            elif isinstance(loc, LocationCurve):
                curve = loc.Curve
                if curve:
                    return curve.Evaluate(0.5, True)
        except:
            pass
        try:
            bb = door.get_BoundingBox(None)
            if bb:
                return XYZ((bb.Min.X + bb.Max.X) * 0.5,
                           (bb.Min.Y + bb.Max.Y) * 0.5,
                           0.0)
        except:
            pass
        return None
    
    def sectors_overlapping_point(pt, boxes):
        if pt is None:
            return []
        hits = []
        eps = 0.01
        for code, bb in boxes:
            if ((bb.Min.X - eps) <= pt.X <= (bb.Max.X + eps)) and \
               ((bb.Min.Y - eps) <= pt.Y <= (bb.Max.Y + eps)):
                hits.append((code, bb))
        return hits
    
    def resolve_owner_sector_at_point(pt, boxes):
        overlaps = sectors_overlapping_point(pt, boxes)
        if not overlaps:
            return None
        overlaps.sort(key=lambda x: (x[1].Min.X, x[1].Min.Y))
        return overlaps[0][0]
    
    def resolve_owner_sector(room, boxes):
        pt = _room_ref_point(room)
        return resolve_owner_sector_at_point(pt, boxes)
    
    # =============================================================
    # Helper Functions
    # =============================================================
    def get_level_code(level):
        try:
            if not level:
                return "???"
            level_name = (level.Name or "").strip().upper()
            for _, code in level_map.items():
                code_str = str(code).upper()
                if level_name == code_str or level_name.endswith(" " + code_str):
                    return code_str
            for _, code in level_map.items():
                code_str = str(code).upper()
                if code_str in level_name:
                    return code_str
            elevation_mm = level.Elevation * 304.8
            closest = min(level_map.keys(), key=lambda k: abs(k - elevation_mm))
            return str(level_map[closest])
        except Exception as e:
            output.print_md("- Level mapping error for `{}`: {}".format(level.Name if level else '?', e))
            return "???"
    
    def get_function_id(name):
        try:
            name = (name or "").strip().upper()
        except:
            name = ""
        return function_map.get(name, None)
    
    def normalize_function_id(fid):
        try:
            if fid is None:
                return None
            if isinstance(fid, int):
                return fid
            s = str(fid).strip()
            if not s:
                return None
            if s.isdigit():
                return int(s)
            if s[0].isdigit():
                return int(s[0])
        except:
            pass
        return None
    
    def extract_function_id_from_number(room_number):
        try:
            parts = (room_number or '').split('-')
            if len(parts) == 3:
                tail = parts[2]
                if tail and tail[0].isdigit():
                    return int(tail[0])
        except:
            pass
        return None
    
    def get_area_category(function_id):
        fid = normalize_function_id(function_id)
        if fid is None:
            return "UNASSIGNED"
        foh_ids = [1, 5, 6, 8, 9]
        boh_ids = [2, 3, 7, 8, 9]
        if fid in foh_ids and fid in boh_ids:
            return "FOH & BOH"
        elif fid in foh_ids:
            return "FOH"
        elif fid in boh_ids:
            return "BOH"
        else:
            return "UNASSIGNED"
    
    # --- Room Validation ---
    def validate_rooms(view):
        output.print_md("## Room Number Validation")
        rooms = list(FilteredElementCollector(doc, view.Id)
                     .OfCategory(BuiltInCategory.OST_Rooms)
                     .WhereElementIsNotElementType())
        if not rooms:
            output.print_md("- No rooms found in this view.")
            return
        
        view_sector = None
        scope_box_param = view.LookupParameter("Scope Box")
        if scope_box_param and scope_box_param.HasValue:
            sb_elem = doc.GetElement(scope_box_param.AsElementId())
            if sb_elem:
                view_sector = parse_sector_code(sb_elem.Name)
        
        if not view_sector:
            view_sector = parse_sector_code(view.Name or "")
        
        if not view_sector:
            output.print_md("- ⚠️ Could not determine sector code for this view.")
            return
        
        output.print_md("- Using View Sector: `{}`".format(view_sector))
        
        valid_count = 0
        error_count = 0
        room_data = []
        
        for room in rooms:
            try:
                rid = room.Id.IntegerValue
                name_param = room.LookupParameter("Name")
                number_param = room.LookupParameter("Number")
                func_param = room.LookupParameter("GIFA NAME")
                
                if not (name_param and number_param and func_param):
                    error_count += 1
                    continue
                
                room_name = (name_param.AsString() or "").strip()
                room_number = (number_param.AsString() or "").strip()
                func_name = (func_param.AsString() or "").strip().upper()
                
                function_id = get_function_id(func_name)
                function_id = normalize_function_id(function_id)
                
                if function_id is None:
                    function_id = normalize_function_id(
                        extract_function_id_from_number(room_number)
                    )
                
                area_cat = get_area_category(function_id)
                
                if validation_mode == "FOH" and "FOH" not in area_cat:
                    continue
                if validation_mode == "BOH" and "BOH" not in area_cat:
                    continue
                
                level = doc.GetElement(room.LevelId)
                if not level:
                    error_count += 1
                    continue
                
                level_code = get_level_code(level)
                owner_sector = resolve_owner_sector(room, all_scope_boxes)
                
                if not owner_sector:
                    error_count += 1
                    continue
                
                if owner_sector != view_sector:
                    continue
                
                sector = owner_sector
                pt = None
                if room.Location:
                    pt = room.Location.Point
                else:
                    bbox = room.get_BoundingBox(None)
                    if bbox:
                        pt = XYZ((bbox.Min.X + bbox.Max.X) / 2.0,
                                 (bbox.Min.Y + bbox.Max.Y) / 2.0,
                                 0)
                
                if not pt:
                    continue
                
                room_data.append((sector, function_id, pt, room,
                                  level_code, area_cat, room_name, room_number))
            except Exception:
                error_count += 1
        
        grouped = defaultdict(list)
        for sector, fid, pt, room, level_code, area_cat, name, number in room_data:
            grouped[(sector, fid)].append(
                (pt, room, level_code, area_cat, fid, name, number)
            )
        
        for key, data in grouped.items():
            data.sort(key=lambda x: -x[0].X)
            band_tol = 3000.0 / 304.8
            bands = []
            current_band = []
            last_x = None
            
            for item in data:
                pt = item[0]
                if last_x is None or abs(pt.X - last_x) <= band_tol:
                    current_band.append(item)
                else:
                    bands.append(current_band)
                    current_band = [item]
                last_x = pt.X
            
            if current_band:
                bands.append(current_band)
            
            sorted_rooms = []
            for band in bands:
                band.sort(key=lambda x: -x[0].Y)
                sorted_rooms.extend(band)
            
            for idx, (pt, room, level_code, area_cat,
                      fid, name, number) in enumerate(sorted_rooms, start=1):
                rid = room.Id.IntegerValue
                function_code = "{}{:02d}".format(fid or 0, idx)
                expected_number = "{}-{}-{}".format(level_code, key[0], function_code)
                
                if number != expected_number:
                    output.print_md(
                        "• Room [{}](revit://element?id={}) '{}' [{}] Expected `{}` | Found `{}`".format(
                            rid, rid, name, area_cat, expected_number, number
                        ))
                    error_count += 1
                else:
                    output.print_md(
                        "• Room [{}](revit://element?id={}) '{}' [{}] OK".format(
                            rid, rid, name, area_cat
                        ))
                    valid_count += 1
        
        output.print_md("")
        output.print_md("Rooms OK: {}, Issues: {}".format(valid_count, error_count))
        output.print_md("")
    
    # --- Door Validation ---
    def get_door_room_with_phase(door, phase_id, from_room=True):
        try:
            if from_room:
                return door.get_ToRoom(phase_id)
            else:
                return door.get_FromRoom(phase_id)
        except:
            try:
                if from_room:
                    return door.ToRoom[new_con_phase]
                else:
                    return door.FromRoom[new_con_phase]
            except:
                return None
    #####################################################################
    def get_reference_room(door):
        try:
            to_room = get_door_room_with_phase(door, new_con_phase_id, from_room=False)
            from_room = get_door_room_with_phase(door, new_con_phase_id, from_room=True)
        except:
            return None
        if to_room and not from_room:
            return to_room
        if from_room and not to_room:
            return from_room
        return from_room
    #######################################################################
    def get_reference_room_for_sector(door, door_sector):
        try:
            to_room = get_door_room_with_phase(door, new_con_phase_id, from_room=False)
            from_room = get_door_room_with_phase(door, new_con_phase_id, from_room=True)
        except:
            return None
        candidates = []
        if from_room:
            candidates.append(from_room)
        if to_room and to_room != from_room:
            candidates.append(from_room)
        for r in candidates:
            rs = resolve_owner_sector(r, all_scope_boxes)
            if rs == door_sector:
                return r
        if to_room and not from_room:
            return to_room
        if from_room and not to_room:
            return from_room
        return from_room 
    
    def validate_doors(view):
        output.print_md("## Door Number Validation")
        print("NOTE: Only check the door number after room numbers are corrected!!")
        doors = list(FilteredElementCollector(doc, view.Id)
                     .OfCategory(BuiltInCategory.OST_Doors)
                     .WhereElementIsNotElementType())
        if not doors:
            output.print_md("- No doors found in this view.")
            return
        
        view_sector = None
        scope_box_param = view.LookupParameter("Scope Box")
        if scope_box_param and scope_box_param.HasValue:
            sb_elem = doc.GetElement(scope_box_param.AsElementId())
            if sb_elem:
                view_sector = parse_sector_code(sb_elem.Name)
        
        if not view_sector:
            view_sector = parse_sector_code(view.Name or "")
        
        if not view_sector:
            output.print_md("- ⚠️ Could not determine sector code for this view (doors).")
            return
        
        valid_count = 0
        error_count = 0
        SKIP_PHRASE = "NOT FOR DOOR SCHEDULE"
        
        for door in doors:
            try:
                did = door.Id.IntegerValue
                skip = False
                try:
                    if door.Symbol:
                        type_comments = door.Symbol.get_Parameter(
                            BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
                        if (type_comments and type_comments.HasValue
                                and SKIP_PHRASE in (type_comments.AsString() or "").upper()):
                            skip = True
                    if not skip:
                        inst_comments = door.LookupParameter("Comments")
                        if (inst_comments and inst_comments.HasValue
                                and SKIP_PHRASE in (inst_comments.AsString() or "").upper()):
                            skip = True
                except:
                    pass
                
                if skip:
                    continue
                
                door_pt = _door_ref_point(door)
                door_sector = resolve_owner_sector_at_point(door_pt, all_scope_boxes)
                
                if not door_sector:
                    output.print_md("- Door [{}](revit://element?id={}) could not resolve sector.".format(did, did))
                    error_count += 1
                    continue
                
                if door_sector != view_sector:
                    continue
                
                mark_param = door.LookupParameter("Mark")
                if not (mark_param and mark_param.HasValue):
                    output.print_md("- Door [{}](revit://element?id={}) has no Mark.".format(did, did))
                    error_count += 1
                    continue
                
                mark = mark_param.AsString() or ""
                
                try:
                    to_room = get_door_room_with_phase(door, new_con_phase_id, from_room=False)
                    from_room = get_door_room_with_phase(door, new_con_phase_id, from_room=True)
                except:
                    to_room = None
                    from_room = None
                
                ref_room = None
                if to_room and not from_room:
                    ref_room = to_room
                elif from_room and not to_room:
                    ref_room = from_room
                elif from_room and to_room:
                    ref_room = from_room
                
                if not ref_room:
                    output.print_md("- Door [{}](revit://element?id={}) has no room reference.".format(did, did))
                    error_count += 1
                    continue
                
                num_param = ref_room.LookupParameter("Number")
                if not (num_param and num_param.HasValue):
                    output.print_md("- Door [{}](revit://element?id={}) reference room missing Number.".format(did, did))
                    error_count += 1
                    continue
                
                room_number = num_param.AsString() or ""
                
                ref_room_name = ""
                try:
                    name_param = ref_room.LookupParameter("Name")
                    if name_param and name_param.HasValue:
                        ref_room_name = name_param.AsString() or ""
                except:
                    pass
                
                pattern = r'^{}[A-Z]?$'.format(re.escape(room_number))
                if not re.match(pattern, mark):
                    output.print_md(
                        "• Door [{}](revit://element?id={}) → Room '{}' [{}] Expected `{}`[A-Z] | Found `{}`".format(
                            did, did, ref_room_name, room_number, room_number, mark))
                    error_count += 1
                else:
                    valid_count += 1
            except Exception as e:
                output.print_md("- Error validating Door [{}]: {}".format(did, e))
                error_count += 1
        
        output.print_md("")
        output.print_md("Doors OK: {}, Issues: {}".format(valid_count, error_count))
        output.print_md("")
    
    # --- Run Validations ---
    output.print_md("### 🔹 Validating View: `{}`".format(view.Name))
    validate_rooms(view)
    validate_doors(view)
    output.print_md("---")
    output.print_md("Validation completed.")

except Exception as e:
    import traceback
    output = script.get_output()
    output.print_md("### Script Error")
    output.print_md("```\n{}\n```".format(traceback.format_exc()))