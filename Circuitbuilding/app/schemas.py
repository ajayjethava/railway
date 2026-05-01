SHEETS = {
    "StationDrawing": [
        "checksum", "station_id", "diagram_name", "station_name", "station_code",
        "version", "date", "drawn_by", "checked_by", "division", "zone",
        "total_sheet", "designation1", "designation2", "designation3"
    ],
    "junction_box": [
        "station_id", "junction_id", "junction_name", "latitude", "longitude",
        "junction_size", "junction_row"
    ],
    "cable": [
        "cable_id", "cable_name", "junction_box", "junction_name", "row",
        "position", "terminal", "start_no"
    ],
    "cable_box": [  # NEW: Add cable_box sheet with same columns as cable plus cable_type
        "cable_id", "cable_name", "junction_box", "junction_name", "row",
        "position", "terminal", "start_no", "cable_type","output"
    ],
    "terminal": [
        "cable_id", "terminal_id", "terminal_no", "symbol", "input_left",
        "input_right", "spare", "input_connected", "output_connected",
        "input_connected_extra", "output_connected_extra",
        "output_left", "output_right"
    ],
    "group": [
        "cable_id", "group_id", "terminal_no", "input_output", "text"
    ],
    "terminal_header": [
        "cable_id", "header_type", "terminal_start", "terminal_end", "input_output", "text"
    ],
    "choketable": [
        "cable_id", "choke_id", "input_terminal", "output_terminal", "terminal_name", "output_type", "output_text", "output_connected"
    ],
    "resistortable": [
        "cable_id", "resistor_id", "input_terminal", "output_terminal", "resistor_name"
    ]
}

HEADER_HINTS = {
    "StationDrawing": "Enter station metadata (checksum, IDs, names, zone, totals, designations).",
    "junction_box": "Enter each Location box or with coordinates and size/rows if available.",
    "cable": "Define each cable: names, junctions, row/position, terminal count, start number.",
    "cable_box": "Define relay boxes: similar to relays but with relay_box type.",  # NEW
    "terminal": "Define terminal details: symbols, connections, spare status, inputs/outputs.",
    "group": "Group terminals by cable with terminal numbers and input/output descriptions.",
    "terminal_header": "Define headers (WIREFROM/WIRETO/RELAY), terminal ranges, and connection notes.",
    "choketable": "Define choke components with input/output connections for cable filtering.",
    "resistortable": "Define resistor components with input/output terminals and resistance values.",
}