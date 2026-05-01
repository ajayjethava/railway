import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import FancyBboxPatch, Circle, Polygon, Rectangle
import numpy as np
import re
import os
from matplotlib.backends.backend_pdf import PdfPages
import sys
from collections import OrderedDict
import hashlib
from datetime import datetime, timezone
import json
import PyPDF2
from pytz import timezone

def generate_pdf_metadata_checksum(df_title, excel_file_path, pdf_file_path):
    """
    Generate checksum for PDF metadata using station code, content size, creation date, and file name
    """
    try:
        # Get current date/time in IST
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        timestamp = current_time.strftime("%Y-%m-%d %H:%M:%S")
    
        # Extract station details
        station_code = ""
        station_name = ""
        if df_title is not None and not df_title.empty:
            title_row = df_title.iloc[0] if hasattr(df_title, 'iloc') else df_title
            station_code = str(title_row.get('station_code', ''))
            station_name = str(title_row.get('station_name', ''))
    
        # Placeholder for content size (will be updated after PDF generation)
        content_size = 0
    
        # Get file name (dynamic)
        file_name = os.path.basename(pdf_file_path)
    
        # Create data string for checksum
        checksum_data = f"{station_code}|{content_size}|{timestamp}|{file_name}"
    
        # Generate MD5 checksum
        checksum = hashlib.md5(checksum_data.encode()).hexdigest()
    
        print(f"Checksum generated: {checksum}")
        print(f"Checksum data: {checksum_data}")
    
        return checksum, checksum_data, content_size, timestamp # Return timestamp too for consistency
    
    except Exception as e:
        print(f"Error generating checksum: {e}")
        return None, None, 0, None

def enhance_pdf_with_metadata(pdf_file_path, checksum, checksum_data, content_size, df_title):
    try:
        # Get station details
        station_code = ""
        if df_title is not None and not df_title.empty:
            title_row = df_title.iloc[0]
            station_code = str(title_row.get('station_code', ''))
        # Read the existing PDF with error handling
        try:
            with open(pdf_file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                pdf_writer = PyPDF2.PdfWriter()
                # Add all pages
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)
        except Exception as e:
            print(f"Warning: Could not read PDF with PyPDF2: {e}")
            print("Skipping metadata enhancement as PDF may be corrupted or in use")
            return False
        # Capture timestamp in IST (same as in update function)
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        now_str = current_time.strftime("D:%Y%m%d%H%M%S+05'30'") # Add +05'30' for IST offset
        # Add comprehensive metadata
        pdf_writer.add_metadata({
            '/Title': f'Terminal Drawing - Station {station_code}',
            '/Author': 'SaltRiver Infosystems Pvt Ltd.',
            '/Subject': f'Railway Project Terminal Diagram - Checksum: {checksum}',
            '/Keywords': f'station_code:{station_code},checksum:{checksum},content_size:{content_size}',
            '/Creator': 'Python Diagram Generator',
            '/Producer': 'Matplotlib/PyPDF2',
            '/CreationDate': now_str,
            '/ModDate': now_str
        })
        # Write to a temporary file
        temp_output = pdf_file_path.replace('.pdf', '_with_metadata.pdf')
        try:
            with open(temp_output, 'wb') as output_file:
                pdf_writer.write(output_file)
        except Exception as e:
            print(f"Error writing enhanced PDF: {e}")
            return False
        # Try to replace original
        try:
            os.remove(pdf_file_path)
            os.rename(temp_output, pdf_file_path)
            print(f"PDF metadata enhanced with checksum: {checksum}")
            return True
        except PermissionError:
            print("Cannot replace the original PDF because it is open or locked. The enhanced PDF is saved as '" + temp_output + "'")
            return False
        except Exception as e:
            print(f"Error replacing original PDF: {e}")
            return False
    except Exception as e:
        print(f"Error enhancing PDF metadata: {e}")
        return False

def update_pdf_checksum_metadata(pdf_file_path, checksum, checksum_data, content_size, df_title):
    """
    Update the PDF metadata with the checksum information
    """
    try:
        # Get station details for metadata
        station_code = ""
        station_name = ""
        if df_title is not None and not df_title.empty:
            title_row = df_title.iloc[0] if hasattr(df_title, 'iloc') else df_title
            station_code = str(title_row.get('station_code', ''))
            station_name = str(title_row.get('station_name', ''))
        # Capture timestamp in IST for consistency
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        timestamp = current_time.strftime("%Y-%m-%d %H:%M:%S")
        # Get actual file size BEFORE enhancement (initial size)
        initial_content_size = os.path.getsize(pdf_file_path)
        # Update checksum data with initial file size and IST timestamp
        updated_checksum_data = f"{station_code}|{initial_content_size}|{timestamp}|{os.path.basename(pdf_file_path)}"
        updated_checksum = hashlib.md5(updated_checksum_data.encode()).hexdigest()
        print(f"Initial checksum with initial file size: {updated_checksum}")
        print(f"Initial file size: {initial_content_size} bytes")
        print(f"Checksum data string: {updated_checksum_data}")
        # Enhance metadata (this may change file size) - with error handling
        enhance_success = enhance_pdf_with_metadata(pdf_file_path, updated_checksum, updated_checksum_data, initial_content_size, df_title)
        # Get final file size AFTER enhancement
        final_content_size = os.path.getsize(pdf_file_path)
        print(f"Final file size after enhancement: {final_content_size} bytes")
        # Always use final size for checksum if it changed
        if final_content_size != initial_content_size:
            updated_checksum_data = f"{station_code}|{final_content_size}|{timestamp}|{os.path.basename(pdf_file_path)}"
            updated_checksum = hashlib.md5(updated_checksum_data.encode()).hexdigest()
            print(f"Updated checksum with final file size: {updated_checksum}")
            print(f"Updated checksum data string: {updated_checksum_data}")
        # Compute full MD5 of the entire PDF file content
        try:
            with open(pdf_file_path, 'rb') as f:
                full_file_md5 = hashlib.md5(f.read()).hexdigest()
            print(f"Full file MD5 hash: {full_file_md5}")
        except Exception as e:
            print(f"Warning: Could not compute full file MD5: {e}")
            full_file_md5 = "N/A"
        # Store as log file (append mode to preserve all logs)
        log_file = 'checksum.log'
        try:
            with open(log_file, 'a') as f:
                f.write("\n" + "-" * 50 + "\n")
                f.write(f"Metadata Checksum: {updated_checksum}\n")
                f.write(f"Metadata Data string: {updated_checksum_data}\n")
                f.write(f"Initial file size: {initial_content_size} bytes\n")
                f.write(f"Final file size: {final_content_size} bytes\n")
                f.write(f"Timestamp (IST): {timestamp}\n")
                f.write(f"Station code: {station_code}\n")
                f.write(f"File name: {os.path.basename(pdf_file_path)}\n")
                f.write(f"Full file MD5: {full_file_md5}\n")
            print(f"Checksum details appended to '{log_file}'")
        except Exception as e:
            print(f"Warning: Could not write to checksum log: {e}")
        return updated_checksum, updated_checksum_data, final_content_size
    except Exception as e:
        print(f"Error updating PDF metadata: {e}")
        return checksum, checksum_data, content_size

# === ENHANCED ROW ORDERING FUNCTION ===
def get_row_order(letter):
    """
    Enhanced row ordering function to ensure proper descending order
    """
    if pd.isna(letter) or str(letter).strip() == '':
        return 0
    
    letter_str = str(letter).strip().upper()
    
    # Define the desired order (highest to lowest)
    row_order = {
        'F': 6, 'E': 5, 'D': 4, 'C': 3, 'B': 2, 'A': 1,
        'H': 8, 'G': 7  # For CTR junctions that have H and G rows
    }
    
    if letter_str in row_order:
        return row_order[letter_str]
    
    # For any other letters, use their ASCII value with offset
    try:
        return 100 - ord(letter_str)  # This makes Z > Y > X > ... > A
    except:
        return 0

# === UPDATED PAGINATION LOGIC WITH CABLE BOX LIMIT ===
def break_cables_into_rows_updated(cable_list, max_terminal_symbols_per_row=36, max_cable_boxes_per_row=6):
    """
    Break cables into multiple rows if they exceed the terminal limit,
    keeping cables from the same letter together.
    Now handles both regular cables and cable boxes with a limit on cable boxes per row.
    WORKS FOR ALL LETTERS (A-Z), not just F row.
    """
    rows = []
    current_row = []
    current_terminal_count = 0
    current_cable_box_count = 0
    
    # Check if df_cable_box exists and is not empty
    df_cable_box_exists = 'df_cable_box' in globals() and df_cable_box is not None and not df_cable_box.empty
    
    for cable_id in cable_list:
        is_cable_box = False
        total_terminals = 0
        
        if df_cable_box_exists:
            # Check if this is a cable box
            cable_box_rows = df_cable_box[df_cable_box['cable_id'] == cable_id]
            is_cable_box = not cable_box_rows.empty
        
        if is_cable_box:
            # This is a cable box - calculate terminal count based on position
            cable_info = cable_box_rows.iloc[0]
            position_val = cable_info.get('position')
            if pd.notna(position_val):
                try:
                    total_terminals = int(float(position_val))
                except:
                    total_terminals = 1
            else:
                total_terminals = 1
        else:
            # Regular cable - calculate from symbols
            group = df_symbols[df_symbols['cable_id'] == cable_id].sort_index().reset_index(drop=True)
            total_terminals = 0
            i = 0
            while i < len(group):
                symbol = str(group.iloc[i].get('symbol', '')).strip().lower()
                if symbol == 'dual_fuse':
                    if i + 1 < len(group):
                        total_terminals += 2
                        i += 2
                    else:
                        total_terminals += 1
                        i += 1
                else:
                    total_terminals += 1
                    i += 1
        
        # Check if we need to break the row due to cable box limit
        if is_cable_box and current_cable_box_count >= max_cable_boxes_per_row and current_row:
            rows.append(current_row)
            current_row = []
            current_terminal_count = 0
            current_cable_box_count = 0
        
        # If adding this cable would exceed the terminal limit, start a new row
        if current_terminal_count + total_terminals > max_terminal_symbols_per_row and current_row:
            rows.append(current_row)
            current_row = []
            current_terminal_count = 0
            current_cable_box_count = 0
        
        current_row.append(cable_id)
        current_terminal_count += total_terminals
        if is_cable_box:
            current_cable_box_count += 1
    
    # Add the last row
    if current_row:
        rows.append(current_row)
    
    return rows

# === UPDATED FUNCTION: Draw cable box row ===
def draw_cable_box_row(ax, x_start, y_center, cable_info, pin_spacing=0.8):
    """
    Draw a SINGLE cable box for cables with cabel_type = 'cabel_box'
    Each cable box has:
    - One rectangle
    - One terminal number above it based on start_no (if provided) or position
    - Inside text: cable_name
    UPDATED: Use start_no if provided, otherwise use position for numbering
    """
    cable_name = cable_info.get('cable_name', '')
    start_no = cable_info.get('start_no')
    position = cable_info.get('position')
    
    # Get the position number for this specific cable box
    try:
        position_num = int(float(position)) if pd.notna(position) and str(position).strip() != '' else 1
    except:
        position_num = 1
    
    # Draw ONE rectangle for this cable box
    x_positions = [x_start + 1.0]
    input_connected_flags = [False]
    output_connected_flags = [False]
    
    # Rectangle dimensions - bigger for better visibility
    rect_width = 1.5
    rect_height = 0.7
    
    # Draw the rectangle
    rect_x = x_start + 1.0 - rect_width / 2
    rect_y = y_center - rect_height / 2
    ax.add_patch(Rectangle((rect_x, rect_y), rect_width, rect_height,
                           edgecolor='black', facecolor='white', linewidth=1.5))
    
    # Generate terminal number: use start_no if provided, otherwise use position
    if pd.notna(start_no) and str(start_no).strip() != '':
        try:
            # Use the start_no directly
            terminal_num = int(float(start_no))
            upper_text = f"{terminal_num:02d}" if terminal_num < 100 else str(terminal_num)
        except:
            # If start_no conversion fails, fall back to position
            terminal_num = position_num
            upper_text = f"{terminal_num:02d}"
    else:
        # No start_no provided, use position
        terminal_num = position_num
        upper_text = f"{terminal_num:02d}"
    
    # Position the terminal number above the rectangle
    ax.text(x_start + 1.0, rect_y + rect_height + 0.15, upper_text,
            fontsize=16, ha='center', va='bottom', fontname='Arial', fontweight='bold')
    
    # Inside text (cable name) - centered in the rectangle
    if pd.notna(cable_name) and str(cable_name).strip() != '':
        cable_text = str(cable_name).strip()
        ax.text(x_start + 1.0, y_center, cable_text,
                fontsize=18, ha='center', va='center', fontname='Arial', fontweight='bold')
    
    return x_positions, input_connected_flags, output_connected_flags

# === UPDATED FUNCTION: Draw extra connections ===
def draw_extra_connections(ax, cable_rows, x_positions, terminal_nos_for_positions,
                          y_top_bus_group, y_bottom_bus_group, capsule_y_center):
    """
    Draw extra connections based on input_connected_extra and output_connected_extra columns
    Format: "start_terminal,end_terminal" e.g., "1,8" means connect terminal 1 to terminal 8
    (Changed delimiter from '.' to ',')
    Draws multiple connections in staggered layers to avoid overlapping
    """
    try:
        # Process input_connected_extra (top connections)
        connections_drawn = 0
        top_connections = []
        bottom_connections = []
     
        # First, collect all connections
        for _, row in cable_rows.iterrows():
            # Collect input connections
            extra_input = row.get('input_connected_extra')
            if pd.notna(extra_input) and str(extra_input).strip() != '':
                try:
                    # CHANGED: Split by comma instead of period
                    pairs = str(extra_input).strip().split(',')
                    if len(pairs) == 2:
                        start_term_raw = pairs[0].strip()
                        end_term_raw = pairs[1].strip()
                     
                        try:
                            start_term = str(int(float(start_term_raw)))
                        except:
                            start_term = start_term_raw
                         
                        try:
                            end_term = str(int(float(end_term_raw)))
                        except:
                            end_term = end_term_raw
                     
                        if start_term in terminal_nos_for_positions and end_term in terminal_nos_for_positions:
                            start_idx = terminal_nos_for_positions.index(start_term)
                            end_idx = terminal_nos_for_positions.index(end_term)
                            x1 = x_positions[start_idx]
                            x2 = x_positions[end_idx]
                            top_connections.append((x1, x2, start_term, end_term))
                except Exception as e:
                    print(f"Error processing input_connected_extra '{extra_input}': {e}")
         
            # Collect output connections
            extra_output = row.get('output_connected_extra')
            if pd.notna(extra_output) and str(extra_output).strip() != '':
                try:
                    # CHANGED: Split by comma instead of period
                    pairs = str(extra_output).strip().split(',')
                    if len(pairs) == 2:
                        start_term_raw = pairs[0].strip()
                        end_term_raw = pairs[1].strip()
                     
                        try:
                            start_term = str(int(float(start_term_raw)))
                        except:
                            start_term = start_term_raw
                         
                        try:
                            end_term = str(int(float(end_term_raw)))
                        except:
                            end_term = end_term_raw
                     
                        if start_term in terminal_nos_for_positions and end_term in terminal_nos_for_positions:
                            start_idx = terminal_nos_for_positions.index(start_term)
                            end_idx = terminal_nos_for_positions.index(end_term)
                            x1 = x_positions[start_idx]
                            x2 = x_positions[end_idx]
                            bottom_connections.append((x1, x2, start_term, end_term))
                except Exception as e:
                    print(f"Error processing output_connected_extra '{extra_output}': {e}")
     
        # Draw top connections in staggered layers
        if top_connections:
        
            # Sort connections by their length (shorter first) to minimize crossing
            top_connections.sort(key=lambda conn: abs(conn[1] - conn[0]))
         
            # Group connections by overlapping and assign layers
            layers = []
            for conn in top_connections:
                x1, x2, start_term, end_term = conn
                placed = False
             
                for layer_idx, layer in enumerate(layers):
                    # Check if this connection can be placed in this layer without overlapping
                    can_place = True
                    for existing_conn in layer:
                        ex_x1, ex_x2 = existing_conn[0], existing_conn[1]
                        # Check for overlap: if the x-intervals overlap
                        if not (x2 <= ex_x1 or x1 >= ex_x2):
                            can_place = False
                            break
                 
                    if can_place:
                        layer.append(conn)
                        placed = True
                        break
             
                if not placed:
                    layers.append([conn])
         
            # Draw each layer
            for layer_idx, layer in enumerate(layers):
                layer_y_offset = 0.8 + (layer_idx * 0.4) # Stagger each layer by 0.4 units
             
                for conn in layer:
                    x1, x2, start_term, end_term = conn
                 
                    # Draw connection line at top (above the bus line)
                    extra_y = y_top_bus_group + layer_y_offset
                    ax.plot([x1, x2], [extra_y, extra_y],
                           color='black', linewidth=1, linestyle='-', zorder=5)
                 
                    vertical_offset = 0.43 # how far down you want it
                    ax.plot([x1, x1], [y_top_bus_group - vertical_offset, extra_y],
                            color='black', linewidth=1, linestyle='-', zorder=5)
                    ax.plot([x2, x2], [y_top_bus_group - vertical_offset, extra_y],
                            color='black', linewidth=1, linestyle='-', zorder=5)
                 
                    connections_drawn += 1
     
        # Draw bottom connections in staggered layers
        if bottom_connections:
        
            # Sort connections by their length (shorter first)
            bottom_connections.sort(key=lambda conn: abs(conn[1] - conn[0]))
         
            # Group connections by overlapping and assign layers
            layers = []
            for conn in bottom_connections:
                x1, x2, start_term, end_term = conn
                placed = False
             
                for layer_idx, layer in enumerate(layers):
                    # Check if this connection can be placed in this layer without overlapping
                    can_place = True
                    for existing_conn in layer:
                        ex_x1, ex_x2 = existing_conn[0], existing_conn[1]
                        # Check for overlap: if the x-intervals overlap
                        if not (x2 <= ex_x1 or x1 >= ex_x2):
                            can_place = False
                            break
                 
                    if can_place:
                        layer.append(conn)
                        placed = True
                        break
             
                if not placed:
                    layers.append([conn])
         
            # Draw each layer
            for layer_idx, layer in enumerate(layers):
                layer_y_offset = 0.8 + (layer_idx * 0.4) # Stagger each layer by 0.4 units
             
                for conn in layer:
                    x1, x2, start_term, end_term = conn
                 
                    # Draw connection line at bottom (below the bus line)
                    extra_y = y_bottom_bus_group - layer_y_offset
                    ax.plot([x1, x2], [extra_y, extra_y],
                           color='black', linewidth=1, linestyle='-', zorder=5)
                 
                    # Add small vertical lines to connect to the bus
                    ax.plot([x1, x1], [y_bottom_bus_group - 0.2, extra_y],
                           color='black', linewidth=1, linestyle='-', zorder=5)
                    ax.plot([x2, x2], [y_bottom_bus_group - 0.2, extra_y],
                           color='black', linewidth=1, linestyle='-', zorder=5)
                 
                    connections_drawn += 1
     
    except Exception as e:
        print(f"Error in draw_extra_connections: {e}")
        import traceback
        traceback.print_exc()

# === Function to merge ranges ===
def merge_ranges(ranges, merge_adjacent=True):
    """
    Merge overlapping (and optionally adjacent) integer index ranges.
    ranges: list of (start, end) tuples where start<=end
    merge_adjacent: if True, ranges like (0,5) and (6,11) will merge into (0,11).
                    if False, adjacent ranges remain separate.
    """
    if not ranges:
        return []
    sorted_ranges = sorted(ranges)
    merged = [list(sorted_ranges[0])]
    for current in sorted_ranges[1:]:
        last = merged[-1]
        if merge_adjacent:
            cond = (current[0] <= last[1] + 1)
        else:
            cond = (current[0] <= last[1]) # only merge if overlapping, not merely adjacent
        if cond:
            last[1] = max(last[1], current[1])
        else:
            merged.append(list(current))
    return [tuple(r) for r in merged]

# === Function to get cable name ===
def get_block_cable_name(df_block):
    if 'row' in df_block.columns:
        s = (
            df_block['row']
            .dropna().astype(str).str.strip()
            .replace('', pd.NA).dropna()
        )
        if not s.empty:
            return s.iloc[0]
    return ""

# === Function to draw cable name with circle ===
def draw_cable_name(ax, x, y, cable_name, x_offset=0.65):
    circle_center = (x + x_offset, y)
    ax.add_patch(Circle(circle_center, radius=0.22,
                        edgecolor='black', facecolor='white', linewidth=0.8))
    ax.text(x + x_offset, y, cable_name, ha='center', va='center',
            fontsize=22, fontname='Arial',fontweight='bold')

# === Function to draw junction box with big text ===
def draw_junction_box(ax, x, y, junction_name, rect_pad=0.2):
    text_raw = str(junction_name).strip()
    if not text_raw:
        return
    s = text_raw
    text_width = len(s) * 0.35
    text_height = 1.0
    font_size = max(22, min(37, 37 - (len(s) - 5) * 0.5))
    rect_width = text_width + rect_pad * 10
    rect_y = y - text_height / 2
    rect = Rectangle((x - rect_width / 2, rect_y), rect_width, text_height,
                     linewidth=2, edgecolor='black', facecolor='none')
    ax.add_patch(rect)
    ax.text(x, y, s, ha='center', va='center',
            fontsize=font_size, fontname='Arial', zorder=5,fontweight='bold')

# === Relay input symbol ===
def draw_relay_input(ax, x_left, x_right, y=0, scale=1.0, text='RELAY', anchor_to_v_tip=False, v_offset=-0.5, is_not_connected_with_bus=False):
    if x_left is None or x_right is None:
        return
    if x_left > x_right:
        x_left, x_right = x_right, x_left
 
    # Apply downward offset ONLY for Scenario 2 (not connected + bus line)
    if is_not_connected_with_bus:
        y += -1 # Move relay down by 0.3 units ONLY in this specific case
 
    span = max(1e-6, float(x_right) - float(x_left))
    center = (x_left + x_right) / 2.0
    tri_base = min(max(span * 0.18, 0.25 * scale), span * 0.45)
    tri_height = tri_base * 0.25 * scale
    v_depth = tri_height * 0.9
    left_notch = (center - tri_base / 2.0, y - v_offset)
    right_notch = (center + tri_base / 2.0, y - v_offset)
    notch_top = (center, y + tri_height - v_offset)
    v_tip = (center, y - v_depth - v_offset)
    pad = min(span * 0.02 + 0.02 * scale, span * 0.05)
    left_wire_x_start = x_left - pad
    right_wire_x_end = x_right + pad
    slant_size = min(0.18 * scale, tri_base * 0.35)
    end_slant_size = slant_size * 0.6
 
    # === COMMENTED OUT: Relay input graphical elements ===
    """
    ax.plot([left_wire_x_start - end_slant_size, left_wire_x_start],
            [y - end_slant_size - v_offset, y - v_offset], color='black', linewidth=1.4)
    ax.plot([right_wire_x_end, right_wire_x_end + end_slant_size],
            [y - v_offset, y + end_slant_size - v_offset], color='black', linewidth=1.4)
    if anchor_to_v_tip:
        ax.plot([left_notch[0], v_tip[0]], [left_notch[1], v_tip[1]], linewidth=1.4, color='black')
        ax.plot([right_notch[0], v_tip[0]], [right_notch[1], v_tip[1]], linewidth=1.4, color='black')
    else:
        ax.plot([left_wire_x_start, left_notch[0]], [y - v_offset, left_notch[1]], linewidth=1.4, color='black')
        ax.plot([right_notch[0], right_wire_x_end], [right_notch[1], y - v_offset], linewidth=1.4, color='black')
    ax.plot([left_notch[0], notch_top[0]], [left_notch[1], notch_top[1]], color='black', linewidth=1.4)
    ax.plot([right_notch[0], notch_top[0]], [right_notch[1], notch_top[1]], color='black', linewidth=1.4)
    """
 
    text_y_offset = 0.15 * scale
    text_y = notch_top[1] + text_y_offset
    ax.text(center, text_y, str(text), ha='center', va='bottom',
            fontsize=int(39 * scale), fontname='Arial')
 
    # === COMMENTED OUT: Small vertical lines for "LIGHT" ===
    """
    # Add two small vertical lines on upper side if text is "LIGHT"
    if str(text).strip().upper() == "LIGHT":
        line_length = 0.3 * scale
        gap = 0.25 * scale # increase gap between lines
        left_line_x = center - gap
        right_line_x = center + gap
        base_y = y - v_offset - 0.4 * scale # move lines slightly down side
        # Left vertical line
        ax.plot([left_line_x, left_line_x], [base_y, base_y + line_length], color='black', linewidth=1.4)
        # Right vertical line
        ax.plot([right_line_x, right_line_x], [base_y, base_y + line_length], color='black', linewidth=1.4)
    """

def draw_relay_output(ax, x_left, x_right, y=0, scale=1.0, text='RELAY', anchor_to_v_tip=False, v_offset=0.5, is_not_connected_with_bus=False):
    if x_left is None or x_right is None:
        return
    if x_left > x_right:
        x_left, x_right = x_right, x_left
 
    # Apply downward offset ONLY for Scenario 2 (not connected + bus line)
    if is_not_connected_with_bus:
        y -= 0.3 # Move relay down by 0.3 units ONLY in this specific case
 
    span = max(1e-6, float(x_right) - float(x_left))
    center = (x_left + x_right) / 2.0
    tri_base = min(max(span * 0.18, 0.25 * scale), span * 0.45)
    tri_height = tri_base * 0.25 * scale
    v_depth = tri_height * 0.9
    left_notch = (center - tri_base / 2.0, y - v_offset)
    right_notch = (center + tri_base / 2.0, y - v_offset)
    notch_bottom = (center, y - tri_height - v_offset)
    v_tip = (center, y + v_depth - v_offset)
    pad = min(span * 0.02 + 0.02 * scale, span * 0.05)
    left_wire_x_start = x_left - pad
    right_wire_x_end = x_right + pad
    slant_size = min(0.18 * scale, tri_base * 0.35)
    end_slant_size = slant_size * 0.6
 
    # === COMMENTED OUT: Relay output graphical elements ===
    """
    ax.plot([left_wire_x_start - end_slant_size, left_wire_x_start],
            [y + end_slant_size - v_offset, y - v_offset], color='black', linewidth=1.4)
    ax.plot([right_wire_x_end, right_wire_x_end + end_slant_size],
            [y - v_offset, y + end_slant_size - v_offset], color='black', linewidth=1.4)
    if anchor_to_v_tip:
        ax.plot([left_notch[0], v_tip[0]], [left_notch[1], v_tip[1]], linewidth=1.4, color='black')
        ax.plot([right_notch[0], v_tip[0]], [right_notch[1], v_tip[1]], linewidth=1.4, color='black')
    else:
        ax.plot([left_wire_x_start, left_notch[0]], [y - v_offset, left_notch[1]], linewidth=1.4, color='black')
        ax.plot([right_notch[0], right_wire_x_end], [right_notch[1], y - v_offset], linewidth=1.4, color='black')
    ax.plot([left_notch[0], notch_bottom[0]], [left_notch[1], notch_bottom[1]], color='black', linewidth=1.4)
    ax.plot([right_notch[0], notch_bottom[0]], [right_notch[1], notch_bottom[1]], color='black', linewidth=1.4)
    """
 
    text_y_offset = 0.15 * scale
    text_y = notch_bottom[1] - text_y_offset
    display_text = str(text)
    if len(display_text) >= 8:
        words = display_text.split()
        if len(words) > 1:
            line1 = " ".join(words[:-1])
            line2 = words[-1]
            return f"{line1}\n{line2}"
        else:
            return display_text[:8] + "\n" + display_text[8:]
    ax.text(center, text_y, display_text,
            ha='center', va='top',
            fontsize=int(39 * scale),
            fontname='Arial', linespacing=1.2)
    # === COMMENTED OUT: Small vertical lines for "LIGHT" ===
    """
    # Add two small vertical lines on upper side if text is "LIGHT"
    if str(text).strip().upper() == "LIGHT":
        line_length = 0.3 * scale
        gap = 0.25 * scale # increase gap between lines
        left_line_x = center - gap
        right_line_x = center + gap
        base_y = y - v_offset + 0.1 * scale # move lines slightly upward
        # Left vertical line
        ax.plot([left_line_x, left_line_x], [base_y, base_y + line_length], color='black', linewidth=1.4)
        # Right vertical line
        ax.plot([right_line_x, right_line_x], [base_y, base_y + line_length], color='black', linewidth=1.4)
    """

def draw_group_top_symbol(
    ax,
    x_start,
    x_end,
    y,
    texts='R1',
    scale=1.0,
    input_connected='N',
    spacing=0.3,
    x_offset=0.3,
    diagonal_length=0.21,
    split_text=True,
    split_length=4,
    draw_diagonal=True,
    draw_vertical=True,
    vertical_linewidth=1.2,
    diagonal_linewidth=1.2,
    text_start_offset=-0.3,
    line_spacing=0.25,
    max_chars_per_line=3,
    auto_wrap=True,
    min_font_scale=0.8
):
    """
    Draw group symbol at top with text in REVERSE order:
    - Last text in list is closest to symbol
    - First text in list is farthest from symbol
    - Text with 4+ characters placed at BOTTOM of text stack
    """
    if isinstance(texts, str):
        texts = [texts]
    
    # Reverse the text order so last item is closest to symbol
    texts = texts[::-1]
    
    # Separate texts into short (<4 chars) and long (>=4 chars)
    short_texts = []
    long_texts = []
    
    for text in texts:
        if len(str(text).strip()) < 4:
            short_texts.append(text)
        else:
            long_texts.append(text)
    
    # Recombine: short texts first (closest to symbol), then long texts (farthest from symbol)
    texts = short_texts + long_texts
    
    x = (x_start + x_end) / 2.0 if abs(x_end - x_start) < 0.1 else x_start
    line_extension = 0.35 * scale
    relay_gap = 0.1 * scale if str(input_connected).strip().upper() == 'Y' else 0.0
    base_y = y + relay_gap
    num_texts = len(texts)
    spacing = spacing * scale
    total_extra = (num_texts - 1) * spacing
    line_extension += total_extra
    
    # Vertical center line (going up)
    ax.plot([x, x], [base_y, base_y + line_extension], color='black', linewidth=1)
    
    # One \______ style segment at the top (diagonal first, horizontal after)
    total_width = x_end - x_start
    horizontal_ratio = 0.85
    y_top = base_y + 0.11 * scale
    drop_height = 0.08 * scale
    y_bottom = y_top - drop_height
    seg_x0 = x_start
    seg_x2 = x_end
    seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio)
    
    # Diagonal down (\ shape) first
    ax.plot([seg_x1, seg_x0], [y_top, y_bottom], color='black', linewidth=1)
    
    # Horizontal line after diagonal (slightly lower)
    horizontal_offset = 0.08 * scale
    ax.plot([seg_x1, seg_x2], [y_bottom + horizontal_offset, y_bottom + horizontal_offset],
            color='black', linewidth=1)
    
    # Stack diagonal bars with small vertical/diagonal connectors if multiple texts
    diagonal_length = diagonal_length * scale
    y_shift = -0.17 * scale
    diag_offset = -0.05 * scale
    left_adjust = -0.04 * scale
    right_adjust = 0.04 * scale
    down_shift = -0.01 * scale
    prev_center_y = None
    x_offset = x_offset * scale
    
    for i in range(num_texts):
        extra_shift = i * spacing
        # Apply x_offset for duplicates (i > 0)
        current_x = x + x_offset if i > 0 else x
        left_y = base_y + line_extension - diagonal_length - y_shift + left_adjust + diag_offset + down_shift + extra_shift
        right_y = base_y + line_extension - y_shift + right_adjust + diag_offset + down_shift + extra_shift
        
        # --- Original diagonal bar (/) ---
        if draw_diagonal:
            ax.plot([current_x - diagonal_length / 2, current_x + diagonal_length / 2],
                    [left_y, right_y],
                    color='black', linewidth=1)
            
            # --- Updated: stacked diagonal connector between symbols ---
            if prev_center_y is not None:
                vertical_shift = -1.3 * scale
                x_shift = -0.19 * scale
                length_extension = 1.8
                # center of current diagonal
                center_diag_y = (left_y + right_y) / 2.0
                # compute endpoints for \ shape with left shift and extended right side
                left_x = current_x - diagonal_length / 2 + x_shift
                right_x = current_x + (diagonal_length / 2 * length_extension) + x_shift
                left_y_shifted = prev_center_y + vertical_shift
                right_y_shifted = center_diag_y + vertical_shift
                # Draw the diagonal connector
                ax.plot([left_x, right_x], [left_y_shifted, right_y_shifted],
                        color='black', linewidth=1)
                # --- Small vertical line at the bottom-right of diagonal (/), extending upward
                small_vert_length_bottom = 1 * scale
                ax.plot([right_x, right_x],
                        [right_y_shifted, right_y_shifted + small_vert_length_bottom],
                        color='black', linewidth=1)
        
        center_y = (left_y + right_y) / 2.0
        
        # --- Vertical connector (|) - EXTENDED DOWNWARD ---
        if draw_vertical and prev_center_y is not None:
            ax.plot([current_x, current_x], [prev_center_y, center_y],
                    color='black', linewidth=1)
        
        prev_center_y = center_y
        
        # --- IMPROVED Text label handling with different positioning ---
        display_text = str(texts[i]).strip()
        
        # Calculate font size with scaling for long text
        base_fontsize = 21 * scale
        if len(display_text) > 8:
            # Scale down font size for very long text
            length_factor = min(1.0, 8.0 / len(display_text))
            fontsize = int(base_fontsize * max(length_factor, min_font_scale))
        else:
            fontsize = int(base_fontsize)
        
        # --- DETERMINE POSITION BASED ON TEXT LENGTH ---
        if len(display_text) < 4:
            # Less than 4 characters: Just a little above the symbol
            text_y = right_y + (0.18 * scale)  # Small offset
            multi_line_y_offset = -0.9  # Tight spacing for multi-line short text
        else:
            # 4 or more characters: Much higher above (at BOTTOM of text stack)
            text_y = right_y + (0.45 * scale)  # Larger offset
            multi_line_y_offset = -1.2  # More spacing for longer text
        
        # Determine if text needs to be split/wrapped
        if auto_wrap and len(display_text) > max_chars_per_line:
            # Split text into lines
            lines = []
            
            # Try to split at natural points first
            if '-' in display_text and len(display_text) > 4:
                # For text with hyphens like "R1-R2"
                parts = display_text.split('-')
                if len(parts) == 2 and len(parts[0]) <= max_chars_per_line and len(parts[1]) <= max_chars_per_line:
                    lines = parts
                else:
                    # Manual splitting
                    for j in range(0, len(display_text), max_chars_per_line):
                        lines.append(display_text[j:j+max_chars_per_line])
            elif len(display_text) <= max_chars_per_line * 2:
                # For moderately long text, split in middle
                mid = len(display_text) // 2
                lines = [display_text[:mid], display_text[mid:]]
            else:
                # For very long text, split into multiple lines
                for j in range(0, len(display_text), max_chars_per_line):
                    lines.append(display_text[j:j+max_chars_per_line])
            
            num_lines = len(lines)
            
            # Calculate text block height based on text length
            line_height = line_spacing * scale
            
            # Adjust spacing multiplier based on text length
            spacing_multiplier = 1.0
            if len(display_text) > 8:
                spacing_multiplier = 1.5  # More spacing for very long text
            
            # Draw each line
            for line_idx, line_text in enumerate(lines):
                # Position lines: for long text, they go DOWN from the starting position
                if len(display_text) < 4:
                    # Short text: lines go UP (negative direction)
                    line_y = text_y + (line_idx * line_height * multi_line_y_offset)
                else:
                    # Long text: lines go DOWN (positive direction)
                    line_y = text_y + (line_idx * line_height * spacing_multiplier)
                
                ax.text(current_x, line_y, line_text, 
                       ha='center', va='bottom',
                       fontsize=fontsize, 
                       fontname='Arial',
                       bbox=dict(boxstyle="round,pad=0.1", 
                                facecolor='white', 
                                edgecolor='none', 
                                alpha=0.7))
        else:
            # Single line text
            # Additional adjustment for longer single-line text
            if len(display_text) >= 4:
                adjusted_text_y = text_y + (0.05 * scale)  # Slightly higher for long text
            else:
                adjusted_text_y = text_y
            
            ax.text(current_x, adjusted_text_y, display_text, 
                   ha='center', va='bottom',
                   fontsize=fontsize, 
                   fontname='Arial',
                   bbox=dict(boxstyle="round,pad=0.1", 
                            facecolor='white', 
                            edgecolor='none', 
                            alpha=0.7))


def draw_group_bottom_symbol(
    ax,
    x_start,
    x_end,
    y,
    texts='R1',
    scale=1.0,
    output_connected='N',
    choke_output_terminal=None,
    spacing=0.3,
    x_offset=0.3,
    diagonal_length=0.21,
    split_text=True,
    split_length=3,
    draw_diagonal=True,
    draw_vertical=True,
    vertical_linewidth=1.2,
    diagonal_linewidth=1.2,
    text_start_offset=0.2,  # Added for consistency
    line_spacing=0.25,  # Added for consistency
    max_chars_per_line=3,  # Default max characters per line
    auto_wrap=True,  # Automatically wrap text
    min_font_scale=0.8  # Minimum font scale factor for long text
):
    """
    Draw group symbol at bottom with improved text handling
    """
    if isinstance(texts, str):
        texts = [texts]
    
    x = (x_start + x_end) / 2.0 if abs(x_end - x_start) < 0.1 else x_start
    line_extension = 0.35 * scale
    relay_gap = 0.1 * scale if str(output_connected).strip().upper() == 'Y' else 0.0
    base_y = y - relay_gap
    num_texts = len(texts)
    spacing = spacing * scale
    total_extra = (num_texts - 1) * spacing
    line_extension += total_extra
    
    if choke_output_terminal is None:
        ax.plot([x, x], [base_y, base_y - line_extension], color='black', linewidth=1)
    
    # One /??? style segment at the bottom (diagonal first, horizontal after)
    total_width = x_end - x_start
    horizontal_ratio = 0.85
    y_bottom = base_y - 0.11 * scale
    rise_height = 0.08 * scale
    y_top = y_bottom + rise_height
    seg_x0 = x_start
    seg_x2 = x_end
    seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio)
    
    # Diagonal up (/ shape) first
    ax.plot([seg_x1, seg_x0], [y_bottom, y_top], color='black', linewidth=1)
    
    # Horizontal line after diagonal (slightly higher for reverse)
    horizontal_offset = -0.08 * scale
    ax.plot([seg_x1, seg_x2], [y_top + horizontal_offset, y_top + horizontal_offset],
            color='black', linewidth=1)
    
    if choke_output_terminal is None:
        # Stack diagonal bars with small vertical/diagonal connectors if multiple texts
        diagonal_length = diagonal_length * scale
        y_shift = 0.17 * scale
        diag_offset = 0.05 * scale
        left_adjust = 0.04 * scale
        right_adjust = -0.04 * scale
        down_shift = 0.01 * scale
        prev_center_y = None
        x_offset = x_offset * scale
        
        for i in range(num_texts):
            extra_shift = -i * spacing
            # Apply x_offset for duplicates (i > 0)
            current_x = x + x_offset if i > 0 else x
            left_y = base_y - line_extension + y_shift + left_adjust + diag_offset + down_shift + extra_shift
            right_y = base_y - line_extension + y_shift + right_adjust + diag_offset + down_shift + extra_shift
            
            # --- Original diagonal bar (\) moved slightly lower ---
            y_offset = 0.235 * scale  # Adjust this value as needed
            
            if draw_diagonal:
                ax.plot(
                    [current_x - diagonal_length / 2, current_x + diagonal_length / 2],
                    [left_y - y_offset, right_y - y_offset],
                    color='black', linewidth=1
                )
                
                # --- Updated: stacked diagonal connector between symbols (\) ---
                if prev_center_y is not None:
                    vertical_shift = 1.3 * scale
                    x_shift = -0.19 * scale
                    length_extension = 1.8
                    center_diag_y = (left_y + right_y) / 2.0
                    left_x = current_x - diagonal_length / 2 + x_shift
                    right_x = current_x + (diagonal_length / 2 * length_extension) + x_shift
                    
                    # Swap y-values for backslash (\)
                    left_y_shifted = prev_center_y + vertical_shift
                    right_y_shifted = center_diag_y + vertical_shift
                    
                    # Draw the diagonal connector (\)
                    y_offset = 0.25 * scale
                    ax.plot(
                        [left_x, right_x],
                        [left_y_shifted - y_offset, right_y_shifted - y_offset],
                        color='black',
                        linewidth=1
                    )
                    
                    # Small vertical line at the top-right of diagonal (\), extending downward
                    small_vert_length_bottom = -1.28 * scale
                    ax.plot(
                        [right_x, right_x],
                        [right_y_shifted - y_offset, right_y_shifted + small_vert_length_bottom - y_offset],
                        color='black', linewidth=1
                    )
            
            center_y = (left_y + right_y) / 2.0
            
            # --- Vertical connector (|) - EXTENDED UPWARD ---
            if draw_vertical and prev_center_y is not None:
                ax.plot([current_x, current_x], [prev_center_y, center_y],
                        color='black', linewidth=1)
            
            prev_center_y = center_y
            
            # --- IMPROVED Text label handling ---
            display_text = str(texts[i]).strip()
            
            # Calculate font size with scaling for long text
            base_fontsize = 21 * scale
            if len(display_text) > 8:
                length_factor = min(1.0, 8.0 / len(display_text))
                fontsize = int(base_fontsize * max(length_factor, min_font_scale))
            else:
                fontsize = int(base_fontsize)
            
            # Determine text positioning
            if auto_wrap and len(display_text) > max_chars_per_line:
                # Split text into lines
                lines = []
                
                # Try to split at natural points first
                if '-' in display_text and len(display_text) > 4:
                    parts = display_text.split('-')
                    if len(parts) == 2 and len(parts[0]) <= max_chars_per_line and len(parts[1]) <= max_chars_per_line:
                        lines = parts
                    else:
                        for j in range(0, len(display_text), max_chars_per_line):
                            lines.append(display_text[j:j+max_chars_per_line])
                elif len(display_text) <= max_chars_per_line * 2:
                    mid = len(display_text) // 2
                    lines = [display_text[:mid], display_text[mid:]]
                else:
                    for j in range(0, len(display_text), max_chars_per_line):
                        lines.append(display_text[j:j+max_chars_per_line])
                
                num_lines = len(lines)
                line_height = line_spacing * scale
                
                # Position for text (bottom symbol)
                text_offset = 0.2 * scale
                if num_lines > 2:
                    text_offset += 0.05 * scale * (num_lines - 1)
                
                first_line_y = base_y - line_extension - 0.1 + y_shift - text_offset + extra_shift
                
                # Draw each line
                for line_idx, line_text in enumerate(lines):
                    line_y = first_line_y - (line_idx * line_height)
                    ax.text(current_x, line_y, line_text, 
                           ha='center', va='top',
                           fontsize=fontsize, 
                           fontname='Arial',
                           bbox=dict(boxstyle="round,pad=0.1", 
                                    facecolor='white', 
                                    edgecolor='none', 
                                    alpha=0.7))
            else:
                # Single line text
                text_offset = 0.2 * scale
                text_y = base_y - line_extension - 0.1 + y_shift - text_offset + extra_shift
                
                # Adjust position if text is long
                if len(display_text) > 4:
                    text_y += 0.05 * scale  # Move slightly higher
                
                ax.text(current_x, text_y, display_text, 
                       ha='center', va='top',
                       fontsize=fontsize, 
                       fontname='Arial',
                       bbox=dict(boxstyle="round,pad=0.1", 
                                facecolor='white', 
                                edgecolor='none', 
                                alpha=0.7))
                                
# === New Relay Box Functions ===
def draw_relay_box_top(ax, x_start, x_end, y, texts='R1', scale=1.0, input_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay box at top with rectangle instead of diagonal bar
    """
    if isinstance(texts, str):
        texts = [texts]
    x = (x_start + x_end) / 2.0 if abs(x_end - x_start) < 0.1 else x_start
    line_extension = 0.35 * scale
    relay_gap = 0.1 * scale if str(input_connected).strip().upper() == 'Y' else 0.0
    base_y = y + relay_gap
    num_texts = len(texts)
    spacing = spacing * scale
    total_extra = (num_texts - 1) * spacing
    line_extension += total_extra
 
    # Extend line upward ONLY
    extra_vertical_extension = 0.16 * scale # <<---- CHANGE THIS TO EXTEND MORE
    # Vertical center line (extended upward)
    ax.plot([x, x],
            [base_y, base_y + line_extension + extra_vertical_extension],
            color='black', linewidth=1)
 
    # One \______ style segment at the top (diagonal first, horizontal after)
    total_width = x_end - x_start
    horizontal_ratio = 0.85
    y_top = base_y + 0.11 * scale
    drop_height = 0.08 * scale
    y_bottom = y_top - drop_height
    seg_x0 = x_start
    seg_x2 = x_end
    seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio) # end of diagonal
 
    # Diagonal down (\ shape) first
    ax.plot([seg_x1, seg_x0], [y_top, y_bottom], color='black', linewidth=1)
 
    # Horizontal line after diagonal (slightly lower)
    horizontal_offset = 0.08 * scale
    ax.plot([seg_x1, seg_x2], [y_bottom + horizontal_offset, y_bottom + horizontal_offset],
            color='black', linewidth=1)
 
    # Stack rectangles with small vertical connectors if multiple texts
    rect_width = 0.45 * scale
    rect_height = 0.35 * scale
    y_shift = -0.17 * scale
    rect_offset = -0.05 * scale
    down_shift = 0.39 * scale
    prev_center_y = None
    x_offset = x_offset * scale
 
    for i in range(num_texts):
        extra_shift = i * spacing
        # Apply x_offset for duplicates (i > 0)
        current_x = x + x_offset if i > 0 else x
     
        # Rectangle position (centered at current_x)
        rect_y = base_y + line_extension - rect_height/2 - y_shift + rect_offset + down_shift + extra_shift
     
        # Draw rectangle
        rect = Rectangle((current_x - rect_width/2, rect_y - rect_height/2),
                        rect_width, rect_height,
                        linewidth=1.2, edgecolor='black', facecolor='white')
        ax.add_patch(rect)
     
        # --- Updated: stacked vertical connector between symbols ---
        if prev_center_y is not None:
            vertical_shift = -1.3 * scale
            x_shift = -0.19 * scale
         
            # Center of current rectangle
            center_rect_y = rect_y
         
            # Draw vertical connector between rectangles
            ax.plot([current_x, current_x], [prev_center_y, center_rect_y],
                    color='black', linewidth=1)
         
            # Small vertical line at the bottom of the vertical connector
            small_vert_length = 0.3 * scale
            ax.plot([current_x, current_x],
                    [center_rect_y, center_rect_y - small_vert_length],
                    color='black', linewidth=1)
     
        prev_center_y = rect_y
     
        # --- Text label inside rectangle ---
        text_y = rect_y
        display_text = str(texts[i]).strip()
        # Split text into two lines if longer than 3 characters
        if len(display_text) > 3:
            mid = len(display_text) // 2
            display_text = display_text[:mid] + '\n' + display_text[mid:]
     
        ax.text(current_x, text_y, display_text, ha='center', va='center',
                fontsize=int(16 * scale), fontname='Arial', linespacing=0.8)

def draw_relay_box_bottom(ax, x_start, x_end, y, texts='R1', scale=1.0,
                          output_connected='N', spacing=0.3, x_offset=0.3):
    if isinstance(texts, str):
        texts = [texts]
    x = (x_start + x_end) / 2.0 if abs(x_end - x_start) < 0.1 else x_start
    line_extension = 0.35 * scale
    relay_gap = 0.1 * scale if str(output_connected).strip().upper() == 'Y' else 0.0
    base_y = y - relay_gap
    num_texts = len(texts)
    spacing = spacing * scale
    total_extra = (num_texts - 1) * spacing
    line_extension += total_extra
    center_extend = 0.16 * scale # you can change 0.25 ? longer or shorter
    # Main center vertical line
    ax.plot(
        [x, x],
        [base_y, base_y - line_extension - center_extend], # extended downward
        color='black', linewidth=1
    )
    # Bottom slanted + horizontal lines
    total_width = x_end - x_start
    horizontal_ratio = 0.85
    y_bottom = base_y - 0.11 * scale
    rise_height = 0.08 * scale
    y_top = y_bottom + rise_height
    seg_x0 = x_start
    seg_x2 = x_end
    seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio)
    ax.plot([seg_x1, seg_x0], [y_bottom, y_top], color='black', linewidth=1)
    ax.plot([seg_x1, seg_x2],
            [y_top - 0.08 * scale, y_top - 0.08 * scale],
            color='black', linewidth=1)
    # Rectangle settings
    rect_width = 0.45 * scale
    rect_height = 0.35 * scale
    y_shift = 0.17 * scale
    rect_offset = 0.05 * scale
    down_shift = -0.18 * scale
    # YOU used this for rectangle shift � keeping as-is
    down_adjust = 0.55 * scale
    prev_center_y = None
    x_offset = x_offset * scale
    # ONLY vertical line extra extension downwards:
    extra_vertical_extension = 0.25 * scale
    for i in range(num_texts):
        extra_shift = -i * spacing
        current_x = x + x_offset if i > 0 else x
        # Rectangle Y (already shifted by your down_adjust)
        rect_y = (base_y - line_extension + rect_height/2 +
                  y_shift + rect_offset + down_shift + extra_shift)
        rect_y -= down_adjust # you already applied rectangle move
        # Draw rectangle
        rect = Rectangle(
            (current_x - rect_width/2, rect_y - rect_height/2),
            rect_width, rect_height,
            linewidth=1.2, edgecolor='black', facecolor='white'
        )
        ax.add_patch(rect)
        # Vertical connectors (extended downward ONLY)
        if prev_center_y is not None:
            # extended vertical drop (downwards only)
            ax.plot(
                [current_x, current_x],
                [prev_center_y - extra_vertical_extension,
                 rect_y],
                color='black',
                linewidth=1
            )
            # Small top vertical line (just shifted down, not extended)
            small_vert_length = 0.3 * scale
            ax.plot(
                [current_x, current_x],
                [rect_y, rect_y + small_vert_length],
                color='black',
                linewidth=1
            )
        prev_center_y = rect_y
        # Text inside rectangle
        display_text = str(texts[i]).strip()
        if len(display_text) > 3:
            mid = len(display_text) // 2
            display_text = display_text[:mid] + "\n" + display_text[mid:]
        ax.text(current_x, rect_y, display_text,
                ha='center', va='center',
                fontsize=int(16 * scale), fontname='Arial', linespacing=0.8)



# === Helper: Find row by terminal number ===
def find_row_by_term(term, df):
    """
    Updated to accept df as a parameter to avoid scope issues
    """
    if pd.isna(term):
        return None
    s = str(term).strip()
    if s.endswith('.0'):
        s = s[:-2]
    col = df['terminal_no'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    matches = df[col == s]
    return matches.iloc[0] if not matches.empty else None

# === New Relay Contact Box Functions ===
def draw_relay_contact_box_top(ax, x_start, x_end, y, texts='R1', scale=1.0, input_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay contact box at top - rectangle above terminals, no horizontal line.
    """
    if isinstance(texts, str):
        texts = [texts]
    # Rectangle dimensions
    rect_width = (0.45 * 2) * scale # double width
    rect_height = 0.35 * scale
    # Move rectangle slightly DOWN (closer to terminal)
    vertical_offset = 0.32 * scale # <---- reduced from 0.35 ? moves downward
    # Rectangle position (centered between terminals, above them)
    rect_x = (x_start + x_end) / 2 - rect_width / 2
    rect_y = y + vertical_offset
    # Draw rectangle
    rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                     linewidth=1.2, edgecolor='black', facecolor='white')
    ax.add_patch(rect)
    # Display text inside rectangle
    display_text = str(texts[0]).strip() if texts else ''
    if len(display_text) > 3:
        mid = len(display_text) // 2
        display_text = display_text[:mid] + '\n' + display_text[mid:]
    ax.text((x_start + x_end) / 2, rect_y + rect_height / 2,
            display_text,
            ha='center', va='center',
            fontsize=int(16 * scale), fontname='Arial', linespacing=0.8)

def draw_relay_contact_box_bottom(ax, x_start, x_end, y, texts='R1', scale=1.0,
                                  output_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay contact box at bottom - rectangle below terminals, no horizontal line.
    Rectangle width automatically adjusts based on terminal spacing.
    """
    if isinstance(texts, str):
        texts = [texts]
    # Calculate terminal gap from XLSX values
    terminal_gap = abs(x_end - x_start)
    # Auto dynamic width
    base_width = 0.90 * scale # width when gap = 1
    extra_width = 0.60 * scale # increase per extra gap
    rect_width = base_width + (terminal_gap - 1) * extra_width
    rect_height = 0.35 * scale
    # Vertical offset (downwards)
    vertical_offset = 0.31 * scale
    # Rectangle position (centered, below terminals)
    rect_x = (x_start + x_end) / 2 - rect_width / 2
    rect_y = y - rect_height - vertical_offset
    # Draw rectangle
    rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                     linewidth=1.2, edgecolor='black', facecolor='white')
    ax.add_patch(rect)
    # Text inside rectangle
    display_text = str(texts[0]).strip() if texts else ''
    if len(display_text) > 3:
        mid = len(display_text) // 2
        display_text = display_text[:mid] + '\n' + display_text[mid:]
    ax.text((x_start + x_end) / 2,
            rect_y + rect_height / 2,
            display_text,
            ha='center', va='center',
            fontsize=int(16 * scale),
            fontname='Arial',
            linespacing=0.8)

def draw_header(ax, cable_id, header_type, x_start, x_end, text, min_symbol_bottom=None,
                first_hook_x=None, last_hook_x=None, y_top_bus_group=0, y_bottom_bus_group=0, special_ha=False):
    """
    Updated to handle text splitting based on gap between terminal start and end
    Splits text into two lines only when terminals are close together (gap < 3.0)
    """
    if pd.isna(text) or str(text).strip() == '':
        return
    
    text = str(text).strip()
    text_length = len(text)
    
    # Calculate the gap between terminals
    gap = x_end - x_start
    
    if str(header_type).strip().upper() == 'WIREFROM':
        x_pos = first_hook_x if first_hook_x is not None else x_start - 0.05
        
        # Split text only if terminals are close together (gap < 3.0)
        if gap < 3.0 and text_length > 5:
            # Try to split at '/' for better formatting (e.g., "Pt. 121/122 nwr")
            if '/' in text:
                # Split at slash for better formatting
                parts = text.split('/', 1)
                if len(parts) == 2:
                    line1 = parts[0].strip()
                    line2 = '/' + parts[1].strip()
                else:
                    # Fallback: split at space
                    words = text.split()
                    if len(words) >= 2:
                        line1 = ' '.join(words[:2])
                        line2 = ' '.join(words[2:])
                    else:
                        # Just split in middle
                        mid = len(text) // 2
                        line1 = text[:mid]
                        line2 = text[mid:]
            else:
                # Split at space if no slash
                words = text.split()
                if len(words) >= 2:
                    line1 = ' '.join(words[:2])
                    line2 = ' '.join(words[2:])
                else:
                    # Just split in middle
                    mid = len(text) // 2
                    line1 = text[:mid]
                    line2 = text[mid:]
            
            # Position for two lines - slightly higher
            y_pos1 = y_top_bus_group + 0.5  # First line position
            y_pos2 = y_top_bus_group + 0.2  # Second line position (below first)
            
            # Draw first line
            ax.text(x_pos, y_pos1, line1, ha='left', va='bottom', fontsize=21, 
                   fontname='Arial')
            
            # Draw second line
            ax.text(x_pos, y_pos2, line2, ha='left', va='bottom', fontsize=21,
                   fontname='Arial')
        else:
            # Original positioning for normal gap or short text
            y_pos = y_top_bus_group + 0.2
            ax.text(x_pos, y_pos, text, ha='left', va='bottom', fontsize=21, fontname='Arial')
            
    elif str(header_type).strip().upper() == 'WIRETO':
        ha = 'center'
        x_pos = last_hook_x if last_hook_x is not None else (x_start + x_end) / 2.0
        if last_hook_x is not None:
            ha = 'left' if special_ha else 'right'
        
        min_symbol_bottom = y_bottom_bus_group - 0.2 if min_symbol_bottom is None else min_symbol_bottom
        
        # Split text only if terminals are close together (gap < 3.0)
        if gap < 3.0 and text_length > 5:
            # Try to split at '/' for better formatting
            if '/' in text:
                # Split at slash for better formatting
                parts = text.split('/', 1)
                if len(parts) == 2:
                    line1 = parts[0].strip()
                    line2 = '/' + parts[1].strip()
                else:
                    # Fallback: split at space
                    words = text.split()
                    if len(words) >= 2:
                        line1 = ' '.join(words[:2])
                        line2 = ' '.join(words[2:])
                    else:
                        # Just split in middle
                        mid = len(text) // 2
                        line1 = text[:mid]
                        line2 = text[mid:]
            else:
                # Split at space if no slash
                words = text.split()
                if len(words) >= 2:
                    line1 = ' '.join(words[:2])
                    line2 = ' '.join(words[2:])
                else:
                    # Just split in middle
                    mid = len(text) // 2
                    line1 = text[:mid]
                    line2 = text[mid:]
            
            # Position for two lines - slightly lower
            text_offset = -0.15
            y_pos1 = min_symbol_bottom - 1.0 + text_offset  # First line position
            y_pos2 = min_symbol_bottom - 1.35 + text_offset  # Second line position (above first)
            
            if special_ha:
                x_pos += 0.3  # Adjust for special alignment
            
            # Draw first line
            ax.text(x_pos, y_pos1, line1, ha=ha, va='top', fontsize=21, 
                   fontname='Arial')
            
            # Draw second line
            ax.text(x_pos, y_pos2, line2, ha=ha, va='top', fontsize=21,
                   fontname='Arial')
        else:
            # Original positioning for normal gap or short text
            text_offset = -0.15
            y_pos = min_symbol_bottom - 1.0 + text_offset
            ax.text(x_pos, y_pos, text, ha=ha, va='top', fontsize=21, fontname='Arial')

# === Symbol Drawers ===
def draw_capsule(ax, x, y_center, terminal_no, input_left, input_right, output_left, output_right,
                 input_connected, output_connected, capsule_type='capsule'):
    """
    Draw capsule with support for different types: 'Ara', 'Wago', 'Ara/Wago'
    All types produce the same output as the original capsule
    """
    capsule_bottom = y_center - SYMBOL_HEIGHT / 2
    capsule_top = capsule_bottom + SYMBOL_HEIGHT
    top_circle_radius = SYMBOL_RADIUS * 0.8
    ax.add_patch(Circle((x, capsule_top), radius=top_circle_radius,
                        edgecolor='black', facecolor='white', linewidth=1))
    bottom_circle_radius = SYMBOL_RADIUS * 0.8
    ax.add_patch(Circle((x, capsule_bottom), radius=bottom_circle_radius,
                        edgecolor='black', facecolor='white', linewidth=1))
    line_offset = SYMBOL_WIDTH / 2
    extend = 0.11
    shift_left = 0.055
    shift_right = 0.055
    ax.plot([x - line_offset + shift_left, x - line_offset + shift_left],
            [capsule_bottom + bottom_circle_radius - extend, capsule_top - SYMBOL_RADIUS + extend],
            color='black', linewidth=1)
    ax.plot([x + line_offset - shift_right, x + line_offset - shift_right],
            [capsule_bottom + bottom_circle_radius - extend, capsule_top - SYMBOL_RADIUS + extend],
            color='black', linewidth=1)
    if pd.notna(terminal_no) and str(terminal_no).strip() != '':
        term_str = str(terminal_no)
        if term_str.endswith('.0'):
            term_str = term_str[:-2]
        ax.text(x, y_center, term_str, fontsize=17, ha='center', va='center', fontname='Arial')
    def format_text(t):
        t = str(t)
        if len(t) >= 8:
            words = t.split()
            if len(words) > 1:
                line1 = " ".join(words[:-1])
                line2 = words[-1]
                return f"{line1}\n{line2}"
            else:
                return t[:7] + "\n" + t[8:]
        return t
    input_left_offset = 0.005
    if pd.notna(input_left) and str(input_left).strip() != "":
        ax.text(x - input_left_offset, capsule_top + 0.18, format_text(input_left),
                fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    input_right_offset = 0.05
    if pd.notna(input_right) and str(input_right).strip() != "":
        ax.text(x + input_right_offset, capsule_top + 0.18, format_text(input_right),
                fontsize=19, ha='left', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    output_left_offset = 0.005
    if pd.notna(output_left) and str(output_left).strip() != "":
        ax.text(x - output_left_offset, capsule_bottom - 0.15, format_text(output_left),
                fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    output_right_offset = 0.05
    if pd.notna(output_right) and str(output_right).strip() != "":
        ax.text(x + output_right_offset, capsule_bottom - 0.18, format_text(output_right),
                fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    top_conn = (x, capsule_top + SYMBOL_RADIUS)
    bottom_conn = (x, capsule_bottom - bottom_circle_radius)
    ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
    oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
    return top_conn, bottom_conn, ic, oc

def draw_s_fuse(ax, x, y_center, terminal_no,
                input_left=None, input_right=None,
                output_left=None, output_right=None,
                input_connected='N', output_connected='N'):
    fuse_top = y_center + SYMBOL_HEIGHT / 2
    fuse_bottom = y_center - SYMBOL_HEIGHT / 2
    top_circle_radius = SYMBOL_RADIUS * 0.8
    bottom_circle_radius = SYMBOL_RADIUS * 0.8
    # Draw circles
    ax.add_patch(Circle((x, fuse_top), top_circle_radius,
                        edgecolor='black', facecolor='white', linewidth=1))
    ax.add_patch(Circle((x, fuse_bottom), bottom_circle_radius,
                        edgecolor='black', facecolor='white', linewidth=1))
    # Curved middle connection
    start = (x, fuse_top - top_circle_radius)
    end = (x, fuse_bottom + bottom_circle_radius)
    ctrl1 = (x + SYMBOL_RADIUS * 2.2, y_center + SYMBOL_HEIGHT * 0.15)
    ctrl2 = (x - SYMBOL_RADIUS * 2.2, y_center - SYMBOL_HEIGHT * 0.15)
    t = np.linspace(0, 1, 100)
    xs = (1 - t)**3 * start[0] + 3 * (1 - t)**2 * t * ctrl1[0] + \
         3 * (1 - t) * t**2 * ctrl2[0] + t**3 * end[0]
    ys = (1 - t)**3 * start[1] + 3 * (1 - t)**2 * t * ctrl1[1] + \
         3 * (1 - t) * t**2 * ctrl2[1] + t**3 * end[1]
    ax.plot(xs, ys, color='black', linewidth=1, solid_capstyle='round')
    # Format function (same as capsule)
    def format_text(t):
        t = str(t)
        if len(t) >= 8:
            words = t.split()
            if len(words) > 1:
                line1 = " ".join(words[:-1])
                line2 = words[-1]
                return f"{line1}\n{line2}"
            else:
                return t[:8] + "\n" + t[8:]
        return t
    # Terminal name
    if pd.notna(terminal_no) and str(terminal_no).strip() != '':
        term_str = str(terminal_no)
        if term_str.endswith('.0'):
            term_str = term_str[:-2]
        ax.text(x - 0.1, y_center + 0.01, term_str,
                ha='center', va='center', fontsize=17, fontname='Arial')
    # Input/Output labels with auto split
    input_left_offset = 0.005
    if pd.notna(input_left) and str(input_left).strip() != "":
        ax.text(x - input_left_offset, fuse_top + 0.18,
                format_text(input_left), fontsize=19,
                ha='right', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    input_right_offset = 0.05
    if pd.notna(input_right) and str(input_right).strip() != "":
        ax.text(x + input_right_offset, fuse_top + 0.18,
                format_text(input_right), fontsize=19,
                ha='left', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    output_left_offset = 0.005
    if pd.notna(output_left) and str(output_left).strip() != "":
        ax.text(x - output_left_offset, fuse_bottom - 0.15,
                format_text(output_left), fontsize=19,
                ha='right', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    output_right_offset = 0.05
    if pd.notna(output_right) and str(output_right).strip() != "":
        ax.text(x + output_right_offset, fuse_bottom - 0.18,
                format_text(output_right), fontsize=19,
                ha='left', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    # Connections
    top_conn = (x, fuse_top + top_circle_radius)
    bottom_conn = (x, fuse_bottom - bottom_circle_radius)
    ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
    oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
    return top_conn, bottom_conn, ic, oc

def draw_choke(ax, x, y_center, terminal_no):
    row = find_row_by_term(terminal_no)
    input_left = input_right = output_left = output_right = None
    input_connected = 'N'
    output_connected = 'N'
    if row is not None:
        input_left = row.get('input_left')
        input_right = row.get('input_right')
        output_left = row.get('output_left')
        output_right = row.get('output_right')
        input_connected = row.get('input_connected', 'N')
        output_connected = row.get('output_connected', 'N')
    choke_top = y_center + SYMBOL_HEIGHT / 2
    choke_bottom = y_center - SYMBOL_HEIGHT / 2
    top_center = np.array([x, choke_top])
    bottom_center = np.array([x, choke_bottom])
    for center in [top_center, bottom_center]:
        circ = Circle(center, SYMBOL_RADIUS, edgecolor='black', facecolor='white', linewidth=1)
        ax.add_patch(circ)
    start_y = choke_top - SYMBOL_RADIUS
    end_y = choke_bottom + SYMBOL_RADIUS
    num_coils = 4
    t = np.linspace(0, num_coils * np.pi, 100)
    xs = x + SYMBOL_RADIUS * 1.5 * np.sin(t)
    ys = start_y - (start_y - end_y) * (t / (num_coils * np.pi))
    ax.plot(xs, ys, color='black', linewidth=1, solid_capstyle='round')
    if pd.notna(terminal_no) and str(terminal_no).strip() != '':
        term_str = str(terminal_no)
        if term_str.endswith('.0'):
            term_str = term_str[:-2]
        ax.text(x - 0.3, y_center + 0.1, term_str, ha='center', va='center', fontsize=12.5, fontname='Arial')
    label_offset = 0.2
    if pd.notna(input_left) and str(input_left).strip() != "":
        ax.text(x - label_offset, choke_top + 0.15, str(input_left), fontsize=10, ha='right', va='bottom', rotation=90, fontname='Arial')
    if pd.notna(input_right) and str(input_right).strip() != "":
        ax.text(x + label_offset, choke_top + 0.15, str(input_right), fontsize=10, ha='left', va='bottom', rotation=90, fontname='Arial')
    if pd.notna(output_left) and str(output_left).strip() != "":
        ax.text(x - label_offset, choke_bottom - 0.15, str(output_left), fontsize=10, ha='right', va='top', rotation=90, fontname='Arial')
    if pd.notna(output_right) and str(output_right).strip() != "":
        ax.text(x + label_offset, choke_bottom - 0.15, str(output_right), fontsize=10, ha='left', va='top', rotation=90, fontname='Arial')
    top_conn = (x, choke_top + SYMBOL_RADIUS)
    bottom_conn = (x, choke_bottom - SYMBOL_RADIUS)
    ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
    oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
    return top_conn, bottom_conn, ic, oc

# === UPDATED draw_horizontal_choke ===
def draw_horizontal_choke(ax, x_center, y_center, label='CHOKE',
                          box_width=0.6, box_height=0.3,
                          special_end=False, output_label='', output_type='terminal', output_text='',
                          output_connected=False, y_top_bus_group=None, y_bottom_bus_group=None):
    """
    Updated to handle output_connected parameter for output direction only.
    Choke position always below (lower side).
    Input (left) connection always to bottom bus/symbol side.
    Output (right) connection to top bus/symbol side if output_connected=True, else to bottom.
    """
    # Always place choke below
    y_shift = -0.5
    y_center = y_center + y_shift
 
    left_x = x_center - box_width / 2
    right_x = x_center + box_width / 2
    bottom_y = y_center - box_height / 2
 
    # Draw rounded box
    choke_box = FancyBboxPatch((left_x, bottom_y),
                               box_width, box_height,
                               boxstyle="round,pad=0.02",
                               edgecolor='black', facecolor='white', linewidth=1.5)
    ax.add_patch(choke_box)
 
    # Draw label
    ax.text(x_center, y_center, label,
            fontsize=16, ha='center', va='center', fontname='Arial')
 
    # Connection lines
    line_length = 0.075
    delta = 0.02
    vert_line_height_left = y_bottom_bus_group - y_center if y_bottom_bus_group is not None else 0.5 # Short up to bottom bus
 
    # Left connection (input): always to bottom side (short up)
    left_horiz_start = left_x - line_length - delta
    left_horiz_end = left_x - delta
    ax.plot([left_horiz_start, left_horiz_end],
            [y_center, y_center],
            color='black', linewidth=1)
 
    v_offset = 0.005
    ax.plot([left_horiz_start - v_offset, left_horiz_start - v_offset],
            [y_center, y_center + vert_line_height_left], # Up to bottom bus
            color='black', linewidth=1)
 
    if not special_end:
        # Right connection (output): non-special
        right_horiz_start = right_x + delta
        right_horiz_end = right_x + line_length + delta
        ax.plot([right_horiz_start, right_horiz_end],
                [y_center, y_center],
                color='black', linewidth=1)
     
        # Right vertical: up to top if connected, else short up to bottom
        if output_connected and y_top_bus_group is not None:
            vert_line_height_right = y_top_bus_group - y_center # Long up to top bus
        else:
            vert_line_height_right = vert_line_height_left # Short up to bottom bus
     
        ax.plot([right_horiz_end + v_offset, right_horiz_end + v_offset],
                [y_center, y_center + vert_line_height_right],
                color='black', linewidth=1)
    else:
        # Special end for output
        horiz_length = 0.5 if output_type.lower() == 'relay' else 0.2
        slant_size = 0.3 if output_type.lower() != 'relay' else 0.15
        vertical_length = 0.5 if output_type.lower() != 'relay' else 0 # No extra vertical for relay
        end_horiz_x = right_x + delta + horiz_length
        ax.plot([right_x + delta, end_horiz_x],
                [y_center, y_center],
                color='black', linewidth=1.4)
     
        if str(output_type).strip().lower() == 'relay':
            # For relay output type - horizontal + diagonal + text
            diag_length = slant_size # Reuse slant_size for diag_length
         
            if output_connected:
                # Upward / from low to high
                diag_start_x = end_horiz_x - 0.1
                diag_start_y = y_center
                diag_end_x = diag_start_x + diag_length
                diag_end_y = diag_start_y + diag_length
                ax.plot([diag_start_x, diag_end_x],
                        [diag_start_y, diag_end_y],
                        color='black', linewidth=1.4)
                # Text at end (higher)
                text_offset = 0.05
                ax.text(diag_end_x + text_offset, diag_end_y, output_text,
                        ha='left', va='center', fontsize=19, fontname='Arial')
            else:
                # Downward \ from high to low
                diag_start_x = end_horiz_x - 0.1
                diag_start_y = y_center + 0.1
                diag_end_x = diag_start_x + diag_length
                diag_end_y = diag_start_y - diag_length
                ax.plot([diag_start_x, diag_end_x],
                        [diag_start_y, diag_end_y],
                        color='black', linewidth=1.4)
                # Text at end (lower)
                text_offset = 0.05
                ax.text(diag_end_x + text_offset, diag_end_y - 0.1, output_text,
                        ha='left', va='center', fontsize=19, fontname='Arial')
         
            return end_horiz_x + diag_length
        else:
            # Terminal output type - horizontal + slant + vertical + label
            if output_connected:
                # Upward slant / + vertical up
                ax.plot([end_horiz_x, end_horiz_x + slant_size],
                        [y_center, y_center + slant_size],
                        color='black', linewidth=1.4)
                vert_top_y = y_center + slant_size
                vert_end_y = vert_top_y + vertical_length # Further up
                ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                        [vert_top_y, vert_end_y],
                        color='black', linewidth=1.4)
             
                # Label
                label_offset = 0.05
                label_y = (vert_top_y + vert_end_y) / 2
                ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                        ha='left', va='center', fontsize=19, rotation=90, fontname='Arial')
                return end_horiz_x + slant_size
            else:
                # Downward slant \ + vertical down (original)
                ax.plot([end_horiz_x, end_horiz_x + slant_size],
                        [y_center, y_center - slant_size],
                        color='black', linewidth=1.4)
                vert_top_y = y_center - slant_size
                vert_end_y = vert_top_y - vertical_length # Further down
                ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                        [vert_top_y, vert_end_y],
                        color='black', linewidth=1.4)
             
                # Label
                label_offset = 0.05
                label_y = (vert_top_y + vert_end_y) / 2
                ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                        ha='left', va='center', fontsize=19, rotation=90, fontname='Arial')
                return end_horiz_x + slant_size
    return None

def draw_dual_fuse(ax, x_left, y_center, left_term, right_term, left_input_left=None, left_input_right=None, left_output_left=None, left_output_right=None, left_input_connected='N', left_output_connected='N', right_input_left=None, right_input_right=None, right_output_left=None, right_output_right=None, right_input_connected='N', right_output_connected='N'):
    INNER_SPACING_MULT = 2.8
    inner_spacing = SYMBOL_WIDTH * INNER_SPACING_MULT
    x_right = x_left + inner_spacing
    def format_text(t):
        t = str(t)
        if len(t) >= 8:
            words = t.split()
            if len(words) > 1:
                line1 = " ".join(words[:-1])
                line2 = words[-1]
                return f"{line1}\n{line2}"
            else:
                return t[:8] + "\n" + t[8:]
        return t
    def _draw_one_s(ax, x_pos, y_c, term, input_left, input_right, output_left, output_right, input_conn, output_conn, term_shift=0.0):
        fuse_top = y_c + SYMBOL_HEIGHT / 2
        fuse_bottom = y_c - SYMBOL_HEIGHT / 2
        top_circle_radius = SYMBOL_RADIUS * 0.8
        bottom_circle_radius = SYMBOL_RADIUS * 0.8
        ax.add_patch(Circle((x_pos, fuse_top), top_circle_radius,
                            edgecolor='black', facecolor='white', linewidth=1))
        ax.add_patch(Circle((x_pos, fuse_bottom), bottom_circle_radius,
                            edgecolor='black', facecolor='white', linewidth=1))
        start = (x_pos, fuse_top - top_circle_radius)
        end = (x_pos, fuse_bottom + bottom_circle_radius)
        ctrl1 = (x_pos + SYMBOL_RADIUS * 2.2, y_c + SYMBOL_HEIGHT * 0.15)
        ctrl2 = (x_pos - SYMBOL_RADIUS * 2.2, y_c - SYMBOL_HEIGHT * 0.15)
        t = np.linspace(0, 1, 100)
        xs = (1 - t)**3 * start[0] + 3 * (1 - t)**2 * t * ctrl1[0] + \
         3 * (1 - t) * t**2 * ctrl2[0] + t**3 * end[0]
        ys = (1 - t)**3 * start[1] + 3 * (1 - t)**2 * t * ctrl1[1] + \
         3 * (1 - t) * t**2 * ctrl2[1] + t**3 * end[1]
        ax.plot(xs, ys, color='black', linewidth=1, solid_capstyle='round')
        if pd.notna(term) and str(term).strip() != '':
            term_str = str(term)
            if term_str.endswith('.0'):
                term_str = term_str[:-2]
            text_x = x_pos + term_shift
            ax.text(text_x, y_c, term_str,
                    ha='center', va='center', fontsize=17, fontname='Arial')
        label_offset = 0.11
        if pd.notna(input_left) and str(input_left).strip() != "":
            ax.text(x_pos - label_offset + 0.11, fuse_top + 0.24, format_text(input_left),
                    fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
        if pd.notna(input_right) and str(input_right).strip() != "":
            ax.text(x_pos + 0.1, fuse_top + 0.27, format_text(input_right),
                    fontsize=19, ha='center', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
        if pd.notna(output_left) and str(output_left).strip() != "":
            ax.text(x_pos - label_offset + 0.11, fuse_bottom - 0.30, format_text(output_left),
                    fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname='Arial')
        if pd.notna(output_right) and str(output_right).strip() != "":
            ax.text(x_pos + label_offset - 0.09, fuse_bottom - 0.28, format_text(output_right),
                    fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname='Arial')
        top_conn = (x_pos, fuse_top + top_circle_radius)
        bottom_conn = (x_pos, fuse_bottom - bottom_circle_radius)
        ic = 'Y' if str(input_conn).strip().upper() == 'Y' else 'N'
        oc = 'Y' if str(output_conn).strip().upper() == 'Y' else 'N'
        return top_conn, bottom_conn, ic, oc
    left_top, left_bottom, left_ic, left_oc = _draw_one_s(ax, x_left, y_center, left_term, left_input_left, left_input_right, left_output_left, left_output_right, left_input_connected, left_output_connected, term_shift=-0.1)
    right_top, right_bottom, right_ic, right_oc = _draw_one_s(ax, x_right, y_center, right_term, right_input_left, right_input_right, right_output_left, right_output_right, right_input_connected, right_output_connected, term_shift=-0.1)
    rail_extension = 0.15
    top_rail_y = max(left_top[1], right_top[1]) + rail_extension
    bottom_rail_y = min(left_bottom[1], right_bottom[1]) - rail_extension
    ax.plot([x_left, x_right], [top_rail_y, top_rail_y], linewidth=1, color='black')
    ax.plot([x_left, x_left], [left_top[1], top_rail_y], linewidth=1, color='black')
    ax.plot([x_right, x_right], [right_top[1], top_rail_y], linewidth=1, color='black')
    ax.plot([x_left, x_right], [bottom_rail_y, bottom_rail_y], linewidth=1, color='black')
    ax.plot([x_left, x_left], [bottom_rail_y, left_bottom[1]], linewidth=1, color='black')
    ax.plot([x_right, x_right], [bottom_rail_y, right_bottom[1]], linewidth=1, color='black')
    top_conn = (x_left, top_rail_y)
    bottom_conn = (x_left, bottom_rail_y)
    return top_conn, bottom_conn, left_ic, left_oc, right_ic, right_oc

def draw_resistor(ax, x, y_center, input_terminal='', output_terminal='', resistor_name='R', input_x_pos=None, output_x_pos=None):
    radius = SYMBOL_RADIUS * 1.5
    ax.add_patch(Circle((x, y_center), radius, edgecolor='black', facecolor='white', linewidth=1))
    ax.text(x, y_center, resistor_name, ha='center', va='center', fontsize=19, fontname='Arial')
    # Upper vertical line
    upper_y_start = y_center + radius
    upper_y_end = y_center + radius * 9.0
    ax.plot([x, x], [upper_y_start, upper_y_end], color='black', linewidth=1)
    # Upper horizontal lines (dynamic to multiple input_x_pos if provided) with vertical drop at start
    upper_labels = [label.strip() for label in str(input_terminal).strip().split(',') if label.strip()]
    if input_x_pos is not None and isinstance(input_x_pos, (list, tuple)) and upper_labels:
        for i in range(len(upper_labels)):
            label = upper_labels[i]
            pos = input_x_pos[i] if i < len(input_x_pos) else x
            if i == 0:
                left_x = min(pos, x)
            else:
                left_x = min(input_x_pos[i-1], pos) if i > 0 and i < len(input_x_pos) else min(pos, x)
            right_x = max(pos, x) if i == len(upper_labels) - 1 else (input_x_pos[i + 1] if i + 1 < len(input_x_pos) else x)
            ax.plot([left_x, right_x], [upper_y_end, upper_y_end], color='black', linewidth=1)
            # Add small vertical line downward from the left end
            vertical_drop_length = 1.2
            ax.plot([left_x, left_x], [upper_y_end, upper_y_end - vertical_drop_length], color='black', linewidth=1)
    else:
        upper_horiz_length = 0. + len(str(input_terminal).strip()) * 0.12
        ax.plot([x - upper_horiz_length, x], [upper_y_end, upper_y_end], color='black', linewidth=1)
    # Lower vertical line
    lower_y_start = y_center - radius
    lower_y_end = y_center - radius * 6
    ax.plot([x, x], [lower_y_start, lower_y_end], color='black', linewidth=1)
    # Lower horizontal lines (dynamic to multiple output_x_pos if provided) with vertical rise at start
    lower_labels = [label.strip() for label in str(output_terminal).strip().split(',') if label.strip()]
    if output_x_pos is not None and isinstance(output_x_pos, (list, tuple)) and lower_labels:
        for i in range(len(lower_labels)):
            label = lower_labels[i]
            pos = output_x_pos[i] if i < len(output_x_pos) else x
            if i == 0:
                left_x = min(pos, x)
            else:
                left_x = min(output_x_pos[i-1], pos) if i > 0 and i < len(output_x_pos) else min(pos, x)
            right_x = max(pos, x) if i == len(lower_labels) - 1 else (output_x_pos[i + 1] if i + 1 < len(output_x_pos) else x)
            ax.plot([left_x, right_x], [lower_y_end, lower_y_end], color='black', linewidth=1)
            # Add small vertical line upward from the left end
            vertical_drop_length = 0.8
            ax.plot([left_x, left_x], [lower_y_end, lower_y_end + vertical_drop_length], color='black', linewidth=1)
    else:
        lower_horiz_length = 0.5 + len(str(output_terminal).strip()) * 0.12
        ax.plot([x - lower_horiz_length, x], [lower_y_end, lower_y_end], color='black', linewidth=1)
    return None, None

def draw_rectangle_symbol(ax, x, y_center, terminal_no, symbol, input_left, input_right, output_left, output_right,
                         input_connected, output_connected):
    rect_width = 0.3 # horizontally small
    rect_height = 0.835 # vertically little big
    rect_x = x - rect_width / 2
    rect_y = y_center - rect_height / 2
    ax.add_patch(Rectangle((rect_x, rect_y), rect_width, rect_height,
                           edgecolor='black', facecolor='white', linewidth=1))
    # Add inner rectangle
    inner_width = 0.22
    inner_height = 0.22
    inner_x = x - inner_width / 2
    inner_offset = 0.25 # Adjust this value to move small rectangle up
    inner_y = y_center - inner_height / 2 + inner_offset
    ax.add_patch(Rectangle((inner_x, inner_y), inner_width, inner_height,
                           edgecolor='black', facecolor='white', linewidth=1))
    # Initialize text_offset with default value
    text_offset = 0 # Default value
    # Add terminal_no text
    if pd.notna(terminal_no) and str(terminal_no).strip() != '':
        term_str = str(terminal_no)
        if term_str.endswith('.0'):
            term_str = term_str[:-2]
        # Move terminal_no text slightly above center
        text_offset = 0 # Adjust this value to control how far above center
        ax.text(x, y_center + text_offset, term_str, fontsize=17, ha='center', va='center', fontname='Arial')
    # Add symbol text above terminal_no
    if pd.notna(symbol) and str(symbol).strip() != '':
        symbol_str = str(symbol)
        if symbol_str.endswith('.0'):
            symbol_str = symbol_str[:-2]
        symbol_offset = 0.25 # Adjust this to position the symbol above terminal_no
        ax.text(x, y_center + text_offset + symbol_offset, symbol_str, fontsize=19, ha='center', va='center', fontname='Arial')
    def format_text(t):
        t = str(t)
        if len(t) >= 8:
            words = t.split()
            if len(words) > 1:
                line1 = " ".join(words[:-1])
                line2 = words[-1]
                return f"{line1}\n{line2}"
            else:
                return t[:8] + "\n" + t[8:]
        return t
    input_left_offset = 0.005
    if pd.notna(input_left) and str(input_left).strip() != "":
        ax.text(x - input_left_offset, y_center + rect_height / 2 + 0.18, format_text(input_left),
                fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    input_right_offset = 0.05
    if pd.notna(input_right) and str(input_right).strip() != "":
        ax.text(x + input_right_offset, y_center + rect_height / 2 + 0.18, format_text(input_right),
                fontsize=19, ha='left', va='bottom', rotation=90, linespacing=1.2, fontname='Arial')
    output_left_offset = 0.005
    if pd.notna(output_left) and str(output_left).strip() != "":
        ax.text(x - output_left_offset, y_center - rect_height / 2 - 0.15, format_text(output_left),
                fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    output_right_offset = 0.05
    if pd.notna(output_right) and str(output_right).strip() != "":
        ax.text(x + output_right_offset, y_center - rect_height / 2 - 0.18, format_text(output_right),
                fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname='Arial')
    top_conn = (x, rect_y + rect_height)
    bottom_conn = (x, rect_y)
    ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
    oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
    return top_conn, bottom_conn, ic, oc

def draw_input_connection(ax, x, symbol_top_y, input_connected_flag, output_connected_flag, y_top_bus_group, has_bus_line=True):
    """
    Draw input connection with three scenarios:
    1. Input connected AND bus line passes: Full connection to bus
    2. Input NOT connected BUT bus line passes: Shortened connection (50%)
    3. Input NOT connected AND NO bus line: Full connection (100%)
    """
    overlap = SYMBOL_RADIUS * 0.15
    start_y = symbol_top_y - overlap
    # FIX: Check for boolean True or string 'Y'
    is_connected = (input_connected_flag is True) or (str(input_connected_flag).strip().upper() == 'Y')
    if is_connected and has_bus_line:
        # Scenario 1: Connected + Bus line - draw full line to bus group
        extended_y = y_top_bus_group + 0.2
        ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
        return True, extended_y
    elif not is_connected and has_bus_line:
        # Scenario 2: Not connected + Bus line - draw 50% shorter line
        total_distance = y_top_bus_group - start_y
        short_length = total_distance * 0.5 # 50% of normal distance
        end_y = start_y + short_length
        ax.plot([x, x], [start_y, end_y], color='black', linewidth=1)
        return False, end_y
    else:
        # Scenario 3: Not connected + No bus line - draw 100% full line
        extended_y = y_top_bus_group + 0.2
        ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
        return False, extended_y

def draw_output_connection(ax, x, symbol_bottom_y, input_connected_flag, output_connected_flag, y_bottom_bus_group, has_bus_line=True):
    """
    Draw output connection with three scenarios:
    1. Output connected AND bus line passes: Full connection to bus
    2. Output NOT connected BUT bus line passes: Shortened connection (50%)
    3. Output NOT connected AND NO bus line: Full connection (100%)
    """
    overlap = SYMBOL_RADIUS * 0.15
    start_y = symbol_bottom_y + overlap
    # FIX: Check for boolean True or string 'Y'
    is_connected = (output_connected_flag is True) or (str(output_connected_flag).strip().upper() == 'Y')
    if is_connected and has_bus_line:
        # Scenario 1: Connected + Bus line - draw full line to bus group
        extended_y = y_bottom_bus_group - 0.2
        ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
        return True, extended_y
    elif not is_connected and has_bus_line:
        # Scenario 2: Not connected + Bus line - draw 50% shorter line
        total_distance = start_y - y_bottom_bus_group
        short_length = total_distance * 0.5 # 50% of normal distance
        end_y = start_y - short_length
        ax.plot([x, x], [start_y, end_y], color='black', linewidth=1)
        return False, end_y
    else:
        # Scenario 3: Not connected + No bus line - draw 100% full line
        extended_y = y_bottom_bus_group - 0.2
        ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
        return False, extended_y

def draw_bus_lines(ax, x_positions, connected_flags, bus_y, gap=0.12, extra=0.12, y_offset=0.15):
    if not x_positions:
        return
    # Apply offset for upper/lower shift
    bus_y = bus_y + y_offset
    connected_indices = [i for i, connected in enumerate(connected_flags) if connected]
    if not connected_indices:
        return
    start_idx = connected_indices[0]
    end_idx = connected_indices[-1]
    x_start = x_positions[start_idx]
    x_end = x_positions[end_idx]
    if start_idx == end_idx:
        small = max(0.06, gap)
        ax.plot([x_start - small, x_start + small], [bus_y, bus_y], color="black", linewidth=1)
    else:
        total_len = x_end - x_start
        shrink = min(gap, total_len / 4.0)
        plot_start = x_start + shrink - extra
        plot_end = x_end - shrink + extra
        if plot_end <= plot_start:
            mid = (x_start + x_end) / 2.0
            small = max(0.06, gap)
            ax.plot([mid - small, mid + small], [bus_y, bus_y], color="black", linewidth=1)
        else:
            ax.plot([plot_start, plot_end], [bus_y, bus_y], color="black", linewidth=1)

# === Helper function for parsing terminal_no field ===
def parse_terminal_no_field(val):
    if pd.isna(val):
        return None, None
    s = str(val).strip()
    if ',' in s:
        parts = s.split(',')
        if len(parts) >= 2:
            try:
                a = parts[0].strip()
                b = parts[1].strip()
                return a, b
            except ValueError:
                pass
    if '-' in s:
        parts = s.split('-')
        if len(parts) >= 2:
            try:
                a = parts[0].strip()
                b = parts[1].strip()
                return a, b
            except ValueError:
                pass
    return s, s

# === UPDATED: Draw horizontal chokes with output_connected support ===
def draw_horizontal_choke_updated(ax, x_center, y_center, label='CHOKE',
                          box_width=0.6, box_height=0.3,
                          special_end=False, output_label='', output_type='terminal', output_text='',
                          output_connected_terminals=[], x_positions=None, terminal_nos_for_positions=None,
                          y_top_bus_group=None, y_bottom_bus_group=None):
    """
    Updated to handle output_connected as a list of terminal numbers.
    Choke position always below (lower side).
    Input (left) connection always to bottom bus/symbol side.
    Output (right) connection to the specified terminals at the bottom.
    """
    # Always place choke below
    y_shift = -0.5
    y_center = y_center + y_shift
 
    left_x = x_center - box_width / 2
    right_x = x_center + box_width / 2
    bottom_y = y_center - box_height / 2
 
    # Draw rounded box
    choke_box = FancyBboxPatch((left_x, bottom_y),
                               box_width, box_height,
                               boxstyle="round,pad=0.02",
                               edgecolor='black', facecolor='white', linewidth=1.5)
    ax.add_patch(choke_box)
 
    # Draw label
    ax.text(x_center, y_center, label,
            fontsize=16, ha='center', va='center', fontname='Arial')
 
    # Connection lines
    line_length = 0.075
    delta = 0.02
    vert_line_height_left = y_bottom_bus_group - y_center if y_bottom_bus_group is not None else 0.5 # Short up to bottom bus
 
    # Left connection (input): always to bottom side (short up)
    left_horiz_start = left_x - line_length - delta
    left_horiz_end = left_x - delta
    ax.plot([left_horiz_start, left_horiz_end],
            [y_center, y_center],
            color='black', linewidth=1)
 
    v_offset = 0.005
    ax.plot([left_horiz_start - v_offset, left_horiz_start - v_offset],
            [y_center, y_center + vert_line_height_left], # Up to bottom bus
            color='black', linewidth=1)
 
    # Handle output connections to multiple terminals
    if output_connected_terminals and x_positions is not None and terminal_nos_for_positions is not None:
        # Find the x-positions of the connected terminals
        connected_x = []
        for term in output_connected_terminals:
            term_str = str(term).strip().replace('.0', '')
            if term_str in terminal_nos_for_positions:
                idx = terminal_nos_for_positions.index(term_str)
                connected_x.append(x_positions[idx])
        
        if connected_x:
            # Draw horizontal bus line connecting all terminals
            min_x = min(connected_x)
            max_x = max(connected_x)
            y_bus = y_bottom_bus_group - 0.8  # Position for the bus line
            
            # Draw the horizontal bus line
            ax.plot([min_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
            
            # Draw vertical connections from each terminal to the bus
            for x in connected_x:
                ax.plot([x, x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
            
            # Draw connection from choke to the bus
            choke_output_x = x_center + box_width/2
            ax.plot([choke_output_x, choke_output_x], [y_center, y_bus], color='black', linewidth=1)
            ax.plot([choke_output_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
            
            return max_x  # Return the rightmost x position for header placement

    # If no output_connected terminals or special end required, use original logic
    if not special_end:
        # Right connection (output): non-special
        right_horiz_start = right_x + delta
        right_horiz_end = right_x + line_length + delta
        ax.plot([right_horiz_start, right_horiz_end],
                [y_center, y_center],
                color='black', linewidth=1)
     
        # Right vertical: up to top if connected, else short up to bottom
        if output_connected_terminals and y_top_bus_group is not None:
            vert_line_height_right = y_top_bus_group - y_center # Long up to top bus
        else:
            vert_line_height_right = vert_line_height_left # Short up to bottom bus
     
        ax.plot([right_horiz_end + v_offset, right_horiz_end + v_offset],
                [y_center, y_center + vert_line_height_right],
                color='black', linewidth=1)
    else:
        # Special end for output (relay type)
        horiz_length = 0.5 if output_type.lower() == 'relay' else 0.2
        slant_size = 0.3 if output_type.lower() != 'relay' else 0.15
        vertical_length = 0.5 if output_type.lower() != 'relay' else 0 # No extra vertical for relay
        end_horiz_x = right_x + delta + horiz_length
        ax.plot([right_x + delta, end_horiz_x],
                [y_center, y_center],
                color='black', linewidth=1.4)
     
        if str(output_type).strip().lower() == 'relay':
            # For relay output type - horizontal + diagonal + text
            diag_length = slant_size # Reuse slant_size for diag_length
         
            if output_connected_terminals:
                # Upward / from low to high
                diag_start_x = end_horiz_x - 0.1
                diag_start_y = y_center
                diag_end_x = diag_start_x + diag_length
                diag_end_y = diag_start_y + diag_length
                ax.plot([diag_start_x, diag_end_x],
                        [diag_start_y, diag_end_y],
                        color='black', linewidth=1.4)
                # Text at end (higher)
                text_offset = 0.05
                ax.text(diag_end_x + text_offset, diag_end_y, output_text,
                        ha='left', va='center', fontsize=19, fontname='Arial')
            else:
                # Downward \ from high to low
                diag_start_x = end_horiz_x - 0.1
                diag_start_y = y_center + 0.1
                diag_end_x = diag_start_x + diag_length
                diag_end_y = diag_start_y - diag_length
                ax.plot([diag_start_x, diag_end_x],
                        [diag_start_y, diag_end_y],
                        color='black', linewidth=1.4)
                # Text at end (lower)
                text_offset = 0.05
                ax.text(diag_end_x + text_offset, diag_end_y - 0.1, output_text,
                        ha='left', va='center', fontsize=19, fontname='Arial')
         
            return end_horiz_x + diag_length
        else:
            # Terminal output type - horizontal + slant + vertical + label
            if output_connected_terminals:
                # Upward slant / + vertical up
                ax.plot([end_horiz_x, end_horiz_x + slant_size],
                        [y_center, y_center + slant_size],
                        color='black', linewidth=1.4)
                vert_top_y = y_center + slant_size
                vert_end_y = vert_top_y + vertical_length # Further up
                ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                        [vert_top_y, vert_end_y],
                        color='black', linewidth=1.4)
             
                # Label
                label_offset = 0.05
                label_y = (vert_top_y + vert_end_y) / 2
                ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                        ha='left', va='center', fontsize=19, rotation=90, fontname='Arial')
                return end_horiz_x + slant_size
            else:
                # Downward slant \ + vertical down (original)
                ax.plot([end_horiz_x, end_horiz_x + slant_size],
                        [y_center, y_center - slant_size],
                        color='black', linewidth=1.4)
                vert_top_y = y_center - slant_size
                vert_end_y = vert_top_y - vertical_length # Further down
                ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                        [vert_top_y, vert_end_y],
                        color='black', linewidth=1.4)
             
                # Label
                label_offset = 0.05
                label_y = (vert_top_y + vert_end_y) / 2
                ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                        ha='left', va='center', fontsize=19, rotation=90, fontname='Arial')
                return end_horiz_x + slant_size
    return None

# === UPDATED MAIN DRAW FUNCTION ===
def draw_symbols(df, ax, page_rows, junction_name, start_x=1, pin_spacing=0.8, cables_per_page=12, page_number=1, max_terminal_symbols_per_row=36, max_rows_visible=3, page_width=None):
    """
    Draw symbols for the provided page_rows on ax.
    UPDATED: Handle cable boxes for ALL letters (A-Z), not just F row
    UPDATED: Proper row management with row breaking for all letters
    UPDATED: Bottom-align content for pages with fewer than max_rows_visible rows
    """
    extra_rows = 0
    max_rows_for_ylim = max_rows_visible + extra_rows
    bottom_margin = 1.0
    top_margin = 3.0
    fixed_ylim_min = CAPSULE_Y_CENTER_BASE + vertical_gap * (1 - max_rows_for_ylim) + y_bottom_bus_offset - 1.8 - bottom_margin - footer_height
    fixed_ylim_max = CAPSULE_Y_CENTER_BASE + y_top_bus_offset + 1.8 + top_margin
    current_x = start_x
    current_terminal_count = 0
    # FIXED: Calculate initial y_offset to bottom-align content
    num_rows = len(page_rows)
    y_offset = vertical_gap * (num_rows - max_rows_visible)
    all_x_positions = []
    all_input_connected_flags = []
    all_output_connected_flags = []
    min_y = float('inf')
    max_y = float('-inf')
    overall_max_x = start_x
    current_row_max_x = start_x
   
    print(f"DEBUG: Processing {len(page_rows)} rows for this page")
   
    # Process each row in the page
    for row_index, (letter, cable_list) in enumerate(page_rows):
        print(f"DEBUG: Processing row {row_index + 1} - Letter '{letter}' with {len(cable_list)} cables")
       
        # Reset for new row
        current_x = start_x
        current_terminal_count = 0
        current_row_max_x = start_x
       
        # Draw group label for this row
        draw_cable_name(ax, start_x - 1.2, CAPSULE_Y_CENTER_BASE + y_offset, letter)
       
        capsule_y_center = CAPSULE_Y_CENTER_BASE + y_offset
        y_top_bus_group = capsule_y_center + y_top_bus_offset
        y_bottom_bus_group = capsule_y_center + y_bottom_bus_offset
       
        # === UPDATED: Process ALL letters for cable boxes ===
        for cable_id in cable_list:
            # Check both regular cables and cable boxes for this cable_id
            cable_rows = df_cable[df_cable['cable_id'] == cable_id]
            cable_box_rows = df_cable_box[df_cable_box['cable_id'] == cable_id] if not df_cable_box.empty else pd.DataFrame()
           
            is_cable_box = not cable_box_rows.empty
           
            if is_cable_box:
                # This is a cable box - handle for ANY letter
                cable_info = cable_box_rows.iloc[0]
               
                # Check both possible column names for cable type
                cable_type = ""
                if pd.notna(cable_info.get('cable_type')):
                    cable_type = str(cable_info.get('cable_type')).strip().lower()
                elif pd.notna(cable_info.get('cabel_type')):
                    cable_type = str(cable_info.get('cabel_type')).strip().lower()
               
                # Draw cable box row for ANY letter
                if cable_type == 'relay_box':
                    position_val = cable_info.get('position')
                    if pd.notna(position_val):
                        try:
                            position_num = int(float(position_val))
                        except:
                            position_num = 1
                    else:
                        position_num = 1
                   
                    # Check space for cable box row
                    if current_terminal_count + position_num > max_terminal_symbols_per_row:
                        print(f"ERROR: Cable boxes in row '{letter}' exceed row limit - this shouldn't happen with proper pagination")
                        continue
                   
                    # Draw cable box row with proper spacing
                    x_positions, input_connected_flags, output_connected_flags = draw_cable_box_row(
                        ax, current_x, capsule_y_center, cable_info, pin_spacing
                    )
                   
                    if x_positions:
                        current_x += len(x_positions) * pin_spacing
                        current_row_max_x = max(current_row_max_x, current_x)
                        current_terminal_count += len(x_positions)
                       
                        all_x_positions.extend(x_positions)
                        all_input_connected_flags.extend(input_connected_flags)
                        all_output_connected_flags.extend(output_connected_flags)
                       
                        # Add extra space after cable boxes for ALL letters
                        current_x += pin_spacing * 2 # Double spacing for cable boxes in any row
                        current_row_max_x = max(current_row_max_x, current_x)
                       
                        min_y = min(min_y, y_bottom_bus_group - 1.8)
                        max_y = max(max_y, y_top_bus_group + 1.8)
               
                # Skip regular cable processing for cable boxes
                current_x += CABLE_GAP
                current_row_max_x = max(current_row_max_x, current_x)
                continue
           
            # === REGULAR CABLE PROCESSING (non-cable-box) ===
            if not cable_rows.empty:
                cable_info = cable_rows.iloc[0]
               
                # Check both possible column names for cable type
                cable_type = ""
                if pd.notna(cable_info.get('cable_type')):
                    cable_type = str(cable_info.get('cable_type')).strip().lower()
                elif pd.notna(cable_info.get('cabel_type')):
                    cable_type = str(cable_info.get('cabel_type')).strip().lower()
               
                # Skip if it's a cable box (shouldn't happen here due to above check, but just in case)
                if cable_type == 'relay_box':
                    continue
               
                group = df_symbols[df_symbols['cable_id'] == cable_id].sort_index().reset_index(drop=True)
               
                if group.empty:
                    current_x += pin_spacing + CABLE_GAP
                    current_row_max_x = max(current_row_max_x, current_x)
                    current_terminal_count += 1
                    min_y = min(min_y, capsule_y_center - 2.0)
                    max_y = max(max_y, capsule_y_center + 2.0)
                    continue
                   
                input_connected_flags = []
                output_connected_flags = []
                x_positions = []
                symbol_bottoms = []
                terminal_nos_for_positions = []
                i = 0
               
                # Calculate how many terminal symbols this cable will take
                symbols_to_add_total = 0
                temp_i = 0
                while temp_i < len(group):
                    symbol = str(group.iloc[temp_i].get('symbol', '')).strip().lower()
                    if symbol == 'dual_fuse':
                        if temp_i + 1 < len(group):
                            symbols_to_add_total += 2
                            temp_i += 2
                        else:
                            symbols_to_add_total += 1
                            temp_i += 1
                    else:
                        symbols_to_add_total += 1
                        temp_i += 1
               
                print(f"DEBUG: Cable {cable_id} requires {symbols_to_add_total} terminal symbols")
               
                # Check if adding this cable would exceed the row limit
                if current_terminal_count + symbols_to_add_total > max_terminal_symbols_per_row:
                    print(f"ERROR: Cable {cable_id} would exceed row limit - this shouldn't happen with proper pagination")
                    continue
               
                # Reset i for actual drawing
                i = 0
                row_x_positions = []
                row_symbol_bottoms = []
                row_symbol_tops = []
                row_terminal_nos = []
                while i < len(group):
                    row = group.iloc[i]
                    symbol = str(row.get('symbol', '')).strip().lower()
                    symbols_to_add = 2 if symbol == 'dual_fuse' else 1
                   
                    # Handle capsule types
                    if symbol in ['capsule', 'ara', 'wago', 'ara/wago']:
                        top_conn, bottom_conn, input_conn, output_conn = draw_capsule(
                            ax, current_x, capsule_y_center,
                            row.get('terminal_no'),
                            row.get('input_left'),
                            row.get('input_right'),
                            row.get('output_left'),
                            row.get('output_right'),
                            row.get('input_connected', 'N'),
                            row.get('output_connected', 'N'),
                            capsule_type=symbol
                        )
                        x_positions.append(current_x)
                        row_x_positions.append(current_x)
                        tname = str(row.get('terminal_no')).strip()
                        if tname.endswith('.0'):
                            tname = tname[:-2]
                        terminal_nos_for_positions.append(tname)
                        input_connected_flags.append(str(input_conn).strip().upper() == "Y")
                        output_connected_flags.append(str(output_conn).strip().upper() == "Y")
                        symbol_bottoms.append(bottom_conn[1])
                        current_x += pin_spacing
                        current_row_max_x = max(current_row_max_x, current_x)
                        current_terminal_count += 1
                        i += 1
                        row_symbol_bottoms.append(bottom_conn[1])
                        row_symbol_tops.append(top_conn[1])
                        row_terminal_nos.append(tname)
                    elif symbol == 'single_fuse':
                        top_conn, bottom_conn, input_conn, output_conn = draw_s_fuse(
                            ax, current_x, capsule_y_center, row.get('terminal_no'),
                            row.get('input_left'), row.get('input_right'), row.get('output_left'), row.get('output_right'),
                            row.get('input_connected', 'N'), row.get('output_connected', 'N')
                        )
                        x_positions.append(current_x)
                        tname = str(row.get('terminal_no')).strip()
                        if tname.endswith('.0'):
                            tname = tname[:-2]
                        terminal_nos_for_positions.append(tname)
                        input_connected_flags.append(str(input_conn).strip().upper() == "Y")
                        output_connected_flags.append(str(output_conn).strip().upper() == "Y")
                        symbol_bottoms.append(bottom_conn[1])
                        current_x += pin_spacing
                        current_row_max_x = max(current_row_max_x, current_x)
                        current_terminal_count += 1
                        i += 1
                    elif symbol == 'choke':
                        i += 1
                        continue # Skip vertical symbol drawing; we'll handle horizontal later
                    elif symbol == 'dual_fuse':
                        ## MODIFIED: Use first row's connection flags for both terminals
                        if i + 1 < len(group):
                            next_row = group.iloc[i+1]
                            LEFT_EXTRA = pin_spacing * 1.0
                            AFTER_SPACING = pin_spacing * 1.5
                            current_x += LEFT_EXTRA
                            dual_start_x = current_x - SYMBOL_WIDTH * 1.25
                       
                            # USE FIRST ROW'S CONNECTION FLAGS FOR BOTH TERMINALS
                            first_input_connected = row.get('input_connected', 'N')
                            first_output_connected = row.get('output_connected', 'N')
                       
                            top_conn, bottom_conn, left_ic, left_oc, right_ic, right_oc = draw_dual_fuse(
                                ax, dual_start_x, capsule_y_center,
                                row.get('terminal_no'),
                                next_row.get('terminal_no'),
                                row.get('input_left'), row.get('input_right'), row.get('output_left'), row.get('output_right'),
                                first_input_connected, first_output_connected, # Use first row for left
                                next_row.get('input_left'), next_row.get('input_right'), row.get('output_left'), next_row.get('output_right'),
                                first_input_connected, first_output_connected # Use first row for right too
                            )
                            tname_left = str(row.get('terminal_no')).strip()
                            tname_right = str(next_row.get('terminal_no')).strip()
                            if tname_left.endswith('.0'): tname_left = tname_left[:-2]
                            if tname_right.endswith('.0'): tname_right = tname_right[:-2]
                            x_positions.append(dual_start_x)
                            x_positions.append(dual_start_x)
                            terminal_nos_for_positions.append(tname_left)
                            terminal_nos_for_positions.append(tname_right)
                            input_connected_flags.append(str(left_ic).strip().upper() == "Y")
                            input_connected_flags.append(str(right_ic).strip().upper() == "Y")
                            output_connected_flags.append(str(left_oc).strip().upper() == "Y")
                            output_connected_flags.append(str(right_oc).strip().upper() == "Y")
                            symbol_bottoms.append(capsule_y_center - SYMBOL_HEIGHT / 2 - SYMBOL_RADIUS)
                            symbol_bottoms.append(capsule_y_center - SYMBOL_HEIGHT / 2 - SYMBOL_RADIUS)
                            current_x += AFTER_SPACING
                            current_row_max_x = max(current_row_max_x, current_x)
                            current_terminal_count += 2
                            i += 2
                        else:
                            top_conn, bottom_conn, input_conn, output_conn = draw_s_fuse(
                                ax, current_x, capsule_y_center, row.get('terminal_no'),
                                row.get('input_left'), row.get('input_right'), row.get('output_left'), row.get('output_right'),
                                row.get('input_connected', 'N'), row.get('output_connected', 'N')
                            )
                            x_positions.append(current_x)
                            tname = str(row.get('terminal_no')).strip()
                            if tname.endswith('.0'):
                                tname = tname[:-2]
                            terminal_nos_for_positions.append(tname)
                            input_connected_flags.append(str(input_conn).strip().upper() == "Y")
                            output_connected_flags.append(str(output_conn).strip().upper() == "Y")
                            symbol_bottoms.append(bottom_conn[1])
                            current_x += pin_spacing
                            current_row_max_x = max(current_row_max_x, current_x)
                            current_terminal_count += 1
                            i += 1
                    else:
                       # For blank or unknown symbols, draw rectangle
                        top_conn, bottom_conn, input_conn, output_conn = draw_rectangle_symbol(
                            ax, current_x, capsule_y_center,
                            row.get('terminal_no'),
                            row.get('symbol', ''),
                            row.get('input_left'),
                            row.get('input_right'),
                            row.get('output_left'),
                            row.get('output_right'),
                            row.get('input_connected', 'N'),
                            row.get('output_connected', 'N')
                        )
                        x_positions.append(current_x)
                        tname = str(row.get('terminal_no')).strip()
                        if tname.endswith('.0'):
                            tname = tname[:-2]
                        terminal_nos_for_positions.append(tname)
                        input_connected_flags.append(str(input_conn).strip().upper() == "Y")
                        output_connected_flags.append(str(output_conn).strip().upper() == "Y")
                        symbol_bottoms.append(bottom_conn[1])
                        current_x += pin_spacing
                        current_row_max_x = max(current_row_max_x, current_x)
                        current_terminal_count += 1
                        i += 1
                # Add Overlay
                R = SYMBOL_RADIUS * 0.8
                # spacing safe
                if len(row_x_positions) >= 2:
                    spacing = row_x_positions[1] - row_x_positions[0]
                else:
                    spacing = 40
                V_DEPTH = spacing * 0.6
                count = min(len(row_x_positions), len(row_symbol_bottoms))
                group_overlay=[
                            (3, 6), (8, 11)
                            
                        ]
                cable_headers_temp = df_header[df_header['cable_id'] == cable_id]

                overlay_headers = cable_headers_temp[
                            cable_headers_temp['header_type'].fillna('').str.upper() == 'OVERLAY'
                        ]
                tpoint_headers = cable_headers_temp[
                            cable_headers_temp['header_type'].fillna('').str.upper() == 'TPOINT'
                        ]        
                
                groups_up = []
                groups_down = []
                groups_tpoint = []
                for _, row in overlay_headers.iterrows():
                            try:
                                start_t = int(float(row['terminal_start']))
                                end_t   = int(float(row['terminal_end']))
                                
                            except Exception as e:
                                print("terminal parse error:", row['terminal_start'], row['terminal_end'], e)
                                continue

                            start_idx = start_t -1
                            end_idx   = end_t -1

                            io_raw = row.get('input_output', '')
                            io = str(io_raw).strip().lower()

                            print("DEBUG io_raw:", repr(io_raw), "-> norm:", io)
                            print("DEBUG range:", start_idx, end_idx)

                            if io == 'input':
                                groups_up.append((start_idx, end_idx))
                            elif io == 'output':
                                groups_down.append((start_idx, end_idx))
                            else:
                                print("Unknown input_output value:", repr(io_raw)) 
                for _, row in tpoint_headers.iterrows():
                            try:
                                start_t = int(float(row['terminal_start']))
                                end_t   = int(float(row['terminal_end']))
                                
                            except Exception as e:
                                print("terminal parse error:", row['terminal_start'], row['terminal_end'], e)
                                continue

                            start_idx = start_t -1
                            end_idx   = end_t -1

                            io_raw = row.get('input_output', '')
                            io = str(io_raw).strip().lower()

                           

                            if io == 'input':
                                groups_tpoint.append((start_idx, end_idx))
                            elif io == 'output':
                                groups_tpoint.append((start_idx, end_idx))
                            else:
                                print("Unknown input_output value:", repr(io_raw)) 
                                                   
               
                   
                for (start, end) in groups_up:
                            start = max(0, start)
                            end = min(count - 1, end)

                            for idx in range(start, end):
                                x1 = row_x_positions[idx]
                                x2 = row_x_positions[idx + 1]
                                y1 = row_symbol_tops[idx]
                                y2 = row_symbol_tops[idx + 1]

                                if y1 is None or y2 is None:
                                    continue

                                xm = (x1 + x2) / 2
                                base_y = max(y1, y2) + R - 0.1   # ⬆️ circle ke upar se start
                                ym = base_y + V_DEPTH           # ⬆️ reverse V upar

                                ax.plot([x1, xm], [base_y, ym], color='black', linewidth=1, clip_on=True)
                                ax.plot([xm, x2], [ym, base_y], color='black', linewidth=1, clip_on=True)
                
                for (start, end) in groups_down:
                                start = max(0, start)
                                end = min(count - 1, end)
                                print("start1 : ",start) 
                                print("end1 : ",end) 
                                for idx in range(start, end):
                                    print("start2 : ",start) 
                                
                                    print("end2 : ",end) 
                                    print("idx : ",idx) 
                                    x1 = row_x_positions[idx]
                                    x2 = row_x_positions[idx + 1]
                                    y1 = row_symbol_bottoms[idx]
                                    y2 = row_symbol_bottoms[idx + 1]

                                    if y1 is None or y2 is None:
                                        continue

                                    xm = (x1 + x2) / 2
                                    base_y = min(y1, y2) - R + 0.1   # ⬇️ circle ke neeche se start
                                    ym = base_y - V_DEPTH           # ⬇️ V neeche

                                    ax.plot([x1, xm], [base_y, ym], color='black', linewidth=1, clip_on=True)
                                    ax.plot([xm, x2], [ym, base_y], color='black', linewidth=1, clip_on=True)
                
                for (start, end) in groups_tpoint:
                    
                    indices_to_draw = [start, end ]  # last valid index = end-1


                    for idx in indices_to_draw:
                        
                        x1 = row_x_positions[idx]
                        x2 = row_x_positions[idx + 1]
                        y1 = row_symbol_bottoms[idx]
                        y2 = row_symbol_bottoms[idx + 1]

                        if y1 is None or y2 is None:
                            continue

                        # midpoint between two terminals
                        xm = (x1 + x2) / 2

                        # start just below the symbol
                        #base_y = min(y1, y2) - R + 1
                        symbol_bottom = min(y1, y2)
                        base_y = symbol_bottom - 0.40  # small offset below the symbol

                        # ---- short diagonal "half V" ----
                        DIAG_LEN = 0.3  # horizontal length of half-V
                        # left point of diagonal
                        x_left = xm - DIAG_LEN
                        y_left = base_y + DIAG_LEN  # upward, adjust sign if axis inverted

                        ax.plot([x_left, xm], [y_left, base_y], color='black', linewidth=1, clip_on=True)

                        # ---- vertical line down from midpoint ----
                        
                        y_down = base_y - 1.5  # use + if axis inverted
                        ax.plot([xm, xm], [base_y, y_down], color='black', linewidth=1, clip_on=True)
                
                # Add middle space
                current_x += pin_spacing # extra space after symbols
                current_row_max_x = max(current_row_max_x, current_x)
            
                # === UPDATED: Draw extra connections for this cable ===
                # Make sure we're using the correct terminal numbers
                draw_extra_connections(ax, group, x_positions, terminal_nos_for_positions,
                                     y_top_bus_group, y_bottom_bus_group, capsule_y_center)
            
                # Add resistor if applicable
                # Resistor logic - draw if cable_id exists in resistortable
                resistor_row = df_resistor[df_resistor['cable_id'] == cable_id]
                special_resistor = False
                if not resistor_row.empty:
                    resistor_label = str(resistor_row['resistor_name'].iloc[0]).strip() if 'resistor_name' in resistor_row.columns and pd.notna(resistor_row['resistor_name'].iloc[0]) else 'R'
                    input_terms = [term.strip() for term in str(resistor_row['input_terminal'].iloc[0]).strip().replace('.0', '').split(',') if term.strip()] if 'input_terminal' in resistor_row.columns else []
                    output_terms = [term.strip() for term in str(resistor_row['output_terminal'].iloc[0]).strip().replace('.0', '').split(',') if term.strip()] if 'output_terminal' in resistor_row.columns else []
               
                    input_x_pos = [x_positions[terminal_nos_for_positions.index(term)] for term in input_terms if term in terminal_nos_for_positions] if input_terms else None
                    output_x_pos = [x_positions[terminal_nos_for_positions.index(term)] for term in output_terms if term in terminal_nos_for_positions] if output_terms else None
               
                    symbols_to_add = 1
                    if current_terminal_count + symbols_to_add > max_terminal_symbols_per_row:
                        # This shouldn't happen in proper pagination, but handle anyway
                        print(f"WARNING: Resistor would exceed row limit, but continuing anyway")
               
                    current_x -= 0.5
                    draw_resistor(ax, current_x, capsule_y_center,
                                input_terminal=','.join(input_terms) if input_terms else '',
                                output_terminal=','.join(output_terms) if output_terms else '',
                                resistor_name=resistor_label,
                                input_x_pos=input_x_pos,
                                output_x_pos=output_x_pos)
                    current_x += pin_spacing
                    current_row_max_x = max(current_row_max_x, current_x)
                    current_terminal_count += 1
                    special_resistor = True
                   
                # === UPDATED: Draw horizontal chokes with output_connected support ===
                choke_rows = df_choke[df_choke['cable_id'] == cable_id]
                special_choke = False
                vert_x = None
                # Process each choke for this cable
                for choke_idx, choke_row in choke_rows.iterrows():
                    input_term = str(choke_row.get('input_terminal', pd.NA)).strip().replace('.0', '') if 'input_terminal' in choke_row else ''
                    output_term = str(choke_row.get('output_terminal', pd.NA)).strip().replace('.0', '') if 'output_terminal' in choke_row else ''
                    output_type = str(choke_row.get('output_type', 'terminal')).strip().lower() if 'output_type' in choke_row else 'terminal'
                    output_text = str(choke_row.get('output_text', '')).strip() if 'output_text' in choke_row else ''
               
                    # === UPDATED: Parse output_connected as list of terminal numbers ===
                    output_connected_terminals = []
                    if 'output_connected' in choke_row and pd.notna(choke_row['output_connected']):
                        output_connected_val = str(choke_row['output_connected']).strip()
                        if output_connected_val and output_connected_val != '':
                            # Split by comma and remove whitespace
                            output_connected_terminals = [t.strip() for t in output_connected_val.split(',')]
                   
                    if input_term and input_term in terminal_nos_for_positions:
                        start_idx = terminal_nos_for_positions.index(input_term)
                        x_left = x_positions[start_idx]
                        choke_label = str(choke_row.get('terminal_name', pd.NA)).strip() if 'terminal_name' in choke_row and pd.notna(choke_row['terminal_name']) else 'CHOKE'
                        # GET SYMBOL TYPE FOR INPUT TERMINAL
                        input_symbol_type = None
                        input_symbol_row = group[group['terminal_no'].astype(str).str.replace('.0', '') == input_term]
                        if not input_symbol_row.empty:
                            input_symbol_type = str(input_symbol_row.iloc[0].get('symbol', '')).strip().lower()
                        # FIXED SIZE FOR ALL SYMBOL TYPES
                        box_width = 0.6
                        box_height = 0.3
                        if output_term and output_term in terminal_nos_for_positions and output_type != 'relay':
                            # Regular choke between two terminals
                            end_idx = terminal_nos_for_positions.index(output_term)
                            x_right = x_positions[end_idx]
                            choke_x_center = (x_left + x_right) / 2
                            # Connection parameters
                            if input_symbol_type == 'dual_fuse':
                                vertical_drop = 0.25
                                horizontal_extension = 0.005
                            else:
                                vertical_drop = 0.3
                                horizontal_extension = 0.0
                            vertical_position = y_bottom_bus_group - vertical_drop
                            # === LEFT VERTICAL (original + downward extension) ===
                            left_extra_up = 0.8
                            original_bottom_left = y_bottom_bus_group + left_extra_up
                            # original left vertical (keeps same)
                            ax.plot([x_left, x_left],
                                    [original_bottom_left, vertical_position],
                                    color='black', linewidth=1)
                            # extra downward extension (below original_bottom_left)
                            # Default extension (when NOT connected)
                            left_extra_extend_down = 1.3
                            # SPECIAL CASE: dual_fuse + output NOT connected ? handle differently
                            if input_symbol_type == 'dual_fuse' and not output_connected_terminals:
                                left_extra_extend_down = 1.24 # <<< your custom smaller size
                            # If output is connected ? extend only a little bit (existing logic)
                            elif output_connected_terminals:
                                left_extra_extend_down = 1.24
                            # Draw the line
                            ax.plot(
                                [x_left, x_left],
                                [original_bottom_left, original_bottom_left - left_extra_extend_down],
                                color='black',
                                linewidth=1
                            )
                            # Left horizontal (slightly cut on the choke side)
                            left_connection_start_x = x_left - horizontal_extension
                            line_down = 0.2
                            cut_amount_left = 0.02
                            ax.plot([left_connection_start_x, (choke_x_center - box_width/2) - cut_amount_left],
                                    [vertical_position - line_down, vertical_position - line_down],
                                    color='black', linewidth=1)
                            # === RIGHT HORIZONTAL (extend when connected; cut from left if connected, otherwise cut from right) ===
                            right_connection_end_x = x_right + horizontal_extension
                            if output_connected_terminals:
                                # make it longer when connected
                                right_connection_end_x += 0.19
                                # CUT FROM LEFT (start moves right)
                                cut_from_left = 0.02
                                right_horizontal_start = (choke_x_center + box_width/2) + cut_from_left
                                right_horizontal_end = right_connection_end_x
                            else:
                                # output NOT connected ? ALSO CUT FROM LEFT
                                cut_from_left = 0.02
                                right_horizontal_start = (choke_x_center + box_width/2) + cut_from_left
                                right_horizontal_end = right_connection_end_x # do NOT cut the right end
                            line_down_right = 0.20
                            # Always draw line
                            ax.plot(
                                [right_horizontal_start, right_horizontal_end],
                                [vertical_position - line_down_right, vertical_position - line_down_right],
                                color='black', linewidth=1
                            )
                            # === RIGHT VERTICAL LINE (original segment + separate downward extension) ===
                            if output_connected_terminals:
                                right_vertical_x = x_right + 0.20 # move vertical line right when connected
                                right_end_y = y_top_bus_group
                            else:
                                right_vertical_x = x_right
                                right_end_y = y_bottom_bus_group
                            # Use distinct name so we don't overwrite left extra_up
                            right_extra_up = 0.4
                            original_top_right = right_end_y + right_extra_up
                            # Draw ORIGINAL right vertical segment (keep unchanged)
                            extend_bottom = 0.2 # make this bigger or smaller
                            ax.plot(
                                [right_vertical_x, right_vertical_x],
                                [vertical_position - extend_bottom, original_top_right], # bottom extended DOWN
                                color='black',
                                linewidth=1
                            )
                            # Draw EXTRA downward extension (only adds below original_top_right)
                            # Default
                            right_extra_extend_down = 0.9
                            # SPECIAL CASE ? dual_fuse + output NOT connected
                            if input_symbol_type == 'dual_fuse' and not output_connected_terminals:
                                right_extra_extend_down = 0.72 # <<< smaller value for dual fuse (same idea as left)
                            # If output is connected (existing logic) � keep your previous behavior
                            elif output_connected_terminals:
                                right_extra_extend_down = 0.85 # <<< whatever you want connected side to be
                            ax.plot(
                                [right_vertical_x, right_vertical_x],
                                [original_top_right, original_top_right - right_extra_extend_down],
                                color='black',
                                linewidth=1
                            )
                            # === small horizontal + small vertical branch from left end of small horizontal when connected ===
                            if output_connected_terminals:
                                small_line_length = 0.2
                                left_end_x = right_vertical_x - small_line_length
                                y_top = right_end_y + right_extra_up
                                # top small horizontal (to left)
                                ax.plot([left_end_x, right_vertical_x],
                                        [y_top, y_top],
                                        color='black', linewidth=1)
                                # small vertical drop from that left end
                                small_vertical_drop = 0.9
                                ax.plot([left_end_x, left_end_x],
                                        [y_top, y_top - small_vertical_drop],
                                        color='black', linewidth=1)
                            # === NEW: Handle output_connected_terminals for regular choke ===
                            if output_connected_terminals:
                                # Find the x-positions of the connected terminals
                                connected_x = []
                                for term in output_connected_terminals:
                                    term_str = str(term).strip().replace('.0', '')
                                    if term_str in terminal_nos_for_positions:
                                        idx = terminal_nos_for_positions.index(term_str)
                                        connected_x.append(x_positions[idx])
                                
                                if connected_x:
                                    # Sort the x-positions from left to right
                                    connected_x_sorted = sorted(connected_x)
                                    
                                    # Draw horizontal bus line connecting all terminals
                                    min_x = min(connected_x_sorted)
                                    max_x = max(connected_x_sorted)
                                    y_bus = y_bottom_bus_group - 0.8 # Position for the bus line
                                    
                                    # Draw the horizontal bus line
                                    offset = 0.2 # slight left extension
                                    drop_length = 0.3 # length of vertical down line (adjust as needed)
                                    # Horizontal bus (extended to left)
                                    ax.plot(
                                        [min_x - offset, max_x],
                                        [y_bus, y_bus],
                                        color='black',
                                        linewidth=1
                                    )
                                    # Small vertical line at left offset going down
                                    ax.plot(
                                        [min_x - offset, min_x - offset],
                                        [y_bus, y_bus - drop_length],
                                        color='black',
                                        linewidth=1
                                    )
                                    
                                    # Draw vertical connections from each terminal to the bus AND add sequential numbers
                                    for i, x in enumerate(connected_x_sorted, 1):
                                        # Draw vertical line from terminal to bus
                                        ax.plot([x, x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
                                        
                                        # Add sequential number (1, 2, 3, ...) below the bus line at this connection point
                                        # Position the number slightly below the bus line
                                        number_y = y_bus - 0.05
                                        ax.text(
                                            x, number_y, str(i),
                                            fontsize=16, ha='center', va='top', fontname='Arial', fontweight='bold'
                                        )
                                    
                                    # Draw connection from choke output to the bus
                                    choke_output_x = x_right
                                    ax.plot([choke_output_x, choke_output_x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
                                    ax.plot([choke_output_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
                                    # === NEW: Draw output_text below the bus (centered across connected range) ===
                                    if output_text:
                                        try:
                                            # place the text just below the bus line (keeping original position)
                                            text_x = (min_x + max_x) / 2.0
                                            # small vertical offset so it sits below the bus
                                            text_y = y_bus - 0.20
                                            ax.text(
                                                text_x, text_y, output_text,
                                                fontsize=22, ha='center', va='top', fontname='Arial',
                                                bbox=dict(boxstyle="round,pad=0.1", facecolor='white', edgecolor='none', alpha=0.0)
                                            )
                                        except Exception:
                                            # fail-safe: if anything goes wrong with text drawing, skip silently
                                            pass
                            # Draw choke box (moved down a bit)
                            down_shift = 0.2
                            choke_box = FancyBboxPatch(
                                (choke_x_center - box_width/2, (vertical_position - box_height/2) - down_shift),
                                box_width, box_height,
                                boxstyle="round,pad=0.02",
                                edgecolor='black', facecolor='white', linewidth=1.5
                            )
                            ax.add_patch(choke_box)
                            # Label (also moved down)
                            ax.text(choke_x_center, vertical_position - down_shift,
                                    choke_label, fontsize=16, ha='center', va='center', fontname='Arial')
                        else:
                            # === SPECIAL CHOKE (relay or output not found) ===
                            special_choke = True
                            if str(output_type).strip().lower() == 'relay':
                                x_center = x_left + box_width/2 + 0.1
                            else:
                                x_center = x_left + box_width/2
                            if input_symbol_type == 'dual_fuse':
                                vertical_drop = 0.25
                                horizontal_extension = 0.3
                            else:
                                vertical_drop = 0.3
                                horizontal_extension = 0.0
                            base_y_position = y_bottom_bus_group
                            current_vert_x = draw_horizontal_choke_updated(
                                ax, x_center, base_y_position, label=choke_label,
                                box_width=box_width, box_height=box_height,
                                special_end=True, output_label=output_term,
                                output_type=output_type, output_text=output_text,
                                output_connected_terminals=output_connected_terminals,
                                x_positions=x_positions,
                                terminal_nos_for_positions=terminal_nos_for_positions,
                                y_top_bus_group=y_top_bus_group,
                                y_bottom_bus_group=y_bottom_bus_group
                            )
                            vert_x = current_vert_x
                # MOVED: Compute bus segments BEFORE drawing connections
                top_ranges = []
                bottom_ranges = []
                cable_headers_temp = df_header[df_header['cable_id'] == cable_id]
               
                # DEBUG: Print header info for this cable
                print(f"DEBUG: Cable {cable_id} has {len(cable_headers_temp)} header(s)")
               
                for _, hrow_temp in cable_headers_temp.iterrows():
                    header_type_temp = str(hrow_temp.get('header_type', '')).strip().upper()
                    terminal_start_temp = hrow_temp.get('terminal_start')
                    terminal_end_temp = hrow_temp.get('terminal_end', terminal_start_temp)
                    start_name_temp = str(terminal_start_temp).strip().replace('.0', '') if pd.notna(terminal_start_temp) else None
                    end_name_temp = str(terminal_end_temp).strip().replace('.0', '') if pd.notna(terminal_end_temp) else None
                   
                    print(f" Header: type={header_type_temp}, start={start_name_temp}, end={end_name_temp}")
                   
                    if pd.isna(start_name_temp) or pd.isna(end_name_temp):
                        print(f" Skipping: start or end is None")
                        continue
                       
                    if start_name_temp not in terminal_nos_for_positions:
                        print(f" Warning: start_name '{start_name_temp}' not found in terminal_nos_for_positions")
                        # Try to find without decimal
                        if '.' in start_name_temp:
                            start_name_temp_no_dec = start_name_temp.split('.')[0]
                            if start_name_temp_no_dec in terminal_nos_for_positions:
                                start_name_temp = start_name_temp_no_dec
                                print(f" Found without decimal: {start_name_temp}")
                            else:
                                continue
                        else:
                            continue
                   
                    if end_name_temp not in terminal_nos_for_positions:
                        print(f" Warning: end_name '{end_name_temp}' not found in terminal_nos_for_positions")
                        # Try to find without decimal
                        if '.' in end_name_temp:
                            end_name_temp_no_dec = end_name_temp.split('.')[0]
                            if end_name_temp_no_dec in terminal_nos_for_positions:
                                end_name_temp = end_name_temp_no_dec
                                print(f" Found without decimal: {end_name_temp}")
                            else:
                                continue
                        else:
                            continue
                   
                    start_idx_temp = terminal_nos_for_positions.index(start_name_temp)
                    end_idx_temp = terminal_nos_for_positions.index(end_name_temp)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                   
                    print(f" Adding range: {start_idx_temp} to {end_idx_temp} for {header_type_temp}")
                   
                    if header_type_temp == 'WIREFROM':
                        top_ranges.append((start_idx_temp, end_idx_temp))
                    elif header_type_temp == 'WIRETO':
                        bottom_ranges.append((start_idx_temp, end_idx_temp))
               
                print(f" Top ranges before merge: {top_ranges}")
                print(f" Bottom ranges before merge: {bottom_ranges}")
               
                # For multiple headers, we need to be careful about merging
                # If we have multiple non-overlapping headers, they should remain separate
                merge_adjacent = False # Changed to False to keep separate headers separate
               
                top_segments = merge_ranges(top_ranges, merge_adjacent=merge_adjacent)
                bottom_segments = merge_ranges(bottom_ranges, merge_adjacent=merge_adjacent)
               
                print(f" Top segments after merge: {top_segments}")
                print(f" Bottom segments after merge: {bottom_segments}")
               
                # Only create default segments if there are no headers at all
                if not top_segments and any(input_connected_flags):
                    top_segments = [(0, len(x_positions)-1)]
                    print(f" Created default top segment: {top_segments}")
                if not bottom_segments and any(output_connected_flags):
                    bottom_segments = [(0, len(x_positions)-1)]
                    print(f" Created default bottom segment: {bottom_segments}")
               
                print(f" Final top segments: {top_segments}")
                print(f" Final bottom segments: {bottom_segments}")
           
                hook_input_flags = []
                hook_output_flags = []
           
                # Determine if bus lines exist for this cable
                has_input_bus = len(top_segments) > 0
                has_output_bus = len(bottom_segments) > 0
           
                for j, x in enumerate(x_positions):
                    if j > 0 and x == x_positions[j-1]:
                        continue
                    symbol_top_y = capsule_y_center + SYMBOL_HEIGHT/2 + SYMBOL_RADIUS
                    symbol_bottom_y = capsule_y_center - SYMBOL_HEIGHT/2 - SYMBOL_RADIUS
               
                    # Pass the has_bus_line parameter to connection functions
                    hooked_in = draw_input_connection(ax, x, symbol_top_y, input_connected_flags[j], output_connected_flags[j], y_top_bus_group, has_input_bus)
                    hooked_out = draw_output_connection(ax, x, symbol_bottom_y, input_connected_flags[j], output_connected_flags[j], y_bottom_bus_group, has_output_bus)
               
                    hook_input_flags.append(hooked_in)
                    hook_output_flags.append(hooked_out)
                UPPER_SHIFT = 0.2
                for min_idx, max_idx in top_segments:
                    sub_x = x_positions[min_idx : max_idx + 1]
                    sub_flags = input_connected_flags[min_idx : max_idx + 1]
                    # SHIFT the top bus upward
                    draw_bus_lines(ax, sub_x, sub_flags, y_top_bus_group, gap=0.12, y_offset=UPPER_SHIFT)
                    connected_local = [i for i, f in enumerate(sub_flags) if f]
                    if connected_local:
                        first_local = connected_local[0]
                        x_first = sub_x[first_local]
                        # ALSO shift hook using same y_offset
                        y_bus = y_top_bus_group + UPPER_SHIFT
                        ax.plot([x_first - 0.3, x_first], [y_bus, y_bus], color='black', linewidth=1)
                        ax.plot([x_first - 0.3, x_first - 0.3], [y_bus, y_bus + 0.2], color='black', linewidth=1)
                LOWER_SHIFT = -0.2
                for min_idx, max_idx in bottom_segments:
                    sub_x = x_positions[min_idx : max_idx + 1]
                    sub_flags = output_connected_flags[min_idx : max_idx + 1]
                    # SHIFT the bottom bus downward
                    draw_bus_lines(ax, sub_x, sub_flags, y_bottom_bus_group, gap=0.12, y_offset=LOWER_SHIFT)
                    connected_local = [i for i, f in enumerate(sub_flags) if f]
                    if connected_local:
                        last_local = connected_local[-1]
                        x_last = sub_x[last_local]
                        if not ((special_choke and max_idx == start_idx and min_idx <= start_idx) or special_resistor):
                            # ALSO shift hook using same y_offset
                            y_bus = y_bottom_bus_group + LOWER_SHIFT
                            ax.plot([x_last, x_last + 0.3], [y_bus, y_bus], color='black', linewidth=1)
                            ax.plot([x_last + 0.3, x_last + 0.3], [y_bus, y_bus - 0.2], color='black', linewidth=1)
                cable_groups = df_group[df_group['cable_id'] == cable_id] if 'cable_id' in df_group.columns else pd.DataFrame()
                name_to_x = {}
                name_to_output_connected = {}
                name_to_input_connected = {}
                for idx, (xval, tname) in enumerate(zip(x_positions, terminal_nos_for_positions)):
                    if tname in name_to_x:
                        continue
                    name_to_x[tname] = xval
                    name_to_output_connected[tname] = output_connected_flags[idx] if idx < len(output_connected_flags) else False
                    name_to_input_connected[tname] = input_connected_flags[idx] if idx < len(input_connected_flags) else False
                x_min = min(x_positions) if x_positions else None
                x_max = max(x_positions) if x_positions else None
                if not cable_groups.empty:
                    min_bottom = min(symbol_bottoms) if symbol_bottoms else y_bottom_bus_group
                    x_start_pos = min(x_positions) if x_positions else current_x - pin_spacing
                    x_end_pos = max(x_positions) if x_positions else current_x
                    for _, grow in cable_groups.iterrows():
                        tn_field = grow.get('terminal_no')
                        start_name, end_name = parse_terminal_no_field(tn_field)
                        x_start_term = name_to_x.get(start_name, x_min)
                        x_end_term = name_to_x.get(end_name, x_max)
                        if x_start_term is None or x_end_term is None:
                            continue
                        label_text = grow.get('text', '')
                        io_field = str(grow.get('input_output', '')).strip().lower()
                       
                        # Get the actual indices for this specific group
                        if start_name in terminal_nos_for_positions and end_name in terminal_nos_for_positions:
                            start_idx_group = terminal_nos_for_positions.index(start_name)
                            end_idx_group = terminal_nos_for_positions.index(end_name)
                            if start_idx_group > end_idx_group:
                                start_idx_group, end_idx_group = end_idx_group, start_idx_group
                           
                            if io_field == 'input':
                                # MODIFIED: Move group top symbol downward when not connected + bus line
                                is_not_connected = not any(name_to_input_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_group:end_idx_group+1])
                                has_bus_line = len(top_segments) > 0
                                is_not_connected_with_bus = is_not_connected and has_bus_line
                           
                                # Adjust y position based on connection scenario
                                if is_not_connected_with_bus:
                                    y_relay = y_top_bus_group + 0.85 # Move further down for Scenario 2
                                else:
                                    y_relay = y_top_bus_group + 0.55 # Original position
                           
                                draw_relay_input(ax, x_start_term, x_end_term, y=y_relay, scale=1.0, text=str(label_text), is_not_connected_with_bus=is_not_connected_with_bus)
                            elif io_field == 'output':
                                # MODIFIED: Move group bottom symbol downward when not connected + bus line
                                is_not_connected = not any(name_to_output_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_group:end_idx_group+1])
                                has_bus_line = len(bottom_segments) > 0
                                is_not_connected_with_bus = is_not_connected and has_bus_line
                           
                                # Adjust y position based on connection scenario
                                if is_not_connected_with_bus:
                                    y_relay = y_bottom_bus_group - 0.85 # Move further down for Scenario 2
                                else:
                                    y_relay = y_bottom_bus_group - 0.55 # Original position
                           
                                draw_relay_output(ax, x_start_term, x_end_term, y=y_relay, scale=1.0, text=str(label_text), is_not_connected_with_bus=is_not_connected_with_bus)
                            else:
                                center_x = (x_start_term + x_end_term) / 2.0
                                ax.text(center_x, y_top_bus_group + 0.2, str(label_text), ha='center', va='bottom', fontsize=10, fontname='Arial')
                cable_headers = df_header[df_header['cable_id'] == cable_id]
                relay_top = {}
                relay_bottom = {}
                relay_box_top = {} # For RELAY_BOX header type at top
                relay_box_bottom = {} # For RELAY_BOX header type at bottom
                relay_contact_box_top = {} # NEW: For RELAY_CONTACT_BOX header type at top
                relay_contact_box_bottom = {} # NEW: For RELAY_CONTACT_BOX header type at bottom
            
                for _, hrow in cable_headers.iterrows():
                    header_type = str(hrow.get('header_type', '')).strip().upper()
                    terminal_start = hrow.get('terminal_start')
                    terminal_end = hrow.get('terminal_end', terminal_start)
                    input_output = str(hrow.get('input_output', '')).strip().lower()
                    text = hrow.get('text', '')
                    if pd.isna(text) or str(text).strip() == '':
                        text = ''
                    else:
                        text = str(text).strip()
                    start_name = str(terminal_start).strip().replace('.0', '') if pd.notna(terminal_start) else None
                    end_name = str(terminal_end).strip().replace('.0', '') if pd.notna(terminal_end) else None
                    if pd.isna(start_name) or pd.isna(end_name) or start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                
                    # Handle RELAY_CONTACT_BOX header type
                    if header_type == 'RELAY_CONTACT_BOX':
                        terminal_start_str = str(terminal_start).strip().replace('.0', '') if pd.notna(terminal_start) else None
                        terminal_end_str = str(terminal_end).strip().replace('.0', '') if pd.notna(terminal_end) else None
                        if terminal_start_str is None or terminal_end_str is None:
                            continue
                        key = (cable_id, terminal_start_str, terminal_end_str)
                        text = str(hrow.get('text', '')).strip()
                        input_output = str(hrow.get('input_output', '')).strip().lower()
                    
                        if input_output == 'input':
                            if key not in relay_contact_box_top:
                                relay_contact_box_top[key] = []
                            relay_contact_box_top[key].append(text)
                        elif input_output == 'output':
                            if key not in relay_contact_box_bottom:
                                relay_contact_box_bottom[key] = []
                            relay_contact_box_bottom[key].append(text)
                        continue
                    # Handle RELAY_BOX header type
                    elif header_type == 'RELAY_BOX':
                        terminal_start_str = str(terminal_start).strip().replace('.0', '') if pd.notna(terminal_start) else None
                        terminal_end_str = str(terminal_end).strip().replace('.0', '') if pd.notna(terminal_end) else None
                        if terminal_start_str is None or terminal_end_str is None:
                            continue
                        key = (cable_id, terminal_start_str, terminal_end_str)
                        text = str(hrow.get('text', '')).strip()
                        input_output = str(hrow.get('input_output', '')).strip().lower()
                    
                        if input_output == 'input':
                            if key not in relay_box_top:
                                relay_box_top[key] = []
                            relay_box_top[key].append(text)
                        elif input_output == 'output':
                            if key not in relay_box_bottom:
                                relay_box_bottom[key] = []
                            relay_box_bottom[key].append(text)
                        continue
                    elif header_type == 'RELAY':
                        terminal_start_str = str(terminal_start).strip().replace('.0', '') if pd.notna(terminal_start) else None
                        terminal_end_str = str(terminal_end).strip().replace('.0', '') if pd.notna(terminal_end) else None
                        if terminal_start_str is None or terminal_end_str is None:
                            continue
                        key = (cable_id, terminal_start_str, terminal_end_str)
                        text = str(hrow.get('text', '')).strip()
                        input_output = str(hrow.get('input_output', '')).strip().lower()
                        if input_output == 'input':
                            if key not in relay_top:
                                relay_top[key] = []
                            relay_top[key].append(text)
                        elif input_output == 'output':
                            if key not in relay_bottom:
                                relay_bottom[key] = []
                            relay_bottom[key].append(text)
                        continue
                    elif header_type in ['WIREFROM', 'WIRETO']:
                        min_symbol_bottom_local = min(symbol_bottoms[start_idx_temp:end_idx_temp+1]) if symbol_bottoms else None
                        if header_type == 'WIREFROM':
                            sub_flags = input_connected_flags[start_idx_temp:end_idx_temp+1]
                            connected_local = [i for i, f in enumerate(sub_flags) if f]
                            first_hook_x_specific = x_positions[start_idx_temp + connected_local[0]] if connected_local else x_left
                            draw_header(ax, cable_id, header_type, x_left, x_right, text,
                                        min_symbol_bottom=min_symbol_bottom_local,
                                        first_hook_x=first_hook_x_specific,
                                        last_hook_x=None,
                                        y_top_bus_group=y_top_bus_group,
                                        y_bottom_bus_group=y_bottom_bus_group)
                        elif header_type == 'WIRETO':
                            sub_flags = output_connected_flags[start_idx_temp:end_idx_temp+1]
                            connected_local = [i for i, f in enumerate(sub_flags) if f]
                            last_hook_x_specific = x_positions[start_idx_temp + connected_local[-1]] if connected_local else None
                            special_ha_local = False
                            if special_choke and end_idx_temp == start_idx:
                                last_hook_x_specific = vert_x
                                special_ha_local = True
                            draw_header(ax, cable_id, header_type, x_left, x_right, text,
                                        min_symbol_bottom=min_symbol_bottom_local,
                                        first_hook_x=None,
                                        last_hook_x=last_hook_x_specific,
                                        y_top_bus_group=y_top_bus_group,
                                        y_bottom_bus_group=y_bottom_bus_group,
                                        special_ha=special_ha_local)
            
                # NEW: Draw relay contact boxes at top
                for key, texts in relay_contact_box_top.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_top_y = capsule_y_center + SYMBOL_HEIGHT/2 + SYMBOL_RADIUS
                    input_conn_flag = any(name_to_input_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # Adjust vertical position based on connection scenario
                    is_not_connected = not input_conn_flag
                    has_bus_line = len(top_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_start = symbol_top_y + stub_length - 0.42
                    elif input_conn_flag:
                        vertical_line_start = y_top_bus_group
                    else:
                        vertical_line_start = symbol_top_y + stub_length
                    
                    draw_relay_contact_box_top(ax, x_left, x_right, vertical_line_start, texts=texts, scale=1.0, input_connected='Y' if input_conn_flag else 'N')
            
                # NEW: Draw relay contact boxes at bottom
                for key, texts in relay_contact_box_bottom.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_bottom_y = capsule_y_center - SYMBOL_HEIGHT/2 - SYMBOL_RADIUS
                    output_conn_flag = any(name_to_output_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # Adjust vertical position based on connection scenario
                    is_not_connected = not output_conn_flag
                    has_bus_line = len(bottom_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_end = symbol_bottom_y - stub_length + 0.38
                    elif output_conn_flag:
                        vertical_line_end = y_bottom_bus_group
                    else:
                        vertical_line_end = symbol_bottom_y - stub_length
                    
                    draw_relay_contact_box_bottom(ax, x_left, x_right, vertical_line_end, texts=texts, scale=1.0, output_connected='Y' if output_conn_flag else 'N')
            
                # Draw relay boxes at top
                for key, texts in relay_box_top.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_top_y = capsule_y_center + SYMBOL_HEIGHT/2 + SYMBOL_RADIUS
                    input_conn_flag = any(name_to_input_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # Adjust vertical position based on connection scenario
                    is_not_connected = not input_conn_flag
                    has_bus_line = len(top_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_start = symbol_top_y + stub_length - 0.42
                    elif input_conn_flag:
                        vertical_line_start = y_top_bus_group
                    else:
                        vertical_line_start = symbol_top_y + stub_length
                    
                    draw_relay_box_top(ax, x_left, x_right, vertical_line_start, texts=texts, scale=1.0, input_connected='Y' if input_conn_flag else 'N')
            
                # Draw relay boxes at bottom
                for key, texts in relay_box_bottom.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_bottom_y = capsule_y_center - SYMBOL_HEIGHT/2 - SYMBOL_RADIUS
                    output_conn_flag = any(name_to_output_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # Adjust vertical position based on connection scenario
                    is_not_connected = not output_conn_flag
                    has_bus_line = len(bottom_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_end = symbol_bottom_y - stub_length + 0.38
                    elif output_conn_flag:
                        vertical_line_end = y_bottom_bus_group
                    else:
                        vertical_line_end = symbol_bottom_y - stub_length
                    
                    draw_relay_box_bottom(ax, x_left, x_right, vertical_line_end, texts=texts, scale=1.0, output_connected='Y' if output_conn_flag else 'N')
            
                # Existing relay top processing
                for key, texts in relay_top.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_top_y = capsule_y_center + SYMBOL_HEIGHT/2 + SYMBOL_RADIUS
                    input_conn_flag = any(name_to_input_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # MODIFIED: Adjust vertical line start position for Scenario 2
                    is_not_connected = not input_conn_flag
                    has_bus_line = len(top_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_start = symbol_top_y + stub_length - 0.42 # Move further down for Scenario 2
                    elif input_conn_flag:
                        vertical_line_start = y_top_bus_group
                    else:
                        vertical_line_start = symbol_top_y + stub_length
                    
                    draw_group_top_symbol(ax, x_left, x_right, vertical_line_start, texts=texts, scale=1.0, input_connected='Y' if input_conn_flag else 'N')
            
                # Existing relay bottom processing
                for key, texts in relay_bottom.items():
                    if not texts:
                        continue
                    cid, start_name, end_name = key
                    if start_name not in terminal_nos_for_positions or end_name not in terminal_nos_for_positions:
                        continue
                    start_idx_temp = terminal_nos_for_positions.index(start_name)
                    end_idx_temp = terminal_nos_for_positions.index(end_name)
                    if start_idx_temp > end_idx_temp:
                        start_idx_temp, end_idx_temp = end_idx_temp, start_idx_temp
                    x_left = x_positions[start_idx_temp]
                    x_right = x_positions[end_idx_temp]
                    symbol_bottom_y = capsule_y_center - SYMBOL_HEIGHT/2 - SYMBOL_RADIUS
                    output_conn_flag = any(name_to_output_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_temp:end_idx_temp+1])
                
                    # MODIFIED: Adjust vertical line end position for Scenario 2
                    is_not_connected = not output_conn_flag
                    has_bus_line = len(bottom_segments) > 0
                    is_not_connected_with_bus = is_not_connected and has_bus_line
                
                    if is_not_connected_with_bus:
                        vertical_line_end = symbol_bottom_y - stub_length + 0.38 # Move further down for Scenario 2
                    elif output_conn_flag:
                        vertical_line_end = y_bottom_bus_group
                    else:
                        vertical_line_end = symbol_bottom_y - stub_length
                    
                    choke_output_terminal = None
                    choke_info = df_choke[df_choke['cable_id'] == cid]
                    if not choke_info.empty:
                        output_terminal = str(choke_info['output_terminal'].iloc[0]).strip()
                        if output_terminal.endswith('.0'):
                            output_terminal = output_terminal[:-2]
                        if output_terminal in [start_name, end_name]:
                            choke_output_terminal = output_terminal
                    draw_group_bottom_symbol(ax, x_left, x_right, vertical_line_end, texts=texts,
                                            scale=1.0, output_connected='Y' if output_conn_flag else 'N',
                                            choke_output_terminal=choke_output_terminal)
                all_x_positions.extend(x_positions)
                all_input_connected_flags.extend(input_connected_flags)
                all_output_connected_flags.extend(output_connected_flags)
                current_x += CABLE_GAP
                current_row_max_x = max(current_row_max_x, current_x)
                min_y = min(min_y, y_bottom_bus_group - 1.8)
                max_y = max(max_y, y_top_bus_group + 1.8)
       
        # FIXED: Move to next row AFTER processing all cables in current row
        y_offset -= vertical_gap
    
    overall_max_x = max(overall_max_x, current_row_max_x)
    if all_x_positions:
        content_min_x = min(all_x_positions)
        content_max_x = max(all_x_positions)
    else:
        content_min_x = start_x
        content_max_x = start_x + (page_width if page_width else fixed_fig_width)
    content_width = max(0.1, content_max_x - content_min_x)
    desired_width = global_max_width
    left = start_x - 1.5
    right = left + desired_width
    page_center_x = (left + right) / 2.0
    ax.set_xlim(left, right)
    ax.set_ylim(fixed_ylim_min, fixed_ylim_max)
   
    # Draw borders
    manual_y = -16.25 - 6.5 * (max_rows_for_ylim - 3)
    ax.plot([left - 2, right + 2], [manual_y, manual_y], 'k-', linewidth=1.0, zorder=10)
    manual_y = 9.4
    ax.plot([left - 2, right + 2], [manual_y, manual_y], 'k-', linewidth=1.0, zorder=10)
    x_vert = left + 0
    y_bottom = -16.65 - 6.5 * (max_rows_for_ylim - 3)
    y_top = 9.4
    ax.plot([x_vert, x_vert], [y_bottom, y_top], 'k-', linewidth=1, zorder=10)
    x_vert_right = right
    ax.plot([x_vert_right, x_vert_right], [y_bottom, y_top], 'k-', linewidth=1, zorder=10)
   
    # Draw junction box
    junction_box_y = CAPSULE_Y_CENTER_BASE + y_top_bus_offset +1.8 + 3.0 -1.0
    draw_junction_box(ax, page_center_x, junction_box_y, junction_name)
   
    return all_x_positions, all_input_connected_flags, all_output_connected_flags


def draw_footer(ax, left, right, fixed_ylim_min, total_pages, page_num, df_title_row, junction_name):
    """
    Draw a compact footer (title-block style) on `ax` occupying the right 50% of [left, right].
    All line widths are set to 1.3 and all text font sizes are set to 20 (as requested).
    """
    if df_title_row is None:
        return

    width = 20.0
    extra_width = 6.0
    height = 3.5  # adjusted footer height
    total_block_width = width + extra_width  # 26

    footer_width = (right - left) / 2.0
    footer_x_start = left + footer_width
    footer_y_start = fixed_ylim_min

    base_height = 3.0
    scale = height / base_height
    x_scale = footer_width / total_block_width

    LINEWIDTH = 1.3
    FONTSIZE = 17

    s = lambda y: y * scale

    outer_width = total_block_width * x_scale
    outer_height = height

    # Full outer rectangle
    ax.add_patch(Rectangle((footer_x_start, footer_y_start), outer_width, outer_height,
                           fill=False, linewidth=LINEWIDTH))

    # Vertical lines (major)
    v_lines = [8, 12, 17.5, width + extra_width / 2]
    for vx in v_lines:
        x = footer_x_start + vx * x_scale
        ax.plot([x, x], [footer_y_start, footer_y_start + outer_height], 'k-', linewidth=LINEWIDTH)

    # Special vertical at x=14.8
    x14_8 = footer_x_start + 14.8 * x_scale
    ax.plot([x14_8, x14_8], [footer_y_start + s(1.5), footer_y_start + outer_height], 'k-', linewidth=LINEWIDTH)

    # Horizontal lines
    horizontal_lines = [
        (0, 8, s(1.5)),
        (0, 8, s(1.0)),
        (12, 17.5, s(2.0)),
        (23, 26, s(0.7)),
        (12, 26, s(1.5)),
        (12, 26, s(2.5))
    ]
    for x0, x1, y in horizontal_lines:
        ax.plot([footer_x_start + x0 * x_scale, footer_x_start + x1 * x_scale],
                [footer_y_start + y, footer_y_start + y],
                'k-', linewidth=LINEWIDTH)

    # Small vertical lines
    ax.plot([footer_x_start + 4 * x_scale, footer_x_start + 4 * x_scale],
            [footer_y_start + 0.0, footer_y_start + s(1.5)], 'k-', linewidth=LINEWIDTH)

    ax.plot([footer_x_start + 25 * x_scale, footer_x_start + 25 * x_scale],
            [footer_y_start + s(-0.02), footer_y_start + s(1.5)], 'k-', linewidth=LINEWIDTH)

    # Company text
    company_y = footer_y_start + s(2.6) - 0.02
    ax.text(footer_x_start + 0.1 * x_scale, company_y,
            "SALTRIVER INFOSYSTEMS\nPRIVATE LTD.,\nAHMEDABAD.",
            va='top', ha='left', fontsize=22, fontname='Arial', linespacing=1.2)

    # Labels
    ax.text(footer_x_start + 1.1 * x_scale, footer_y_start + s(0.5),
            "DRAWN BY", va='center', ha='left', fontsize=22, fontname='Arial')

    ax.text(footer_x_start + 5.1 * x_scale, footer_y_start + s(0.5),
            "CHECKED BY", va='center', ha='left', fontsize=22, fontname='Arial')

    # Values
    drawn_by = df_title_row.get('drawn_by')
    checked_by = df_title_row.get('checked_by')

    if pd.notna(drawn_by):
        ax.text(footer_x_start + 1.4 * x_scale, footer_y_start + s(1.2),
                str(drawn_by), va='bottom', ha='left', fontsize=FONTSIZE, fontname='Arial')

    if pd.notna(checked_by):
        ax.text(footer_x_start + 5.4 * x_scale, footer_y_start + s(1.2),
                str(checked_by), va='bottom', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Designations
    ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(2.75),
            str(df_title_row.get('designation1', '')), va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')
    ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(2.25),
            str(df_title_row.get('designation2', '')), va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')
    ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(1.82),
            str(df_title_row.get('designation3', '')), va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Station info
    ax.text(footer_x_start + 19 * x_scale, footer_y_start + s(2.75),
            str(df_title_row.get('station_name', '')), va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    ax.text(footer_x_start + 18 * x_scale, footer_y_start + s(2.0),
            junction_name, va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    ax.text(footer_x_start + 18.5 * x_scale, footer_y_start + s(0.9),
            f"DRG. NO. {df_title_row.get('station_code', '')}", va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Formatter
    def format_text(t):
        t = str(t)
        if len(t) > 12:
            words = t.split()
            if len(words) > 1:
                return " ".join(words[:-1]) + "\n" + words[-1]
            return t[:12] + "\n" + t[12:]
        return t

    # Zone & Division
    ax.text(footer_x_start + 24 * x_scale, footer_y_start + s(2.8),
            format_text(df_title_row.get('zone', '')),
            va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    ax.text(footer_x_start + 23.3 * x_scale, footer_y_start + s(2.3),
            format_text(f"{df_title_row.get('division', '')} DIVISION"),
            va='top', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Total Pages
    ax.text(footer_x_start + 25.25 * x_scale, footer_y_start + s(0.25),
            str(total_pages), va='bottom', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Sheet labels
    ax.text(footer_x_start + (width + 3.6) * x_scale, footer_y_start + s(1.1),
            "SHEET\nNO", va='center', ha='left', fontsize=FONTSIZE,
            fontname='Arial', linespacing=1.0)

    ax.text(footer_x_start + (width + 3.5) * x_scale, footer_y_start + s(0.35),
            "TOTAL\nSHEETS", va='center', ha='left', fontsize=FONTSIZE,
            fontname='Arial', linespacing=1.0)

    ax.text(footer_x_start + (width + 5.25) * x_scale, footer_y_start + s(1.1),
            str(page_num), va='center', ha='left', fontsize=FONTSIZE, fontname='Arial')

    # Date
    # if pd.notna(df_title_row.get('date')):
    # ax.text(footer_x_start + (width + 5.5) * x_scale, footer_y_start + s(0.35),
    # str(df_title_row.get('date')), va='center', ha='left', fontsize=FONTSIZE)

# === NEW FUNCTIONS FOR PAGINATION AND ROW ORGANIZATION ===

def organize_junction_rows(junction_cables_regular, junction_cables_box):
    """
    Organize junction rows with proper sorting and grouping
    """
    # Combine regular cables and cable boxes if both exist
    if not junction_cables_regular.empty and not junction_cables_box.empty:
        junction_cables = pd.concat([junction_cables_regular, junction_cables_box], ignore_index=True)
    elif not junction_cables_regular.empty:
        junction_cables = junction_cables_regular.copy()
    elif not junction_cables_box.empty:
        junction_cables = junction_cables_box.copy()
    else:
        return OrderedDict()  # Return empty OrderedDict
    
    # Enhanced sorting with proper row ordering
    junction_cables = junction_cables.sort_values(
        by=['row', 'position'], 
        key=lambda x: x.map(get_row_order) if x.name == 'row' else x,
        ascending=[False, False]  # Row descending, position descending
    )
    
    # Group by row letter
    letter_groups = OrderedDict()
    
    for _, cable_row in junction_cables.iterrows():
        letter = str(cable_row.get('row', '')).strip()
        if not letter:  # Skip if row is empty
            continue
        cable_id = cable_row['cable_id']
        
        if letter not in letter_groups:
            letter_groups[letter] = []
        letter_groups[letter].append(cable_id)
    
    return letter_groups

def create_pages_for_junction(junction, letter_groups, max_rows_per_page=3):
    """
    Create pages for a junction with exactly max_rows_per_page rows per page
    """
    # Sort letters using the enhanced row ordering function
    sorted_letters = sorted(letter_groups.keys(), key=get_row_order, reverse=True)
    
    print(f"\n=== Processing Junction: {junction} ===")
    print(f"All letters in junction: {sorted_letters}")
    print(f"Letter groups: {[(letter, len(cables)) for letter, cables in letter_groups.items()]}")
    
    # Break each letter group into multiple rows if needed with cable box limit
    all_rows = []
    for letter in sorted_letters:
        cable_list = letter_groups[letter]
        letter_rows = break_cables_into_rows_updated(cable_list, max_terminal_symbols_per_row=36, max_cable_boxes_per_row=6)
        
        # REVERSE THE CABLE ORDER WITHIN EACH ROW to get left-to-right descending
        for i, row_cables in enumerate(letter_rows):
            reversed_row_cables = list(reversed(row_cables))  # This makes highest positions on left
            all_rows.append((letter, reversed_row_cables))
    
    # === FIXED PAGINATION LOGIC ===
    # Create pages with exactly max_rows_per_page rows per page
    # But distribute rows properly across pages
    pages = []
    
    # Reverse the list to process from bottom to top
    all_rows_reversed = list(reversed(all_rows))
    
    # Break into chunks of max_rows_per_page
    chunks = [all_rows_reversed[i:i+max_rows_per_page] for i in range(0, len(all_rows_reversed), max_rows_per_page)]
    
    # Reverse the chunks to maintain proper order
    chunks_reversed = list(reversed(chunks))
    
    # Now reverse each chunk to get the original order of rows in the chunk
    for chunk in chunks_reversed:
        chunk_reversed = list(reversed(chunk))
        pages.append((junction, chunk_reversed))
    
    return pages

# === STANDARDIZED DIMENSIONS ===
SYMBOL_HEIGHT = 0.6
SYMBOL_WIDTH = 0.35
SYMBOL_RADIUS = 0.15
CAPSULE_Y_CENTER_BASE = 3.6
y_top_bus_offset = 1.3
y_bottom_bus_offset = -1.3
stub_length = 0.74
CABLE_GAP = 1.5
vertical_gap = 6.5
# Footer dimensions (adjusted to match row spacing)
footer_height = 2.75 # Adjusted footer height
footer_inch_add = 4.0 # Reduced additional inches for footer

# ADD THIS LINE - Define pin_spacing
pin_spacing = 0.8

# === Load Excel file path from command line (or prompt) ===
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
else:
    EXCEL_FILE = input("Enter Excel file path (e.g. C:\\Diagram\\RAILWAYPROJECT.xlsx) or press Enter to exit: ").strip()
    if not EXCEL_FILE:
        print("No Excel file provided. Exiting.")
        sys.exit(1)

if not os.path.exists(EXCEL_FILE):
    print(f"Error: Excel file not found at: {EXCEL_FILE}")
    sys.exit(1)

# Validate required sheets exist
try:
    xls = pd.ExcelFile(EXCEL_FILE)
except Exception as e:
    print(f"Unable to open Excel file: {e}")
    sys.exit(1)

required_sheets = ['terminal', 'junction_box', 'terminal_header', 'group', 'cable']
available_sheets = [s.strip() for s in xls.sheet_names]
missing = [s for s in required_sheets if s not in available_sheets]
if missing:
    print(f"Excel file is missing required sheets: {missing}")
    print(f"Available sheets: {available_sheets}")
    sys.exit(1)

# === UPDATED: Load both regular cables and cable boxes ===
try:
    df_cable = pd.read_excel(EXCEL_FILE, sheet_name='cable')
    df_cable.columns = df_cable.columns.str.strip()
    
    # Load cable boxes from separate sheet (if exists) or filter from main cable sheet
    try:
        df_cable_box = pd.read_excel(EXCEL_FILE, sheet_name='cable_box')
        df_cable_box.columns = df_cable_box.columns.str.strip()
        print("Loaded relay boxes from separate 'relay_box' sheet")
    except Exception as e:
        # Fallback: extract cable boxes from main cable sheet
        print(f"No separate 'relay_box' sheet found: {e}. Creating empty relay_box DataFrame.")
        
        # Create empty DataFrame with required structure
        df_cable_box = pd.DataFrame(columns=['cable_id', 'cable_name', 'junction_name', 'row', 
                                           'position', 'terminal', 'start_no', 'cabel_type', 
                                           'cable_letter', 'letter_order'])
        
        print("Created empty relay_box DataFrame.")
        
except Exception as e:
    print(f"Error reading cable sheets from Excel file: {e}")
    sys.exit(1)

# Load other required sheets
try:
    df = pd.read_excel(EXCEL_FILE, sheet_name='terminal')
    df.columns = df.columns.str.strip()
    # Replace terminal_name with terminal_no in the terminal sheet
    if 'terminal_name' in df.columns:
        df.rename(columns={'terminal_name': 'terminal_no'}, inplace=True)
    df_junction = pd.read_excel(EXCEL_FILE, sheet_name='junction_box')
    df_junction.columns = df_junction.columns.str.strip()
    df_header = pd.read_excel(EXCEL_FILE, sheet_name='terminal_header')
    df_header.columns = df_header.columns.str.strip()
    df_group = pd.read_excel(EXCEL_FILE, sheet_name='group')
    df_group.columns = df_group.columns.str.strip()
    df_choke = pd.read_excel(EXCEL_FILE, sheet_name='choketable')
    df_choke.columns = df_choke.columns.str.strip()
    df_resistor = pd.read_excel(EXCEL_FILE, sheet_name='resistortable')
    df_resistor.columns = df_resistor.columns.str.strip()
    
    # Load StationDrawing for footer if available
    df_title = None
    try:
        df_title = pd.read_excel(EXCEL_FILE, sheet_name='StationDrawing')
        df_title.columns = df_title.columns.str.strip()
        print("Loaded StationDrawing sheet for footer.")
    except Exception as e:
        print(f"Warning: Could not load StationDrawing sheet for footer: {e}. Footer will be skipped.")
        
except Exception as e:
    print(f"Error reading required sheets from Excel file: {e}")
    sys.exit(1)
finally:
    try:
        xls.close()
    except Exception:
        pass

if 'spare' in df.columns:
    df.loc[df['spare'].astype(str).str.upper() == 'Y', 'input_left'] = 'SP'

# === Prepare Plotting ===
df_symbols = df.reset_index(drop=True) # Include all rows, even with blank or unknown symbols

# === SAFETY CHECK: Ensure df_cable_box has required columns ===
if df_cable_box is None or df_cable_box.empty:
    print("No cable box data found. Creating empty DataFrame with required structure.")
    df_cable_box = pd.DataFrame(columns=['cable_id', 'cable_name', 'junction_name', 'row', 
                                       'position', 'terminal', 'start_no', 'cabel_type', 
                                       'cable_letter', 'letter_order'])
else:
    # Add missing columns if they don't exist
    required_columns = ['row', 'cable_letter', 'letter_order', 'junction_name']
    for col in required_columns:
        if col not in df_cable_box.columns:
            if col == 'cable_letter':
                df_cable_box[col] = ''
            elif col == 'letter_order':
                df_cable_box[col] = -1
            else:
                df_cable_box[col] = pd.NA

# === Prepare cable data with separate cable boxes ===
# For regular cables, use 'row' column directly for cable_letter
df_cable['cable_letter'] = df_cable['row'].astype(str).str.strip()
df_cable['letter_order'] = df_cable['cable_letter'].apply(
    lambda x: ord(x.upper()) - ord('A') if pd.notna(x) and x.strip() != '' and x.strip() in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' else -1
)

# For cable boxes, handle only if DataFrame is not empty
if not df_cable_box.empty and 'row' in df_cable_box.columns:
    df_cable_box['cable_letter'] = df_cable_box['row'].astype(str).str.strip()
    df_cable_box['letter_order'] = df_cable_box['cable_letter'].apply(
        lambda x: ord(x.upper()) - ord('A') if pd.notna(x) and x.strip() != '' and x.strip() in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' else -1
    )
else:
    # Ensure the columns exist even if empty
    if 'cable_letter' not in df_cable_box.columns:
        df_cable_box['cable_letter'] = ''
    if 'letter_order' not in df_cable_box.columns:
        df_cable_box['letter_order'] = -1

# === FIXED: Get unique junction names in sheet-order (preserve first-seen order) ===
junction_names = pd.unique(df_cable['junction_name'].astype(str).str.strip())

# === FIXED: Check for cable boxes in the cable sheet and add them to junction_names ===
if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
    cable_box_junctions = pd.unique(df_cable_box['junction_name'].astype(str).str.strip())
    for junction in cable_box_junctions:
        if junction not in junction_names:
            junction_names = np.append(junction_names, junction)
else:
    print("No cable box data to process for junction names.")

print(f"DEBUG: Found {len(junction_names)} junctions: {list(junction_names)}")

# === FIXED PAGINATION LOGIC FOR SEPARATE CABLE BOXES ===
# Compute max_row_width for each junction to determine page size
junction_row_widths = {}
for junction in junction_names:
    current_x_pre = 1
    current_row_max_x_pre = 1
    current_terminal_count_pre = 0
    current_cable_box_count_pre = 0
    current_letter = None
    max_row_width = 0
    
    # Get regular cables for this junction
    junction_mask_regular = df_cable['junction_name'].astype(str).str.strip() == junction
    junction_cables_regular = df_cable[junction_mask_regular].copy()
    
    # Get cable boxes for this junction
    junction_mask_box = df_cable_box['junction_name'].astype(str).str.strip() == junction if not df_cable_box.empty else pd.Series([False] * len(df_cable_box))
    junction_cables_box = df_cable_box[junction_mask_box].copy()
    
    # Combine regular cables and cable boxes
    junction_cables = pd.concat([junction_cables_regular, junction_cables_box], ignore_index=True)
    
    # FIXED: Enhanced sorting with proper row ordering
    junction_cables = junction_cables.sort_values(
        by=['row', 'position'], 
        key=lambda x: x.map(get_row_order) if x.name == 'row' else x,
        ascending=[False, False]  # Row descending, position descending
    )
    
    cable_list = junction_cables['cable_id'].tolist()
    
    print(f"\n=== Junction: {junction} ===")
    print("Cable order for width calculation:")
    for cable_id in cable_list:
        # Check both regular cables and cable boxes
        cable_row_regular = df_cable[df_cable['cable_id'] == cable_id]
        cable_row_box = df_cable_box[df_cable_box['cable_id'] == cable_id] if not df_cable_box.empty else pd.DataFrame()
        
        if not cable_row_regular.empty:
            cable_row = cable_row_regular.iloc[0]
            cable_type = "regular"
        elif not cable_row_box.empty:
            cable_row = cable_row_box.iloc[0]
            cable_type = "cable_box"
        else:
            continue
            
        print(f"  Cable {cable_id} ({cable_type}): Row '{cable_row.get('row', '')}', Position {cable_row.get('position', '')}")
    
    for cable_id_pre in cable_list:
        # Check both regular cables and cable boxes for row letter
        r_regular = df_cable[df_cable['cable_id'] == cable_id_pre]
        r_box = df_cable_box[df_cable_box['cable_id'] == cable_id_pre] if not df_cable_box.empty else pd.DataFrame()
        
        is_cable_box = not r_box.empty
        
        if is_cable_box:
            r = r_box
            letter = r['row'].iloc[0] if not r.empty and 'row' in r.columns else ""
        else:
            r = r_regular
            letter = r['row'].iloc[0] if not r.empty and 'row' in r.columns else ""
            
        if letter != current_letter and current_terminal_count_pre > 0:
            max_row_width = max(max_row_width, current_row_max_x_pre - 1)
            current_row_max_x_pre = 1
            current_x_pre = 1
            current_terminal_count_pre = 0
            current_cable_box_count_pre = 0
        current_letter = letter
        
        # STEP 5: Updated width calculation for cable boxes
        # Check if this is a cable box
        cable_box_rows = df_cable_box[df_cable_box['cable_id'] == cable_id_pre] if not df_cable_box.empty else pd.DataFrame()
        if not cable_box_rows.empty:
            # Cable box - use position for terminal count
            cable_info = cable_box_rows.iloc[0]
            position_val = cable_info.get('position')
            if pd.notna(position_val):
                try:
                    total_terminals = int(float(position_val))
                except:
                    total_terminals = 1
            else:
                total_terminals = 1
            added_width = total_terminals * pin_spacing
        else:
            # Regular cable - calculate from symbols
            group_pre = df_symbols[df_symbols['cable_id'] == cable_id_pre].sort_index().reset_index(drop=True)
            total_terminals = 0
            added_width = 0
            i = 0
            while i < len(group_pre):
                symbol = str(group_pre.iloc[i].get('symbol', '')).strip().lower()
                if symbol == 'dual_fuse':
                    if i + 1 < len(group_pre):
                        added_width += pin_spacing * 1.0 + pin_spacing * 1.5
                        total_terminals += 2
                        i += 2
                    else:
                        added_width += pin_spacing
                        total_terminals += 1
                        i += 1
                else:
                    added_width += pin_spacing
                    total_terminals += 1
                    i += 1
        
        # Check if we need to break the row due to cable box limit
        if is_cable_box and current_cable_box_count_pre >= 6 and current_terminal_count_pre > 0:
            max_row_width = max(max_row_width, current_row_max_x_pre - 1)
            current_row_max_x_pre = 1
            current_x_pre = 1
            current_terminal_count_pre = 0
            current_cable_box_count_pre = 0
        
        # If adding this cable would exceed the terminal limit, start a new row
        if current_terminal_count_pre + total_terminals > 36 and current_terminal_count_pre > 0:
            max_row_width = max(max_row_width, current_row_max_x_pre - 1)
            current_row_max_x_pre = 1
            current_x_pre = 1
            current_terminal_count_pre = 0
            current_cable_box_count_pre = 0
        
        current_x_pre += added_width + CABLE_GAP
        current_row_max_x_pre = max(current_row_max_x_pre, current_x_pre)
        current_terminal_count_pre += total_terminals
        if is_cable_box:
            current_cable_box_count_pre += 1
            
    max_row_width = max(max_row_width, current_row_max_x_pre - 1)
    junction_row_widths[junction] = max_row_width

global_max_width = max(junction_row_widths.values()) + 2.0 if junction_row_widths else 30.0

# === FIXED: Compute page dimensions ===
max_rows_visible = 3
max_terminal_symbols_per_row = 36
max_cable_boxes_per_row = 6

fixed_fig_width = 42.8
fixed_fig_height = 31.0
bottom_margin = 1.0
top_margin = 3.0
fixed_ylim_min = CAPSULE_Y_CENTER_BASE + vertical_gap * (1 - max_rows_visible) + y_bottom_bus_offset - 1.8 - bottom_margin - footer_height
fixed_ylim_max = CAPSULE_Y_CENTER_BASE + y_top_bus_offset + 1.8 + top_margin

# === UPDATED PAGINATION LOGIC ===
pages = []
for junction in junction_names:
    junction_mask = df_cable['junction_name'].astype(str).str.strip() == junction
    junction_cables_regular = df_cable[junction_mask].copy()
    
    # Get cable boxes for this junction - only if DataFrame exists and has data
    junction_cables_box = pd.DataFrame()
    if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
        junction_mask_box = df_cable_box['junction_name'].astype(str).str.strip() == junction
        junction_cables_box = df_cable_box[junction_mask_box].copy()
    else:
        print(f"No cable box data for junction: {junction}")
    
    # Check if we have any data to process
    if junction_cables_regular.empty and junction_cables_box.empty:
        print(f"Warning: No data found for junction '{junction}'. Skipping.")
        continue
    
    # Organize rows with proper sorting
    letter_groups = organize_junction_rows(junction_cables_regular, junction_cables_box)
    
    # Create pages for this junction
    junction_pages = create_pages_for_junction(junction, letter_groups, max_rows_per_page=3)
    pages.extend(junction_pages)

total_pages = len(pages)
print(f"\nTotal pages: {total_pages}")

# Print page details for debugging
print("\n=== PAGE BREAKDOWN ===")
for i, (junction, page_rows) in enumerate(pages, 1):
    row_info = [f"{letter}({len(cables)} cables)" for letter, cables in page_rows]
    print(f"Page {i}: {junction} - Rows: {row_info}")

title_row = df_title.iloc[0] if df_title is not None and not df_title.empty else None

# Generate PDF with fixed dimensions
output_base = 'Terminal_Symbols_Centered_Fixed_Size'
# Generate timestamp suffix in UTC for uniqueness
ist_tz = timezone('Asia/Kolkata')
current_time = datetime.now(ist_tz)
timestamp_suffix = current_time.strftime("%Y-%m-%d_%H-%M-%S")
output_file = f"{output_base}_{timestamp_suffix}.pdf"

# Generate initial checksum with the unique timestamped filename
checksum, checksum_data, content_size, timestamp_ist = generate_pdf_metadata_checksum(df_title, EXCEL_FILE, output_file)

if os.path.exists(output_file):
    print(f"Warning: '{output_file}' already exists (timestamp collision). Adding checksum suffix.")
    short_checksum = checksum[:8] # First 8 chars for brevity
    output_file = f"{output_base}_{timestamp_suffix}_{short_checksum}.pdf"
    # Regenerate checksum with new filename
    checksum, checksum_data, content_size, timestamp_ist = generate_pdf_metadata_checksum(df_title, EXCEL_FILE, output_file)

# Generate PDF with metadata
with PdfPages(output_file) as pdf:
    for page_num, (junction_name, page_rows) in enumerate(pages, 1):
        # Calculate page-specific dimensions
        page_max_width = 0
        current_x_page = 1
        current_row_max_x_page = 1
        current_terminal_count_page = 0
        current_letter = None
    
        # Calculate page width for centering - iterate through page_rows instead of page_cable_ids
        for letter, cable_list in page_rows:
            for cid in cable_list:
                # Check if cable box
                r_box = df_cable_box[df_cable_box['cable_id'] == cid] if not df_cable_box.empty else pd.DataFrame()
                is_cable_box = not r_box.empty
                
                if is_cable_box:
                    r = r_box
                    letter = r['row'].iloc[0] if not r.empty and 'row' in r.columns else ""
                else:
                    r = df_cable[df_cable['cable_id'] == cid]
                    letter = r['row'].iloc[0] if not r.empty and 'row' in r.columns else ""
                    
                if letter != current_letter and current_terminal_count_page > 0:
                    page_max_width = max(page_max_width, current_row_max_x_page - 1)
                    current_row_max_x_page = 1
                    current_x_page = 1
                    current_terminal_count_page = 0
                current_letter = letter
            
                group = df_symbols[df_symbols['cable_id'] == cid].sort_index().reset_index(drop=True)
                total_terminals = 0
                added_width = 0
                i = 0
                while i < len(group):
                    symbol = str(group.iloc[i].get('symbol', '')).strip().lower()
                    if symbol == 'dual_fuse':
                        if i + 1 < len(group):
                            added_width += pin_spacing * 1.0 + pin_spacing * 1.5
                            total_terminals += 2
                            i += 2
                        else:
                            added_width += pin_spacing
                            total_terminals += 1
                            i += 1
                    else:
                        added_width += pin_spacing
                        total_terminals += 1
                        i += 1
                    
                if current_terminal_count_page + total_terminals > 36:
                    page_max_width = max(page_max_width, current_row_max_x_page - 1)
                    current_row_max_x_page = 1
                    current_x_page = 1
                    current_terminal_count_page = 0
                
                current_x_page += added_width + CABLE_GAP
                current_row_max_x_page = max(current_row_max_x_page, current_x_page)
                current_terminal_count_page += total_terminals
        
        page_max_width = max(page_max_width, current_row_max_x_page - 1)
    
        # Calculate page_start_x for centering
        shift = (global_max_width - (page_max_width + 1.2)) / 2
        page_start_x = 1 + shift
      
        # Set PDF metadata with checksum
        pdf_info = pdf.infodict()
        pdf_info['Title'] = f'Terminal Drawing - {junction_name}'
        pdf_info['Author'] = 'SaltRiver Infosystems'
        pdf_info['Subject'] = f'Station Code: {df_title.iloc[0].get("station_code", "") if df_title is not None else ""}'
        pdf_info['Keywords'] = f'checksum:{checksum}' if checksum else ''
        pdf_info['CreationDate'] = current_time.strftime("D:%Y%m%d%H%M%S+05'30'") # Use IST with offset for CreationDate
      
        # Create figure
        fixed_fig_width = 42.8
        fixed_fig_height = 31.0
        fig, ax = plt.subplots(figsize=(fixed_fig_width, fixed_fig_height))
        ax.set_facecolor('white')
        ax.axis('off')
      
        # Draw symbols - UPDATED: pass page_rows instead of page_cable_ids
        # Also adjust y_offset based on number of rows
        y_offset = - (max_rows_visible - len(page_rows)) * vertical_gap
        
        x_positions, input_connected_flags, output_connected_flags = draw_symbols(
            df_symbols, ax, page_rows, junction_name,  # Note: page_rows instead of page_cable_ids
            start_x=page_start_x,
            pin_spacing=pin_spacing,
            cables_per_page=sum(len(cables) for _, cables in page_rows),  # Total cables in all rows
            page_number=page_num,
            max_terminal_symbols_per_row=36,
            max_rows_visible=3,
            page_width=global_max_width
        )
      
        # Draw footer
        left = page_start_x - 1.5
        right = left + global_max_width
        max_rows_visible = 3
        bottom_margin = 1.0
        top_margin = 3.0
        fixed_ylim_min = CAPSULE_Y_CENTER_BASE + vertical_gap * (1 - max_rows_visible) + y_bottom_bus_offset - 1.8 - bottom_margin - footer_height
        draw_footer(ax, left, right, fixed_ylim_min, total_pages, page_num, title_row, junction_name)
      
        fig.subplots_adjust(left=0.04, right=0.99, top=0.98, bottom=0.02)
        pdf.savefig(fig, dpi=300, facecolor='white')
        plt.close(fig)
        print(f"Page {page_num} (Junction: {junction_name}) added to '{output_file}'")
  
    # After all pages are written, update with actual file size
    final_checksum, final_checksum_data, final_content_size = update_pdf_checksum_metadata(
        output_file, checksum, checksum_data, content_size, df_title
    )

print(f"Multi-page PDF saved as '{output_file}' (with date-time suffix for uniqueness).")

# Enhance PDF with proper metadata if PyPDF2 is available
try:
    import PyPDF2
    enhance_pdf_with_metadata(output_file, final_checksum, final_checksum_data, final_content_size, df_title)
except ImportError:
    print("PyPDF2 not available. Using basic PDF metadata.")

print("Code updated successfully with proper pagination and row ordering!")