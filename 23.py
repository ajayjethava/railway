#!/usr/bin/env python3
"""
Railway Project Terminal Diagram Generator
24x7 Directory Monitor with Auto File Processing
Updated with font fixes and robust error handling
"""

import matplotlib
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
import psycopg2
from psycopg2 import sql
import traceback
import time
import shutil
import signal
import threading

# === FONT FIX: Configure matplotlib to use available fonts ===
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Liberation Sans', 'Arial', 'Helvetica']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['pdf.fonttype'] = 42  # Output Type 42 (TrueType) fonts

# Global font constant
DEFAULT_FONT = 'sans-serif'

# === GLOBAL FLAG FOR GRACEFUL SHUTDOWN ===
running = True

def signal_handler(sig, frame):
    """Handle shutdown signals gracefully"""
    global running
    print("\n" + "="*60)
    print("SHUTDOWN SIGNAL RECEIVED")
    print("Waiting for current processing to complete...")
    print("="*60)
    running = False

# Register signal handlers
signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# === CONFIGURATION ===
XLSX_INPUT_DIR = "/root/srv/local/git/xlsx_download"
PDF_OUTPUT_DIR = "/root/srv/local/git/uploads"
PROCESSED_EXCEL_DIR = "/root/srv/local/git/processed_excel"
ERROR_EXCEL_DIR = "/root/srv/local/git/error_excel"

# Create directories if they don't exist
os.makedirs(XLSX_INPUT_DIR, exist_ok=True)
os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
os.makedirs(PROCESSED_EXCEL_DIR, exist_ok=True)
os.makedirs(ERROR_EXCEL_DIR, exist_ok=True)

# Database configuration
DB_CONFIG = {
    "host": "pso.cellapps.com",
    "port": "5432",
    "database": "postgrestest",
    "user": "postgres",
    "password": "Omhari@8899"
}

# === HELPER FUNCTIONS ===
def get_latest_xlsx_file(directory):
    """
    Get the most recent XLSX file from the specified directory
    """
    try:
        # Get all XLSX files in the directory
        xlsx_files = [f for f in os.listdir(directory) if f.lower().endswith('.xlsx')]
        
        if not xlsx_files:
            print(f"No XLSX files found in {directory}")
            return None
        
        # Get the latest file based on modification time
        latest_file = max(xlsx_files, key=lambda f: os.path.getmtime(os.path.join(directory, f)))
        latest_path = os.path.join(directory, latest_file)
        
        print(f"Found XLSX file: {latest_file}")
        print(f"Last modified: {datetime.fromtimestamp(os.path.getmtime(latest_path))}")
        
        return latest_path
        
    except Exception as e:
        print(f"Error finding latest XLSX file: {e}")
        traceback.print_exc()
        return None

def get_db_connection():
    """Establish database connection"""
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        return conn
    except Exception as e:
        print(f"Database connection error: {e}")
        traceback.print_exc()
        return None

def store_pdf_metadata(db_conn, pdf_metadata):
    """
    Store PDF metadata in generatedpdf table
    """
    try:
        cursor = db_conn.cursor()
        
        # Parse project_id from xlsx_filename
        xlsx_filename = pdf_metadata.get('xlsx_filename', '')
        project_id = None
        
        if xlsx_filename:
            # Multiple patterns to try for extracting project_id
            patterns = [
                r'RAILWAYPROJECT_ID(\d+)',
                r'PROJECT_(\d+)',
                r'(?i)project[_-](\d+)',
                r'ID(\d+)',
                r'_(\d+)_'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, xlsx_filename)
                if match:
                    try:
                        project_id = int(match.group(1))
                        print(f"Extracted project_id: {project_id} from filename: {xlsx_filename}")
                        break
                    except ValueError:
                        continue
        
        # If still None, try to get from metadata or use default
        if project_id is None:
            project_id = pdf_metadata.get('project_id')
            if project_id is None:
                # Try to extract from filename using simpler method
                numbers = re.findall(r'\d+', xlsx_filename)
                if numbers:
                    try:
                        for num in numbers:
                            if len(num) >= 2:
                                project_id = int(num)
                                print(f"Using fallback project_id: {project_id} from numbers in filename")
                                break
                    except ValueError:
                        pass
            
            if project_id is None:
                project_id = 0
                print(f"Warning: Could not extract project_id from filename. Using default: {project_id}")
        
        # Prepare junction data JSON
        junction_data_json = json.dumps(pdf_metadata.get('junction_data', []))
        
        # SQL query
        if project_id is not None:
            insert_query = """
            INSERT INTO generated_pdf(
                project_id, pdf_filename, xlsx_filename, file_size,
                checksum_md5, metadata_checksum, metadata_data,
                initial_size_bytes, final_size_bytes, metadata_ts_ist,
                station_code, source_pdf_name, full_file_md5, remarks,
                checksum_algo, created_at, level1_status, level2_status,
                level3_status, version, junction_data
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s
            )
            """
            
            values = (
                project_id,
                pdf_metadata.get('pdf_filename'),
                xlsx_filename,
                pdf_metadata.get('final_size_bytes'),
                pdf_metadata.get('full_file_md5'),
                pdf_metadata.get('metadata_checksum'),
                pdf_metadata.get('metadata_data'),
                pdf_metadata.get('initial_size_bytes'),
                pdf_metadata.get('final_size_bytes'),
                pdf_metadata.get('metadata_ts_ist'),
                pdf_metadata.get('station_code'),
                pdf_metadata.get('pdf_filename'),
                pdf_metadata.get('full_file_md5'),
                'Generated by Python script',
                'md5',
                datetime.now(timezone('Asia/Kolkata')).strftime("%Y-%m-%d %H:%M:%S.%f"),
                'pending',
                'pending',
                'pending',
                1,
                junction_data_json
            )
        else:
            insert_query = """
            INSERT INTO generated_pdf(
                pdf_filename, xlsx_filename, file_size,
                checksum_md5, metadata_checksum, metadata_data,
                initial_size_bytes, final_size_bytes, metadata_ts_ist,
                station_code, source_pdf_name, full_file_md5, remarks,
                checksum_algo, created_at, level1_status, level2_status,
                level3_status, version, junction_data
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s
            )
            """
            
            values = (
                pdf_metadata.get('pdf_filename'),
                xlsx_filename,
                pdf_metadata.get('final_size_bytes'),
                pdf_metadata.get('full_file_md5'),
                pdf_metadata.get('metadata_checksum'),
                pdf_metadata.get('metadata_data'),
                pdf_metadata.get('initial_size_bytes'),
                pdf_metadata.get('final_size_bytes'),
                pdf_metadata.get('metadata_ts_ist'),
                pdf_metadata.get('station_code'),
                pdf_metadata.get('pdf_filename'),
                pdf_metadata.get('full_file_md5'),
                'Generated by Python script',
                'md5',
                datetime.now(timezone('Asia/Kolkata')).strftime("%Y-%m-%d %H:%M:%S.%f"),
                'pending',
                'pending',
                'pending',
                1,
                junction_data_json
            )
        
        cursor.execute(insert_query, values)
        db_conn.commit()
        cursor.close()
        print(f"PDF metadata stored in database with project_id: {project_id}")
        return True
        
    except Exception as e:
        print(f"Error storing PDF metadata in database: {e}")
        traceback.print_exc()
        db_conn.rollback()
        return False

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
    
        return checksum, checksum_data, content_size, timestamp
    
    except Exception as e:
        print(f"Error generating checksum: {e}")
        traceback.print_exc()
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
        
        # Capture timestamp in IST
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        now_str = current_time.strftime("D:%Y%m%d%H%M%S+05'30'")
        
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
            traceback.print_exc()
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
            traceback.print_exc()
            return False
    except Exception as e:
        print(f"Error enhancing PDF metadata: {e}")
        traceback.print_exc()
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
        
        # Enhance metadata with error handling
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
            traceback.print_exc()
            full_file_md5 = "N/A"
        
        # Store as log file
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
            traceback.print_exc()
        
        return updated_checksum, updated_checksum_data, final_content_size, full_file_md5, timestamp, initial_content_size
    except Exception as e:
        print(f"Error updating PDF metadata: {e}")
        traceback.print_exc()
        return checksum, checksum_data, content_size, "N/A", None, 0

# === ENHANCED ROW ORDERING FUNCTION ===
def get_row_order(letter):
    """
    Enhanced row ordering function to ensure proper descending order
    """
    try:
        if pd.isna(letter) or str(letter).strip() == '':
            return 0
        
        letter_str = str(letter).strip().upper()
        
        # Define the desired order (highest to lowest)
        row_order = {
            'F': 6, 'E': 5, 'D': 4, 'C': 3, 'B': 2, 'A': 1,
            'H': 8, 'G': 7
        }
        
        if letter_str in row_order:
            return row_order[letter_str]
        
        # For any other letters, use their ASCII value with offset
        try:
            return 100 - ord(letter_str)
        except:
            return 0
    except Exception as e:
        print(f"Error in get_row_order for letter '{letter}': {e}")
        return 0

# === UPDATED PAGINATION LOGIC WITH CABLE BOX LIMIT ===
def break_cables_into_rows_updated(cable_list, max_terminal_symbols_per_row=36, max_cable_boxes_per_row=6):
    """
    Break cables into multiple rows if they exceed the terminal limit,
    keeping cables from the same letter together.
    """
    try:
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
    except Exception as e:
        print(f"Error in break_cables_into_rows_updated: {e}")
        traceback.print_exc()
        return []

# === UPDATED FUNCTION: Draw cable box row ===
def draw_cable_box_row(ax, x_start, y_center, cable_info, pin_spacing=0.8):
    """
    Draw a SINGLE cable box for cables with cabel_type = 'cabel_box'
    """
    try:
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
        
        # Rectangle dimensions
        rect_width = 1.5
        rect_height = 0.7
        
        # Draw the rectangle
        rect_x = x_start + 1.0 - rect_width / 2
        rect_y = y_center - rect_height / 2
        ax.add_patch(Rectangle((rect_x, rect_y), rect_width, rect_height,
                            edgecolor='black', facecolor='white', linewidth=1.5))
        
        # Generate terminal number
        if pd.notna(start_no) and str(start_no).strip() != '':
            try:
                terminal_num = int(float(start_no))
                upper_text = f"{terminal_num:02d}" if terminal_num < 100 else str(terminal_num)
            except:
                terminal_num = position_num
                upper_text = f"{terminal_num:02d}"
        else:
            terminal_num = position_num
            upper_text = f"{terminal_num:02d}"
        
        # Position the terminal number above the rectangle
        ax.text(x_start + 1.0, rect_y + rect_height + 0.15, upper_text,
                fontsize=16, ha='center', va='bottom', fontname=DEFAULT_FONT, fontweight='bold')
        
        # Inside text (cable name)
        if pd.notna(cable_name) and str(cable_name).strip() != '':
            cable_text = str(cable_name).strip()
            ax.text(x_start + 1.0, y_center, cable_text,
                    fontsize=18, ha='center', va='center', fontname=DEFAULT_FONT, fontweight='bold')
        
        return x_positions, input_connected_flags, output_connected_flags
    except Exception as e:
        print(f"Error in draw_cable_box_row: {e}")
        traceback.print_exc()
        return [], [], []

# === UPDATED FUNCTION: Draw extra connections ===
def draw_extra_connections(ax, cable_rows, x_positions, terminal_nos_for_positions,
                        y_top_bus_group, y_bottom_bus_group, capsule_y_center):
    """
    Draw extra connections based on input_connected_extra and output_connected_extra columns
    """
    try:
        connections_drawn = 0
        top_connections = []
        bottom_connections = []
    
        # First, collect all connections
        for _, row in cable_rows.iterrows():
            # Collect input connections
            extra_input = row.get('input_connected_extra')
            if pd.notna(extra_input) and str(extra_input).strip() != '':
                try:
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
                    traceback.print_exc()
        
            # Collect output connections
            extra_output = row.get('output_connected_extra')
            if pd.notna(extra_output) and str(extra_output).strip() != '':
                try:
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
                    traceback.print_exc()
    
        # Draw top connections in staggered layers
        if top_connections:
            top_connections.sort(key=lambda conn: abs(conn[1] - conn[0]))
            layers = []
            
            for conn in top_connections:
                x1, x2, start_term, end_term = conn
                placed = False
            
                for layer_idx, layer in enumerate(layers):
                    can_place = True
                    for existing_conn in layer:
                        ex_x1, ex_x2 = existing_conn[0], existing_conn[1]
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
                layer_y_offset = 0.8 + (layer_idx * 0.4)
            
                for conn in layer:
                    x1, x2, start_term, end_term = conn
                
                    # Draw connection line at top
                    extra_y = y_top_bus_group + layer_y_offset
                    ax.plot([x1, x2], [extra_y, extra_y],
                        color='black', linewidth=1, linestyle='-', zorder=5)
                
                    vertical_offset = 0.43
                    ax.plot([x1, x1], [y_top_bus_group - vertical_offset, extra_y],
                            color='black', linewidth=1, linestyle='-', zorder=5)
                    ax.plot([x2, x2], [y_top_bus_group - vertical_offset, extra_y],
                            color='black', linewidth=1, linestyle='-', zorder=5)
                
                    connections_drawn += 1
    
        # Draw bottom connections in staggered layers
        if bottom_connections:
            bottom_connections.sort(key=lambda conn: abs(conn[1] - conn[0]))
            layers = []
            
            for conn in bottom_connections:
                x1, x2, start_term, end_term = conn
                placed = False
            
                for layer_idx, layer in enumerate(layers):
                    can_place = True
                    for existing_conn in layer:
                        ex_x1, ex_x2 = existing_conn[0], existing_conn[1]
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
                layer_y_offset = 0.8 + (layer_idx * 0.4)
            
                for conn in layer:
                    x1, x2, start_term, end_term = conn
                
                    # Draw connection line at bottom
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
        traceback.print_exc()

# === Function to merge ranges ===
def merge_ranges(ranges, merge_adjacent=True):
    """
    Merge overlapping integer index ranges.
    """
    try:
        if not ranges:
            return []
        sorted_ranges = sorted(ranges)
        merged = [list(sorted_ranges[0])]
        for current in sorted_ranges[1:]:
            last = merged[-1]
            if merge_adjacent:
                cond = (current[0] <= last[1] + 1)
            else:
                cond = (current[0] <= last[1])
            if cond:
                last[1] = max(last[1], current[1])
            else:
                merged.append(list(current))
        return [tuple(r) for r in merged]
    except Exception as e:
        print(f"Error in merge_ranges: {e}")
        traceback.print_exc()
        return []

# === Function to get cable name ===
def get_block_cable_name(df_block):
    try:
        if 'row' in df_block.columns:
            s = (
                df_block['row']
                .dropna().astype(str).str.strip()
                .replace('', pd.NA).dropna()
            )
            if not s.empty:
                return s.iloc[0]
        return ""
    except Exception as e:
        print(f"Error in get_block_cable_name: {e}")
        traceback.print_exc()
        return ""

# === Function to draw cable name with circle ===
def draw_cable_name(ax, x, y, cable_name, x_offset=0.65):
    try:
        circle_center = (x + x_offset, y)
        ax.add_patch(Circle(circle_center, radius=0.22,
                            edgecolor='black', facecolor='white', linewidth=0.8))
        ax.text(x + x_offset, y, cable_name, ha='center', va='center',
                fontsize=22, fontname=DEFAULT_FONT, fontweight='bold')
    except Exception as e:
        print(f"Error in draw_cable_name: {e}")
        traceback.print_exc()

# === Function to draw junction box with big text ===
def draw_junction_box(ax, x, y, junction_name, rect_pad=0.2):
    try:
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
                fontsize=font_size, fontname=DEFAULT_FONT, zorder=5, fontweight='bold')
    except Exception as e:
        print(f"Error in draw_junction_box: {e}")
        traceback.print_exc()

# === Relay input symbol ===
def draw_relay_input(ax, x_left, x_right, y=0, scale=1.0, text='RELAY', anchor_to_v_tip=False, v_offset=-0.5, is_not_connected_with_bus=False):
    try:
        if x_left is None or x_right is None:
            return
        if x_left > x_right:
            x_left, x_right = x_right, x_left
    
        if is_not_connected_with_bus:
            y += -1
    
        span = max(1e-6, float(x_right) - float(x_left))
        center = (x_left + x_right) / 2.0
        tri_base = min(max(span * 0.18, 0.25 * scale), span * 0.45)
        tri_height = tri_base * 0.25 * scale
        v_depth = tri_height * 0.9
        left_notch = (center - tri_base / 2.0, y - v_offset)
        right_notch = (center + tri_base / 2.0, y - v_offset)
        notch_top = (center, y + tri_height - v_offset)
        v_tip = (center, y - v_depth - v_offset)
    
        text_y_offset = 0.15 * scale
        text_y = notch_top[1] + text_y_offset
        ax.text(center, text_y, str(text), ha='center', va='bottom',
                fontsize=int(39 * scale), fontname=DEFAULT_FONT)
    except Exception as e:
        print(f"Error in draw_relay_input: {e}")
        traceback.print_exc()

def draw_relay_output(ax, x_left, x_right, y=0, scale=1.0, text='RELAY', anchor_to_v_tip=False, v_offset=0.5, is_not_connected_with_bus=False):
    try:
        if x_left is None or x_right is None:
            return
        if x_left > x_right:
            x_left, x_right = x_right, x_left
    
        if is_not_connected_with_bus:
            y -= 0.3
    
        span = max(1e-6, float(x_right) - float(x_left))
        center = (x_left + x_right) / 2.0
        tri_base = min(max(span * 0.18, 0.25 * scale), span * 0.45)
        tri_height = tri_base * 0.25 * scale
        v_depth = tri_height * 0.9
        left_notch = (center - tri_base / 2.0, y - v_offset)
        right_notch = (center + tri_base / 2.0, y - v_offset)
        notch_bottom = (center, y - tri_height - v_offset)
        v_tip = (center, y + v_depth - v_offset)
    
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
                fontname=DEFAULT_FONT, linespacing=1.2)
    except Exception as e:
        print(f"Error in draw_relay_output: {e}")
        traceback.print_exc()

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
    split_length=3,
    draw_diagonal=True,
    draw_vertical=True,
    vertical_linewidth=1.2,
    diagonal_linewidth=1.2
):
    try:
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
        # Vertical center line (going up)
        ax.plot([x, x], [base_y, base_y + line_extension], color='black', linewidth=1)
        # One \______ style segment at the top
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
        # Stack diagonal bars with small vertical/diagonal connectors
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
            current_x = x + x_offset if i > 0 else x
            left_y = base_y + line_extension - diagonal_length - y_shift + left_adjust + diag_offset + down_shift + extra_shift
            right_y = base_y + line_extension - y_shift + right_adjust + diag_offset + down_shift + extra_shift
            
            if draw_diagonal:
                ax.plot([current_x - diagonal_length / 2, current_x + diagonal_length / 2],
                        [left_y, right_y],
                        color='black', linewidth=1)
                
                if prev_center_y is not None:
                    vertical_shift = -1.3 * scale
                    x_shift = -0.19 * scale
                    length_extension = 1.8
                    center_diag_y = (left_y + right_y) / 2.0
                    left_x = current_x - diagonal_length / 2 + x_shift
                    right_x = current_x + (diagonal_length / 2 * length_extension) + x_shift
                    left_y_shifted = prev_center_y + vertical_shift
                    right_y_shifted = center_diag_y + vertical_shift
                    
                    ax.plot([left_x, right_x], [left_y_shifted, right_y_shifted],
                            color='black', linewidth=1)
                    
                    small_vert_length_bottom = 1 * scale
                    ax.plot([right_x, right_x],
                            [right_y_shifted, right_y_shifted + small_vert_length_bottom],
                            color='black', linewidth=1)
            
            center_y = (left_y + right_y) / 2.0
            
            if draw_vertical and prev_center_y is not None:
                ax.plot([current_x, current_x], [prev_center_y, center_y],
                        color='black', linewidth=1)
            
            prev_center_y = center_y
            
            text_offset = -0.1 * scale
            text_y = base_y + line_extension + 0.1 - y_shift + text_offset + extra_shift
            display_text = str(texts[i]).strip()
            
            if split_text and len(display_text) > split_length:
                mid = len(display_text) // 2
                display_text = display_text[:mid] + '\n' + display_text[mid:]
            
            ax.text(current_x, text_y, display_text, ha='center', va='bottom',
                    fontsize=int(21 * scale), fontname=DEFAULT_FONT)
    except Exception as e:
        print(f"Error in draw_group_top_symbol: {e}")
        traceback.print_exc()

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
    diagonal_linewidth=1.2
):
    try:
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
        
        total_width = x_end - x_start
        horizontal_ratio = 0.85
        y_bottom = base_y - 0.11 * scale
        rise_height = 0.08 * scale
        y_top = y_bottom + rise_height
        seg_x0 = x_start
        seg_x2 = x_end
        seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio)
        
        ax.plot([seg_x1, seg_x0], [y_bottom, y_top], color='black', linewidth=1)
        horizontal_offset = -0.08 * scale
        ax.plot([seg_x1, seg_x2], [y_top + horizontal_offset, y_top + horizontal_offset],
                color='black', linewidth=1)
        
        if choke_output_terminal is None:
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
                current_x = x + x_offset if i > 0 else x
                left_y = base_y - line_extension + y_shift + left_adjust + diag_offset + down_shift + extra_shift
                right_y = base_y - line_extension + y_shift + right_adjust + diag_offset + down_shift + extra_shift
                
                if draw_diagonal:
                    y_offset = 0.235
                    ax.plot(
                        [current_x - diagonal_length / 2, current_x + diagonal_length / 2],
                        [left_y - y_offset, right_y - y_offset],
                        color='black', linewidth=1
                    )
                    
                    if prev_center_y is not None:
                        vertical_shift = 1.3 * scale
                        x_shift = -0.19 * scale
                        length_extension = 1.8
                        center_diag_y = (left_y + right_y) / 2.0
                        left_x = current_x - diagonal_length / 2 + x_shift
                        right_x = current_x + (diagonal_length / 2 * length_extension) + x_shift
                        left_y_shifted = prev_center_y + vertical_shift
                        right_y_shifted = center_diag_y + vertical_shift
                        y_offset = 0.25 * scale
                        
                        ax.plot(
                            [left_x, right_x],
                            [left_y_shifted - y_offset, right_y_shifted - y_offset],
                            color='black',
                            linewidth=1
                        )
                        
                        small_vert_length_bottom = -1.28 * scale
                        ax.plot(
                            [right_x, right_x],
                            [right_y_shifted - y_offset, right_y_shifted + small_vert_length_bottom - y_offset],
                            color='black', linewidth=1
                        )
                
                center_y = (left_y + right_y) / 2.0
                
                if draw_vertical and prev_center_y is not None:
                    ax.plot([current_x, current_x], [prev_center_y, center_y],
                            color='black', linewidth=1)
                
                prev_center_y = center_y
                
                text_offset = 0.2 * scale
                text_y = base_y - line_extension - 0.1 + y_shift - text_offset + extra_shift
                display_text = str(texts[i]).strip()
                
                if split_text and len(display_text) > split_length:
                    mid = len(display_text) // 2
                    display_text = display_text[:mid] + '\n' + display_text[mid:]
                
                ax.text(current_x, text_y, display_text, ha='center', va='top',
                        fontsize=int(21 * scale), fontname=DEFAULT_FONT)
        else:
            display_text = str(texts[0]).strip() if texts else ''
            diagonal_length = 0.21 * scale
            y_shift = 0.07 * scale
            diag_offset = 0.05 * scale
            left_adjust = 0.04 * scale
            right_adjust = -0.04 * scale
            diagonal_down_shift = 0.02 * scale
            
            left_y = base_y - line_extension - diagonal_length + y_shift + left_adjust + diag_offset - diagonal_down_shift
            right_y = base_y - line_extension + y_shift + right_adjust + diag_offset - diagonal_down_shift
            
            dx = 0.5
            ax.plot([x - diagonal_length/2 + dx, x + diagonal_length/2 + dx],
                    [left_y, right_y],
                    color='black', linewidth=1.2)
            
            horiz_length = 0.5 * scale
            left_shift = 0.1
            end_x = x + diagonal_length/2 + horiz_length - left_shift
            start_x = x + diagonal_length/2 - left_shift
            line_y_shift = 2.78
            line_y = right_y + line_y_shift
            
            x_shift_left = 0.005 * scale
            small_vertical = 0.95 * scale
            ax.plot(
                [start_x - x_shift_left, start_x - x_shift_left],
                [line_y, line_y - small_vertical],
                color='black',
                linewidth=1.2
            )
            
            ax.plot([start_x, end_x], [line_y, line_y], color='black', linewidth=1.2)
            
            vertical_length = 2.84 * scale
            ax.plot([end_x, end_x], [line_y, line_y - vertical_length], color='black', linewidth=1.2)
            
            text_offset = 0.2 * scale
            text_y = right_y - text_offset - 0.05
            text_x = x + 0.55
            
            if len(display_text) > 3:
                mid = len(display_text) // 2
                display_text = display_text[:mid] + '\n' + display_text[mid:]
            
            ax.text(text_x, text_y, display_text, ha='center', va='top',
                    fontsize=int(21 * scale), fontname=DEFAULT_FONT, linespacing=1.2)
    except Exception as e:
        print(f"Error in draw_group_bottom_symbol: {e}")
        traceback.print_exc()

def draw_relay_box_top(ax, x_start, x_end, y, texts='R1', scale=1.0, input_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay box at top with rectangle
    """
    try:
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
        
        extra_vertical_extension = 0.16 * scale
        ax.plot([x, x],
                [base_y, base_y + line_extension + extra_vertical_extension],
                color='black', linewidth=1)
        
        total_width = x_end - x_start
        horizontal_ratio = 0.85
        y_top = base_y + 0.11 * scale
        drop_height = 0.08 * scale
        y_bottom = y_top - drop_height
        seg_x0 = x_start
        seg_x2 = x_end
        seg_x1 = seg_x0 + total_width * (1 - horizontal_ratio)
        
        ax.plot([seg_x1, seg_x0], [y_top, y_bottom], color='black', linewidth=1)
        
        horizontal_offset = 0.08 * scale
        ax.plot([seg_x1, seg_x2], [y_bottom + horizontal_offset, y_bottom + horizontal_offset],
                color='black', linewidth=1)
        
        rect_width = 0.45 * scale
        rect_height = 0.35 * scale
        y_shift = -0.17 * scale
        rect_offset = -0.05 * scale
        down_shift = 0.39 * scale
        prev_center_y = None
        x_offset = x_offset * scale
        
        for i in range(num_texts):
            extra_shift = i * spacing
            current_x = x + x_offset if i > 0 else x
            rect_y = base_y + line_extension - rect_height/2 - y_shift + rect_offset + down_shift + extra_shift
            
            rect = Rectangle((current_x - rect_width/2, rect_y - rect_height/2),
                            rect_width, rect_height,
                            linewidth=1.2, edgecolor='black', facecolor='white')
            ax.add_patch(rect)
            
            if prev_center_y is not None:
                vertical_shift = -1.3 * scale
                x_shift = -0.19 * scale
                center_rect_y = rect_y
                ax.plot([current_x, current_x], [prev_center_y, center_rect_y],
                        color='black', linewidth=1)
                
                small_vert_length = 0.3 * scale
                ax.plot([current_x, current_x],
                        [center_rect_y, center_rect_y - small_vert_length],
                        color='black', linewidth=1)
            
            prev_center_y = rect_y
            
            text_y = rect_y
            display_text = str(texts[i]).strip()
            if len(display_text) > 3:
                mid = len(display_text) // 2
                display_text = display_text[:mid] + '\n' + display_text[mid:]
            
            ax.text(current_x, text_y, display_text, ha='center', va='center',
                    fontsize=int(16 * scale), fontname=DEFAULT_FONT, linespacing=0.8)
    except Exception as e:
        print(f"Error in draw_relay_box_top: {e}")
        traceback.print_exc()

def draw_relay_box_bottom(ax, x_start, x_end, y, texts='R1', scale=1.0,
                        output_connected='N', spacing=0.3, x_offset=0.3):
    try:
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
        center_extend = 0.16 * scale
        
        ax.plot(
            [x, x],
            [base_y, base_y - line_extension - center_extend],
            color='black', linewidth=1
        )
        
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
        
        rect_width = 0.45 * scale
        rect_height = 0.35 * scale
        y_shift = 0.17 * scale
        rect_offset = 0.05 * scale
        down_shift = -0.18 * scale
        down_adjust = 0.55 * scale
        prev_center_y = None
        x_offset = x_offset * scale
        extra_vertical_extension = 0.25 * scale
        
        for i in range(num_texts):
            extra_shift = -i * spacing
            current_x = x + x_offset if i > 0 else x
            rect_y = (base_y - line_extension + rect_height/2 +
                    y_shift + rect_offset + down_shift + extra_shift)
            rect_y -= down_adjust
            
            rect = Rectangle(
                (current_x - rect_width/2, rect_y - rect_height/2),
                rect_width, rect_height,
                linewidth=1.2, edgecolor='black', facecolor='white'
            )
            ax.add_patch(rect)
            
            if prev_center_y is not None:
                ax.plot(
                    [current_x, current_x],
                    [prev_center_y - extra_vertical_extension, rect_y],
                    color='black',
                    linewidth=1
                )
                
                small_vert_length = 0.3 * scale
                ax.plot(
                    [current_x, current_x],
                    [rect_y, rect_y + small_vert_length],
                    color='black',
                    linewidth=1
                )
            
            prev_center_y = rect_y
            
            display_text = str(texts[i]).strip()
            if len(display_text) > 3:
                mid = len(display_text) // 2
                display_text = display_text[:mid] + "\n" + display_text[mid:]
            
            ax.text(current_x, rect_y, display_text,
                    ha='center', va='center',
                    fontsize=int(16 * scale), fontname=DEFAULT_FONT, linespacing=0.8)
    except Exception as e:
        print(f"Error in draw_relay_box_bottom: {e}")
        traceback.print_exc()

def draw_relay_contact_box_top(ax, x_start, x_end, y, texts='R1', scale=1.0, input_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay contact box at top - rectangle above terminals
    """
    try:
        if isinstance(texts, str):
            texts = [texts]
        rect_width = (0.45 * 2) * scale
        rect_height = 0.35 * scale
        vertical_offset = 0.32 * scale
        rect_x = (x_start + x_end) / 2 - rect_width / 2
        rect_y = y + vertical_offset
        
        rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                        linewidth=1.2, edgecolor='black', facecolor='white')
        ax.add_patch(rect)
        
        display_text = str(texts[0]).strip() if texts else ''
        if len(display_text) > 3:
            mid = len(display_text) // 2
            display_text = display_text[:mid] + '\n' + display_text[mid:]
        
        ax.text((x_start + x_end) / 2, rect_y + rect_height / 2,
                display_text,
                ha='center', va='center',
                fontsize=int(16 * scale), fontname=DEFAULT_FONT, linespacing=0.8)
    except Exception as e:
        print(f"Error in draw_relay_contact_box_top: {e}")
        traceback.print_exc()

def draw_relay_contact_box_bottom(ax, x_start, x_end, y, texts='R1', scale=1.0,
                                output_connected='N', spacing=0.3, x_offset=0.3):
    """
    Draw relay contact box at bottom - rectangle below terminals
    """
    try:
        if isinstance(texts, str):
            texts = [texts]
        terminal_gap = abs(x_end - x_start)
        base_width = 0.90 * scale
        extra_width = 0.60 * scale
        rect_width = base_width + (terminal_gap - 1) * extra_width
        rect_height = 0.35 * scale
        vertical_offset = 0.31 * scale
        rect_x = (x_start + x_end) / 2 - rect_width / 2
        rect_y = y - rect_height - vertical_offset
        
        rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                        linewidth=1.2, edgecolor='black', facecolor='white')
        ax.add_patch(rect)
        
        display_text = str(texts[0]).strip() if texts else ''
        if len(display_text) > 3:
            mid = len(display_text) // 2
            display_text = display_text[:mid] + '\n' + display_text[mid:]
        
        ax.text((x_start + x_end) / 2,
                rect_y + rect_height / 2,
                display_text,
                ha='center', va='center',
                fontsize=int(16 * scale),
                fontname=DEFAULT_FONT,
                linespacing=0.8)
    except Exception as e:
        print(f"Error in draw_relay_contact_box_bottom: {e}")
        traceback.print_exc()

# === Updated draw_header ===
def draw_header(ax, cable_id, header_type, x_start, x_end, text, min_symbol_bottom=None,
                first_hook_x=None, last_hook_x=None, y_top_bus_group=0, y_bottom_bus_group=0, special_ha=False):
    try:
        top_y_offset = 0.2
        bottom_y_offset = 1
        if pd.isna(text) or str(text).strip() == '':
            return
        text = str(text).strip()
        if str(header_type).strip().upper() == 'WIREFROM':
            x_pos = first_hook_x if first_hook_x is not None else x_start - 0.05
            y_pos = y_top_bus_group + top_y_offset
            ax.text(x_pos, y_pos, text, ha='left', va='bottom', fontsize=21, fontname=DEFAULT_FONT)
        elif str(header_type).strip().upper() == 'WIRETO':
            ha = 'center'
            x_pos = last_hook_x if last_hook_x is not None else (x_start + x_end) / 2.0
            if last_hook_x is not None:
                ha = 'left' if special_ha else 'right'
            min_symbol_bottom = y_bottom_bus_group - 0.2 if min_symbol_bottom is None else min_symbol_bottom
            text_offset = -0.15
            y_pos = min_symbol_bottom - bottom_y_offset + text_offset
            ax.text(x_pos, y_pos, text, ha=ha, va='top', fontsize=21, fontname=DEFAULT_FONT)
    except Exception as e:
        print(f"Error in draw_header: {e}")
        traceback.print_exc()

# === Helper: Find row by terminal number ===
def find_row_by_term(term):
    try:
        if pd.isna(term):
            return None
        s = str(term).strip()
        if s.endswith('.0'):
            s = s[:-2]
        col = df['terminal_no'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        matches = df[col == s]
        return matches.iloc[0] if not matches.empty else None
    except Exception as e:
        print(f"Error in find_row_by_term: {e}")
        traceback.print_exc()
        return None

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
footer_height = 2.75
footer_inch_add = 4.0
pin_spacing = 0.8

# === Symbol Drawers ===
def draw_capsule(ax, x, y_center, terminal_no, input_left, input_right, output_left, output_right,
                input_connected, output_connected, capsule_type='capsule'):
    """
    Draw capsule with support for different types
    """
    try:
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
            ax.text(x, y_center, term_str, fontsize=17, ha='center', va='center', fontname=DEFAULT_FONT)
        
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
                    fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        input_right_offset = 0.05
        if pd.notna(input_right) and str(input_right).strip() != "":
            ax.text(x + input_right_offset, capsule_top + 0.18, format_text(input_right),
                    fontsize=19, ha='left', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_left_offset = 0.005
        if pd.notna(output_left) and str(output_left).strip() != "":
            ax.text(x - output_left_offset, capsule_bottom - 0.15, format_text(output_left),
                    fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_right_offset = 0.05
        if pd.notna(output_right) and str(output_right).strip() != "":
            ax.text(x + output_right_offset, capsule_bottom - 0.18, format_text(output_right),
                    fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        top_conn = (x, capsule_top + SYMBOL_RADIUS)
        bottom_conn = (x, capsule_bottom - bottom_circle_radius)
        ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
        oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
        
        return top_conn, bottom_conn, ic, oc
    except Exception as e:
        print(f"Error in draw_capsule: {e}")
        traceback.print_exc()
        return (x, y_center), (x, y_center), 'N', 'N'

def draw_s_fuse(ax, x, y_center, terminal_no,
                input_left=None, input_right=None,
                output_left=None, output_right=None,
                input_connected='N', output_connected='N'):
    try:
        fuse_top = y_center + SYMBOL_HEIGHT / 2
        fuse_bottom = y_center - SYMBOL_HEIGHT / 2
        top_circle_radius = SYMBOL_RADIUS * 0.8
        bottom_circle_radius = SYMBOL_RADIUS * 0.8
        
        ax.add_patch(Circle((x, fuse_top), top_circle_radius,
                            edgecolor='black', facecolor='white', linewidth=1))
        ax.add_patch(Circle((x, fuse_bottom), bottom_circle_radius,
                            edgecolor='black', facecolor='white', linewidth=1))
        
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
        
        if pd.notna(terminal_no) and str(terminal_no).strip() != '':
            term_str = str(terminal_no)
            if term_str.endswith('.0'):
                term_str = term_str[:-2]
            ax.text(x - 0.1, y_center + 0.01, term_str,
                    ha='center', va='center', fontsize=17, fontname=DEFAULT_FONT)
        
        input_left_offset = 0.005
        if pd.notna(input_left) and str(input_left).strip() != "":
            ax.text(x - input_left_offset, fuse_top + 0.18,
                    format_text(input_left), fontsize=19,
                    ha='right', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        input_right_offset = 0.05
        if pd.notna(input_right) and str(input_right).strip() != "":
            ax.text(x + input_right_offset, fuse_top + 0.18,
                    format_text(input_right), fontsize=19,
                    ha='left', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_left_offset = 0.005
        if pd.notna(output_left) and str(output_left).strip() != "":
            ax.text(x - output_left_offset, fuse_bottom - 0.15,
                    format_text(output_left), fontsize=19,
                    ha='right', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_right_offset = 0.05
        if pd.notna(output_right) and str(output_right).strip() != "":
            ax.text(x + output_right_offset, fuse_bottom - 0.18,
                    format_text(output_right), fontsize=19,
                    ha='left', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        top_conn = (x, fuse_top + top_circle_radius)
        bottom_conn = (x, fuse_bottom - bottom_circle_radius)
        ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
        oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
        
        return top_conn, bottom_conn, ic, oc
    except Exception as e:
        print(f"Error in draw_s_fuse: {e}")
        traceback.print_exc()
        return (x, y_center), (x, y_center), 'N', 'N'

def draw_choke(ax, x, y_center, terminal_no):
    try:
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
            ax.text(x - 0.3, y_center + 0.1, term_str, ha='center', va='center', fontsize=12.5, fontname=DEFAULT_FONT)
        
        label_offset = 0.2
        if pd.notna(input_left) and str(input_left).strip() != "":
            ax.text(x - label_offset, choke_top + 0.15, str(input_left), fontsize=10, ha='right', va='bottom', rotation=90, fontname=DEFAULT_FONT)
        
        if pd.notna(input_right) and str(input_right).strip() != "":
            ax.text(x + label_offset, choke_top + 0.15, str(input_right), fontsize=10, ha='left', va='bottom', rotation=90, fontname=DEFAULT_FONT)
        
        if pd.notna(output_left) and str(output_left).strip() != "":
            ax.text(x - label_offset, choke_bottom - 0.15, str(output_left), fontsize=10, ha='right', va='top', rotation=90, fontname=DEFAULT_FONT)
        
        if pd.notna(output_right) and str(output_right).strip() != "":
            ax.text(x + label_offset, choke_bottom - 0.15, str(output_right), fontsize=10, ha='left', va='top', rotation=90, fontname=DEFAULT_FONT)
        
        top_conn = (x, choke_top + SYMBOL_RADIUS)
        bottom_conn = (x, choke_bottom - SYMBOL_RADIUS)
        ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
        oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
        
        return top_conn, bottom_conn, ic, oc
    except Exception as e:
        print(f"Error in draw_choke: {e}")
        traceback.print_exc()
        return (x, y_center), (x, y_center), 'N', 'N'

def draw_horizontal_choke(ax, x_center, y_center, label='CHOKE',
                        box_width=0.6, box_height=0.3,
                        special_end=False, output_label='', output_type='terminal', output_text='',
                        output_connected=False, y_top_bus_group=None, y_bottom_bus_group=None):
    """
    Draw horizontal choke with output_connected parameter
    """
    try:
        y_shift = -0.5
        y_center = y_center + y_shift
    
        left_x = x_center - box_width / 2
        right_x = x_center + box_width / 2
        bottom_y = y_center - box_height / 2
    
        choke_box = FancyBboxPatch((left_x, bottom_y),
                                box_width, box_height,
                                boxstyle="round,pad=0.02",
                                edgecolor='black', facecolor='white', linewidth=1.5)
        ax.add_patch(choke_box)
    
        ax.text(x_center, y_center, label,
                fontsize=16, ha='center', va='center', fontname=DEFAULT_FONT)
    
        line_length = 0.075
        delta = 0.02
        vert_line_height_left = y_bottom_bus_group - y_center if y_bottom_bus_group is not None else 0.5
    
        left_horiz_start = left_x - line_length - delta
        left_horiz_end = left_x - delta
        ax.plot([left_horiz_start, left_horiz_end],
                [y_center, y_center],
                color='black', linewidth=1)
    
        v_offset = 0.005
        ax.plot([left_horiz_start - v_offset, left_horiz_start - v_offset],
                [y_center, y_center + vert_line_height_left],
                color='black', linewidth=1)
    
        if not special_end:
            right_horiz_start = right_x + delta
            right_horiz_end = right_x + line_length + delta
            ax.plot([right_horiz_start, right_horiz_end],
                    [y_center, y_center],
                    color='black', linewidth=1)
        
            if output_connected and y_top_bus_group is not None:
                vert_line_height_right = y_top_bus_group - y_center
            else:
                vert_line_height_right = vert_line_height_left
        
            ax.plot([right_horiz_end + v_offset, right_horiz_end + v_offset],
                    [y_center, y_center + vert_line_height_right],
                    color='black', linewidth=1)
        else:
            horiz_length = 0.5 if output_type.lower() == 'relay' else 0.2
            slant_size = 0.3 if output_type.lower() != 'relay' else 0.15
            vertical_length = 0.5 if output_type.lower() != 'relay' else 0
            end_horiz_x = right_x + delta + horiz_length
            ax.plot([right_x + delta, end_horiz_x],
                    [y_center, y_center],
                    color='black', linewidth=1.4)
        
            if str(output_type).strip().lower() == 'relay':
                diag_length = slant_size
            
                if output_connected:
                    diag_start_x = end_horiz_x - 0.1
                    diag_start_y = y_center
                    diag_end_x = diag_start_x + diag_length
                    diag_end_y = diag_start_y + diag_length
                    ax.plot([diag_start_x, diag_end_x],
                            [diag_start_y, diag_end_y],
                            color='black', linewidth=1.4)
                    text_offset = 0.05
                    ax.text(diag_end_x + text_offset, diag_end_y, output_text,
                            ha='left', va='center', fontsize=19, fontname=DEFAULT_FONT)
                else:
                    diag_start_x = end_horiz_x - 0.1
                    diag_start_y = y_center + 0.1
                    diag_end_x = diag_start_x + diag_length
                    diag_end_y = diag_start_y - diag_length
                    ax.plot([diag_start_x, diag_end_x],
                            [diag_start_y, diag_end_y],
                            color='black', linewidth=1.4)
                    text_offset = 0.05
                    ax.text(diag_end_x + text_offset, diag_end_y - 0.1, output_text,
                            ha='left', va='center', fontsize=19, fontname=DEFAULT_FONT)
            
                return end_horiz_x + diag_length
            else:
                if output_connected:
                    ax.plot([end_horiz_x, end_horiz_x + slant_size],
                            [y_center, y_center + slant_size],
                            color='black', linewidth=1.4)
                    vert_top_y = y_center + slant_size
                    vert_end_y = vert_top_y + vertical_length
                    ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                            [vert_top_y, vert_end_y],
                            color='black', linewidth=1.4)
                
                    label_offset = 0.05
                    label_y = (vert_top_y + vert_end_y) / 2
                    ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                            ha='left', va='center', fontsize=19, rotation=90, fontname=DEFAULT_FONT)
                    return end_horiz_x + slant_size
                else:
                    ax.plot([end_horiz_x, end_horiz_x + slant_size],
                            [y_center, y_center - slant_size],
                            color='black', linewidth=1.4)
                    vert_top_y = y_center - slant_size
                    vert_end_y = vert_top_y - vertical_length
                    ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                            [vert_top_y, vert_end_y],
                            color='black', linewidth=1.4)
                
                    label_offset = 0.05
                    label_y = (vert_top_y + vert_end_y) / 2
                    ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                            ha='left', va='center', fontsize=19, rotation=90, fontname=DEFAULT_FONT)
                    return end_horiz_x + slant_size
        return None
    except Exception as e:
        print(f"Error in draw_horizontal_choke: {e}")
        traceback.print_exc()
        return None

def draw_dual_fuse(ax, x_left, y_center, left_term, right_term, left_input_left=None, left_input_right=None, left_output_left=None, left_output_right=None, left_input_connected='N', left_output_connected='N', right_input_left=None, right_input_right=None, right_output_left=None, right_output_right=None, right_input_connected='N', right_output_connected='N'):
    try:
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
                        ha='center', va='center', fontsize=17, fontname=DEFAULT_FONT)
            
            label_offset = 0.11
            if pd.notna(input_left) and str(input_left).strip() != "":
                ax.text(x_pos - label_offset + 0.11, fuse_top + 0.24, format_text(input_left),
                        fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
            
            if pd.notna(input_right) and str(input_right).strip() != "":
                ax.text(x_pos + 0.1, fuse_top + 0.27, format_text(input_right),
                        fontsize=19, ha='center', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
            
            if pd.notna(output_left) and str(output_left).strip() != "":
                ax.text(x_pos - label_offset + 0.11, fuse_bottom - 0.30, format_text(output_left),
                        fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
            
            if pd.notna(output_right) and str(output_right).strip() != "":
                ax.text(x_pos + label_offset - 0.09, fuse_bottom - 0.28, format_text(output_right),
                        fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
            
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
    except Exception as e:
        print(f"Error in draw_dual_fuse: {e}")
        traceback.print_exc()
        return (x_left, y_center), (x_left, y_center), 'N', 'N', 'N', 'N'

def draw_resistor(ax, x, y_center, input_terminal='', output_terminal='', resistor_name='R', input_x_pos=None, output_x_pos=None):
    try:
        radius = SYMBOL_RADIUS * 1.5
        ax.add_patch(Circle((x, y_center), radius, edgecolor='black', facecolor='white', linewidth=1))
        ax.text(x, y_center, resistor_name, ha='center', va='center', fontsize=19, fontname=DEFAULT_FONT)
        
        upper_y_start = y_center + radius
        upper_y_end = y_center + radius * 9.0
        ax.plot([x, x], [upper_y_start, upper_y_end], color='black', linewidth=1)
        
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
                vertical_drop_length = 1.2
                ax.plot([left_x, left_x], [upper_y_end, upper_y_end - vertical_drop_length], color='black', linewidth=1)
        else:
            upper_horiz_length = 0. + len(str(input_terminal).strip()) * 0.12
            ax.plot([x - upper_horiz_length, x], [upper_y_end, upper_y_end], color='black', linewidth=1)
        
        lower_y_start = y_center - radius
        lower_y_end = y_center - radius * 6
        ax.plot([x, x], [lower_y_start, lower_y_end], color='black', linewidth=1)
        
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
                vertical_drop_length = 0.8
                ax.plot([left_x, left_x], [lower_y_end, lower_y_end + vertical_drop_length], color='black', linewidth=1)
        else:
            lower_horiz_length = 0.5 + len(str(output_terminal).strip()) * 0.12
            ax.plot([x - lower_horiz_length, x], [lower_y_end, lower_y_end], color='black', linewidth=1)
        
        return None, None
    except Exception as e:
        print(f"Error in draw_resistor: {e}")
        traceback.print_exc()
        return None, None

def draw_rectangle_symbol(ax, x, y_center, terminal_no, symbol, input_left, input_right, output_left, output_right,
                        input_connected, output_connected):
    try:
        rect_width = 0.3
        rect_height = 0.835
        rect_x = x - rect_width / 2
        rect_y = y_center - rect_height / 2
        ax.add_patch(Rectangle((rect_x, rect_y), rect_width, rect_height,
                            edgecolor='black', facecolor='white', linewidth=1))
        
        inner_width = 0.22
        inner_height = 0.22
        inner_x = x - inner_width / 2
        inner_offset = 0.25
        inner_y = y_center - inner_height / 2 + inner_offset
        ax.add_patch(Rectangle((inner_x, inner_y), inner_width, inner_height,
                            edgecolor='black', facecolor='white', linewidth=1))
        
        text_offset = 0
        if pd.notna(terminal_no) and str(terminal_no).strip() != '':
            term_str = str(terminal_no)
            if term_str.endswith('.0'):
                term_str = term_str[:-2]
            ax.text(x, y_center + text_offset, term_str, fontsize=17, ha='center', va='center', fontname=DEFAULT_FONT)
        
        if pd.notna(symbol) and str(symbol).strip() != '':
            symbol_str = str(symbol)
            if symbol_str.endswith('.0'):
                symbol_str = symbol_str[:-2]
            symbol_offset = 0.25
            ax.text(x, y_center + text_offset + symbol_offset, symbol_str, fontsize=19, ha='center', va='center', fontname=DEFAULT_FONT)
        
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
                    fontsize=19, ha='right', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        input_right_offset = 0.05
        if pd.notna(input_right) and str(input_right).strip() != "":
            ax.text(x + input_right_offset, y_center + rect_height / 2 + 0.18, format_text(input_right),
                    fontsize=19, ha='left', va='bottom', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_left_offset = 0.005
        if pd.notna(output_left) and str(output_left).strip() != "":
            ax.text(x - output_left_offset, y_center - rect_height / 2 - 0.15, format_text(output_left),
                    fontsize=19, ha='right', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        output_right_offset = 0.05
        if pd.notna(output_right) and str(output_right).strip() != "":
            ax.text(x + output_right_offset, y_center - rect_height / 2 - 0.18, format_text(output_right),
                    fontsize=19, ha='left', va='top', rotation=90, linespacing=1.2, fontname=DEFAULT_FONT)
        
        top_conn = (x, rect_y + rect_height)
        bottom_conn = (x, rect_y)
        ic = 'Y' if str(input_connected).strip().upper() == 'Y' else 'N'
        oc = 'Y' if str(output_connected).strip().upper() == 'Y' else 'N'
        
        return top_conn, bottom_conn, ic, oc
    except Exception as e:
        print(f"Error in draw_rectangle_symbol: {e}")
        traceback.print_exc()
        return (x, y_center), (x, y_center), 'N', 'N'

def draw_input_connection(ax, x, symbol_top_y, input_connected_flag, output_connected_flag, y_top_bus_group, has_bus_line=True):
    """
    Draw input connection with three scenarios
    """
    try:
        overlap = SYMBOL_RADIUS * 0.15
        start_y = symbol_top_y - overlap
        is_connected = (input_connected_flag is True) or (str(input_connected_flag).strip().upper() == 'Y')
        
        if is_connected and has_bus_line:
            extended_y = y_top_bus_group + 0.2
            ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
            return True, extended_y
        elif not is_connected and has_bus_line:
            total_distance = y_top_bus_group - start_y
            short_length = total_distance * 0.5
            end_y = start_y + short_length
            ax.plot([x, x], [start_y, end_y], color='black', linewidth=1)
            return False, end_y
        else:
            extended_y = y_top_bus_group + 0.2
            ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
            return False, extended_y
    except Exception as e:
        print(f"Error in draw_input_connection: {e}")
        traceback.print_exc()
        return False, symbol_top_y

def draw_output_connection(ax, x, symbol_bottom_y, input_connected_flag, output_connected_flag, y_bottom_bus_group, has_bus_line=True):
    """
    Draw output connection with three scenarios
    """
    try:
        overlap = SYMBOL_RADIUS * 0.15
        start_y = symbol_bottom_y + overlap
        is_connected = (output_connected_flag is True) or (str(output_connected_flag).strip().upper() == 'Y')
        
        if is_connected and has_bus_line:
            extended_y = y_bottom_bus_group - 0.2
            ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
            return True, extended_y
        elif not is_connected and has_bus_line:
            total_distance = start_y - y_bottom_bus_group
            short_length = total_distance * 0.5
            end_y = start_y - short_length
            ax.plot([x, x], [start_y, end_y], color='black', linewidth=1)
            return False, end_y
        else:
            extended_y = y_bottom_bus_group - 0.2
            ax.plot([x, x], [start_y, extended_y], color='black', linewidth=1)
            return False, extended_y
    except Exception as e:
        print(f"Error in draw_output_connection: {e}")
        traceback.print_exc()
        return False, symbol_bottom_y

def draw_bus_lines(ax, x_positions, connected_flags, bus_y, gap=0.12, extra=0.12, y_offset=0.15):
    try:
        if not x_positions:
            return
        
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
    except Exception as e:
        print(f"Error in draw_bus_lines: {e}")
        traceback.print_exc()

# === Helper function for parsing terminal_no field ===
def parse_terminal_no_field(val):
    try:
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
    except Exception as e:
        print(f"Error in parse_terminal_no_field: {e}")
        traceback.print_exc()
        return None, None

def draw_horizontal_choke_updated(ax, x_center, y_center, label='CHOKE',
                        box_width=0.6, box_height=0.3,
                        special_end=False, output_label='', output_type='terminal', output_text='',
                        output_connected_terminals=[], x_positions=None, terminal_nos_for_positions=None,
                        y_top_bus_group=None, y_bottom_bus_group=None):
    """
    Updated horizontal choke with output_connected as list of terminal numbers
    """
    try:
        y_shift = -0.5
        y_center = y_center + y_shift
    
        left_x = x_center - box_width / 2
        right_x = x_center + box_width / 2
        bottom_y = y_center - box_height / 2
    
        choke_box = FancyBboxPatch((left_x, bottom_y),
                                box_width, box_height,
                                boxstyle="round,pad=0.02",
                                edgecolor='black', facecolor='white', linewidth=1.5)
        ax.add_patch(choke_box)
    
        ax.text(x_center, y_center, label,
                fontsize=16, ha='center', va='center', fontname=DEFAULT_FONT)
    
        line_length = 0.075
        delta = 0.02
        vert_line_height_left = y_bottom_bus_group - y_center if y_bottom_bus_group is not None else 0.5
    
        left_horiz_start = left_x - line_length - delta
        left_horiz_end = left_x - delta
        ax.plot([left_horiz_start, left_horiz_end],
                [y_center, y_center],
                color='black', linewidth=1)
    
        v_offset = 0.005
        ax.plot([left_horiz_start - v_offset, left_horiz_start - v_offset],
                [y_center, y_center + vert_line_height_left],
                color='black', linewidth=1)
    
        if output_connected_terminals and x_positions is not None and terminal_nos_for_positions is not None:
            connected_x = []
            for term in output_connected_terminals:
                term_str = str(term).strip().replace('.0', '')
                if term_str in terminal_nos_for_positions:
                    idx = terminal_nos_for_positions.index(term_str)
                    connected_x.append(x_positions[idx])
            
            if connected_x:
                min_x = min(connected_x)
                max_x = max(connected_x)
                y_bus = y_bottom_bus_group - 0.8
                
                ax.plot([min_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
                
                for x in connected_x:
                    ax.plot([x, x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
                
                choke_output_x = x_center + box_width/2
                ax.plot([choke_output_x, choke_output_x], [y_center, y_bus], color='black', linewidth=1)
                ax.plot([choke_output_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
                
                return max_x

        if not special_end:
            right_horiz_start = right_x + delta
            right_horiz_end = right_x + line_length + delta
            ax.plot([right_horiz_start, right_horiz_end],
                    [y_center, y_center],
                    color='black', linewidth=1)
        
            if output_connected_terminals and y_top_bus_group is not None:
                vert_line_height_right = y_top_bus_group - y_center
            else:
                vert_line_height_right = vert_line_height_left
        
            ax.plot([right_horiz_end + v_offset, right_horiz_end + v_offset],
                    [y_center, y_center + vert_line_height_right],
                    color='black', linewidth=1)
        else:
            horiz_length = 0.5 if output_type.lower() == 'relay' else 0.2
            slant_size = 0.3 if output_type.lower() != 'relay' else 0.15
            vertical_length = 0.5 if output_type.lower() != 'relay' else 0
            end_horiz_x = right_x + delta + horiz_length
            ax.plot([right_x + delta, end_horiz_x],
                    [y_center, y_center],
                    color='black', linewidth=1.4)
        
            if str(output_type).strip().lower() == 'relay':
                diag_length = slant_size
            
                if output_connected_terminals:
                    diag_start_x = end_horiz_x - 0.1
                    diag_start_y = y_center
                    diag_end_x = diag_start_x + diag_length
                    diag_end_y = diag_start_y + diag_length
                    ax.plot([diag_start_x, diag_end_x],
                            [diag_start_y, diag_end_y],
                            color='black', linewidth=1.4)
                    text_offset = 0.05
                    ax.text(diag_end_x + text_offset, diag_end_y, output_text,
                            ha='left', va='center', fontsize=19, fontname=DEFAULT_FONT)
                else:
                    diag_start_x = end_horiz_x - 0.1
                    diag_start_y = y_center + 0.1
                    diag_end_x = diag_start_x + diag_length
                    diag_end_y = diag_start_y - diag_length
                    ax.plot([diag_start_x, diag_end_x],
                            [diag_start_y, diag_end_y],
                            color='black', linewidth=1.4)
                    text_offset = 0.05
                    ax.text(diag_end_x + text_offset, diag_end_y - 0.1, output_text,
                            ha='left', va='center', fontsize=19, fontname=DEFAULT_FONT)
            
                return end_horiz_x + diag_length
            else:
                if output_connected_terminals:
                    ax.plot([end_horiz_x, end_horiz_x + slant_size],
                            [y_center, y_center + slant_size],
                            color='black', linewidth=1.4)
                    vert_top_y = y_center + slant_size
                    vert_end_y = vert_top_y + vertical_length
                    ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                            [vert_top_y, vert_end_y],
                            color='black', linewidth=1.4)
                
                    label_offset = 0.05
                    label_y = (vert_top_y + vert_end_y) / 2
                    ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                            ha='left', va='center', fontsize=19, rotation=90, fontname=DEFAULT_FONT)
                    return end_horiz_x + slant_size
                else:
                    ax.plot([end_horiz_x, end_horiz_x + slant_size],
                            [y_center, y_center - slant_size],
                            color='black', linewidth=1.4)
                    vert_top_y = y_center - slant_size
                    vert_end_y = vert_top_y - vertical_length
                    ax.plot([end_horiz_x + slant_size, end_horiz_x + slant_size],
                            [vert_top_y, vert_end_y],
                            color='black', linewidth=1.4)
                
                    label_offset = 0.05
                    label_y = (vert_top_y + vert_end_y) / 2
                    ax.text(end_horiz_x + slant_size + label_offset, label_y, output_label,
                            ha='left', va='center', fontsize=19, rotation=90, fontname=DEFAULT_FONT)
                    return end_horiz_x + slant_size
        return None
    except Exception as e:
        print(f"Error in draw_horizontal_choke_updated: {e}")
        traceback.print_exc()
        return None

def draw_symbols(df, ax, page_rows, junction_name, start_x=1, pin_spacing=0.8, cables_per_page=12, page_number=1, max_terminal_symbols_per_row=36, max_rows_visible=3, page_width=None, global_max_width=None):
    """
    Draw symbols for the provided page_rows on ax.
    """
    try:
        extra_rows = 0
        max_rows_for_ylim = max_rows_visible + extra_rows
        bottom_margin = 1.0
        top_margin = 3.0
        fixed_ylim_min = CAPSULE_Y_CENTER_BASE + vertical_gap * (1 - max_rows_for_ylim) + y_bottom_bus_offset - 1.8 - bottom_margin - footer_height
        fixed_ylim_max = CAPSULE_Y_CENTER_BASE + y_top_bus_offset + 1.8 + top_margin
        current_x = start_x
        current_terminal_count = 0
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
    
        for row_index, (letter, cable_list) in enumerate(page_rows):
            print(f"DEBUG: Processing row {row_index + 1} - Letter '{letter}' with {len(cable_list)} cables")
        
            current_x = start_x
            current_terminal_count = 0
            current_row_max_x = start_x
        
            draw_cable_name(ax, start_x - 1.2, CAPSULE_Y_CENTER_BASE + y_offset, letter)
        
            capsule_y_center = CAPSULE_Y_CENTER_BASE + y_offset
            y_top_bus_group = capsule_y_center + y_top_bus_offset
            y_bottom_bus_group = capsule_y_center + y_bottom_bus_offset
        
            for cable_id in cable_list:
                cable_rows = df_cable[df_cable['cable_id'] == cable_id]
                cable_box_rows = df_cable_box[df_cable_box['cable_id'] == cable_id] if not df_cable_box.empty else pd.DataFrame()
            
                is_cable_box = not cable_box_rows.empty
            
                if is_cable_box:
                    cable_info = cable_box_rows.iloc[0]
                    cable_type = ""
                    if pd.notna(cable_info.get('cable_type')):
                        cable_type = str(cable_info.get('cable_type')).strip().lower()
                    elif pd.notna(cable_info.get('cabel_type')):
                        cable_type = str(cable_info.get('cabel_type')).strip().lower()
                
                    if cable_type == 'relay_box':
                        position_val = cable_info.get('position')
                        if pd.notna(position_val):
                            try:
                                position_num = int(float(position_val))
                            except:
                                position_num = 1
                        else:
                            position_num = 1
                    
                        if current_terminal_count + position_num > max_terminal_symbols_per_row:
                            print(f"ERROR: Cable boxes in row '{letter}' exceed row limit")
                            continue
                    
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
                        
                            current_x += pin_spacing * 2
                            current_row_max_x = max(current_row_max_x, current_x)
                        
                            min_y = min(min_y, y_bottom_bus_group - 1.8)
                            max_y = max(max_y, y_top_bus_group + 1.8)
                
                    current_x += CABLE_GAP
                    current_row_max_x = max(current_row_max_x, current_x)
                    continue
            
                if not cable_rows.empty:
                    cable_info = cable_rows.iloc[0]
                    cable_type = ""
                    if pd.notna(cable_info.get('cable_type')):
                        cable_type = str(cable_info.get('cable_type')).strip().lower()
                    elif pd.notna(cable_info.get('cabel_type')):
                        cable_type = str(cable_info.get('cabel_type')).strip().lower()
                
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
                
                    if current_terminal_count + symbols_to_add_total > max_terminal_symbols_per_row:
                        print(f"ERROR: Cable {cable_id} would exceed row limit")
                        continue
                
                    i = 0
                    while i < len(group):
                        row = group.iloc[i]
                        symbol = str(row.get('symbol', '')).strip().lower()
                        symbols_to_add = 2 if symbol == 'dual_fuse' else 1
                    
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
                            continue
                        elif symbol == 'dual_fuse':
                            if i + 1 < len(group):
                                next_row = group.iloc[i+1]
                                LEFT_EXTRA = pin_spacing * 1.0
                                AFTER_SPACING = pin_spacing * 1.5
                                current_x += LEFT_EXTRA
                                dual_start_x = current_x - SYMBOL_WIDTH * 1.25
                           
                                first_input_connected = row.get('input_connected', 'N')
                                first_output_connected = row.get('output_connected', 'N')
                           
                                top_conn, bottom_conn, left_ic, left_oc, right_ic, right_oc = draw_dual_fuse(
                                    ax, dual_start_x, capsule_y_center,
                                    row.get('terminal_no'),
                                    next_row.get('terminal_no'),
                                    row.get('input_left'), row.get('input_right'), row.get('output_left'), row.get('output_right'),
                                    first_input_connected, first_output_connected,
                                    next_row.get('input_left'), next_row.get('input_right'), row.get('output_left'), next_row.get('output_right'),
                                    first_input_connected, first_output_connected
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
                    
                    current_x += pin_spacing
                    current_row_max_x = max(current_row_max_x, current_x)
                
                    draw_extra_connections(ax, group, x_positions, terminal_nos_for_positions,
                                        y_top_bus_group, y_bottom_bus_group, capsule_y_center)
                
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
                            print(f"WARNING: Resistor would exceed row limit")
                   
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
                       
                    choke_rows = df_choke[df_choke['cable_id'] == cable_id]
                    special_choke = False
                    vert_x = None
                    
                    for choke_idx, choke_row in choke_rows.iterrows():
                        input_term = str(choke_row.get('input_terminal', pd.NA)).strip().replace('.0', '') if 'input_terminal' in choke_row else ''
                        output_term = str(choke_row.get('output_terminal', pd.NA)).strip().replace('.0', '') if 'output_terminal' in choke_row else ''
                        output_type = str(choke_row.get('output_type', 'terminal')).strip().lower() if 'output_type' in choke_row else 'terminal'
                        output_text = str(choke_row.get('output_text', '')).strip() if 'output_text' in choke_row else ''
                   
                        output_connected_terminals = []
                        if 'output_connected' in choke_row and pd.notna(choke_row['output_connected']):
                            output_connected_val = str(choke_row['output_connected']).strip()
                            if output_connected_val and output_connected_val != '':
                                output_connected_terminals = [t.strip() for t in output_connected_val.split(',')]
                           
                        if input_term and input_term in terminal_nos_for_positions:
                            start_idx = terminal_nos_for_positions.index(input_term)
                            x_left = x_positions[start_idx]
                            choke_label = str(choke_row.get('terminal_name', pd.NA)).strip() if 'terminal_name' in choke_row and pd.notna(choke_row['terminal_name']) else 'CHOKE'
                            input_symbol_type = None
                            input_symbol_row = group[group['terminal_no'].astype(str).str.replace('.0', '') == input_term]
                            if not input_symbol_row.empty:
                                input_symbol_type = str(input_symbol_row.iloc[0].get('symbol', '')).strip().lower()
                            
                            box_width = 0.6
                            box_height = 0.3
                            
                            if output_term and output_term in terminal_nos_for_positions and output_type != 'relay':
                                end_idx = terminal_nos_for_positions.index(output_term)
                                x_right = x_positions[end_idx]
                                choke_x_center = (x_left + x_right) / 2
                                
                                if input_symbol_type == 'dual_fuse':
                                    vertical_drop = 0.25
                                    horizontal_extension = 0.005
                                else:
                                    vertical_drop = 0.3
                                    horizontal_extension = 0.0
                                
                                vertical_position = y_bottom_bus_group - vertical_drop
                                left_extra_up = 0.8
                                original_bottom_left = y_bottom_bus_group + left_extra_up
                                ax.plot([x_left, x_left],
                                        [original_bottom_left, vertical_position],
                                        color='black', linewidth=1)
                                
                                left_extra_extend_down = 1.3
                                if input_symbol_type == 'dual_fuse' and not output_connected_terminals:
                                    left_extra_extend_down = 1.24
                                elif output_connected_terminals:
                                    left_extra_extend_down = 1.24
                                
                                ax.plot(
                                    [x_left, x_left],
                                    [original_bottom_left, original_bottom_left - left_extra_extend_down],
                                    color='black',
                                    linewidth=1
                                )
                                
                                left_connection_start_x = x_left - horizontal_extension
                                line_down = 0.2
                                cut_amount_left = 0.02
                                ax.plot([left_connection_start_x, (choke_x_center - box_width/2) - cut_amount_left],
                                        [vertical_position - line_down, vertical_position - line_down],
                                        color='black', linewidth=1)
                                
                                right_connection_end_x = x_right + horizontal_extension
                                if output_connected_terminals:
                                    right_connection_end_x += 0.19
                                    cut_from_left = 0.02
                                    right_horizontal_start = (choke_x_center + box_width/2) + cut_from_left
                                    right_horizontal_end = right_connection_end_x
                                else:
                                    cut_from_left = 0.02
                                    right_horizontal_start = (choke_x_center + box_width/2) + cut_from_left
                                    right_horizontal_end = right_connection_end_x
                                
                                line_down_right = 0.20
                                ax.plot(
                                    [right_horizontal_start, right_horizontal_end],
                                    [vertical_position - line_down_right, vertical_position - line_down_right],
                                    color='black', linewidth=1
                                )
                                
                                if output_connected_terminals:
                                    right_vertical_x = x_right + 0.20
                                    right_end_y = y_top_bus_group
                                else:
                                    right_vertical_x = x_right
                                    right_end_y = y_bottom_bus_group
                                
                                right_extra_up = 0.4
                                original_top_right = right_end_y + right_extra_up
                                extend_bottom = 0.2
                                ax.plot(
                                    [right_vertical_x, right_vertical_x],
                                    [vertical_position - extend_bottom, original_top_right],
                                    color='black',
                                    linewidth=1
                                )
                                
                                right_extra_extend_down = 0.9
                                if input_symbol_type == 'dual_fuse' and not output_connected_terminals:
                                    right_extra_extend_down = 0.72
                                elif output_connected_terminals:
                                    right_extra_extend_down = 0.85
                                
                                ax.plot(
                                    [right_vertical_x, right_vertical_x],
                                    [original_top_right, original_top_right - right_extra_extend_down],
                                    color='black',
                                    linewidth=1
                                )
                                
                                if output_connected_terminals:
                                    small_line_length = 0.2
                                    left_end_x = right_vertical_x - small_line_length
                                    y_top = right_end_y + right_extra_up
                                    ax.plot([left_end_x, right_vertical_x],
                                            [y_top, y_top],
                                            color='black', linewidth=1)
                                    small_vertical_drop = 0.9
                                    ax.plot([left_end_x, left_end_x],
                                            [y_top, y_top - small_vertical_drop],
                                            color='black', linewidth=1)
                                
                                if output_connected_terminals:
                                    connected_x = []
                                    for term in output_connected_terminals:
                                        term_str = str(term).strip().replace('.0', '')
                                        if term_str in terminal_nos_for_positions:
                                            idx = terminal_nos_for_positions.index(term_str)
                                            connected_x.append(x_positions[idx])
                                    
                                    if connected_x:
                                        connected_x_sorted = sorted(connected_x)
                                        min_x = min(connected_x_sorted)
                                        max_x = max(connected_x_sorted)
                                        y_bus = y_bottom_bus_group - 0.8
                                        
                                        offset = 0.2
                                        drop_length = 0.3
                                        ax.plot(
                                            [min_x - offset, max_x],
                                            [y_bus, y_bus],
                                            color='black',
                                            linewidth=1
                                        )
                                        ax.plot(
                                            [min_x - offset, min_x - offset],
                                            [y_bus, y_bus - drop_length],
                                            color='black',
                                            linewidth=1
                                        )
                                        
                                        for i, x in enumerate(connected_x_sorted, 1):
                                            ax.plot([x, x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
                                            number_y = y_bus - 0.05
                                            ax.text(
                                                x, number_y, str(i),
                                                fontsize=16, ha='center', va='top', fontname=DEFAULT_FONT, fontweight='bold'
                                            )
                                        
                                        choke_output_x = x_right
                                        ax.plot([choke_output_x, choke_output_x], [y_bottom_bus_group, y_bus], color='black', linewidth=1)
                                        ax.plot([choke_output_x, max_x], [y_bus, y_bus], color='black', linewidth=1)
                                        
                                        if output_text:
                                            try:
                                                text_x = (min_x + max_x) / 2.0
                                                text_y = y_bus - 0.20
                                                ax.text(
                                                    text_x, text_y, output_text,
                                                    fontsize=22, ha='center', va='top', fontname=DEFAULT_FONT,
                                                    bbox=dict(boxstyle="round,pad=0.1", facecolor='white', edgecolor='none', alpha=0.0)
                                                )
                                            except Exception:
                                                pass
                                
                                down_shift = 0.2
                                choke_box = FancyBboxPatch(
                                    (choke_x_center - box_width/2, (vertical_position - box_height/2) - down_shift),
                                    box_width, box_height,
                                    boxstyle="round,pad=0.02",
                                    edgecolor='black', facecolor='white', linewidth=1.5
                                )
                                ax.add_patch(choke_box)
                                
                                ax.text(choke_x_center, vertical_position - down_shift,
                                        choke_label, fontsize=16, ha='center', va='center', fontname=DEFAULT_FONT)
                            else:
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
                    
                    cable_headers_temp = df_header[df_header['cable_id'] == cable_id]
                    top_ranges = []
                    bottom_ranges = []
                   
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
                   
                    merge_adjacent = False
                    top_segments = merge_ranges(top_ranges, merge_adjacent=merge_adjacent)
                    bottom_segments = merge_ranges(bottom_ranges, merge_adjacent=merge_adjacent)
                   
                    print(f" Top segments after merge: {top_segments}")
                    print(f" Bottom segments after merge: {bottom_segments}")
                   
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
               
                    has_input_bus = len(top_segments) > 0
                    has_output_bus = len(bottom_segments) > 0
               
                    for j, x in enumerate(x_positions):
                        if j > 0 and x == x_positions[j-1]:
                            continue
                        symbol_top_y = capsule_y_center + SYMBOL_HEIGHT/2 + SYMBOL_RADIUS
                        symbol_bottom_y = capsule_y_center - SYMBOL_HEIGHT/2 - SYMBOL_RADIUS
                    
                        hooked_in = draw_input_connection(ax, x, symbol_top_y, input_connected_flags[j], output_connected_flags[j], y_top_bus_group, has_input_bus)
                        hooked_out = draw_output_connection(ax, x, symbol_bottom_y, input_connected_flags[j], output_connected_flags[j], y_bottom_bus_group, has_output_bus)
                    
                        hook_input_flags.append(hooked_in)
                        hook_output_flags.append(hooked_out)
                    
                    UPPER_SHIFT = 0.2
                    for min_idx, max_idx in top_segments:
                        sub_x = x_positions[min_idx : max_idx + 1]
                        sub_flags = input_connected_flags[min_idx : max_idx + 1]
                        draw_bus_lines(ax, sub_x, sub_flags, y_top_bus_group, gap=0.12, y_offset=UPPER_SHIFT)
                        connected_local = [i for i, f in enumerate(sub_flags) if f]
                        if connected_local:
                            first_local = connected_local[0]
                            x_first = sub_x[first_local]
                            y_bus = y_top_bus_group + UPPER_SHIFT
                            ax.plot([x_first - 0.3, x_first], [y_bus, y_bus], color='black', linewidth=1)
                            ax.plot([x_first - 0.3, x_first - 0.3], [y_bus, y_bus + 0.2], color='black', linewidth=1)
                    
                    LOWER_SHIFT = -0.2
                    for min_idx, max_idx in bottom_segments:
                        sub_x = x_positions[min_idx : max_idx + 1]
                        sub_flags = output_connected_flags[min_idx : max_idx + 1]
                        draw_bus_lines(ax, sub_x, sub_flags, y_bottom_bus_group, gap=0.12, y_offset=LOWER_SHIFT)
                        connected_local = [i for i, f in enumerate(sub_flags) if f]
                        if connected_local:
                            last_local = connected_local[-1]
                            x_last = sub_x[last_local]
                            if not ((special_choke and max_idx == start_idx and min_idx <= start_idx) or special_resistor):
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
                           
                            if start_name in terminal_nos_for_positions and end_name in terminal_nos_for_positions:
                                start_idx_group = terminal_nos_for_positions.index(start_name)
                                end_idx_group = terminal_nos_for_positions.index(end_name)
                                if start_idx_group > end_idx_group:
                                    start_idx_group, end_idx_group = end_idx_group, start_idx_group
                               
                                if io_field == 'input':
                                    is_not_connected = not any(name_to_input_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_group:end_idx_group+1])
                                    has_bus_line = len(top_segments) > 0
                                    is_not_connected_with_bus = is_not_connected and has_bus_line
                               
                                    if is_not_connected_with_bus:
                                        y_relay = y_top_bus_group + 0.85
                                    else:
                                        y_relay = y_top_bus_group + 0.55
                               
                                    draw_relay_input(ax, x_start_term, x_end_term, y=y_relay, scale=1.0, text=str(label_text), is_not_connected_with_bus=is_not_connected_with_bus)
                                elif io_field == 'output':
                                    is_not_connected = not any(name_to_output_connected.get(term, False) for term in terminal_nos_for_positions[start_idx_group:end_idx_group+1])
                                    has_bus_line = len(bottom_segments) > 0
                                    is_not_connected_with_bus = is_not_connected and has_bus_line
                               
                                    if is_not_connected_with_bus:
                                        y_relay = y_bottom_bus_group - 0.85
                                    else:
                                        y_relay = y_bottom_bus_group - 0.55
                               
                                    draw_relay_output(ax, x_start_term, x_end_term, y=y_relay, scale=1.0, text=str(label_text), is_not_connected_with_bus=is_not_connected_with_bus)
                                else:
                                    center_x = (x_start_term + x_end_term) / 2.0
                                    ax.text(center_x, y_top_bus_group + 0.2, str(label_text), ha='center', va='bottom', fontsize=10, fontname=DEFAULT_FONT)
                    
                    cable_headers = df_header[df_header['cable_id'] == cable_id]
                    relay_top = {}
                    relay_bottom = {}
                    relay_box_top = {}
                    relay_box_bottom = {}
                    relay_contact_box_top = {}
                    relay_contact_box_bottom = {}
                
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
                    
                        is_not_connected = not input_conn_flag
                        has_bus_line = len(top_segments) > 0
                        is_not_connected_with_bus = is_not_connected and has_bus_line
                    
                        if is_not_connected_with_bus:
                            vertical_line_start = symbol_top_y + stub_length - 0.42
                        elif input_conn_flag:
                            vertical_line_start = y_top_bus_group
                        else:
                            vertical_line_start = symbol_top_y + stub_length
                        
                        draw_group_top_symbol(ax, x_left, x_right, vertical_line_start, texts=texts, scale=1.0, input_connected='Y' if input_conn_flag else 'N')
                
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
                    
                        is_not_connected = not output_conn_flag
                        has_bus_line = len(bottom_segments) > 0
                        is_not_connected_with_bus = is_not_connected and has_bus_line
                    
                        if is_not_connected_with_bus:
                            vertical_line_end = symbol_bottom_y - stub_length + 0.38
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
           
            y_offset -= vertical_gap
    
        overall_max_x = max(overall_max_x, current_row_max_x)
        if all_x_positions:
            content_min_x = min(all_x_positions)
            content_max_x = max(all_x_positions)
        else:
            content_min_x = start_x
            content_max_x = start_x + (page_width if page_width else fixed_fig_width)
        
        content_width = max(0.1, content_max_x - content_min_x)
        desired_width = global_max_width if global_max_width is not None else (page_width if page_width else 30.0)
        left = start_x - 1.5
        right = left + desired_width
        page_center_x = (left + right) / 2.0
        ax.set_xlim(left, right)
        ax.set_ylim(fixed_ylim_min, fixed_ylim_max)
       
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
       
        junction_box_y = CAPSULE_Y_CENTER_BASE + y_top_bus_offset +1.8 + 3.0 -1.0
        draw_junction_box(ax, page_center_x, junction_box_y, junction_name)
       
        return all_x_positions, all_input_connected_flags, all_output_connected_flags
    except Exception as e:
        print(f"Error in draw_symbols: {e}")
        traceback.print_exc()
        return [], [], []

def draw_footer(ax, left, right, fixed_ylim_min, total_pages, page_num, df_title_row, junction_name):
    """
    Draw a compact footer (title-block style)
    """
    try:
        if df_title_row is None:
            return

        width = 20.0
        extra_width = 6.0
        height = 3.5
        total_block_width = width + extra_width

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

        ax.add_patch(Rectangle((footer_x_start, footer_y_start), outer_width, outer_height,
                            fill=False, linewidth=LINEWIDTH))

        v_lines = [8, 12, 17.5, width + extra_width / 2]
        for vx in v_lines:
            x = footer_x_start + vx * x_scale
            ax.plot([x, x], [footer_y_start, footer_y_start + outer_height], 'k-', linewidth=LINEWIDTH)

        x14_8 = footer_x_start + 14.8 * x_scale
        ax.plot([x14_8, x14_8], [footer_y_start + s(1.5), footer_y_start + outer_height], 'k-', linewidth=LINEWIDTH)

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

        ax.plot([footer_x_start + 4 * x_scale, footer_x_start + 4 * x_scale],
                [footer_y_start + 0.0, footer_y_start + s(1.5)], 'k-', linewidth=LINEWIDTH)

        ax.plot([footer_x_start + 25 * x_scale, footer_x_start + 25 * x_scale],
                [footer_y_start + s(-0.02), footer_y_start + s(1.5)], 'k-', linewidth=LINEWIDTH)

        company_y = footer_y_start + s(2.6) - 0.02
        ax.text(footer_x_start + 0.1 * x_scale, company_y,
                "SALTRIVER INFOSYSTEMS\nPRIVATE LTD.,\nAHMEDABAD.",
                va='top', ha='left', fontsize=22, fontname=DEFAULT_FONT, linespacing=1.2)

        ax.text(footer_x_start + 1.1 * x_scale, footer_y_start + s(0.5),
                "DRAWN BY", va='center', ha='left', fontsize=22, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 5.1 * x_scale, footer_y_start + s(0.5),
                "CHECKED BY", va='center', ha='left', fontsize=22, fontname=DEFAULT_FONT)

        drawn_by = df_title_row.get('drawn_by')
        checked_by = df_title_row.get('checked_by')

        if pd.notna(drawn_by):
            ax.text(footer_x_start + 1.4 * x_scale, footer_y_start + s(1.2),
                    str(drawn_by), va='bottom', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        if pd.notna(checked_by):
            ax.text(footer_x_start + 5.4 * x_scale, footer_y_start + s(1.2),
                    str(checked_by), va='bottom', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(2.75),
                str(df_title_row.get('designation1', '')), va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)
        ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(2.25),
                str(df_title_row.get('designation2', '')), va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)
        ax.text(footer_x_start + 15 * x_scale, footer_y_start + s(1.82),
                str(df_title_row.get('designation3', '')), va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 19 * x_scale, footer_y_start + s(2.75),
                str(df_title_row.get('station_name', '')), va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 18 * x_scale, footer_y_start + s(2.0),
                junction_name, va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 18.5 * x_scale, footer_y_start + s(0.9),
                f"DRG. NO. {df_title_row.get('station_code', '')}", va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        def format_text(t):
            t = str(t)
            if len(t) > 12:
                words = t.split()
                if len(words) > 1:
                    return " ".join(words[:-1]) + "\n" + words[-1]
                return t[:12] + "\n" + t[12:]
            return t

        ax.text(footer_x_start + 24 * x_scale, footer_y_start + s(2.8),
                format_text(df_title_row.get('zone', '')),
                va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 23.3 * x_scale, footer_y_start + s(2.3),
                format_text(f"{df_title_row.get('division', '')} DIVISION"),
                va='top', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + 25.25 * x_scale, footer_y_start + s(0.25),
                str(total_pages), va='bottom', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)

        ax.text(footer_x_start + (width + 3.6) * x_scale, footer_y_start + s(1.1),
                "SHEET\nNO", va='center', ha='left', fontsize=FONTSIZE,
                fontname=DEFAULT_FONT, linespacing=1.0)

        ax.text(footer_x_start + (width + 3.5) * x_scale, footer_y_start + s(0.35),
                "TOTAL\nSHEETS", va='center', ha='left', fontsize=FONTSIZE,
                fontname=DEFAULT_FONT, linespacing=1.0)

        ax.text(footer_x_start + (width + 5.25) * x_scale, footer_y_start + s(1.1),
                str(page_num), va='center', ha='left', fontsize=FONTSIZE, fontname=DEFAULT_FONT)
    except Exception as e:
        print(f"Error in draw_footer: {e}")
        traceback.print_exc()

# === NEW FUNCTIONS FOR PAGINATION AND ROW ORGANIZATION ===

def organize_junction_rows(junction_cables_regular, junction_cables_box):
    """
    Organize junction rows with proper sorting and grouping
    """
    try:
        # Combine regular cables and cable boxes if both exist
        if not junction_cables_regular.empty and not junction_cables_box.empty:
            junction_cables = pd.concat([junction_cables_regular, junction_cables_box], ignore_index=True)
        elif not junction_cables_regular.empty:
            junction_cables = junction_cables_regular.copy()
        elif not junction_cables_box.empty:
            junction_cables = junction_cables_box.copy()
        else:
            return OrderedDict()
        
        # Enhanced sorting with proper row ordering
        junction_cables = junction_cables.sort_values(
            by=['row', 'position'], 
            key=lambda x: x.map(get_row_order) if x.name == 'row' else x,
            ascending=[False, False]
        )
        
        # Group by row letter
        letter_groups = OrderedDict()
        
        for _, cable_row in junction_cables.iterrows():
            letter = str(cable_row.get('row', '')).strip()
            if not letter:
                continue
            cable_id = cable_row['cable_id']
            
            if letter not in letter_groups:
                letter_groups[letter] = []
            letter_groups[letter].append(cable_id)
        
        return letter_groups
    except Exception as e:
        print(f"Error in organize_junction_rows: {e}")
        traceback.print_exc()
        return OrderedDict()

def create_pages_for_junction(junction, letter_groups, max_rows_per_page=3):
    """
    Create pages for a junction with exactly max_rows_per_page rows per page
    """
    try:
        sorted_letters = sorted(letter_groups.keys(), key=get_row_order, reverse=True)
        
        print(f"\n=== Processing Junction: {junction} ===")
        print(f"All letters in junction: {sorted_letters}")
        print(f"Letter groups: {[(letter, len(cables)) for letter, cables in letter_groups.items()]}")
        
        all_rows = []
        for letter in sorted_letters:
            cable_list = letter_groups[letter]
            letter_rows = break_cables_into_rows_updated(cable_list, max_terminal_symbols_per_row=36, max_cable_boxes_per_row=6)
            
            for i, row_cables in enumerate(letter_rows):
                reversed_row_cables = list(reversed(row_cables))
                all_rows.append((letter, reversed_row_cables))
        
        pages = []
        all_rows_reversed = list(reversed(all_rows))
        chunks = [all_rows_reversed[i:i+max_rows_per_page] for i in range(0, len(all_rows_reversed), max_rows_per_page)]
        chunks_reversed = list(reversed(chunks))
        
        for chunk in chunks_reversed:
            chunk_reversed = list(reversed(chunk))
            pages.append((junction, chunk_reversed))
        
        return pages
    except Exception as e:
        print(f"Error in create_pages_for_junction for junction '{junction}': {e}")
        traceback.print_exc()
        return []

def process_excel_file(excel_file_path):
    """
    Process a single Excel file and generate PDF
    """
    try:
        print("\n" + "="*60)
        print(f"PROCESSING FILE: {excel_file_path}")
        print("="*60)
        
        excel_filename = os.path.basename(excel_file_path)
        print(f"Processing Excel file: {excel_filename}")

        # Validate required sheets exist
        try:
            xls = pd.ExcelFile(excel_file_path)
        except Exception as e:
            print(f"Unable to open Excel file: {e}")
            traceback.print_exc()
            return False, "Error opening Excel file"

        required_sheets = ['terminal', 'junction_box', 'terminal_header', 'group', 'cable']
        available_sheets = [s.strip() for s in xls.sheet_names]
        missing = [s for s in required_sheets if s not in available_sheets]
        if missing:
            print(f"Excel file is missing required sheets: {missing}")
            print(f"Available sheets: {available_sheets}")
            return False, f"Missing required sheets: {missing}"

        # Load data
        global df_cable, df_cable_box, df, df_junction, df_header, df_group, df_choke, df_resistor, df_title, df_symbols
        
        try:
            df_cable = pd.read_excel(excel_file_path, sheet_name='cable')
            df_cable.columns = df_cable.columns.str.strip()
            
            try:
                df_cable_box = pd.read_excel(excel_file_path, sheet_name='cable_box')  # FIXED: Changed from 'relay_box' to 'cable_box'
                df_cable_box.columns = df_cable_box.columns.str.strip()
                print("Loaded relay boxes from 'cable_box' sheet")
            except Exception as e:
                # Fallback: extract cable boxes from main cable sheet
                print(f"No 'cable_box' sheet found: {e}. Creating empty relay_box DataFrame.")
                
                # Create empty DataFrame with required structure
                df_cable_box = pd.DataFrame(columns=['cable_id', 'cable_name', 'junction_name', 'row', 
                                                'position', 'terminal', 'start_no', 'cabel_type', 
                                                'cable_letter', 'letter_order'])
                
                print("Created empty relay_box DataFrame.")
                
        except Exception as e:
            print(f"Error reading cable sheets from Excel file: {e}")
            traceback.print_exc()
            return False, "Error reading cable sheets"

        # Load other required sheets
        try:
            df = pd.read_excel(excel_file_path, sheet_name='terminal')
            df.columns = df.columns.str.strip()
            if 'terminal_name' in df.columns:
                df.rename(columns={'terminal_name': 'terminal_no'}, inplace=True)
            df_junction = pd.read_excel(excel_file_path, sheet_name='junction_box')
            df_junction.columns = df_junction.columns.str.strip()
            df_header = pd.read_excel(excel_file_path, sheet_name='terminal_header')
            df_header.columns = df_header.columns.str.strip()
            df_group = pd.read_excel(excel_file_path, sheet_name='group')
            df_group.columns = df_group.columns.str.strip()
            df_choke = pd.read_excel(excel_file_path, sheet_name='choketable')
            df_choke.columns = df_choke.columns.str.strip()
            df_resistor = pd.read_excel(excel_file_path, sheet_name='resistortable')
            df_resistor.columns = df_resistor.columns.str.strip()
            
            try:
                df_title = pd.read_excel(excel_file_path, sheet_name='StationDrawing')
                df_title.columns = df_title.columns.str.strip()
                print("Loaded StationDrawing sheet for footer.")
            except Exception as e:
                print(f"Warning: Could not load StationDrawing sheet for footer: {e}. Footer will be skipped.")
                df_title = None
                
        except Exception as e:
            print(f"Error reading required sheets from Excel file: {e}")
            traceback.print_exc()
            return False, "Error reading required sheets"
        finally:
            try:
                xls.close()
            except Exception:
                pass

        if 'spare' in df.columns:
            try:
                df.loc[df['spare'].astype(str).str.upper() == 'Y', 'input_left'] = 'SP'
            except Exception as e:
                print(f"Error processing spare column: {e}")
                traceback.print_exc()

        # Prepare Plotting
        df_symbols = df.reset_index(drop=True)

        # Safety check for df_cable_box
        if df_cable_box is None or df_cable_box.empty:
            print("No cable box data found. Creating empty DataFrame with required structure.")
            df_cable_box = pd.DataFrame(columns=['cable_id', 'cable_name', 'junction_name', 'row', 
                                            'position', 'terminal', 'start_no', 'cabel_type', 
                                            'cable_letter', 'letter_order'])

        # Prepare cable data
        try:
            df_cable['cable_letter'] = df_cable['row'].astype(str).str.strip()
            df_cable['letter_order'] = df_cable['cable_letter'].apply(
                lambda x: ord(x.upper()) - ord('A') if pd.notna(x) and x.strip() != '' and x.strip() in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' else -1
            )

            if not df_cable_box.empty and 'row' in df_cable_box.columns:
                df_cable_box['cable_letter'] = df_cable_box['row'].astype(str).str.strip()
                df_cable_box['letter_order'] = df_cable_box['cable_letter'].apply(
                    lambda x: ord(x.upper()) - ord('A') if pd.notna(x) and x.strip() != '' and x.strip() in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' else -1
                )
            else:
                if 'cable_letter' not in df_cable_box.columns:
                    df_cable_box['cable_letter'] = ''
                if 'letter_order' not in df_cable_box.columns:
                    df_cable_box['letter_order'] = -1
        except Exception as e:
            print(f"Error preparing cable data: {e}")
            traceback.print_exc()

        # Get unique junction names
        junction_names = pd.unique(df_cable['junction_name'].astype(str).str.strip())

        if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
            try:
                cable_box_junctions = pd.unique(df_cable_box['junction_name'].astype(str).str.strip())
                for junction in cable_box_junctions:
                    if junction not in junction_names:
                        junction_names = np.append(junction_names, junction)
            except Exception as e:
                print(f"Error processing cable box junction names: {e}")
                traceback.print_exc()
        else:
            print("No cable box data to process for junction names.")

        print(f"DEBUG: Found {len(junction_names)} junctions: {list(junction_names)}")

        # Collect junction data for database
        junction_data = []
        for junction in junction_names:
            junction_data.append({
                "junction_name": junction,
                "junction_size": "Full",
                "junction_row": str(get_row_order(junction[0]) if junction and junction[0].isalpha() else 0)
            })

        # Compute max_row_width for each junction
        junction_row_widths = {}
        for junction in junction_names:
            try:
                current_x_pre = 1
                current_row_max_x_pre = 1
                current_terminal_count_pre = 0
                current_cable_box_count_pre = 0
                current_letter = None
                max_row_width = 0
                
                junction_mask_regular = df_cable['junction_name'].astype(str).str.strip() == junction
                junction_cables_regular = df_cable[junction_mask_regular].copy()
                
                junction_mask_box = df_cable_box['junction_name'].astype(str).str.strip() == junction if not df_cable_box.empty else pd.Series([False] * len(df_cable_box))
                junction_cables_box = df_cable_box[junction_mask_box].copy()
                
                junction_cables = pd.concat([junction_cables_regular, junction_cables_box], ignore_index=True)
                junction_cables = junction_cables.sort_values(
                    by=['row', 'position'], 
                    key=lambda x: x.map(get_row_order) if x.name == 'row' else x,
                    ascending=[False, False]
                )
                
                cable_list = junction_cables['cable_id'].tolist()
                
                print(f"\n=== Junction: {junction} ===")
                print("Cable order for width calculation:")
                for cable_id in cable_list:
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
                    
                    cable_box_rows = df_cable_box[df_cable_box['cable_id'] == cable_id_pre] if not df_cable_box.empty else pd.DataFrame()
                    if not cable_box_rows.empty:
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
                    
                    if is_cable_box and current_cable_box_count_pre >= 6 and current_terminal_count_pre > 0:
                        max_row_width = max(max_row_width, current_row_max_x_pre - 1)
                        current_row_max_x_pre = 1
                        current_x_pre = 1
                        current_terminal_count_pre = 0
                        current_cable_box_count_pre = 0
                    
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
            except Exception as e:
                print(f"Error computing row widths for junction '{junction}': {e}")
                traceback.print_exc()
                junction_row_widths[junction] = 30.0

        global_max_width = max(junction_row_widths.values()) + 2.0 if junction_row_widths else 30.0

        # Compute page dimensions
        max_rows_visible = 3
        max_terminal_symbols_per_row = 36
        max_cable_boxes_per_row = 6

        fixed_fig_width = 42.8
        fixed_fig_height = 31.0
        bottom_margin = 1.0
        top_margin = 3.0
        fixed_ylim_min = CAPSULE_Y_CENTER_BASE + vertical_gap * (1 - max_rows_visible) + y_bottom_bus_offset - 1.8 - bottom_margin - footer_height
        fixed_ylim_max = CAPSULE_Y_CENTER_BASE + y_top_bus_offset + 1.8 + top_margin

        # Pagination logic
        pages = []
        for junction in junction_names:
            try:
                junction_mask = df_cable['junction_name'].astype(str).str.strip() == junction
                junction_cables_regular = df_cable[junction_mask].copy()
                
                junction_cables_box = pd.DataFrame()
                if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
                    junction_mask_box = df_cable_box['junction_name'].astype(str).str.strip() == junction
                    junction_cables_box = df_cable_box[junction_mask_box].copy()
                else:
                    print(f"No cable box data for junction: {junction}")
                
                if junction_cables_regular.empty and junction_cables_box.empty:
                    print(f"Warning: No data found for junction '{junction}'. Skipping.")
                    continue
                
                letter_groups = organize_junction_rows(junction_cables_regular, junction_cables_box)
                junction_pages = create_pages_for_junction(junction, letter_groups, max_rows_per_page=3)
                pages.extend(junction_pages)
            except Exception as e:
                print(f"Error processing junction '{junction}': {e}")
                traceback.print_exc()

        total_pages = len(pages)
        print(f"\nTotal pages: {total_pages}")

        print("\n=== PAGE BREAKDOWN ===")
        for i, (junction, page_rows) in enumerate(pages, 1):
            row_info = [f"{letter}({len(cables)} cables)" for letter, cables in page_rows]
            print(f"Page {i}: {junction} - Rows: {row_info}")

        title_row = df_title.iloc[0] if df_title is not None and not df_title.empty else None

        # Generate PDF filename
        excel_basename = os.path.splitext(excel_filename)[0]
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        timestamp_suffix = current_time.strftime("%Y-%m-%d_%H-%M-%S")
        output_filename = f"{excel_basename}_{timestamp_suffix}.pdf"
        output_file = os.path.join(PDF_OUTPUT_DIR, output_filename)

        # Generate initial checksum
        checksum, checksum_data, content_size, timestamp_ist = generate_pdf_metadata_checksum(df_title, excel_file_path, output_file)

        if os.path.exists(output_file):
            print(f"Warning: '{output_file}' already exists. Adding checksum suffix.")
            short_checksum = checksum[:8]
            output_filename = f"{excel_basename}_{timestamp_suffix}_{short_checksum}.pdf"
            output_file = os.path.join(PDF_OUTPUT_DIR, output_filename)
            checksum, checksum_data, content_size, timestamp_ist = generate_pdf_metadata_checksum(df_title, excel_file_path, output_file)

        # Generate PDF
        try:
            with PdfPages(output_file) as pdf:
                for page_num, (junction_name, page_rows) in enumerate(pages, 1):
                    try:
                        page_max_width = 0
                        current_x_page = 1
                        current_row_max_x_page = 1
                        current_terminal_count_page = 0
                        current_letter = None
                    
                        for letter, cable_list in page_rows:
                            for cid in cable_list:
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
                    
                        shift = (global_max_width - (page_max_width + 1.2)) / 2
                        page_start_x = 1 + shift
                    
                        pdf_info = pdf.infodict()
                        pdf_info['Title'] = f'Terminal Drawing - {junction_name}'
                        pdf_info['Author'] = 'SaltRiver Infosystems'
                        pdf_info['Subject'] = f'Station Code: {df_title.iloc[0].get("station_code", "") if df_title is not None else ""}'
                        pdf_info['Keywords'] = f'checksum:{checksum}' if checksum else ''
                        pdf_info['CreationDate'] = current_time.strftime("D:%Y%m%d%H%M%S+05'30'")
                    
                        fixed_fig_width = 42.8
                        fixed_fig_height = 31.0
                        fig, ax = plt.subplots(figsize=(fixed_fig_width, fixed_fig_height))
                        ax.set_facecolor('white')
                        ax.axis('off')
                    
                        y_offset = - (max_rows_visible - len(page_rows)) * vertical_gap
                        
                        x_positions, input_connected_flags, output_connected_flags = draw_symbols(
                            df_symbols, ax, page_rows, junction_name,
                            start_x=page_start_x,
                            pin_spacing=pin_spacing,
                            cables_per_page=sum(len(cables) for _, cables in page_rows),
                            page_number=page_num,
                            max_terminal_symbols_per_row=36,
                            max_rows_visible=3,
                            page_width=global_max_width,
                            global_max_width=global_max_width
                        )
                    
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
                    except Exception as e:
                        print(f"Error processing page {page_num} for junction '{junction_name}': {e}")
                        traceback.print_exc()
                        continue
          
            # Update with actual file size
            final_checksum, final_checksum_data, final_content_size, full_file_md5, metadata_ts_ist, initial_content_size = update_pdf_checksum_metadata(
                output_file, checksum, checksum_data, content_size, df_title
            )
        except Exception as e:
            print(f"Error creating PDF: {e}")
            traceback.print_exc()
            return False, "Error creating PDF"

        print(f"Multi-page PDF saved as '{output_file}'")

        # Enhance PDF with proper metadata
        try:
            enhance_pdf_with_metadata(output_file, final_checksum, final_checksum_data, final_content_size, df_title)
        except Exception as e:
            print(f"Error enhancing PDF metadata: {e}")
            traceback.print_exc()

        # Store PDF metadata in database
        try:
            db_conn = get_db_connection()
            if db_conn:
                pdf_metadata = {
                    'pdf_filename': output_filename,
                    'xlsx_filename': excel_filename,
                    'final_size_bytes': final_content_size,
                    'full_file_md5': full_file_md5,
                    'metadata_checksum': final_checksum,
                    'metadata_data': final_checksum_data,
                    'initial_size_bytes': initial_content_size,
                    'metadata_ts_ist': metadata_ts_ist,
                    'station_code': df_title.iloc[0].get('station_code', '') if df_title is not None and not df_title.empty else '',
                    'junction_data': junction_data
                }
                
                store_pdf_metadata(db_conn, pdf_metadata)
                db_conn.close()
                print("PDF metadata successfully stored in database.")
            else:
                print("Could not connect to database. Skipping database storage.")
        except Exception as e:
            print(f"Error storing PDF metadata in database: {e}")
            traceback.print_exc()

        print("=" * 60)
        print("PROCESSING COMPLETED SUCCESSFULLY!")
        print(f"Input Excel: {excel_filename}")
        print(f"Output PDF: {output_filename}")
        print(f"PDF saved to: {output_file}")
        print(f"Total pages generated: {total_pages}")
        print("=" * 60)
        
        return True, "Success"
        
    except Exception as e:
        print(f"CRITICAL ERROR: An unexpected error occurred in processing: {e}")
        traceback.print_exc()
        return False, str(e)

def move_excel_file(excel_file_path, success=True):
    """
    Move Excel file to appropriate directory after processing
    """
    try:
        excel_filename = os.path.basename(excel_file_path)
        
        if success:
            destination = os.path.join(PROCESSED_EXCEL_DIR, excel_filename)
            if os.path.exists(destination):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(excel_filename)
                new_filename = f"{name}_{timestamp}{ext}"
                destination = os.path.join(PROCESSED_EXCEL_DIR, new_filename)
            
            shutil.move(excel_file_path, destination)
            print(f"Moved Excel file to: {destination}")
        else:
            destination = os.path.join(ERROR_EXCEL_DIR, excel_filename)
            if os.path.exists(destination):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(excel_filename)
                new_filename = f"{name}_{timestamp}{ext}"
                destination = os.path.join(ERROR_EXCEL_DIR, new_filename)
            
            shutil.move(excel_file_path, destination)
            print(f"Moved Excel file to error directory: {destination}")
            
    except Exception as e:
        print(f"Error moving Excel file: {e}")
        traceback.print_exc()

def monitor_directory():
    """
    Monitor the input directory for new Excel files and process them
    """
    print("\n" + "="*60)
    print("STARTING 24x7 DIRECTORY MONITOR")
    print(f"Input Directory: {XLSX_INPUT_DIR}")
    print(f"PDF Output Directory: {PDF_OUTPUT_DIR}")
    print(f"Processed Excel Directory: {PROCESSED_EXCEL_DIR}")
    print(f"Error Excel Directory: {ERROR_EXCEL_DIR}")
    print("="*60 + "\n")
    
    processed_files = set()
    
    while running:
        try:
            excel_files = [f for f in os.listdir(XLSX_INPUT_DIR) if f.lower().endswith('.xlsx')]
            
            if excel_files:
                excel_files.sort(key=lambda f: os.path.getmtime(os.path.join(XLSX_INPUT_DIR, f)))
                
                for excel_file in excel_files:
                    if not running:
                        break
                        
                    excel_path = os.path.join(XLSX_INPUT_DIR, excel_file)
                    
                    if excel_path in processed_files:
                        continue
                    
                    # Check if file is still being written
                    try:
                        with open(excel_path, 'rb+'):
                            pass
                    except IOError:
                        print(f"File {excel_file} is still being written or locked. Skipping...")
                        continue
                    
                    print(f"\n" + "="*60)
                    print(f"STARTING PROCESSING FOR: {excel_file}")
                    print("="*60)
                    
                    processed_files.add(excel_path)
                    
                    success, message = process_excel_file(excel_path)
                    
                    if success:
                        print(f"\n✓ SUCCESS: {excel_file} processed successfully")
                        move_excel_file(excel_path, success=True)
                    else:
                        print(f"\n✗ ERROR: Failed to process {excel_file}: {message}")
                        move_excel_file(excel_path, success=False)
                    
                    print(f"\n✓ MOVED: {excel_file} to appropriate directory")
                    print("="*60 + "\n")
                    
                    time.sleep(2)
            else:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] No Excel files found. Waiting...")
                time.sleep(10)
                
        except KeyboardInterrupt:
            print("\nKeyboard interrupt received. Shutting down gracefully...")
            break
        except Exception as e:
            print(f"Error in monitor loop: {e}")
            traceback.print_exc()
            time.sleep(30)
    
    print("\n" + "="*60)
    print("DIRECTORY MONITOR STOPPED")
    print("="*60)

# === MAIN EXECUTION ===
if __name__ == "__main__":
    try:
        print("\n" + "="*60)
        print("RAILWAY PROJECT TERMINAL DIAGRAM GENERATOR")
        print("24x7 Directory Monitor with Auto File Processing")
        print("="*60)
        
        monitor_directory()
        
        print("\n" + "="*60)
        print("PROGRAM SHUTDOWN COMPLETE")
        print("="*60)
        
    except Exception as e:
        print(f"\nFATAL ERROR in main execution: {e}")
        traceback.print_exc()
        sys.exit(1)