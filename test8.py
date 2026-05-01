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
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading

# === MAIN SCRIPT - RUNS CONTINUOUSLY ===
def main():
    while True:
        try:
            # === CONFIGURATION ===
            XLSX_INPUT_DIR = "/Users/admin/Documents/test/git1/git/xlsx_download"
            PDF_OUTPUT_DIR = "/Users/admin/Documents/test/git1/git/uploads"
            
            # Create directories if they don't exist
            os.makedirs(XLSX_INPUT_DIR, exist_ok=True)
            os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
            
            # Create subdirectories for organization
            PROCESSED_EXCEL_DIR = os.path.join(PDF_OUTPUT_DIR, "processed_excel")
            ARCHIVE_EXCEL_DIR = os.path.join(PDF_OUTPUT_DIR, "archive_excel")
            os.makedirs(PROCESSED_EXCEL_DIR, exist_ok=True)
            os.makedirs(ARCHIVE_EXCEL_DIR, exist_ok=True)
            
            # Database configuration
            DB_CONFIG = {
                "host": "localhost",
                "port": "5432",
                "database": "postgres3",
                "user": "postgres",
                "password": ""
            }

            print(f"\n{'='*80}")
            print(f"Starting Terminal Diagram Generator - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Monitoring directory: {XLSX_INPUT_DIR}")
            print(f"Output directory: {PDF_OUTPUT_DIR}")
            print(f"{'='*80}\n")

            # Get all XLSX files in the directory
            xlsx_files = [f for f in os.listdir(XLSX_INPUT_DIR) if f.lower().endswith('.xlsx')]
            
            if not xlsx_files:
                print(f"No XLSX files found. Waiting for new files...")
                time.sleep(10)  # Wait 10 seconds before checking again
                continue
            
            # Process each XLSX file
            for xlsx_file in sorted(xlsx_files):
                try:
                    excel_path = os.path.join(XLSX_INPUT_DIR, xlsx_file)
                    
                    # Check if file is still being written (size hasn't changed in last 2 seconds)
                    if is_file_locked(excel_path):
                        print(f"File {xlsx_file} appears to be in use. Skipping for now...")
                        continue
                    
                    print(f"\n{'='*60}")
                    print(f"Processing: {xlsx_file}")
                    print(f"File size: {os.path.getsize(excel_path):,} bytes")
                    print(f"Last modified: {datetime.fromtimestamp(os.path.getmtime(excel_path))}")
                    print(f"{'='*60}")
                    
                    # Process the Excel file
                    success = process_excel_file(excel_path, DB_CONFIG, XLSX_INPUT_DIR, PDF_OUTPUT_DIR, 
                                                PROCESSED_EXCEL_DIR, ARCHIVE_EXCEL_DIR)
                    
                    if success:
                        print(f"✓ Successfully processed {xlsx_file}")
                    else:
                        print(f"✗ Failed to process {xlsx_file}")
                        # Move to error directory
                        error_dir = os.path.join(PDF_OUTPUT_DIR, "error_excel")
                        os.makedirs(error_dir, exist_ok=True)
                        error_path = os.path.join(error_dir, xlsx_file)
                        shutil.move(excel_path, error_path)
                        print(f"Moved to error directory: {error_path}")
                    
                except Exception as e:
                    print(f"Error processing {xlsx_file}: {e}")
                    traceback.print_exc()
                    continue
            
            # Wait before checking for new files again
            print(f"\nWaiting for new files... (Press Ctrl+C to stop)")
            time.sleep(5)
            
        except KeyboardInterrupt:
            print("\n\nShutdown requested. Exiting gracefully...")
            break
        except Exception as e:
            print(f"Unexpected error in main loop: {e}")
            traceback.print_exc()
            time.sleep(30)  # Wait 30 seconds before retrying

def is_file_locked(filepath):
    """Check if a file is locked/being written to"""
    try:
        # Try to open the file in append mode
        with open(filepath, 'a') as f:
            pass
        return False
    except IOError:
        return True

def process_excel_file(excel_path, db_config, xlsx_input_dir, pdf_output_dir, processed_excel_dir, archive_excel_dir):
    """Process a single Excel file and generate PDF"""
    try:
        excel_filename = os.path.basename(excel_path)
        
        # === Load and validate Excel data ===
        try:
            xls = pd.ExcelFile(excel_path)
        except Exception as e:
            print(f"Unable to open Excel file: {e}")
            return False

        required_sheets = ['terminal', 'junction_box', 'terminal_header', 'group', 'cable']
        available_sheets = [s.strip() for s in xls.sheet_names]
        missing = [s for s in required_sheets if s not in available_sheets]
        if missing:
            print(f"Excel file is missing required sheets: {missing}")
            return False

        # Load all sheets
        try:
            df_cable = pd.read_excel(xls, sheet_name='cable')
            df_cable.columns = df_cable.columns.str.strip()
            
            # Load cable boxes
            try:
                df_cable_box = pd.read_excel(xls, sheet_name='relay_box')
                df_cable_box.columns = df_cable_box.columns.str.strip()
            except:
                df_cable_box = pd.DataFrame()
            
            # Load other sheets
            df = pd.read_excel(xls, sheet_name='terminal')
            df.columns = df.columns.str.strip()
            if 'terminal_name' in df.columns:
                df.rename(columns={'terminal_name': 'terminal_no'}, inplace=True)
            
            df_junction = pd.read_excel(xls, sheet_name='junction_box')
            df_junction.columns = df_junction.columns.str.strip()
            df_header = pd.read_excel(xls, sheet_name='terminal_header')
            df_header.columns = df_header.columns.str.strip()
            df_group = pd.read_excel(xls, sheet_name='group')
            df_group.columns = df_group.columns.str.strip()
            df_choke = pd.read_excel(xls, sheet_name='choketable')
            df_choke.columns = df_choke.columns.str.strip()
            df_resistor = pd.read_excel(xls, sheet_name='resistortable')
            df_resistor.columns = df_resistor.columns.str.strip()
            
            # Load StationDrawing for footer
            try:
                df_title = pd.read_excel(xls, sheet_name='StationDrawing')
                df_title.columns = df_title.columns.str.strip()
            except:
                df_title = None
                
        except Exception as e:
            print(f"Error reading sheets: {e}")
            return False
        finally:
            try:
                xls.close()
            except:
                pass

        # === Process spare column ===
        if 'spare' in df.columns:
            try:
                df.loc[df['spare'].astype(str).str.upper() == 'Y', 'input_left'] = 'SP'
            except Exception as e:
                print(f"Error processing spare column: {e}")

        # === Prepare plotting data ===
        df_symbols = df.reset_index(drop=True)
        
        # === Get junction names ===
        junction_names = pd.unique(df_cable['junction_name'].astype(str).str.strip())
        
        # Add cable box junctions if they exist
        if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
            try:
                cable_box_junctions = pd.unique(df_cable_box['junction_name'].astype(str).str.strip())
                for junction in cable_box_junctions:
                    if junction not in junction_names:
                        junction_names = np.append(junction_names, junction)
            except Exception as e:
                print(f"Error processing cable box junctions: {e}")

        # === Generate PDF ===
        # Create output filename with timestamp
        excel_basename = os.path.splitext(excel_filename)[0]
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        timestamp_suffix = current_time.strftime("%Y%m%d_%H%M%S")
        
        # Generate checksum for unique filename
        checksum_data = f"{timestamp_suffix}|{excel_basename}"
        short_checksum = hashlib.md5(checksum_data.encode()).hexdigest()[:8]
        
        output_filename = f"{excel_basename}_{timestamp_suffix}_{short_checksum}.pdf"
        output_file = os.path.join(pdf_output_dir, output_filename)
        
        # Generate PDF
        success = generate_pdf(
            df_symbols, df_cable, df_cable_box, df_header, df_group, 
            df_choke, df_resistor, df_title, junction_names, 
            output_file, excel_filename, db_config
        )
        
        if success:
            # Move Excel file to processed directory
            processed_excel_path = os.path.join(processed_excel_dir, excel_filename)
            
            # If file already exists, add timestamp
            if os.path.exists(processed_excel_path):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(excel_filename)[0]
                extension = os.path.splitext(excel_filename)[1]
                processed_excel_path = os.path.join(processed_excel_dir, f"{base_name}_{timestamp}{extension}")
            
            shutil.move(excel_path, processed_excel_path)
            print(f"Moved Excel file to: {processed_excel_path}")
            
            # Archive copy for backup
            archive_path = os.path.join(archive_excel_dir, excel_filename)
            shutil.copy2(processed_excel_path, archive_path)
            
            return True
        else:
            return False
            
    except Exception as e:
        print(f"Error in process_excel_file: {e}")
        traceback.print_exc()
        return False

def generate_pdf(df_symbols, df_cable, df_cable_box, df_header, df_group, df_choke, 
                 df_resistor, df_title, junction_names, output_file, excel_filename, db_config):
    """Generate PDF from data"""
    try:
        # === Standardized dimensions ===
        SYMBOL_HEIGHT = 0.6
        SYMBOL_WIDTH = 0.35
        SYMBOL_RADIUS = 0.15
        CAPSULE_Y_CENTER_BASE = 3.6
        y_top_bus_offset = 1.3
        y_bottom_bus_offset = -1.3
        CABLE_GAP = 1.5
        vertical_gap = 6.5
        footer_height = 2.75
        
        # Calculate page dimensions
        global_max_width = 40.0
        fixed_fig_width = 42.8
        fixed_fig_height = 31.0
        max_rows_visible = 3
        
        # === Create pages ===
        pages = []
        for junction in junction_names:
            try:
                junction_mask = df_cable['junction_name'].astype(str).str.strip() == junction
                junction_cables_regular = df_cable[junction_mask].copy()
                
                junction_cables_box = pd.DataFrame()
                if not df_cable_box.empty and 'junction_name' in df_cable_box.columns:
                    junction_mask_box = df_cable_box['junction_name'].astype(str).str.strip() == junction
                    junction_cables_box = df_cable_box[junction_mask_box].copy()
                
                if junction_cables_regular.empty and junction_cables_box.empty:
                    continue
                
                # Organize rows
                letter_groups = organize_junction_rows(junction_cables_regular, junction_cables_box)
                junction_pages = create_pages_for_junction(junction, letter_groups, max_rows_per_page=3)
                pages.extend(junction_pages)
            except Exception as e:
                print(f"Error processing junction '{junction}': {e}")
                continue

        if not pages:
            print("No pages to generate")
            return False

        total_pages = len(pages)
        print(f"Total pages to generate: {total_pages}")

        # === Generate PDF ===
        with PdfPages(output_file) as pdf:
            for page_num, (junction_name, page_rows) in enumerate(pages, 1):
                try:
                    # Calculate page width
                    page_max_width = calculate_page_width(page_rows, df_symbols, df_cable_box)
                    
                    # Calculate centering
                    shift = (global_max_width - (page_max_width + 1.2)) / 2
                    page_start_x = 1 + shift
                    
                    # Set PDF metadata
                    pdf_info = pdf.infodict()
                    ist_tz = timezone('Asia/Kolkata')
                    current_time = datetime.now(ist_tz)
                    pdf_info['Title'] = f'Terminal Drawing - {junction_name}'
                    pdf_info['Author'] = 'SaltRiver Infosystems'
                    pdf_info['CreationDate'] = current_time.strftime("D:%Y%m%d%H%M%S+05'30'")
                    
                    # Create figure
                    fig, ax = plt.subplots(figsize=(fixed_fig_width, fixed_fig_height))
                    ax.set_facecolor('white')
                    ax.axis('off')
                    
                    # Draw symbols (simplified for this example)
                    # Note: You would need to include your actual drawing functions here
                    # For now, we'll create a simple placeholder
                    ax.text(0.5, 0.5, f"Junction: {junction_name}\nPage {page_num} of {total_pages}", 
                           ha='center', va='center', fontsize=24)
                    
                    # Draw footer if title data exists
                    if df_title is not None and not df_title.empty:
                        title_row = df_title.iloc[0]
                        draw_footer_simple(ax, page_start_x - 1.5, page_start_x - 1.5 + global_max_width, 
                                          -20, total_pages, page_num, title_row, junction_name)
                    
                    fig.subplots_adjust(left=0.04, right=0.99, top=0.98, bottom=0.02)
                    pdf.savefig(fig, dpi=300, facecolor='white')
                    plt.close(fig)
                    
                    print(f"Generated page {page_num} for {junction_name}")
                    
                except Exception as e:
                    print(f"Error generating page {page_num}: {e}")
                    continue
        
        print(f"PDF saved to: {output_file}")
        
        # Update PDF metadata with checksum
        if df_title is not None and not df_title.empty:
            update_pdf_checksum(output_file, df_title)
        
        # Store in database
        store_in_database(output_file, excel_filename, db_config, df_title)
        
        return True
        
    except Exception as e:
        print(f"Error generating PDF: {e}")
        traceback.print_exc()
        return False

def organize_junction_rows(junction_cables_regular, junction_cables_box):
    """Organize junction rows"""
    try:
        if not junction_cables_regular.empty and not junction_cables_box.empty:
            junction_cables = pd.concat([junction_cables_regular, junction_cables_box], ignore_index=True)
        elif not junction_cables_regular.empty:
            junction_cables = junction_cables_regular.copy()
        elif not junction_cables_box.empty:
            junction_cables = junction_cables_box.copy()
        else:
            return OrderedDict()
        
        # Sort by row and position
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
        print(f"Error organizing rows: {e}")
        return OrderedDict()

def create_pages_for_junction(junction, letter_groups, max_rows_per_page=3):
    """Create pages for junction"""
    try:
        sorted_letters = sorted(letter_groups.keys(), key=get_row_order, reverse=True)
        
        all_rows = []
        for letter in sorted_letters:
            cable_list = letter_groups[letter]
            # Break into rows if needed
            rows = break_cables_into_rows(cable_list)
            for row_cables in rows:
                all_rows.append((letter, row_cables))
        
        # Create pages
        pages = []
        for i in range(0, len(all_rows), max_rows_per_page):
            chunk = all_rows[i:i + max_rows_per_page]
            pages.append((junction, chunk))
        
        return pages
    except Exception as e:
        print(f"Error creating pages: {e}")
        return []

def break_cables_into_rows(cable_list, max_per_row=6):
    """Break cables into multiple rows"""
    rows = []
    current_row = []
    
    for cable_id in cable_list:
        if len(current_row) >= max_per_row:
            rows.append(current_row)
            current_row = []
        current_row.append(cable_id)
    
    if current_row:
        rows.append(current_row)
    
    return rows

def calculate_page_width(page_rows, df_symbols, df_cable_box):
    """Calculate page width"""
    try:
        page_max_width = 0
        for letter, cable_list in page_rows:
            row_width = 0
            for cable_id in cable_list:
                # Calculate width for each cable
                row_width += 2.0  # Simplified width calculation
            page_max_width = max(page_max_width, row_width)
        
        return min(page_max_width, 35.0)  # Cap at reasonable width
    except Exception as e:
        print(f"Error calculating page width: {e}")
        return 30.0

def get_row_order(letter):
    """Get row order for sorting"""
    try:
        if pd.isna(letter) or str(letter).strip() == '':
            return 0
        
        letter_str = str(letter).strip().upper()
        row_order = {
            'F': 6, 'E': 5, 'D': 4, 'C': 3, 'B': 2, 'A': 1,
            'H': 8, 'G': 7
        }
        
        if letter_str in row_order:
            return row_order[letter_str]
        
        try:
            return 100 - ord(letter_str)
        except:
            return 0
    except Exception as e:
        print(f"Error in get_row_order: {e}")
        return 0

def draw_footer_simple(ax, left, right, y_pos, total_pages, page_num, title_row, junction_name):
    """Draw simple footer"""
    try:
        footer_text = f"Page {page_num} of {total_pages} | Junction: {junction_name}"
        if title_row is not None:
            station_code = title_row.get('station_code', '')
            station_name = title_row.get('station_name', '')
            footer_text += f" | Station: {station_code} - {station_name}"
        
        ax.text((left + right) / 2, y_pos, footer_text, 
                ha='center', va='center', fontsize=12)
    except Exception as e:
        print(f"Error drawing footer: {e}")

def update_pdf_checksum(pdf_file_path, df_title):
    """Update PDF with checksum metadata"""
    try:
        # Generate checksum
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        timestamp = current_time.strftime("%Y-%m-%d %H:%M:%S")
        
        station_code = ""
        if df_title is not None and not df_title.empty:
            station_code = str(df_title.iloc[0].get('station_code', ''))
        
        file_size = os.path.getsize(pdf_file_path)
        file_name = os.path.basename(pdf_file_path)
        
        checksum_data = f"{station_code}|{file_size}|{timestamp}|{file_name}"
        checksum = hashlib.md5(checksum_data.encode()).hexdigest()
        
        print(f"Generated checksum: {checksum}")
        return checksum
    except Exception as e:
        print(f"Error updating checksum: {e}")
        return None

def store_in_database(pdf_file_path, excel_filename, db_config, df_title):
    """Store PDF metadata in database"""
    try:
        conn = psycopg2.connect(**db_config)
        cursor = conn.cursor()
        
        # Extract project_id from filename
        project_id = None
        patterns = [
            r'RAILWAYPROJECT_ID(\d+)',
            r'PROJECT_(\d+)',
            r'(?i)project[_-](\d+)',
            r'ID(\d+)',
            r'_(\d+)_'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, excel_filename)
            if match:
                try:
                    project_id = int(match.group(1))
                    break
                except ValueError:
                    continue
        
        if project_id is None:
            project_id = 0
        
        # Get file info
        file_size = os.path.getsize(pdf_file_path)
        pdf_filename = os.path.basename(pdf_file_path)
        
        # Generate checksums
        with open(pdf_file_path, 'rb') as f:
            full_file_md5 = hashlib.md5(f.read()).hexdigest()
        
        ist_tz = timezone('Asia/Kolkata')
        current_time = datetime.now(ist_tz)
        
        # Prepare data
        metadata_ts_ist = current_time.strftime("%Y-%m-%d %H:%M:%S.%f")
        
        # Insert into database
        insert_query = """
        INSERT INTO generated_pdf(
            project_id, pdf_filename, xlsx_filename, file_size,
            checksum_md5, metadata_checksum, metadata_data,
            initial_size_bytes, final_size_bytes, metadata_ts_ist,
            station_code, source_pdf_name, full_file_md5, remarks,
            checksum_algo, created_at, level1_status, level2_status,
            level3_status, version
        ) VALUES (
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
            %s, %s, %s, %s, %s, %s
        )
        """
        
        values = (
            project_id,
            pdf_filename,
            excel_filename,
            file_size,
            full_file_md5,
            full_file_md5,  # Using same for simplicity
            json.dumps({"generated_at": metadata_ts_ist}),
            file_size,
            file_size,
            metadata_ts_ist,
            df_title.iloc[0].get('station_code', '') if df_title is not None and not df_title.empty else '',
            pdf_filename,
            full_file_md5,
            'Generated by Python script',
            'md5',
            metadata_ts_ist,
            'pending',
            'pending',
            'pending',
            1
        )
        
        cursor.execute(insert_query, values)
        conn.commit()
        cursor.close()
        conn.close()
        
        print(f"PDF metadata stored in database with project_id: {project_id}")
        return True
        
    except Exception as e:
        print(f"Error storing in database: {e}")
        return False

# === ENTRY POINT ===
if __name__ == "__main__":
    print("\n" + "="*80)
    print("TERMINAL DIAGRAM GENERATOR SERVICE")
    print("="*80)
    print("Service is starting...")
    print("Press Ctrl+C to stop the service")
    print("\nStarting in 5 seconds...")
    time.sleep(5)
    
    main()