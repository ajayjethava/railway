#!/usr/bin/env python3
""" Railway Project Terminal Diagram Generator - 24x7 Directory Monitor
Monitors input directory for Excel files and processes them

NOTE: PDF generation call has been commented out as requested. This script
now only checks for .xlsx files every 10 seconds and logs their presence.
"""

import os
import time
import shutil
import signal
import traceback
import logging
import logging.handlers
from datetime import datetime

# === CONFIGURATION ===
XLSX_INPUT_DIR = "/root/srv/local/git/xlsx_download"
PROCESSED_EXCEL_DIR = "/root/srv/local/git/processed_excel"
ERROR_EXCEL_DIR = "/root/srv/local/git/error_excel"
LOG_DIR = "/root/srv/local/git/logs"

# Create directories if they don't exist
os.makedirs(XLSX_INPUT_DIR, exist_ok=True)
os.makedirs(PROCESSED_EXCEL_DIR, exist_ok=True)
os.makedirs(ERROR_EXCEL_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# === SETUP LOGGING ===
def setup_logging():
    """Setup logging configuration"""
    try:
        # Try to import from logging_config module
        try:
            from logging_config import setup_logging as setup_logging_external
            logger = setup_logging_external(
                log_name="directory_monitor",
                log_dir=LOG_DIR,
                max_log_size=10*1024*1024,  # 10 MB
                backup_count=10
            )
            return logger
        except ImportError as e:
            print(f"Warning: Could not import from logging_config: {e}")
            print("Falling back to built-in logging setup...")

        # Fallback: built-in logging setup
        log_file = os.path.join(LOG_DIR, "directory_monitor.log")

        # Create logger
        logger = logging.getLogger("directory_monitor")
        logger.setLevel(logging.DEBUG)

        # Clear existing handlers
        logger.handlers.clear()

        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # File handler with rotation
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=10*1024*1024,  # 10 MB
            backupCount=10
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        return logger

    except Exception as e:
        print(f"Error setting up logging: {e}")
        # Last resort: basic logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        return logging.getLogger("directory_monitor")

# Initialize logger
logger = setup_logging()

# === GLOBAL FLAG FOR GRACEFUL SHUTDOWN ===
running = True

def signal_handler(sig, frame):
    """Handle shutdown signals gracefully"""
    global running
    logger.info("="*60)
    logger.info("SHUTDOWN SIGNAL RECEIVED")
    logger.info("Stopping directory monitor...")
    logger.info("="*60)
    running = False

# Register signal handlers
signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

def move_excel_file(excel_file_path, success=True):
    """
    Move Excel file to appropriate directory after processing
    (kept for future use; not used in the current "watch-only" mode)
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
            logger.info(f"Moved Excel file to: {destination}")
        else:
            destination = os.path.join(ERROR_EXCEL_DIR, excel_filename)
            if os.path.exists(destination):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(excel_filename)
                new_filename = f"{name}_{timestamp}{ext}"
                destination = os.path.join(ERROR_EXCEL_DIR, new_filename)

            shutil.move(excel_file_path, destination)
            logger.warning(f"Moved Excel file to error directory: {destination}")

    except Exception as e:
        logger.error(f"Error moving Excel file: {e}")
        logger.error(traceback.format_exc())

def monitor_directory():
    """
    Monitor the input directory for new Excel files and log their presence.
    NOTE: The PDF generation / processing logic has been commented out.
    """
    logger.info("="*60)
    logger.info("STARTING 24x7 DIRECTORY MONITOR (watch-only mode)")
    logger.info(f"Input Directory: {XLSX_INPUT_DIR}")
    logger.info(f"Checked every 10 seconds. PDF generation is disabled.")
    logger.info("="*60)

    # If you prefer to log only when changes happen, you can keep track of previous state.
    # For now we will log current contents each interval to satisfy "see just xlsx come or not".
    previous_listing = None

    while running:
        try:
            xlsx_files = [f for f in os.listdir(XLSX_INPUT_DIR) if f.lower().endswith('.xlsx')]

            if xlsx_files:
                # Build a readable list with last-modified times
                files_info = []
                for f in sorted(xlsx_files):
                    path = os.path.join(XLSX_INPUT_DIR, f)
                    try:
                        mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        mtime = "unknown"
                    files_info.append(f"{f} (mtime: {mtime})")

                listing_text = "; ".join(files_info)
                # Log info every cycle. If you want to log only on change, compare to previous_listing.
                if listing_text != previous_listing:
                    logger.info(f"XLSX files present: {listing_text}")
                    previous_listing = listing_text
                else:
                    logger.debug(f"XLSX files unchanged: {listing_text}")
            else:
                logger.info("No XLSX files found.")
                previous_listing = None

            # -----------------------------------------------------------------
            # PDF generation and processing section HAS BEEN COMMENTED OUT
            # -----------------------------------------------------------------
            # The following code used to import and call the pdf_generator:
            #
            #     from pdf_generator import generate_pdf_from_excel
            #     success, message = generate_pdf_from_excel(excel_path)
            #     ...
            #
            # Per your request this is disabled so the monitor only logs arrivals.
            # -----------------------------------------------------------------

            # Wait 10 seconds between checks (watch-only mode)
            for i in range(10):
                if not running:
                    break
                time.sleep(1)

        except KeyboardInterrupt:
            logger.info("Keyboard interrupt received. Shutting down gracefully...")
            break
        except Exception as e:
            logger.error(f"Error in monitor loop: {e}")
            logger.error(traceback.format_exc())
            # Wait a bit before retrying to avoid spinning on persistent errors
            for i in range(10):
                if not running:
                    break
                time.sleep(1)

    logger.info("="*60)
    logger.info("DIRECTORY MONITOR STOPPED")
    logger.info("="*60)

# === MAIN EXECUTION ===
if __name__ == "__main__":
    try:
        monitor_directory()
    except Exception as e:
        logger.critical(f"FATAL ERROR in directory monitor: {e}")
        logger.critical(traceback.format_exc())
        import sys
        sys.exit(1)
