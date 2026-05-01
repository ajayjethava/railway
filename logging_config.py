#!/usr/bin/env python3
"""
Logging configuration for Railway Project Terminal Diagram Generator
"""

import logging
import logging.handlers
import os
from datetime import datetime

def setup_logging(
    log_name="railway_project",
    log_dir="/root/srv/local/git/logs",
    max_log_size=10*1024*1024,  # 10 MB
    backup_count=10
):
    """
    Set up comprehensive logging configuration
    
    Args:
        log_name: Base name for log files
        log_dir: Directory to store log files
        max_log_size: Maximum size of each log file in bytes
        backup_count: Number of backup log files to keep
    """
    
    # Create log directory if it doesn't exist
    os.makedirs(log_dir, exist_ok=True)
    
    # Create logger
    logger = logging.getLogger(log_name)
    logger.setLevel(logging.DEBUG)
    
    # Clear any existing handlers
    logger.handlers.clear()
    
    # Log format
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # File handler with rotation
    log_file = os.path.join(log_dir, f"{log_name}.log")
    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=max_log_size,
        backupCount=backup_count
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

def get_module_logger(module_name):
    """
    Get a logger for a specific module
    """
    return logging.getLogger(f"railway_project.{module_name}")

# Example usage:
# logger = get_module_logger(__name__)
# logger.info("Message")