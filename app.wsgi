import sys
import logging

# Enable logging for troubleshooting
logging.basicConfig(stream=sys.stderr)

# Add your project path
sys.path.insert(0, "/var/www/html/git")

# Import the Flask app object from run.py
from run import app as application
