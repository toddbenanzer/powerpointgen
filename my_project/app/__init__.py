# This file makes 'app' a Python package

import logging
import os

# --- Logger Setup ---
# Create logger
app_logger = logging.getLogger('app_logger')
app_logger.setLevel(logging.INFO)

# Create file handler
# Log file will be in the 'my_project' directory, which is the parent of the 'app' directory
log_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'app.log')
file_handler = logging.FileHandler(log_file_path)
file_handler.setLevel(logging.INFO)

# Create formatter and add it to the handler
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Add the handler to the logger
# Prevent adding multiple handlers if this init is loaded multiple times in some testing scenarios
if not app_logger.handlers:
    app_logger.addHandler(file_handler)
