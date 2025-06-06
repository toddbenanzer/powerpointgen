import configparser
import os
import logging

# Utility functions for the project
logger = logging.getLogger('app_logger')

def greet(name: str = None) -> str:
  """
  Greets the person with the given name.
  If no name is provided, it uses a default name from config.ini.

  Raises:
    ValueError: If the name is provided but is not a string or is an empty string.

  Args:
    name: The name of the person to greet. Defaults to None.

  Returns:
    A greeting string.
  """
  logger.info(f"greet function called with name: {name}")
  if name is None:
    logger.info("Name is None, attempting to load default name from config.")
    config = configparser.ConfigParser()
    # Construct the path to config.ini relative to this file's directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_path = os.path.join(base_dir, 'config.ini')
    config.read(config_path)
    name = config.get('greeting', 'default_name', fallback='User') # Default to 'User' if not found
    logger.info(f"Default name loaded: {name}")
  elif not isinstance(name, str):
    logger.error(f"ValueError: Name must be a string. Received type: {type(name)}")
    raise ValueError("Name must be a string.")
  elif not name:
    logger.error("ValueError: Name cannot be empty.")
    raise ValueError("Name cannot be empty.")
  return f"Hello, {name}!"
