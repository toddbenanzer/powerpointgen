import logging

# Calculator module with basic arithmetic operations
logger = logging.getLogger('app_logger')

def add(a, b):
  """
  Adds two numbers.

  Args:
    a: The first number.
    b: The second number.

  Returns:
    The sum of a and b.
  """
  logger.info(f"add function called with a: {a}, b: {b}")
  return a + b
