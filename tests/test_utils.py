import unittest
from unittest.mock import patch, MagicMock
from my_project.app.utils import greet
import logging

# Get the logger used by the utils module
utils_logger = logging.getLogger('app_logger')

class TestUtils(unittest.TestCase):

    @patch.object(utils_logger, 'info')
    def test_greet_success(self, mock_info_log):
        """Test greet function with valid input and check logging."""
        self.assertEqual(greet("World"), "Hello, World!")
        mock_info_log.assert_any_call("greet function called with name: World")

    @patch.object(utils_logger, 'error')
    @patch.object(utils_logger, 'info')
    def test_greet_invalid_type(self, mock_info_log, mock_error_log):
        """Test greet function with non-string input and check error logging."""
        with self.assertRaises(ValueError) as context:
            greet(123)
        self.assertEqual(str(context.exception), "Name must be a string.")
        mock_info_log.assert_any_call("greet function called with name: 123")
        mock_error_log.assert_called_once_with("ValueError: Name must be a string. Received type: <class 'int'>")

    @patch.object(utils_logger, 'error')
    @patch.object(utils_logger, 'info')
    def test_greet_empty_string(self, mock_info_log, mock_error_log):
        """Test greet function with an empty string and check error logging."""
        with self.assertRaises(ValueError) as context:
            greet("")
        self.assertEqual(str(context.exception), "Name cannot be empty.")
        mock_info_log.assert_any_call("greet function called with name: ")
        mock_error_log.assert_called_once_with("ValueError: Name cannot be empty.")

    @patch.object(utils_logger, 'info')
    def test_greet_default_name(self, mock_info_log):
        """Test greet function with no name provided, using default from config and check logging."""
        self.assertEqual(greet(None), "Hello, User!") # Assumes default_name = User in config.ini
        mock_info_log.assert_any_call("greet function called with name: None")
        mock_info_log.assert_any_call("Name is None, attempting to load default name from config.")
        mock_info_log.assert_any_call("Default name loaded: User")

        # Reset mock for the second call if needed, or use separate tests
        mock_info_log.reset_mock()
        self.assertEqual(greet(), "Hello, User!") # Also test with no argument
        mock_info_log.assert_any_call("greet function called with name: None")
        mock_info_log.assert_any_call("Name is None, attempting to load default name from config.")
        mock_info_log.assert_any_call("Default name loaded: User")


if __name__ == "__main__":
    # If you want to see log output from the logger itself during tests (not just mocks)
    # you might need to configure the logger for the test environment here.
    # However, for checking calls, mocks are usually sufficient.
    # logging.basicConfig(level=logging.INFO, handlers=[logging.StreamHandler()])
    unittest.main()
