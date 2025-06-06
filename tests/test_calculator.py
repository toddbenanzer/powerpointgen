import unittest
from unittest.mock import patch
from my_project.app.calculator import add
import logging

# Get the logger used by the calculator module
calculator_logger = logging.getLogger('app_logger')

class TestCalculator(unittest.TestCase):
    @patch.object(calculator_logger, 'info')
    def test_add(self, mock_info_log):
        """Test add function and check logging."""
        self.assertEqual(add(1, 2), 3)
        mock_info_log.assert_called_with("add function called with a: 1, b: 2")

        self.assertEqual(add(-1, 1), 0)
        mock_info_log.assert_called_with("add function called with a: -1, b: 1")

        self.assertEqual(add(-1, -1), -2)
        mock_info_log.assert_called_with("add function called with a: -1, b: -1")

        self.assertEqual(add(0, 0), 0)
        mock_info_log.assert_called_with("add function called with a: 0, b: 0")

        self.assertEqual(add(1.5, 2.5), 4.0)
        mock_info_log.assert_called_with("add function called with a: 1.5, b: 2.5")

if __name__ == "__main__":
    unittest.main()
