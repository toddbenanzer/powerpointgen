import unittest
import subprocess
import sys
import os

# Determine the path to main.py
# Assuming this test file is in 'tests/' and main.py is in the parent directory 'my_project/'
# The 'my_project' directory is what we cd into in bash typically.
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # This should be /app
MAIN_PY_PATH = os.path.join(BASE_DIR, "my_project", "main.py")
PYTHON_EXECUTABLE = sys.executable # Use the same python that's running the tests

class TestCLI(unittest.TestCase):

    def run_cli(self, command_args):
        """Helper function to run main.py with specified arguments."""
        cmd = [PYTHON_EXECUTABLE, MAIN_PY_PATH] + command_args
        # Running from the /app directory as the working directory
        return subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)

    def test_greet_command_with_name(self):
        """Test the 'greet --name <name>' command."""
        result = self.run_cli(["greet", "--name", "Alice"])
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout.strip(), "Hello, Alice!")
        self.assertEqual(result.stderr.strip(), "")

    def test_greet_command_no_name(self):
        """Test the 'greet' command without a name (should use default from config)."""
        # This relies on my_project/config.ini having default_name = User
        result = self.run_cli(["greet"])
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout.strip(), "Hello, User!")
        self.assertEqual(result.stderr.strip(), "")

    def test_greet_command_help(self):
        """Test the 'greet --help' command."""
        result = self.run_cli(["greet", "--help"])
        self.assertEqual(result.returncode, 0)
        self.assertIn("usage: main.py greet [-h] [--name NAME]", result.stdout)
        self.assertIn("Greets a user with their name or a default name.", result.stdout) # Changed this line
        self.assertEqual(result.stderr.strip(), "")

    def test_main_help(self):
        """Test the main help message (-h or --help)."""
        for help_arg in ["-h", "--help"]:
            with self.subTest(help_arg=help_arg):
                result = self.run_cli([help_arg])
                self.assertEqual(result.returncode, 0)
                self.assertIn("usage: main.py [-h] {greet,add} ...", result.stdout)
                self.assertIn("Available commands", result.stdout)
                self.assertEqual(result.stderr.strip(), "")

    def test_add_command_success(self):
        """Test the 'add <a> <b>' command with valid integers."""
        result = self.run_cli(["add", "10", "5"])
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout.strip(), "15")
        self.assertEqual(result.stderr.strip(), "")

        result = self.run_cli(["add", "-3", "5"])
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout.strip(), "2")
        self.assertEqual(result.stderr.strip(), "")

    def test_add_command_invalid_input(self):
        """Test the 'add' command with non-integer input."""
        result = self.run_cli(["add", "10", "abc"])
        self.assertNotEqual(result.returncode, 0) # Should fail
        self.assertIn("usage: main.py add [-h] a b", result.stderr) # argparse error
        self.assertIn("main.py add: error: argument b: invalid int value: 'abc'", result.stderr)
        self.assertEqual(result.stdout.strip(), "")

    def test_add_command_missing_arguments(self):
        """Test the 'add' command with missing arguments."""
        result = self.run_cli(["add", "10"])
        self.assertNotEqual(result.returncode, 0) # Should fail
        self.assertIn("usage: main.py add [-h] a b", result.stderr) # argparse error
        self.assertIn("main.py add: error: the following arguments are required: b", result.stderr)
        self.assertEqual(result.stdout.strip(), "")

    def test_add_command_help(self):
        """Test the 'add --help' command."""
        result = self.run_cli(["add", "--help"])
        self.assertEqual(result.returncode, 0)
        self.assertIn("usage: main.py add [-h] a b", result.stdout)
        self.assertIn("Adds two integer numbers provided as arguments.", result.stdout) # Changed this line
        self.assertEqual(result.stderr.strip(), "")

if __name__ == "__main__":
    unittest.main()
