import argparse
import sys
# Ensure the app directory is in the Python path
# This might be needed if running main.py directly from my_project/
# and the my_project directory itself isn't in PYTHONPATH
sys.path.insert(0, '.')

from app.utils import greet
from app.calculator import add
# Initialize logger (it's configured in app/__init__.py)
# We just need to make sure app modules are imported so logger is set up.
import app

def main():
    parser = argparse.ArgumentParser(description="A simple CLI for my_project.")
    subparsers = parser.add_subparsers(dest="command", help="Available commands", required=True)

    # Greet command
    greet_parser = subparsers.add_parser("greet", help="Greets a user.", description="Greets a user with their name or a default name.")
    greet_parser.add_argument("--name", type=str, help="The name of the person to greet.")

    # Add command
    add_parser = subparsers.add_parser("add", help="Adds two numbers.", description="Adds two integer numbers provided as arguments.")
    add_parser.add_argument("a", type=int, help="The first number.")
    add_parser.add_argument("b", type=int, help="The second number.")

    args = parser.parse_args()

    if args.command == "greet":
        if args.name:
            result = greet(args.name)
        else:
            result = greet(None) # Or greet() which now handles None
        print(result)
    elif args.command == "add":
        result = add(args.a, args.b)
        print(result)

if __name__ == "__main__":
    main()
