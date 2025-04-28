#!/usr/bin/env python3

import os
import re

EXCLUDE_DIRS = {"env", "__pycache__"}


def gather_python_files(root="."):
    """
    Recursively gather all .py files from the root directory,
    skipping directories in EXCLUDE_DIRS.
    """
    py_files = []
    for dirpath, dirnames, filenames in os.walk(root):
        # Filter out excluded directories
        dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]
        for fname in filenames:
            if fname.endswith(".py"):
                full_path = os.path.join(dirpath, fname)
                py_files.append(full_path)
    return py_files

def count_lines_of_code(filepath):
    """
    Return total number of lines in the given file.
    """
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        return sum(1 for _ in f)

def find_classes(filepath):
    """
    Use a simple regex to find class definitions in a .py file
    and return them as a list of class names.
    """
    classes = []
    class_pattern = re.compile(r"^\s*class\s+(\w+)\s*[\(:]", re.MULTILINE)
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            classes = class_pattern.findall(content)
    except Exception as e:
        pass
    return classes

def main():
    print("Collecting Python file info...\n")
    py_files = gather_python_files(".")

    total_files = len(py_files)
    grand_total_lines = 0
    class_dict = {}  # filepath -> [class names]

    for pf in py_files:
        loc = count_lines_of_code(pf)
        classes_found = find_classes(pf)
        grand_total_lines += loc
        class_dict[pf] = classes_found

    print(f"Total .py files found: {total_files}")
    print(f"Total lines of Python code: {grand_total_lines}")
    print("\nClasses found (file -> classes):\n")
    for pf, classes in class_dict.items():
        if classes:
            print(f"  {pf}: {classes}")
    print("\nDone. You can copy/paste this info back to me!")

if __name__ == "__main__":
    main()