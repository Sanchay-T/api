import os
import sys
import tempfile


def get_saved_pdf_dir():
    TEMP_SAVED_PDF_DIR = os.path.join(tempfile.gettempdir(), "saved_pdf")

    return TEMP_SAVED_PDF_DIR


def get_saved_excel_dir():
    """Return a directory called 'output' located at the project root (same level as backend/)."""

    # Determine project root as the parent of the directory that contains this utils.py file
    # utils.py lives in backend/, so root is one level up.
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

    OUTPUT_DIR = os.path.join(project_root, "output")

    return OUTPUT_DIR


def get_base_dir():
    """
    Determine the base directory of the application.
    - Use sys.executable if running as an executable
    - Use __file__ if running as a script

    """
    if hasattr(sys, "_MEIPASS"):
        print("MEIPASS : ", sys._MEIPASS)
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))
