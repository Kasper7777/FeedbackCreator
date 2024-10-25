from distutils.core import setup
import py2exe
import os
# eg    'pyinstaller --onefile --windowed --icon=mortor.ico FeedbackCreator.py'
# Explicitly define the script to be included
script = os.path.join(os.path.dirname(__file__), 'FeedbackCreator.py')

setup(
    windows=[{
        "script": script,
        "icon_resources": [(0, "mortor.ico")]
    }],
    options={
        "py2exe": {
            "packages": ["os", "tkinter", "pandas", "docx", "reportlab"],  # Include necessary packages
            "bundle_files": 1,  # Bundle everything into a single EXE
            "compressed": True,  # Compress the library archive
            "excludes": ["gui_version", "main", "main001", "main002", "gui_version01"],  # Exclude unnecessary modules
        }
    },
    zipfile=None,  # Do not create a separate library zip file
    py_modules=['FeedbackCreator'],  # Explicit module
)