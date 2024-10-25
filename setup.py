import sys
from cx_Freeze import setup, Executable
# This is a setup script for cx_Freeze to create an executable. Use 'python setup.py build' to run it.
# Replace 'file.py' with the name of your script
# Replace 'ico.ico' with the path to your icon file

# Dependencies are automatically detected, but some modules need help.
build_exe_options = {
    "packages": ["os", "tkinter", "pandas", "docx", "reportlab"],  # Add any other packages you are using here
    "include_files": ["mortor.ico"],  # Make sure your icon is in the same directory
}

# Base must be "Win32GUI" to suppress the console window for a GUI application on Windows.
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="FeedbackCreator",
    version="1.1",
    description="A tool to generate feedback forms",
    options={
        "build_exe": build_exe_options
    },
    executables=[Executable("FeedbackCreator.py", base=base, icon="mortor.ico")]
)