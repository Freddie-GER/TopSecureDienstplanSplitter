"""
This is a setup.py script for the DienstplanSplitter app.
"""

from setuptools import setup
import os

APP = ['dienstplan_splitter_gui.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['tkinter', 'PIL', 'pandas', 'openpyxl'],
    'includes': ['tkinter', 'PIL', 'pandas', 'openpyxl', 'numpy'],
    'excludes': ['matplotlib', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6'],
    'frameworks': [],
    'iconfile': 'icon.icns' if os.path.exists('icon.icns') else None,
    'plist': {
        'CFBundleName': 'Dienstplan Splitter',
        'CFBundleDisplayName': 'Dienstplan Splitter',
        'CFBundleIdentifier': 'com.dienstplan.splitter',
        'CFBundleVersion': '1.0.0',
        'LSMinimumSystemVersion': '10.13.0',
        'NSHighResolutionCapable': True,
    }
}

setup(
    name="Dienstplan Splitter",
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
) 