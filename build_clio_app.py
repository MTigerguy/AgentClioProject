from setuptools import setup

APP = ["clio_app.py"]  # Your main script file

# Py2app options
OPTIONS = {
    "packages": ["PyPDF2", "fitz"],  # Added PyMuPDF for fitz support
    "includes": ["tkinter", "PyPDF2", "fitz"],  # tkinter for GUI, fitz for PDF text extraction
    "iconfile": "clio.icns",  # Placeholder for your app icon file (update as needed)
    "plist": {
        "CFBundleName": "Agent Clio",
        "CFBundleDisplayName": "Agent Clio",
        "CFBundleIdentifier": "com.agentclio.app",
        "CFBundleVersion": "0.1.0",
        "CFBundleShortVersionString": "0.1.0",
        "LSMinimumSystemVersion": "10.14.0",  # Minimum macOS version
    },
}

setup(
    app=APP,
    name="Agent Clio",
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
    script_args=["py2app"]
)
