from setuptools import setup

APP = ["run_desktop.py"]
OPTIONS = {
    "argv_emulation": False,
    "packages": ["flask", "werkzeug", "jinja2", "markupsafe", "docx", "openpyxl", "webview"],
    "includes": ["socket", "tempfile", "re", "shutil", "time", "threading"],
    "resources": ["templates", "uploads", "app.py", "extract_employee_shifts.py"],
    "iconfile": "app_icon.icns",  # Add your own icon file if you have one
    "plist": {
        "CFBundleName": "Shift Manager",
        "CFBundleDisplayName": "Shift Manager",
        "CFBundleIdentifier": "com.yourcompany.shiftmanager",
        "CFBundleVersion": "1.0.0",
        "CFBundleShortVersionString": "1.0.0",
        "LSBackgroundOnly": False,
        "NSHumanReadableCopyright": "Copyright © 2025 Your Company"
    }
}

setup(
    app=APP,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
