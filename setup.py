import sys
from cx_Freeze import setup, Executable
includefiles = ["static", "templates", "pdf", "Excel", "wkhtmltox"]
includes = []
packages = ['flask', 'xlrd', 'pdfkit', 'os', 'shutil', 'wkhtmltopdf', 'jinja2', 'pandas', 'locale', 'numpy', 'logging']
base = None
if sys.platform == "win32":
    base = "Win32GUI"
if sys.platform == 'win64':
    base = "Win64GUI"

setup(
    name="excel_to_pdf",
    version="3.1",
    description="A Excel_to_pdf convertor help tool.",
    options={"build_exe": {"packages": packages, "include_files": includefiles}},
    executables=[Executable("excel_to_pdf.py", base=base)]
        )
