import os

# --- Configuration Paths ---
ENV_FILE_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env")
LOADER_PATH = r"C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin"
ADD_IN_PATH = r"C:\Program Files\Ninety One\FOXL\v8\_91ExcelAddin"

# Local Paths
BASE_LOCAL_PATH = r"C:\ExcelAddIn"
LOCAL_TEST_PATH = r"C:\ExcelAddIn\_91ExcelAddIn"
REG_PATH = r"Software\Microsoft\Office\16.0\excel\options"

# The Standard FOXL Loader
FOXL_LOADER_PATH = r"C:\Program Files\Ninety One\FOXL\v8\NinetyOne.ExcelAddIn.Loader-AddIn64.xll"

# Network Paths
NETWORK_PATH_8 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn"
NETWORK_PATH_6 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\InvestmentTechExcelAddIn"

CONFIG_FILE_NAME = "NinetyOne.ExcelAddIn.config.json"
