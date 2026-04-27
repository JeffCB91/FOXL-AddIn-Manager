import os

# --- Configuration Paths ---
ENV_FILE_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env")
LOADER_PATH = r"C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin"
ADD_IN_PATH = r"C:\Program Files\Ninety One\FOXL\v8\_91ExcelAddin"

# Local Paths
BASE_LOCAL_PATH = r"C:\ExcelAddIn"
LOCAL_TEST_PATH = r"C:\ExcelAddIn\_91ExcelAddIn"
REG_PATH = r"Software\Microsoft\Office\16.0\excel\options"
# --- Logging Paths ---
LOG_DIR_PATH = r"C:\ProgramData\NinetyOne.ExcelAddIn\Logs"
# LOG_DIR_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\Logs")

# The Standard FOXL Loader
FOXL_LOADER_PATH = r"C:\Program Files\Ninety One\FOXL\v8\NinetyOne.ExcelAddIn.Loader-AddIn64.xll"

# Network Paths
NETWORK_PATH_8 = r"C:\Users\jcrewebrown\OneDrive - Ninety One\Documents\FOXL\admin\InvestmentTechExcelAddIn"
NETWORK_PATH_8_BAK = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn"
NETWORK_PATH_6 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\InvestmentTechExcelAddIn"
TEMPLATES_PATH = r"\\uranus\FMC\Data\Internal Data\FOXL Templates"

CONFIG_FILE_NAME = "NinetyOne.ExcelAddIn.config.json"

# Azure DevOps — FOXL build pipeline
ADO_ORG = "Ninety-One"
ADO_PROJECT = "Ninety-One"
ADO_PIPELINE_ID = 556
ADO_ARTIFACT_NAME = "Binaries"
ADO_ARTIFACT_FILE = "NetworkShareFiles.zip"
