"""
Constants and configuration for EDM Library Wizard
"""

# Feature availability flags (set at runtime based on imports)
ANTHROPIC_AVAILABLE = False
FUZZYWUZZY_AVAILABLE = False

# PAS API Configuration
PAS_API_URL = "https://api.pas.partquest.com"
PAS_AUTH_URL = "https://samauth.us-east-1.sws.siemens.com/token"
PAS_SEARCH_PROVIDER_ID = 44
PAS_SEARCH_PROVIDER_VERSION = 2
PAS_SUPPLY_CHAIN_ENRICHER_ID = 33
PAS_SUPPLY_CHAIN_ENRICHER_VERSION = 1

# PAS API Property IDs
PAS_PROPERTY_MANUFACTURER_NAME = "6230417e"
PAS_PROPERTY_MANUFACTURER_PN = "d8ac8dcc"
PAS_PROPERTY_DATASHEET_URL = "750a45c8"
PAS_PROPERTY_FINDCHIPS_URL = "2a2b1476"
PAS_PROPERTY_LIFECYCLE_STATUS = "e5434e21"
PAS_PROPERTY_LIFECYCLE_STATUS_CODE = "a189d244"
PAS_PROPERTY_PART_ID = "e1aa6f26"

# Claude AI Models
CLAUDE_MODELS = {
    "Sonnet 4.5 (Recommended)": "claude-sonnet-4-5-20250929",
    "Haiku 4.5 (Fastest)": "claude-haiku-4-5-20250929",
    "Opus 4.1 (Most Capable)": "claude-opus-4-1-20250514"
}

# Default values
DEFAULT_PROJECT_NAME = "VarTrainingLab"
DEFAULT_CATALOG = "VV"
DEFAULT_MAX_MATCHES = 10
DEFAULT_AI_MAX_RETRIES = 5
DEFAULT_PREVIEW_ROWS = 10

# Excel Configuration
EXCEL_MAX_SHEET_NAME_LENGTH = 31
EXCEL_INVALID_SHEET_CHARS = ['\\', '/', '*', '?', ':', '[', ']']

# XML Configuration
XML_CLASS_MFG = "090"
XML_CLASS_MFGPN = "060"
XML_SPECIAL_CHARS = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&apos;'
}

# Match Types
MATCH_TYPE_FOUND = "Found"
MATCH_TYPE_MULTIPLE = "Multiple"
MATCH_TYPE_NEED_REVIEW = "Need user review"
MATCH_TYPE_NONE = "None"
MATCH_TYPE_ERROR = "Error"

# QSettings keys
SETTINGS_ORG = "VarIndustries"
SETTINGS_APP = "EDMWizard"
SETTINGS_CLAUDE_API_KEY = "claude_api_key"
SETTINGS_CLAUDE_MODEL = "claude_model"
SETTINGS_PAS_CLIENT_ID = "pas_client_id"
SETTINGS_PAS_CLIENT_SECRET = "pas_client_secret"
SETTINGS_PAS_ACCESS_TOKEN = "pas_access_token"
SETTINGS_PAS_TOKEN_EXPIRY = "pas_token_expiry"
SETTINGS_MGLAUNCH_PATH = "mglaunch_path"
SETTINGS_OUTPUT_FOLDER = "output_folder"

# File extensions
SUPPORTED_ACCESS_EXTENSIONS = ['.mdb', '.accdb']
SUPPORTED_SQLITE_EXTENSIONS = ['.db', '.sqlite', '.sqlite3']
SUPPORTED_EXCEL_EXTENSIONS = ['.xlsx', '.xls']

# Common MG Launch paths to search
MGLAUNCH_SEARCH_PATHS = [
    r"C:\MentorGraphics",
    r"C:\SiemensEDA"
]
