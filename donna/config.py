import os
from dotenv import load_dotenv

load_dotenv()

# Base directory of the project (DONNA_AGENT)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# -------------------------------
# Directory paths
# -------------------------------
TEMPLATES_DIR = os.getenv("TEMPLATES_DIR", os.path.join(BASE_DIR, "templates"))
RESULTS_DIR = os.getenv("RESULTS_DIR", os.path.join(BASE_DIR, "results"))
import os

print("BASE_DIR:", os.path.dirname(os.path.abspath(__file__)))
print("TEMPLATES_DIR:", TEMPLATES_DIR)
print("RESULTS_DIR:", RESULTS_DIR)
# -------------------------------
# File validation settings
# -------------------------------
MAX_FILENAME_LENGTH = 200
MAX_VARIABLE_VALUE_LENGTH = 10000
MAX_VARIABLES_COUNT = 100
ALLOWED_EXTENSIONS = {'.docx'}
FORBIDDEN_CHARS = ['<', '>', ':', '"', '/', '\\', '|', '?']

# -------------------------------
# Other constants
# -------------------------------
EMAIL_PATTERN = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
PHONE_MIN_LENGTH = 7
PHONE_MAX_LENGTH = 15
ACCEPTED_DATE_FORMATS = [
    '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%m-%d-%Y',
    '%d/%m/%y', '%m/%d/%y', '%Y-%m-%d'
]
