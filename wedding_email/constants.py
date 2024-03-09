import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# Paths
PROJECT_PATH = os.getenv("PROJECT_PATH")

# Data
DATA_PATH = Path(PROJECT_PATH) / "data"
RAW_DATA_PATH = DATA_PATH / "raw"
INTERIM_DATA_PATH = DATA_PATH / "interim"
PROCESSED_DATA_PATH = DATA_PATH / "processed"

# Email
MAIL_APP_PASSWORD = os.environ.get("MAIL_APP_PASSWORD")
