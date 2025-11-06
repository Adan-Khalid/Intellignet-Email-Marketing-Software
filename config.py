# -------------------------
# config.py
# -------------------------
# This is the central configuration file for the Bulk Email Sender.
# Edit the settings below to match your preferences.

# 1. File and Directory Settings
# These are the names of the data files the application will create and use.
SMTP_FILE = "smtp_list.json"
SUBJECTS_FILE = "subjects.json"
EMAIL_BODIES_FILE = "bodies.json"
FOLLOWUP_BODIES_FILE = "followup_bodies.json"
BLACKLIST_FILE = "blacklist.json"
NOTIFICATIONS_FILE = "notifications.json"

# These directories will be created to store logs and template files.
BODIES_DIR = "bodies"
LOG_DIR = "logs"

# 2. Admin & Notification Settings
# !!! IMPORTANT !!!
# Change this to the email address where you want to receive reply notifications.
ADMIN_EMAIL = "your-admin-email@example.com"

# 3. Application Settings
# Interval (in seconds) to check for new replies in the background.
# 900 seconds = 15 minutes
REPLY_CHECK_INTERVAL = 900

# Default UI settings
DEFAULT_APPEARANCE_MODE = "Dark" # "Dark" or "Light"
DEFAULT_COLOR_THEME = "blue"     # "blue", "green", "dark-blue"
