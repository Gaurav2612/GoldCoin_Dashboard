
from pathlib import Path

# ============================================================
# BASE DIRECTORY
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent


# ============================================================
# EXCEL AUTOMATIC IMPORT DIRECTORY
# ============================================================

EXCEL_IMPORT_DIR = BASE_DIR / "data" / "excel_import"


# ============================================================
# SECURITY SETTINGS
# ============================================================

SECRET_KEY = "change-this-secret-key-before-production"

DEBUG = True

ALLOWED_HOSTS = [
    "127.0.0.1",
    "localhost",
]


# ============================================================
# APPLICATIONS
# ============================================================

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",

    "calls",
]


# ============================================================
# MIDDLEWARE
# ============================================================

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]


# ============================================================
# URL CONFIGURATION
# ============================================================

ROOT_URLCONF = "goldcoin_dashboard.urls"


# ============================================================
# TEMPLATES
# ============================================================

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",

        "DIRS": [
            BASE_DIR / "templates",
        ],

        "APP_DIRS": True,

        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]


# ============================================================
# WSGI
# ============================================================

WSGI_APPLICATION = "goldcoin_dashboard.wsgi.application"


# ============================================================
# DATABASE
# ============================================================

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
        "OPTIONS": {
            "timeout": 30,
        },
    }
}


# ============================================================
# PASSWORD VALIDATION
# ============================================================

AUTH_PASSWORD_VALIDATORS = []


# ============================================================
# LANGUAGE AND TIMEZONE
# ============================================================

LANGUAGE_CODE = "en-us"

TIME_ZONE = "Asia/Kolkata"

USE_I18N = True

USE_TZ = True


# ============================================================
# STATIC FILES
# ============================================================

STATIC_URL = "/static/"

STATICFILES_DIRS = [
    BASE_DIR / "static",
]


# ============================================================
# DEFAULT PRIMARY KEY
# ============================================================

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"


# ============================================================
# LOGIN / LOGOUT
# ============================================================

LOGIN_URL = "login"

LOGIN_REDIRECT_URL = "dashboard"

LOGOUT_REDIRECT_URL = "login"


# ============================================================
# UPLOAD SETTINGS
# ============================================================

DATA_UPLOAD_MAX_NUMBER_FIELDS = 20000


# ============================================================
# SHARED FILE-BASED CACHE
# ============================================================

CACHES = {
    "default": {
        "BACKEND": "django.core.cache.backends.filebased.FileBasedCache",

        "LOCATION": str(BASE_DIR / ".django_cache"),

        "TIMEOUT": None,

        "OPTIONS": {
            "MAX_ENTRIES": 1000,
        },
    },
}
