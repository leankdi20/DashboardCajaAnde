from .base import *

DEBUG = True
ALLOWED_HOSTS = os.getenv("ALLOWED_HOSTS", "localhost").split(",")
SESSION_COOKIE_SECURE = False
CSRF_COOKIE_SECURE = False

# Fuerza que el token CSRF se regenere en cada sesión nueva
CSRF_COOKIE_SAMESITE = 'Lax'
CSRF_COOKIE_HTTPONLY = False
SESSION_COOKIE_SAMESITE = 'Lax'


# Durante desarrollo, acepta el origen local
CSRF_TRUSTED_ORIGINS = [
    "http://127.0.0.1:8000",
    "http://localhost:8000",
]
# Logs más detallados en desarrollo
LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,
    "handlers": {
        "console": {"class": "logging.StreamHandler"},
    },
    "root": {
        "handlers": ["console"],
        "level": "DEBUG",
    },
}