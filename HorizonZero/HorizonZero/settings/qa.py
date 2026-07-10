from .base import *

DEBUG = env_bool("DEBUG", False)
ALLOWED_HOSTS = os.getenv("ALLOWED_HOSTS", "").split(",")
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True
