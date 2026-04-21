from .base import *

DEBUG = False
ALLOWED_HOSTS = os.getenv("ALLOWED_HOSTS", "").split(",")
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True