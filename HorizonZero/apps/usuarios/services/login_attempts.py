import math
import time

from django.conf import settings
from django.core.cache import cache

MAX_LOGIN_ATTEMPTS = getattr(settings, "LOGIN_MAX_ATTEMPTS", 5)
BLOCK_MINUTES = getattr(settings, "LOGIN_BLOCK_MINUTES", 5)


def _get_client_ip(request):
    forwarded_for = request.META.get("HTTP_X_FORWARDED_FOR")
    if forwarded_for:
        return forwarded_for.split(",")[0].strip()
    return request.META.get("REMOTE_ADDR", "")


def _fail_key(username, ip):
    return f"login_fail:{username}:{ip}"


def _block_key(username, ip):
    return f"login_block:{username}:{ip}"


def is_login_blocked(username, request):
    ip = _get_client_ip(request)
    blocked_until = cache.get(_block_key(username, ip))
    return bool(blocked_until and blocked_until > time.time())


def get_login_block_remaining_seconds(username, request):
    ip = _get_client_ip(request)
    blocked_until = cache.get(_block_key(username, ip))

    if not blocked_until:
        return 0

    remaining = int(blocked_until - time.time())
    return max(0, remaining)


def get_login_block_remaining_minutes(username, request):
    remaining_seconds = get_login_block_remaining_seconds(username, request)
    if remaining_seconds <= 0:
        return 0
    return math.ceil(remaining_seconds / 60)


def register_failed_attempt(username, request):
    ip = _get_client_ip(request)
    fail_key = _fail_key(username, ip)
    block_key = _block_key(username, ip)

    attempts = cache.get(fail_key, 0) + 1
    cache.set(fail_key, attempts, timeout=BLOCK_MINUTES * 60)

    if attempts >= MAX_LOGIN_ATTEMPTS:
        blocked_until = time.time() + (BLOCK_MINUTES * 60)
        cache.set(block_key, blocked_until, timeout=BLOCK_MINUTES * 60)


def clear_failed_attempts(username, request):
    ip = _get_client_ip(request)
    cache.delete(_fail_key(username, ip))
    cache.delete(_block_key(username, ip))
