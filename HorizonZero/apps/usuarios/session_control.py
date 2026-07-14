from urllib.parse import urlencode

from django.conf import settings
from django.http import JsonResponse
from django.shortcuts import redirect


COOKIE_NAME = "hz_token"
FORCED_LOGOUT_QUERY_PARAM = "forced_logout"
SESSION_REPLACED_REASON = "session_replaced"
SESSION_INVALID_REASON = "session_invalid"

FORCED_LOGOUT_MESSAGES = {
    SESSION_REPLACED_REASON: "Su sesion fue iniciada en otro navegador o dispositivo.",
    SESSION_INVALID_REASON: "Su sesion ya no es valida. Inicie sesion de nuevo.",
}


class APISessionExpiredError(Exception):
    def __init__(self, reason=SESSION_REPLACED_REASON, message=None):
        self.reason = reason or SESSION_REPLACED_REASON
        self.message = message or get_forced_logout_message(self.reason)
        super().__init__(self.message)


def get_forced_logout_message(reason):
    return FORCED_LOGOUT_MESSAGES.get(reason, FORCED_LOGOUT_MESSAGES[SESSION_INVALID_REASON])


def is_async_request(request):
    requested_with = request.headers.get("X-Requested-With", "")
    accept = request.headers.get("Accept", "")
    return requested_with == "XMLHttpRequest" or "application/json" in accept.lower()


def _clear_server_session(request):
    try:
        request.session.flush()
    except Exception:
        try:
            request.session.clear()
        except Exception:
            pass


def _delete_auth_cookie(response):
    response.delete_cookie(
        COOKIE_NAME,
        path="/",
        samesite="Lax",
    )
    return response


def build_login_url(reason=None, next_url=None):
    params = {}
    if reason:
        params[FORCED_LOGOUT_QUERY_PARAM] = reason
    if next_url:
        params["next"] = next_url
    query = urlencode(params)
    return f"{settings.LOGIN_URL}?{query}" if query else settings.LOGIN_URL


def build_forced_logout_response(request, reason=SESSION_REPLACED_REASON, message=None, status=401):
    _clear_server_session(request)
    message = message or get_forced_logout_message(reason)

    if is_async_request(request):
        response = JsonResponse(
            {
                "detail": message,
                "reason": reason,
                "forced_logout": True,
                "login_url": build_login_url(reason=reason),
            },
            status=status,
        )
        response["X-Force-Logout"] = "1"
        response["X-Force-Logout-Reason"] = reason
        response["Cache-Control"] = "no-store"
        return _delete_auth_cookie(response)

    response = redirect(build_login_url(reason=reason))
    return _delete_auth_cookie(response)
