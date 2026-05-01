# apps/dashboard/views/home.py
from django.contrib.auth.decorators import login_required
from django.shortcuts import render


@login_required
def dashboard_home(request):
    for attr in ('_perm_cache', '_user_perm_cache', '_group_perm_cache'):
        try:
            delattr(request.user, attr)
        except AttributeError:
            pass
    return render(request, "dashboard/home.html", {
        "show_welcome": True,
    })
