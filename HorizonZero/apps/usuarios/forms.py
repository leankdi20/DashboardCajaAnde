from django import forms

class LoginForm(forms.Form):
    usuario = forms.CharField(
        max_length=150,
        required=True,
        widget=forms.TextInput(
            attrs={
                "id": "usuario",
                "name": "usuario",
                "placeholder": "Ingrese su usuario",
                "class": "w-full pl-12 pr-4 py-4 rounded-xl border-none bg-slate-100 dark:bg-slate-800 focus:ring-2 focus:ring-primary shadow-sm transition-all text-slate-900 dark:text-slate-100 placeholder-slate-400",
            }
        ),
    )

    password = forms.CharField(
        required=True,
        widget=forms.PasswordInput(
            attrs={
                "id": "password",
                "name": "password",
                "placeholder": "Ingrese su contraseña",
                "class": "w-full pl-12 pr-12 py-4 rounded-xl border-none bg-slate-100 dark:bg-slate-800 focus:ring-2 focus:ring-primary shadow-sm transition-all text-slate-900 dark:text-slate-100 placeholder-slate-400",
            }
        ),
    )