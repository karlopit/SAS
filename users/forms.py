from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django import forms
from django.contrib.auth import get_user_model

User = get_user_model()

class RegisterForm(UserCreationForm):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('staff', 'Staff'),
    ]
    role = forms.ChoiceField(choices=ROLE_CHOICES)

    class Meta:
        model = User
        fields = ['username', 'email', 'role', 'password1', 'password2']

class LoginForm(AuthenticationForm):
    username = forms.CharField()
    password = forms.CharField(widget=forms.PasswordInput)

class EditUserForm(forms.ModelForm):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('staff', 'Staff'),
    ]
    role = forms.ChoiceField(choices=ROLE_CHOICES)
    is_active = forms.BooleanField(required=False, label='Active')

    class Meta:
        model = User
        fields = ['username', 'email', 'role', 'is_active']

class ResetPasswordForm(forms.Form):
    new_password = forms.CharField(
        widget=forms.PasswordInput,
        min_length=6,
        label='New Password'
    )
    confirm_password = forms.CharField(
        widget=forms.PasswordInput,
        label='Confirm Password'
    )

    def clean(self):
        cleaned_data = super().clean()
        p1 = cleaned_data.get('new_password')
        p2 = cleaned_data.get('confirm_password')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError("Passwords do not match.")
        return cleaned_data