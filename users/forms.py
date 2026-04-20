from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django import forms
from django.contrib.auth import get_user_model

User = get_user_model()


class AddUserForm(forms.ModelForm):
    """
    Used by the Admin 'Add User' popup modal in user_list.html.
    Collects: username, password, first_name, middle_initial, last_name, role.
    """
    password = forms.CharField(
        widget=forms.PasswordInput,
        min_length=6,
        label='Password',
    )
    confirm_password = forms.CharField(
        widget=forms.PasswordInput,
        label='Confirm Password',
    )

    class Meta:
        model = User
        fields = ['username', 'first_name', 'middle_initial', 'last_name', 'role']

    def clean(self):
        cleaned = super().clean()
        p1 = cleaned.get('password')
        p2 = cleaned.get('confirm_password')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError('Passwords do not match.')
        return cleaned

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data['password'])
        if commit:
            user.save()
        return user


class RegisterForm(UserCreationForm):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('staff', 'Staff'),
    ]
    role           = forms.ChoiceField(choices=ROLE_CHOICES)
    first_name     = forms.CharField(max_length=150, required=False, label='First Name')
    middle_initial = forms.CharField(max_length=5,   required=False, label='Middle Initial')
    last_name      = forms.CharField(max_length=150, required=False, label='Last Name')

    class Meta:
        model = User
        fields = ['username', 'first_name', 'middle_initial', 'last_name', 'role', 'password1', 'password2']

    def save(self, commit=True):
        user = super().save(commit=False)
        user.middle_initial = self.cleaned_data.get('middle_initial', '')
        if commit:
            user.save()
        return user


class LoginForm(AuthenticationForm):
    username = forms.CharField()
    password = forms.CharField(widget=forms.PasswordInput)


class EditUserForm(forms.ModelForm):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('staff', 'Staff'),
    ]
    role           = forms.ChoiceField(choices=ROLE_CHOICES)
    is_active      = forms.BooleanField(required=False, label='Active')
    middle_initial = forms.CharField(max_length=5, required=False, label='Middle Initial')

    class Meta:
        model = User
        fields = ['username', 'first_name', 'middle_initial', 'last_name', 'email', 'role', 'is_active']


class ResetPasswordForm(forms.Form):
    new_password = forms.CharField(
        widget=forms.PasswordInput,
        min_length=6,
        label='New Password',
    )
    confirm_password = forms.CharField(
        widget=forms.PasswordInput,
        label='Confirm Password',
    )

    def clean(self):
        cleaned_data = super().clean()
        p1 = cleaned_data.get('new_password')
        p2 = cleaned_data.get('confirm_password')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError('Passwords do not match.')
        return cleaned_data