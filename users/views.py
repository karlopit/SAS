from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.core.exceptions import PermissionDenied
from .forms import RegisterForm, LoginForm, EditUserForm, ResetPasswordForm
from django.contrib.auth import get_user_model

User = get_user_model()

def register_view(request):
    form = RegisterForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.save()
        login(request, user)
        return redirect('login')
    return render(request, 'users/register.html', {'form': form})

def login_view(request):
    form = LoginForm(data=request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.get_user()
        login(request, user)
        return redirect('index')
    return render(request, 'users/login.html', {'form': form})

def logout_view(request):
    logout(request)
    return redirect('login')

@login_required
def profile_view(request):
    return render(request, 'users/profile.html', {'user': request.user})

@login_required
def user_list_view(request):
    if request.user.role != 'admin':
        raise PermissionDenied
    users = User.objects.all().order_by('username')
    return render(request, 'users/user_list.html', {'users': users})

@login_required
def edit_user_view(request, user_id):
    if request.user.role != 'admin':
        raise PermissionDenied
    target = get_object_or_404(User, id=user_id)
    form = EditUserForm(request.POST or None, instance=target)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, f"User '{target.username}' updated successfully.")
        return redirect('user_list')
    return render(request, 'users/edit_user.html', {'form': form, 'target': target})

@login_required
def reset_password_view(request, user_id):
    if request.user.role != 'admin':
        raise PermissionDenied
    target = get_object_or_404(User, id=user_id)
    form = ResetPasswordForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        target.set_password(form.cleaned_data['new_password'])
        target.save()
        messages.success(request, f"Password for '{target.username}' has been reset.")
        return redirect('user_list')
    return render(request, 'users/reset_password.html', {'form': form, 'target': target})