from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login, logout
from django.contrib.auth.decorators import login_required
from django.views.decorators.cache import never_cache
from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.http import JsonResponse
from django.views.decorators.http import require_POST
from .forms import RegisterForm, LoginForm, EditUserForm, ResetPasswordForm, AddUserForm
from django.contrib.auth import get_user_model

User = get_user_model()


def register_view(request):
    form = RegisterForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.save()
        login(request, user)
        return redirect('login')
    return render(request, 'users/register.html', {'form': form})


@never_cache
def login_view(request):
    form = LoginForm(data=request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.get_user()
        login(request, user)
        return redirect('index')
    return render(request, 'users/login.html', {'form': form})


@never_cache
def logout_view(request):
    logout(request)
    response = redirect('login')
    response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response['Pragma']  = 'no-cache'
    response['Expires'] = '0'
    return response


@never_cache
@login_required
def profile_view(request):
    return render(request, 'users/profile.html', {'user': request.user})


@never_cache
@login_required
def user_list_view(request):
    if request.user.role != 'admin':
        raise PermissionDenied

    add_form = AddUserForm()
    add_errors = None

    if request.method == 'POST' and request.POST.get('_action') == 'add_user':
        add_form = AddUserForm(request.POST)
        if add_form.is_valid():
            add_form.save()
            messages.success(request, f"User '{add_form.cleaned_data['username']}' created successfully.")
            return redirect('user_list')
        else:
            # Re-render page with modal errors
            add_errors = add_form.errors

    users = User.objects.all().order_by('username')
    return render(request, 'users/user_list.html', {
        'users':      users,
        'add_form':   add_form,
        'add_errors': add_errors,
        # Re-open modal if there were errors
        'show_add_modal': bool(add_errors),
    })


@never_cache
@login_required
def edit_user_view(request, user_id):
    if request.user.role != 'admin':
        raise PermissionDenied
    target = get_object_or_404(User, id=user_id)

    if request.method == 'POST':
        target.username       = request.POST.get('username',       target.username)
        target.email          = request.POST.get('email',          target.email)
        target.first_name     = request.POST.get('first_name',     target.first_name)
        target.middle_initial = request.POST.get('middle_initial', target.middle_initial)
        target.last_name      = request.POST.get('last_name',      target.last_name)
        target.role           = request.POST.get('role',           target.role)
        target.is_active      = request.POST.get('is_active') == '1'
        target.save()
        messages.success(request, f"User '{target.username}' updated successfully.")
        return redirect('user_list')

    return render(request, 'users/edit_user.html', {'target': target})


@never_cache
@login_required
def reset_password_view(request, user_id):
    if request.user.role != 'admin':
        raise PermissionDenied
    target = get_object_or_404(User, id=user_id)
    form   = ResetPasswordForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        target.set_password(form.cleaned_data['new_password'])
        target.save()
        messages.success(request, f"Password for '{target.username}' has been reset.")
        return redirect('user_list')
    return render(request, 'users/reset_password.html', {'form': form, 'target': target})