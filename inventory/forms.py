from django import forms
from .models import Item, BorrowRequest, Transaction


class BorrowRequestForm(forms.ModelForm):
    borrower_role = forms.ChoiceField(
        choices=[('student', 'Student'), ('employee', 'Employee')],
        widget=forms.RadioSelect,
        required=True,
        label="Borrower Type"
    )
    
    # Student name fields (split)
    student_last_name = forms.CharField(max_length=100, required=False, label="Last Name")
    student_first_name = forms.CharField(max_length=100, required=False, label="First Name")
    student_middle_initial = forms.CharField(max_length=2, required=False, label="Middle Initial")
    
    # Student fields
    year_section = forms.CharField(max_length=100, required=False, label="Year & Section")
    student_id = forms.CharField(max_length=50, required=False, label="Student ID")
    college = forms.CharField(max_length=200, required=False, label="College")
    academic_year = forms.CharField(max_length=50, required=False, label="Academic Year")
    
    # Employee name fields (split)
    employee_last_name = forms.CharField(max_length=100, required=False, label="Last Name")
    employee_first_name = forms.CharField(max_length=100, required=False, label="First Name")
    employee_middle_initial = forms.CharField(max_length=2, required=False, label="Middle Initial")
    
    # Employee fields
    employee_id = forms.CharField(max_length=50, required=False, label="Employee ID")
    office = forms.CharField(max_length=200, required=False, label="Office")
    
    # Common fields
    item = forms.ModelChoiceField(
        queryset=Item.objects.filter(available_quantity__gt=0),
        required=True,
        label="Item to Borrow"
    )
    quantity = forms.IntegerField(min_value=1, required=True, label="Quantity Needed")
    
    class Meta:
        model = BorrowRequest
        fields = ['item', 'quantity']
    
    def clean(self):
        cleaned_data = super().clean()
        borrower_role = cleaned_data.get('borrower_role')
        
        if borrower_role == 'student':
            required_fields = ['student_last_name', 'student_first_name', 'year_section', 
                              'student_id', 'college', 'academic_year']
            for field in required_fields:
                value = cleaned_data.get(field)
                if not value or value.strip() == '':
                    self.add_error(field, f'This field is required for students.')
        
        elif borrower_role == 'employee':
            required_fields = ['employee_last_name', 'employee_first_name', 'employee_id', 'office']
            for field in required_fields:
                value = cleaned_data.get(field)
                if not value or value.strip() == '':
                    self.add_error(field, f'This field is required for employees.')
        
        # Validate quantity against available stock
        item = cleaned_data.get('item')
        quantity = cleaned_data.get('quantity')
        if item and quantity:
            if quantity > item.available_quantity:
                self.add_error('quantity', f'Only {item.available_quantity} item(s) available.')
        
        return cleaned_data
    
    def save(self, commit=True):
        instance = super().save(commit=False)
        
        # Populate borrower name and type based on role
        if self.cleaned_data['borrower_role'] == 'student':
            # Combine name: First Name + Middle Initial + Last Name
            first_name = self.cleaned_data['student_first_name']
            middle_initial = self.cleaned_data['student_middle_initial']
            last_name = self.cleaned_data['student_last_name']
            
            # Build full name
            full_name = first_name
            if middle_initial:
                full_name += f" {middle_initial}."
            full_name += f" {last_name}"
            
            instance.borrower_name = full_name
            instance.borrower_type = 'student'
            instance.student_id = self.cleaned_data['student_id']
            instance.year_section = self.cleaned_data['year_section']
            instance.college = self.cleaned_data['college']
            instance.academic_year = self.cleaned_data['academic_year']
            # Clear employee fields
            instance.employee_id = None
            instance.office = None
            # Set office_college to college for students
            instance.office_college = self.cleaned_data['college']
        else:
            # Combine name: First Name + Middle Initial + Last Name
            first_name = self.cleaned_data['employee_first_name']
            middle_initial = self.cleaned_data['employee_middle_initial']
            last_name = self.cleaned_data['employee_last_name']
            
            # Build full name
            full_name = first_name
            if middle_initial:
                full_name += f" {middle_initial}."
            full_name += f" {last_name}"
            
            instance.borrower_name = full_name
            instance.borrower_type = 'employee'
            instance.employee_id = self.cleaned_data['employee_id']
            instance.office = self.cleaned_data['office']
            # Clear student fields
            instance.student_id = None
            instance.year_section = None
            instance.college = None
            instance.academic_year = None
            # Set office_college to office for employees
            instance.office_college = self.cleaned_data['office']
        
        if commit:
            instance.save()
        return instance


class StaffBorrowForm(forms.ModelForm):
    # Dynamic field for multiple serial numbers
    serial_numbers = forms.CharField(
        required=True,
        label="Device Serial Numbers",
        help_text="Enter one serial number per line. You need to enter exactly the same number as the quantity requested.",
        widget=forms.Textarea(attrs={
            'rows': 4,
            'placeholder': 'SN-001\nSN-002\nSN-003',
            'class': 'form-control'
        })
    )
    
    class Meta:
        model = Transaction
        fields = ['item', 'quantity_borrowed', 'office_college']
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['item'].queryset = Item.objects.filter(available_quantity__gt=0)
        self.fields['quantity_borrowed'].widget.attrs['min'] = 1
    
    def clean_serial_numbers(self):
        serial_numbers_text = self.cleaned_data.get('serial_numbers')
        quantity = self.cleaned_data.get('quantity_borrowed')
        
        if not serial_numbers_text:
            raise forms.ValidationError('Please enter serial numbers.')
        
        # Split by newline and strip whitespace
        serials = [s.strip() for s in serial_numbers_text.split('\n') if s.strip()]
        
        if len(serials) != quantity:
            raise forms.ValidationError(f'You entered {len(serials)} serial number(s), but requested quantity is {quantity}. Please enter exactly {quantity} serial number(s).')
        
        # Check for duplicate serial numbers
        if len(serials) != len(set(serials)):
            raise forms.ValidationError('Duplicate serial numbers found. Each serial number must be unique.')
        
        # Check if any serial number is already borrowed and not returned
        existing_serials = Transaction.objects.filter(
            serial_number__in=serials,
            status='borrowed'
        ).values_list('serial_number', flat=True)
        
        if existing_serials:
            raise forms.ValidationError(f'The following serial numbers are already borrowed: {", ".join(existing_serials)}')
        
        return serials


class ItemForm(forms.ModelForm):
    class Meta:
        model = Item
        fields = ['name', 'description', 'serial', 'quantity']
    
    def save(self, commit=True):
        instance = super().save(commit=False)
        instance.available_quantity = instance.quantity
        if commit:
            instance.save()
        return instance


class TransactionConditionForm(forms.ModelForm):
    class Meta:
        model = Transaction
        fields = ['serviceable', 'unserviceable', 'sealed', 'lent_to_students', 'box_only']