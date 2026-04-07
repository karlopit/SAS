from django import forms
from .models import Item, Transaction

class ItemForm(forms.ModelForm):
    class Meta:
        model = Item
        fields = ['name', 'description', 'serial', 'quantity']

class BorrowForm(forms.ModelForm):
    class Meta:
        model = Transaction
        fields = ['item', 'quantity_borrowed', 'office_college']

    def clean(self):
        cleaned_data = super().clean()
        item = cleaned_data.get('item')
        qty = cleaned_data.get('quantity_borrowed')
        if item and qty and qty > item.available_quantity:
            raise forms.ValidationError(f"Only {item.available_quantity} units available.")
        return cleaned_data

class TransactionConditionForm(forms.ModelForm):
    class Meta:
        model = Transaction
        fields = ['serviceable', 'unserviceable', 'sealed', 'lent_to_students', 'box_only']
        widgets = {
            'serviceable': forms.NumberInput(attrs={'min': '0'}),
            'unserviceable': forms.NumberInput(attrs={'min': '0'}),
            'sealed': forms.NumberInput(attrs={'min': '0'}),
            'lent_to_students': forms.NumberInput(attrs={'min': '0'}),
            'box_only': forms.NumberInput(attrs={'min': '0'}),
        }