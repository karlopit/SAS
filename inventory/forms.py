from django import forms
from .models import Item, Transaction, BorrowRequest

class ItemForm(forms.ModelForm):
    class Meta:
        model = Item
        fields = ['name', 'description', 'quantity']

class BorrowRequestForm(forms.ModelForm):
    item = forms.ModelChoiceField(
        queryset=Item.objects.filter(available_quantity__gt=0),
        empty_label="— Select an item —",
        required=True,
    )

    class Meta:
        model = BorrowRequest
        fields = ['borrower_name', 'office_college', 'item', 'quantity']

class StaffBorrowForm(forms.ModelForm):
    class Meta:
        model = Transaction
        fields = ['item', 'quantity_borrowed', 'office_college']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['item'].queryset = Item.objects.filter(available_quantity__gt=0)

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
        fields = ['serviceable', 'unserviceable', 'sealed', 'lent_to_students', 'box_only', 'returned_qty']
        widgets = {
            'serviceable': forms.NumberInput(attrs={'min': '0'}),
            'unserviceable': forms.NumberInput(attrs={'min': '0'}),
            'sealed': forms.NumberInput(attrs={'min': '0'}),
            'lent_to_students': forms.NumberInput(attrs={'min': '0'}),
            'box_only': forms.NumberInput(attrs={'min': '0'}),
            'returned_qty': forms.NumberInput(attrs={'min': '0'}),
        }