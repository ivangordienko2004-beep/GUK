from django import forms


class ExcelUploadForm(forms.Form):
    files = forms.FileField(
        label='Excel файлы',
        widget=forms.ClearableFileInput(attrs={'multiple': True, 'accept': '.xlsx,.xls'})
    )

    def clean_files(self):
        files = self.files.getlist('files')
        if not files:
            raise forms.ValidationError('Выберите хотя бы один Excel файл.')
        for file in files:
            if not file.name.lower().endswith(('.xlsx', '.xls')):
                raise forms.ValidationError(f'Файл {file.name} не является Excel.')
        return files
