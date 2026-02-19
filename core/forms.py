from django import forms


class MultipleFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class MultipleFileField(forms.FileField):
    def __init__(self, *args, **kwargs):
        kwargs.setdefault('widget', MultipleFileInput(attrs={'accept': '.xlsx,.xls'}))
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        single_file_clean = super().clean
        if isinstance(data, (list, tuple)):
            return [single_file_clean(item, initial) for item in data]
        return [single_file_clean(data, initial)]


class ExcelUploadForm(forms.Form):
    files = MultipleFileField(label='Excel файлы')

    def clean_files(self):
        files = self.cleaned_data.get('files', [])
        if not files:
            raise forms.ValidationError('Выберите хотя бы один Excel файл.')

        for file in files:
            if not file.name.lower().endswith(('.xlsx', '.xls')):
                raise forms.ValidationError(f'Файл {file.name} не является Excel.')

        return files
