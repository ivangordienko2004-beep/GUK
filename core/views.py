from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.contrib.auth.decorators import login_required
from django.http import FileResponse, Http404
from django.shortcuts import redirect, render

from .forms import ExcelUploadForm
from .services import create_report, decode_for_admin, merge_excel_files


SESSION_KEY = 'merged_file_path'
UPLOADED_FILES_KEY = 'uploaded_file_names'


def _session_path(request) -> Path | None:
    value = request.session.get(SESSION_KEY)
    if not value:
        return None
    path = Path(value)
    if not path.exists() or not str(path).startswith(str(settings.MEDIA_ROOT)):
        return None
    return path


@login_required
def dashboard(request):
    merged_file = _session_path(request)
    context = {
        'form': ExcelUploadForm(),
        'merged_file': merged_file,
        'uploaded_file_names': request.session.get(UPLOADED_FILES_KEY, []),
    }
    return render(request, 'core/dashboard.html', context)


@login_required
def questionnaire(request):
    return render(request, 'core/questionnaire.html')


@login_required
def upload_files(request):
    if request.method != 'POST':
        return redirect('dashboard')

    form = ExcelUploadForm(request.POST, request.FILES)
    uploaded_files = request.FILES.getlist('files')
    request.session[UPLOADED_FILES_KEY] = [file.name for file in uploaded_files]

    if not form.is_valid():
        for error in form.errors.get('files', []):
            messages.error(request, error)
        return redirect('dashboard')

    try:
        merged = merge_excel_files(form.cleaned_data['files'])
    except Exception as exc:
        messages.error(request, f'Не удалось объединить файлы: {exc}')
        return redirect('dashboard')

    request.session[SESSION_KEY] = str(merged)
    messages.success(
        request,
        f'Файлы успешно объединены: {", ".join(request.session.get(UPLOADED_FILES_KEY, []))}.'
    )
    return redirect('dashboard')


@login_required
def download_merged(request):
    path = _session_path(request)
    if not path:
        raise Http404('Нет подготовленного файла для скачивания.')
    return FileResponse(path.open('rb'), as_attachment=True, filename=path.name)


@staff_member_required
def decode_vus(request):
    path = _session_path(request)
    if not path:
        raise Http404('Сначала загрузите и объедините файлы.')

    decoded = decode_for_admin(path)
    request.session[SESSION_KEY] = str(decoded)
    messages.success(request, 'Расшифровка выполнена (доступно только администратору).')
    return redirect('dashboard')


@login_required
def create_report(request):
    path = _session_path(request)
    if not path:
        raise Http404('Сначала загрузите и объедините файлы.')

    report = create_report(path)
    return FileResponse(report.open('rb'), as_attachment=True, filename=report.name)