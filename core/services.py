from __future__ import annotations

from pathlib import Path
import uuid

import pandas as pd
from django.conf import settings
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

VUS_DECODING = {
    '021500': 'Командир мотострелкового подразделения',
    '101000': 'Специалист связи',
    '201200': 'Инженер вооружения',
}

POSITION_DECODING = {
    '001': 'Командир отделения',
    '002': 'Заместитель командира взвода',
    '003': 'Старшина роты',
}


REQUIRED_COLUMNS = [
    'okrug_vuza',
    'ovu_otv_podgotovku',
    'nazvanie_vuza',
    'vus_no',
    'vus_naimenovanie',
    'doljnost_no',
    'doljnost_naimenovanie',
    'sbor_stazhirovka',
    'programma_podgotovki',
    'mesto_provedeniya_uchebnogo_sbora',
    'planiruetsya_prepodavatelej',
    'planiruetsya_studentov',
    'srok_provedeniya_nachalo',
    'srok_provedeniya_okonchanie',
    'fio_otvetstvennogo',
    'mobilnyy',
]


HEADER_ALIASES = {
    'округ вуза': 'okrug_vuza',
    'ову, отв. за подготовку': 'ovu_otv_podgotovku',
    'наименование вуза': 'nazvanie_vuza',
    'вуз': 'nazvanie_vuza',
    'код вус': 'vus_no',
    '№ вус': 'vus_no',
    'наименование вус': 'vus_naimenovanie',
    '№ должности': 'doljnost_no',
    'номер должности': 'doljnost_no',
    'наименование должности': 'doljnost_naimenovanie',
    'сбор/стаж': 'sbor_stazhirovka',
    'сбор/стажировка': 'sbor_stazhirovka',
    'программа': 'programma_podgotovki',
    'программа подготовки': 'programma_podgotovki',
    'место проведения': 'mesto_provedeniya_uchebnogo_sbora',
    'место проведения сбора': 'mesto_provedeniya_uchebnogo_sbora',
    'преподавателей': 'planiruetsya_prepodavatelej',
    'планируем преподавателей': 'planiruetsya_prepodavatelej',
    'студентов': 'planiruetsya_studentov',
    'планируем студентов': 'planiruetsya_studentov',
    'начало': 'srok_provedeniya_nachalo',
    'окончание': 'srok_provedeniya_okonchanie',
    'фио ответ': 'fio_otvetstvennogo',
    'мобильный': 'mobilnyy',
}


def _normalize(value: str) -> str:
    return ''.join(ch for ch in str(value).lower() if ch.isalnum() or ch == ' ')


def _harmonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for col in df.columns:
        normalized = _normalize(col)
        for alias, target in HEADER_ALIASES.items():
            if alias in normalized:
                renamed[col] = target
                break
    out = df.rename(columns=renamed)
    for column in REQUIRED_COLUMNS:
        if column not in out.columns:
            out[column] = ''
    return out[REQUIRED_COLUMNS]


def merge_excel_files(files) -> Path:
    frames = []
    for file in files:
        sheet_df = pd.read_excel(file, dtype=str).fillna('')
        frames.append(_harmonize_columns(sheet_df))

    merged = pd.concat(frames, ignore_index=True).fillna('')
    merged['planiruetsya_studentov'] = pd.to_numeric(merged['planiruetsya_studentov'], errors='coerce').fillna(0).astype(int)
    merged['planiruetsya_prepodavatelej'] = pd.to_numeric(merged['planiruetsya_prepodavatelej'], errors='coerce').fillna(0).astype(int)

    output_dir = Path(settings.MEDIA_ROOT) / 'exports'
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f'merged_{uuid.uuid4().hex}.xlsx'
    merged.to_excel(path, index=False)
    return path


def decode_for_admin(path: Path) -> Path:
    df = pd.read_excel(path, dtype=str).fillna('')
    df['vus_naimenovanie'] = df.apply(
        lambda row: VUS_DECODING.get(str(row['vus_no']).strip(), row['vus_naimenovanie']),
        axis=1,
    )

    officer_mask = df['programma_podgotovki'].str.contains('офицер', case=False, na=False)
    df.loc[officer_mask, ['doljnost_no', 'doljnost_naimenovanie']] = ''

    sergeant_mask = ~officer_mask
    df.loc[sergeant_mask, 'doljnost_naimenovanie'] = df.loc[sergeant_mask, 'doljnost_no'].map(
        lambda val: POSITION_DECODING.get(str(val).strip(), '')
    )

    decoded_path = path.with_name(f'{path.stem}_decoded.xlsx')
    df.to_excel(decoded_path, index=False)
    return decoded_path


def create_report(path: Path) -> Path:
    df = pd.read_excel(path).fillna('')
    for col in ['planiruetsya_studentov', 'planiruetsya_prepodavatelej']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'

    border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000'),
    )
    fill = PatternFill('solid', fgColor='4B5320')
    title_fill = PatternFill('solid', fgColor='6B8E23')
    white_bold = Font(color='FFFFFF', bold=True)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.merge_cells('A1:H1')
    ws['A1'] = 'ИТОГОВЫЙ ПЛАН СБОРОВ'
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=14)
    ws['A1'].fill = fill
    ws['A1'].alignment = center

    headers = ['Округ', 'ОВУ', 'ВУЗ', 'ВУС', 'Наименование ВУС', 'Программа', 'Студентов', 'Преподавателей']
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=idx, value=header)
        cell.font = white_bold
        cell.fill = title_fill
        cell.alignment = center
        cell.border = border

    current = 3
    for _, row in df.iterrows():
        values = [
            row['okrug_vuza'],
            row['ovu_otv_podgotovku'],
            row['nazvanie_vuza'],
            row['vus_no'],
            row['vus_naimenovanie'],
            row['programma_podgotovki'],
            int(row['planiruetsya_studentov']),
            int(row['planiruetsya_prepodavatelej']),
        ]
        for idx, value in enumerate(values, start=1):
            cell = ws.cell(row=current, column=idx, value=value)
            cell.border = border
            cell.alignment = center
        current += 1

    ws.cell(row=current, column=6, value='ИТОГО').font = Font(bold=True)
    ws.cell(row=current, column=7, value=int(df['planiruetsya_studentov'].sum())).font = Font(bold=True)
    ws.cell(row=current, column=8, value=int(df['planiruetsya_prepodavatelej'].sum())).font = Font(bold=True)
    for c in range(1, 9):
        ws.cell(row=current, column=c).border = border
        ws.cell(row=current, column=c).alignment = center

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws.column_dimensions[col].width = 24

    report_path = path.with_name(f'{path.stem}_report.xlsx')
    wb.save(report_path)
    return report_path
