from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
import uuid

import pandas as pd
from django.conf import settings
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MAPPING_DIR = Path(settings.BASE_DIR) / 'Mapping'
OFFICER_VUS_PATH = MAPPING_DIR / 'officer_vus.json'
NONOFFICE_VUS_PATH = MAPPING_DIR / 'nonoffice_vus.json'
NONOFFICE_POSITION_PATH = MAPPING_DIR / 'nonoffice_position.json'


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


OUTPUT_COLUMNS = [
    ('', 'ОКРУГ ВУЗа', '1', 'okrug_vuza'),
    ('', 'ОВУ, отв. за подготовку', '2', 'ovu_otv_podgotovku'),
    ('', 'Наименование ВУЗа', '3', 'nazvanie_vuza'),
    ('ВУС', '№', '4', 'vus_no'),
    ('ВУС', 'Наименование', '5', 'vus_naimenovanie'),
    ('Должность', '№', '6', 'doljnost_no'),
    ('Должность', 'Наименование', '7', 'doljnost_naimenovanie'),
    ('', 'Сбор/стаж', '8', 'sbor_stazhirovka'),
    ('', 'Программа', '9', 'programma_podgotovki'),
    ('', 'Место проведения', '10', 'mesto_provedeniya_uchebnogo_sbora'),
    ('Планируется', 'преподавателей', '11', 'planiruetsya_prepodavatelej'),
    ('Планируется', 'студентов', '12', 'planiruetsya_studentov'),
    ('Сроки проведения', 'начало', '13', 'srok_provedeniya_nachalo'),
    ('Сроки проведения', 'окончание', '14', 'srok_provedeniya_okonchanie'),
    ('', 'ФИО ответственного', '15', 'fio_otvetstvennogo'),
    ('', 'Мобильный', '16', 'mobilnyy'),
]


HEADER_ALIASES = {
    'округ вуза': 'okrug_vuza',
    'ову отв за подготовку': 'ovu_otv_podgotovku',
    'наименование вуза': 'nazvanie_vuza',
    'вуз': 'nazvanie_vuza',
    'код вус': 'vus_no',
    'вус': 'vus_no',
    'вус наименование': 'vus_naimenovanie',
    'должность': 'doljnost_no',
    'должность наименование': 'doljnost_naimenovanie',
    'сбор стаж': 'sbor_stazhirovka',
    'сбор стажировка': 'sbor_stazhirovka',
    'программа': 'programma_podgotovki',
    'программа подготовки': 'programma_podgotovki',
    'место проведения': 'mesto_provedeniya_uchebnogo_sbora',
    'место проведения сбора': 'mesto_provedeniya_uchebnogo_sbora',
    'преподавателей': 'planiruetsya_prepodavatelej',
    'студентов': 'planiruetsya_studentov',
    'начало': 'srok_provedeniya_nachalo',
    'окончание': 'srok_provedeniya_okonchanie',
    'фио ответственного': 'fio_otvetstvennogo',
    'мобильный': 'mobilnyy',
}


def _normalize(value: str) -> str:
    return ''.join(ch for ch in str(value).lower() if ch.isalnum() or ch.isspace()).strip()


def _normalize_code(value) -> str:
    if value is None:
        return ''

    text = str(value).strip()
    if not text:
        return ''

    if text.endswith('.0') and text[:-2].isdigit():
        text = text[:-2]

    return text


def _code_candidates(value, width: int) -> list[str]:
    text = _normalize_code(value)
    if not text:
        return []

    candidates = [text]
    if text.isdigit():
        padded = text.zfill(width)
        if padded not in candidates:
            candidates.insert(0, padded)
    return candidates


@lru_cache(maxsize=1)
def _load_json_mapping(path: str) -> dict[str, str]:
    file_path = Path(path)
    if not file_path.exists():
        return {}

    try:
        with file_path.open('r', encoding='utf-8') as file_obj:
            raw_mapping = json.load(file_obj)
    except (OSError, json.JSONDecodeError):
        return {}

    if not isinstance(raw_mapping, dict):
        return {}

    normalized: dict[str, str] = {}
    for key, value in raw_mapping.items():
        normalized_key = _normalize_code(key)
        if normalized_key:
            normalized[normalized_key] = str(value)
    return normalized


def _load_decoding_maps() -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    return (
        _load_json_mapping(str(OFFICER_VUS_PATH)),
        _load_json_mapping(str(NONOFFICE_VUS_PATH)),
        _load_json_mapping(str(NONOFFICE_POSITION_PATH)),
    )


def _lookup_decoding(value, mapping: dict[str, str], width: int) -> str | None:
    for candidate in _code_candidates(value, width):
        if candidate in mapping:
            return mapping[candidate]
    return None


def _vus_code_length(value) -> int:
    text = _normalize_code(value)
    digits = ''.join(ch for ch in text if ch.isdigit())
    return len(digits)


def _harmonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    direct_map = {_normalize(column): column for column in REQUIRED_COLUMNS}
    mapped_columns: dict[str, int] = {}

    for col_idx, col in enumerate(df.columns):
        normalized = _normalize(col)
        target = None
        if normalized in direct_map:
            target = direct_map[normalized]
        else:
            for alias, alias_target in HEADER_ALIASES.items():
                if alias in normalized:
                    target = alias_target
                    break

        if target and target not in mapped_columns:
            mapped_columns[target] = col_idx

    out = pd.DataFrame(index=df.index)
    for column in REQUIRED_COLUMNS:
        if column in mapped_columns:
            out[column] = df.iloc[:, mapped_columns[column]].fillna('').astype(str)
        else:
            out[column] = ''

    return out[REQUIRED_COLUMNS]


def _detect_header_depth(raw_df: pd.DataFrame) -> int:
    for idx in range(min(6, len(raw_df))):
        row = raw_df.iloc[idx].astype(str).str.strip()
        numeric_values = [value for value in row if value.isdigit()]
        if len(numeric_values) >= 5 and '1' in numeric_values:
            return idx + 1
    return 1


def _load_tabular_data(file) -> pd.DataFrame:
    if hasattr(file, 'seek'):
        file.seek(0)

    raw_df = pd.read_excel(file, header=None, dtype=str).fillna('')
    header_depth = _detect_header_depth(raw_df)

    if header_depth <= 1:
        if hasattr(file, 'seek'):
            file.seek(0)
        single_header_df = pd.read_excel(file, dtype=str).fillna('')
        return _harmonize_columns(single_header_df)

    header_rows = raw_df.iloc[:header_depth]
    filled_headers = []
    for idx, row in header_rows.iterrows():
        row_values = row.astype(str).str.strip()
        if idx < header_depth - 1:
            row_values = row_values.replace('', pd.NA).ffill().fillna('')
        filled_headers.append(row_values)

    combined_headers = []
    for col_idx in range(raw_df.shape[1]):
        pieces = []
        for row_idx in range(header_depth - 1):
            value = str(filled_headers[row_idx].iloc[col_idx]).strip()
            if value and value not in pieces:
                pieces.append(value)
        combined_headers.append(' '.join(pieces).strip())

    body_df = raw_df.iloc[header_depth:].copy().reset_index(drop=True)
    body_df.columns = combined_headers
    return _harmonize_columns(body_df)


def _export_merged_table(df: pd.DataFrame, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Объединение'

    for col_idx, (group, title, number, column_name) in enumerate(OUTPUT_COLUMNS, start=1):
        ws.cell(row=1, column=col_idx, value=group)
        ws.cell(row=2, column=col_idx, value=title)
        ws.cell(row=3, column=col_idx, value=number)
        ws.column_dimensions[ws.cell(row=2, column=col_idx).column_letter].width = 22

    group_start = None
    current_group = None
    for col_idx, (group, _, _, _) in enumerate(OUTPUT_COLUMNS, start=1):
        if group != current_group:
            if current_group and group_start and col_idx - 1 > group_start:
                ws.merge_cells(start_row=1, start_column=group_start, end_row=1, end_column=col_idx - 1)
            current_group = group
            group_start = col_idx
    if current_group and group_start and len(OUTPUT_COLUMNS) > group_start:
        ws.merge_cells(start_row=1, start_column=group_start, end_row=1, end_column=len(OUTPUT_COLUMNS))

    for row_idx, (_, row) in enumerate(df.iterrows(), start=4):
        for col_idx, (_, _, _, column_name) in enumerate(OUTPUT_COLUMNS, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row[column_name])

    wb.save(path)


def _read_harmonized_dataframe(path_or_file) -> pd.DataFrame:
    return _load_tabular_data(path_or_file)


def _remove_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    cleaned = df.copy().fillna('')
    mask = cleaned.apply(lambda row: any(str(value).strip() for value in row), axis=1)
    return cleaned.loc[mask].reset_index(drop=True)


def merge_excel_files(files) -> Path:
    files = list(files)
    if not files:
        raise ValueError('Не переданы файлы для объединения.')

    frames: list[pd.DataFrame] = []

    for file_obj in files:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)

        df = _read_harmonized_dataframe(file_obj)
        df = _remove_empty_rows(df)

        if not df.empty:
            frames.append(df)

    if not frames:
        raise ValueError('Не удалось прочитать данные из файлов.')

    merged_df = pd.concat(frames, ignore_index=True)
    merged_df = _remove_empty_rows(merged_df)

    output_dir = Path(settings.MEDIA_ROOT) / 'exports'
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f'merged_{uuid.uuid4().hex}.xlsx'

    _export_merged_table(merged_df, path)
    return path


def decode_for_admin(path: Path) -> Path:
    df = _read_harmonized_dataframe(path)
    officer_vus, nonoffice_vus, nonoffice_positions = _load_decoding_maps()

    def _decode_vus(row) -> str:
        if _vus_code_length(row['vus_no']) == 6:
            officer_name = _lookup_decoding(row['vus_no'], officer_vus, 6)
            if officer_name:
                return officer_name
        elif _vus_code_length(row['vus_no']) == 3:
            nonoffice_name = _lookup_decoding(row['vus_no'], nonoffice_vus, 3)
            if nonoffice_name:
                return nonoffice_name

        return row['vus_naimenovanie']

    df['vus_naimenovanie'] = df.apply(_decode_vus, axis=1)

    officer_mask = df['vus_no'].map(lambda value: _vus_code_length(value) == 6)
    nonofficer_mask = df['vus_no'].map(lambda value: _vus_code_length(value) == 3)

    df.loc[officer_mask, ['doljnost_no', 'doljnost_naimenovanie']] = ''
    df.loc[nonofficer_mask, 'doljnost_naimenovanie'] = df.loc[nonofficer_mask, 'doljnost_no'].map(
        lambda val: _lookup_decoding(val, nonoffice_positions, 3) or ''
    )

    decoded_path = path.with_name(f'{path.stem}_decoded.xlsx')
    _export_merged_table(df, decoded_path)
    return decoded_path


def create_report(path: Path) -> Path:
    df = _read_harmonized_dataframe(path)
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