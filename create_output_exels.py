import json
import openpyxl
from format_exel import format_exel
from utils import open_exel
from settings import DATA_PATH, INPUTFILEPATH
from convert_logic import write_headers_to_exel, write_row_to_exel


def find_column(type_request: str) -> int:
    """Находит индекс необходимой для извлечения данных колонки."""
    sheet = open_exel(INPUTFILEPATH)
    for cell in sheet['1']:
        if cell.value == type_request:
            column_num_from_book = sheet[cell.coordinate].column - 1  # уменьшаем на 1 ,т.к. далее используется
            # нумерация колонок с индекса 0
            return column_num_from_book

def convert_json_to_exel(column_num_from_book: int, value_column_name: str) -> str:
    """Формирует exel файл с данными на основе json."""
    sheet = open_exel(INPUTFILEPATH)
    file_bodes = []

    for row in range(2, sheet.max_row+1):
        if sheet[row][column_num_from_book].value is not None:
            try:
                file_bodes.append(json.loads(sheet[row][column_num_from_book].value))
            except Exception:
                print(f'Ячейка (строка {row}, колонка {column_num_from_book+1}) содержит некорректный json')
    json_numerated_bodes = dict(enumerate(file_bodes, start=2))
    # записываем наименования колонок в exel файл
    book = openpyxl.Workbook()
    write_headers_to_exel(sheet=book.active)
    # записываем строки в exel файл
    for row_number, json_body in json_numerated_bodes.items():
        write_row_to_exel(row_number, json_body, sheet=book.active)
    # записываем название файла
    file_name_suffix = sheet[1][column_num_from_book].value
    doc_class_name = ''
    for cell in sheet['1']:
        if cell.value == 'document_type':
            doc_class_name = str(sheet[2][sheet[cell.coordinate].column-1].value)
            break
    new_book_name = doc_class_name + '_' + file_name_suffix + '.xlsx'
    book.save(DATA_PATH / new_book_name)
    book.close()
    return format_exel(value_column_name, new_book_name)
