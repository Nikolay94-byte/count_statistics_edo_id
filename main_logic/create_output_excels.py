import json
import logging
import openpyxl
import pandas as pd
from format_excel import format_excel
from utils.constants import INPUT_FILE_PATH, DOCUMENT_TYPE
from utils.utils import open_excel
from settings import DATA_PATH
from utils.convert_logic import write_headers_to_excel, write_rows_to_excel


def find_column_index(type_request: str) -> int:
    """Находит индекс необходимой колонки для извлечения данных."""
    sheet = open_excel(INPUT_FILE_PATH)
    for cell in sheet['1']:
        if cell.value == type_request:
            column_num_from_book = sheet[cell.coordinate].column - 1  # уменьшаем на 1 ,т.к. далее используется
            # нумерация колонок с индекса 0
            return column_num_from_book

def convert_json_to_excel(column_num_from_book: int, value_column_name: str) -> str:
    """Формирует excel-файл с данными на основе json."""
    sheet = open_excel(INPUT_FILE_PATH)
    file_bodes = []

    for row in range(2, sheet.max_row+1):
        if sheet[row][column_num_from_book].value is not None:
            try:
                file_bodes.append(json.loads(sheet[row][column_num_from_book].value))
            except Exception:
                logging.error(f'Ячейка (строка {row}, колонка {column_num_from_book+1}) содержит некорректный json')
    json_numerated_bodes = dict(enumerate(file_bodes, start=2))
    # записываем наименования колонок в excel-файл
    book = openpyxl.Workbook()
    write_headers_to_excel(sheet=book.active)
    # записываем строки в excel-файл
    for row_number, json_body in json_numerated_bodes.items():
        write_rows_to_excel(row_number, json_body, sheet=book.active)
    # записываем название файла
    file_name_suffix = sheet[1][column_num_from_book].value
    doc_class_name = ''
    for cell in sheet['1']:
        if cell.value == DOCUMENT_TYPE:
            doc_class_name = str(sheet[2][sheet[cell.coordinate].column-1].value)
            break
    new_book_name = f"{doc_class_name}_{file_name_suffix}.xlsx"

    # Преобразуем DataFrame для последующего форматирования
    data = list(book.active.values)
    dataframe_for_formating = pd.DataFrame(data[1:], columns=data[0])
    formatted_dataframe = format_excel(value_column_name, dataframe_for_formating)
    formatted_dataframe.to_excel(DATA_PATH / new_book_name, index=False)
    return new_book_name
