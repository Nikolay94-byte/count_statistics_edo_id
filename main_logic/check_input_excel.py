import logging
from pathlib import Path

from utils import constants
from utils.utils import open_excel


def check_and_clean_file(filepath: str):
    """Проверяет, соответствует ли excel-файл необходимому формату:
    1 Есть ли 5 обязательных колонок:
    - regnumber
    - attribute_name
    - rus_attribute_name
    - text_normalized
    - text_verification
    2 Есть ли ячейки, содержащие 32767 символов (32767 - максимально допустимое количество символов в ячейке excel,
    что означает, что json не полный, сломанный. Удаляет строки с такими ячейками.
    """

    # Проверка на нужные колонки
    sheet = open_excel(filepath)
    book = sheet.parent
    mandatory_column_name_list = [
        constants.REGNUMBER, constants.ATTRIBUTE_NAME,
        constants.RUS_ATTRIBUTE_NAME, constants.TEXT_NORMALIZED, constants.TEXT_VERIFICATION]
    current_column_name_list = [cell.value for cell in sheet['1'] if cell.value]
    if not set(mandatory_column_name_list).issubset(current_column_name_list):
        raise ValueError(
            f'В файле {filepath} отсутствуют обязательные колонки, убедитесь что все 5 есть в файле '
            f'{constants.REGNUMBER},{constants.ATTRIBUTE_NAME}, {constants.RUS_ATTRIBUTE_NAME}'
            f'{constants.TEXT_NORMALIZED},{constants.TEXT_VERIFICATION}'
        )

    # Проверка на битые ячейки
    rows_to_delete = set()
    problematic_cells = []
    for row in range(2, sheet.max_row+1):
        for cell in sheet[str(row)]:
            if len(str(cell.value)) == 32767:
                rows_to_delete.add(row)
                problematic_cells.append(
                    f"строка {row}, колонка {sheet[cell.coordinate].column}"
                )

    # Удаляем строки с битыми ячейками если есть
    if rows_to_delete:
        for row in sorted(rows_to_delete, reverse=True):
            sheet.delete_rows(row)
        book.save(filepath)

        cells_info = "\n".join(problematic_cells)
        log_message = (
            f"В файле {Path(filepath).stem} были удалены строки с ячейками, содержащими 32767 символов:\n"
            f"{cells_info}\n"
            f"Всего удалено строк: {len(rows_to_delete)}"
        )
        logging.warning(log_message)
    else:
        logging.info(f"Файл {Path(filepath).stem} не содержит проблемных строк (32767 символов в ячейке)")

    book.close()
