import logging
from pathlib import Path

from utils import constants
from utils.utils import open_excel


def check_and_clean_file(filepath: str):
    """Проверяет, соответствует ли excel-файл необходимому формату:
    1 Есть ли 3 обязательные колонки:
    - reg_number
    - document_input_request
    - document_verification_request
    2 Есть ли ячейки, содержащие 32767 символов (32767 - максимально допустимое количество символов в ячейке excel,
    что означает, что json не полный, сломанный. Удаляет строки с такими ячейками.
    3 Проверяет, если дубли в колонке reg_number. Удаляет строки с такими ячейками.
    """

    # Проверка на нужные колонки
    sheet = open_excel(filepath)
    book = sheet.parent
    mandatory_column_name_list = [
        constants.REG_NUMBER, constants.DOCUMENT_INPUT_REQUEST,
        constants.DOCUMENT_VERIFICATION_REQUEST]
    current_column_name_list = [cell.value for cell in sheet['1'] if cell.value]
    if not set(mandatory_column_name_list).issubset(current_column_name_list):
        raise ValueError(
            f'В файле {filepath} отсутствуют обязательные колонки, убедитесь что все 3 есть в файле '
            f'{constants.REG_NUMBER},'
            f'{constants.DOCUMENT_INPUT_REQUEST}, {constants.DOCUMENT_VERIFICATION_REQUEST}'
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


    # Проверка на дубликаты reg_number
    reg_numbers = {}
    duplicates = {}
    rows_to_delete = set()
    reg_number_col = None

    for idx, cell in enumerate(sheet['1'], 1):
        if cell.value == constants.REG_NUMBER:
            reg_number_col = idx
            break

    if reg_number_col is not None:
        # Собираем дубликаты
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=reg_number_col)
            reg_number = cell.value

            if reg_number in reg_numbers:
                if reg_number not in duplicates:
                    duplicates[reg_number] = [reg_numbers[reg_number]]
                duplicates[reg_number].append(row)
                rows_to_delete.add(row)  # Добавляем строку в список на удаление
            else:
                reg_numbers[reg_number] = row  # Сохраняем первое вхождение

        # Удаляем дубликаты (кроме первого вхождения)
        if rows_to_delete:
            # Удаляем строки в обратном порядке
            for row in sorted(rows_to_delete, reverse=True):
                sheet.delete_rows(row)
            book.save(filepath)

            dup_info = []
            for reg_number, rows in duplicates.items():
                dup_info.append(
                    f"Значение '{reg_number}': первое вхождение в строке {rows[0]}, "
                    f"дубликаты в строках {', '.join(map(str, rows[1:]))}"
                )
            joined_info = '\n'.join(dup_info)
            log_message = (
                rf"В файле {Path(filepath).stem} удалены дубликаты reg_number:" + "\n"
                f"{joined_info}\n"
                f"Всего удалено строк: {len(rows_to_delete)}"
            )
            logging.warning(log_message)
        else:
            logging.info(f"Файл {Path(filepath).stem} не содержит дубликатов reg_number")

    book.close()
