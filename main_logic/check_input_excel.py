import logging
from pathlib import Path

from utils import constants
from utils.utils import open_excel


def check_and_clean_file(filepath: str):
    """Проверяет, соответствует ли excel-файл необходимому формату:
    1 Есть ли 6 обязательных колонок:
    - regnumber
    - doc_class
    - attribute_name
    - rus_attribute_name
    - text_normalized
    - text_verification
    2 Есть ли ячейки, содержащие 32767 символов (32767 - максимально допустимое количество символов в ячейке excel,
    что означает, что json не полный, сломанный. Удаляет строки с такими ячейками.
    3 Удаляет все строки документа (по regnumber), если у него нет ни одного значения в text_normalized.
    """

    # Проверка на нужные колонки
    sheet = open_excel(filepath)
    book = sheet.parent
    mandatory_column_name_list = [
        constants.REGNUMBER, constants.DOC_CLASS, constants.ATTRIBUTE_NAME,
        constants.RUS_ATTRIBUTE_NAME, constants.TEXT_NORMALIZED, constants.TEXT_VERIFICATION]
    current_column_name_list = [cell.value for cell in sheet['1'] if cell.value]
    if not set(mandatory_column_name_list).issubset(current_column_name_list):
        raise ValueError(
            f'В файле {filepath} отсутствуют обязательные колонки, убедитесь что все 6 есть в файле '
            f'{constants.REGNUMBER},{constants.DOC_CLASS}, {constants.ATTRIBUTE_NAME}, {constants.RUS_ATTRIBUTE_NAME}'
            f'{constants.TEXT_NORMALIZED},{constants.TEXT_VERIFICATION}'
        )

    # Получаем индексы колонок
    header_row = [cell.value for cell in sheet[1]]
    regnumber_col_idx = header_row.index(constants.REGNUMBER) + 1  # +1 потому что в openpyxl колонки с 1
    text_normalized_col_idx = header_row.index(constants.TEXT_NORMALIZED) + 1

    # Собираем regnumber документов без text_normalized
    regnumbers_without_text_normalized = set()
    regnumbers_with_text_normalized = set()

    for row in range(2, sheet.max_row + 1):
        regnumber = sheet.cell(row=row, column=regnumber_col_idx).value
        text_normalized = sheet.cell(row=row, column=text_normalized_col_idx).value

        if text_normalized:
            regnumbers_with_text_normalized.add(regnumber)
        else:
            if regnumber not in regnumbers_with_text_normalized:
                regnumbers_without_text_normalized.add(regnumber)

    # Находим regnumber, которые есть в without, но нет в with
    regnumbers_to_delete = regnumbers_without_text_normalized - regnumbers_with_text_normalized

    # Удаляем строки с битыми ячейками (32767 символов)
    rows_to_delete = set()
    problematic_cells = []
    for row in range(2, sheet.max_row + 1):
        for cell in sheet[str(row)]:
            if len(str(cell.value)) == 32767:
                rows_to_delete.add(row)
                problematic_cells.append(
                    f"строка {row}, колонка {sheet[cell.coordinate].column}"
                )

    # Добавляем строки для удаления по regnumber без text_normalized
    if regnumbers_to_delete:
        for row in range(2, sheet.max_row + 1):
            regnumber = sheet.cell(row=row, column=regnumber_col_idx).value
            if regnumber in regnumbers_to_delete:
                rows_to_delete.add(row)

    # Удаляем строки если есть что удалять
    if rows_to_delete:
        for row in sorted(rows_to_delete, reverse=True):
            sheet.delete_rows(row)
        book.save(filepath)

        # Формируем сообщение для лога
        log_messages = []

        if problematic_cells:
            cells_info = "\n".join(problematic_cells)
            log_messages.append(
                f"Удалены строки с ячейками, содержащими 32767 символов:\n"
                f"{cells_info}\n"
                f"Всего удалено строк с битыми ячейками: {len([r for r in rows_to_delete if r not in regnumbers_to_delete])}"
            )

        if regnumbers_to_delete:
            regnumbers_info = ", ".join(str(r) for r in regnumbers_to_delete)
            log_messages.append(
                f"Удалены все строки документов со следующими regnumber, так как у них отсутствуют значения в text_normalized: {regnumbers_info}\n"
                f"Всего удалено строк документов без text_normalized: {len(regnumbers_to_delete)}"
            )

        log_message = (
                f"В файле {Path(filepath).stem} были выполнены следующие действия:\n"
                + "\n".join(log_messages)
        )
        logging.warning(log_message)
    else:
        logging.info(
            f"Файл {Path(filepath).stem} не содержит проблемных строк (32767 символов в ячейке или документов без text_normalized)")

    book.close()
