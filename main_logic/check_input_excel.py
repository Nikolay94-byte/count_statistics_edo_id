import os.path


from utils.utils import open_excel
from utils import constants


def check_excel(filepath: str):
    """Проверяет, соответствует ли excel-файл необходимому формату:
    1 Есть ли 3 обязательные колонки:
    - reg_number
    - document_input_request
    - document_verification_request
    3 Есть ли ячейки, содержащие 32767 символов (32767 - максимально допустимое количество символов в ячейке excel,
    что означает, что json не полный, сломанный
    4 Проверяет, если дубли в колонке reg_number
    """

    # if not os.path.exists(filepath):
    #    raise FileNotFoundError(
    #         f'Необходимый файл отсутствует, убедитесь, что файл назван верно и находится по пути {INPUT_FILE_PATH}'
    #     )

    sheet = open_excel(filepath)
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

    for row in range(2, sheet.max_row+1):
        for cell in sheet[str(row)]:
            if len(str(cell.value)) == 32767:
                raise ValueError(
                    f'Ячейка (строка {row}, колонка {sheet[cell.coordinate].column}) содержит 32767 символов'
                )

    reg_numbers = {}
    duplicates = {}
    reg_number_col = None

    # Находим индекс колонки reg_number
    for idx, cell in enumerate(sheet['1'], 1):
        if cell.value == constants.REG_NUMBER:
            reg_number_col = idx
            break

    if reg_number_col is not None:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=reg_number_col)
            reg_number = cell.value

            if reg_number in reg_numbers:
                if reg_number not in duplicates:
                    duplicates[reg_number] = [reg_numbers[reg_number]]  # Добавляем первую строку с этим значением
                duplicates[reg_number].append(row)  # Добавляем текущую строку
            else:
                reg_numbers[reg_number] = row  # Сохраняем номер строки для этого значения

        if duplicates:
            error_message = "Найдены дубликаты reg_number:\n"
            for reg_number, rows in duplicates.items():
                error_message += f"- Значение '{reg_number}' повторяется в строках: {', '.join(map(str, rows))}\n"
            raise ValueError(error_message)
